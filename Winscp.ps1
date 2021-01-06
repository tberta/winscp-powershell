<#
.SYNOPSIS
Upload or download files to or from remote server.

.DESCRIPTION
Upload or download files to or from a remote server. Based on WinSCP .NET library
Protocols supported by WinSCP : SFTP / SCP / S3 / FTP / WebDAV

.PARAMETER sessionURL
sessionURL are one string containing several or all information needed for remote site connection
sessionURL Syntax:
<protocol> :// [ <username> [ : <password> ] [ ; <advanced> ] @ ] <host> [ : <port> ] /
sessionURL Example 

.EXAMPLE
.\Winscp.ps1 -protocol sftp -hostname remotehost.com -port 2222 -user mylogin -pass mypass -remotePath "/incoming" -localPath "C:\to_send" -filemask "*"

.EXAMPLE
$UserName = "mylogin"
$pass = "mypass"
$bucket = "s3-my-bucketname-001"
$winscp = "C:\Programs Files\WinSCP\WinSCPnet.dll"
$sessionURL = "s3://s3.amazonaws.com/$($bucket)/"
$localPath = "C:\To_upload"
$remotePath = "incoming"
$filemask = "*.txt"
$command = "upload"
PS> Winscp.ps1 -winscpPath $winscp -sessionURL $sessionURL -localPath $localPath -filemask $filemask -remotePath $remotePath -command $command -password $pass -user $UserName

.EXAMPLE
$UserName = "mylogin"
$pass = "mypass"
$pass = [uri]::EscapeDataString($pass) # To make it valid in sessionURL
$bucket = "s3-my-bucketname-001"
$sessionURL = "s3://$($UserName):$($pass)@s3.amazonaws.com/$($bucket)/"
$localPath = "C:\To_upload"
$remotePath = "incoming"
$command = "upload"
PS> Winscp.ps1 -winscpPath $winscp -sessionURL $sessionURL -localPath $localPath -filemask $filemask -remotePath $remotePath -command $command -password $pass -user $UserName

.INPUTS

None. You cannot pipe objects to Winscp1.ps1.

.OUTPUTS

None. Besides some Console output, Winscp1.ps1 does not generate any output object.

.LINK
https://github.com/tberta/winscp-powershell

#>
# [CmdletBinding(DefaultParameterSetName = 'combined')]
param (
    # [Parameter(ParameterSetName = 'combined')]
    # [Parameter(ParameterSetName = 'splitted')]
    [ValidateScript( {
            Test-Path -Path $_ -PathType Leaf
        })]
    [string]
    $winscpPath = "C:\Program Files (x86)\WinSCP\WinSCPnet.dll",

    # [Parameter(ParameterSetName = 'combined')]
    [string]
    $sessionURL, 

    # [Parameter(ParameterSetName = 'combined', Mandatory = $true)]
    # [Parameter(ParameterSetName = 'splitted', Mandatory = $true)]
    [ValidateScript( {
            Test-Path -Path $_
        })]
    [string]
    $localPath,

    # [Parameter(ParameterSetName = 'combined')]
    # [Parameter(ParameterSetName = 'splitted')]
    [string]
    $remotePath,

    # [Parameter(ParameterSetName = 'splitted', Mandatory = $true)]
    [string]
    $HostName,
    
    # [Parameter(ParameterSetName = 'splitted')]
    [Alias("Port")]
    [int]
    $PortNumber,

    # [Parameter(ParameterSetName = 'combined')]
    # [Parameter(ParameterSetName = 'splitted')]
    [Alias("User")]
    [string]
    $UserName,

    # [Parameter(ParameterSetName = 'combined')]
    # [Parameter(ParameterSetName = 'splitted')]
    [Alias("Password")]
    [SecureString] $SecurePassword,

    # [Parameter(ParameterSetName = 'combined')]
    # [Parameter(ParameterSetName = 'splitted')]
    [string]
    $CSEntryName,

    # [Parameter(ParameterSetName = 'combined')]
    # [Parameter(ParameterSetName = 'splitted')]
    [string]
    $Filemask = $null,

    # [Parameter(ParameterSetName = 'combined', Mandatory = $true)]
    # [Parameter(ParameterSetName = 'splitted', Mandatory = $true)]
    [ValidateSet('download', 'upload')]
    [string] 
    $command,

    # [Parameter(ParameterSetName = 'splitted')]
    [ValidateSet('sftp', 'ftp', 's3', 'scp', 'webdav')]
    [string]
    $Protocol,

    # [Parameter(ParameterSetName = 'splitted')]
    [Alias("serverFingerprint")]
    [string]
    $SshHostKeyFingerprint,

    [Parameter()]
    [ValidateScript( {
            Test-Path -Path $_ -PathType Leaf 
        })]
    [String]
    $SshPrivateKeyPath,

    # [Parameter(ParameterSetName = 'combined')]
    # [Parameter(ParameterSetName = 'splitted')]
    [string]
    $SecurePrivateKeyCSEntryName,

    [Parameter()]
    [Switch]
    $GiveUpSecurityAndAcceptAnySshHostKey,

    [Parameter()]
    [Switch]
    $GiveUpSecurityAndAcceptAnyTlsHostCertificate,

    # [Parameter(ParameterSetName = 'combined')]
    # [Parameter(ParameterSetName = 'splitted')]
    [switch] $deleteSourceFile = $false


    # [Parameter(ParameterSetName = 'combined')]
    # [Parameter(ParameterSetName = 'splitted')]
    #     [switch] $verbose = $false
)
         
try {
    # Load WinSCP .NET assembly
    Add-Type -Path "$winscpPath"
}
catch [Exception] {
    Write-Host "Error: "$_.Exception.Message
    Exit 1
}

# if (-not (Test-Path -Path $localPath)) {
#     Write-Error "Local folder doesn't exist" -Category ObjectNotFound
#     Write-Information "When running in OpCon, you need an extra *ending* slash for localPath : -localPath=""C:\Test\\"" " -Verbose
#     Write-Information "Received argument: '$localPath'" -Verbose
#     Exit 2
# }


Function Get-Cred {
    Param(
        [Parameter(Mandatory)]
        [string]
        $EntryName
    )

    if (-not (Get-Module -ListAvailable -Name CredentialStore)) {
        throw [System.Management.Automation.RuntimeException] "Module CredentialStore not present.`r`n" +
        "Please install it first"
        Exit 3
    }
    try {
        Import-Module -Name CredentialStore -ErrorAction Stop
        [SecureString] $Password = Get-CsPassword -Name $EntryName -ErrorAction Stop
    }
    catch {
        throw [System.Management.Automation.RuntimeException] "Entry '${_}' does not exist.`r`n" +
        "Please set it first with :`r`n" + 
        "  Import-Module CredentialStore`r`n" +
        "  Set-CsEntry -Name $EntryName`r`n" +
        "Can't continue. Exiting."
        Exit 3
    }
    
    return $Password
}


Function Get-UserName {
    Param(
        [Parameter(Mandatory)]
        [string]
        $EntryName
    )

    if (-not (Get-Module -ListAvailable -Name CredentialStore)) {
        throw [System.Management.Automation.RuntimeException] "Module CredentialStore not present.`r`n" +
        "Please install it first"
        Exit 3
    }
    try {
        Import-Module -Name CredentialStore -ErrorAction Stop
        $UserName = (Get-CsCredential -Name $EntryName -ErrorAction Stop).UserName
    }
    catch {
        throw "Error $_`r`n" + $_.Exception.Message
    }
    
    return $UserName
}

$sessionOptions = New-Object WinSCP.SessionOptions

# User or Password may need to be HTML encoded in a sessionURL
# So for convenience, User and Password value can be passed as arguments
[System.Collections.Generic.List[string]] $PSBoundKeys = $PSBoundParameters.Keys
switch ($PSBoundKeys) {
    
    'CSEntryName' {
        $PSBoundParameters["UserName"] = Get-UserName $CSEntryName
        $PSBoundParameters["SecurePassword"] = Get-Cred $CSEntryName
    }
    
    'sessionURL' {
        try {
            $sessionOptions.ParseURL($sessionURL)    
        }
        catch [Exception] {
            Write-Host "Error while parsing provided sessionURL argument : '$sessionURL'"
            Write-Host $_.Exception.Message
            Exit 3
        }
    }
    
    'Protocol' {
        switch ($Protocol) {

            "sftp" { $PSBoundParameters["Protocol"] = [WinSCP.Protocol]::sftp   ; break }
            "ftp" { $PSBoundParameters["Protocol"] = [WinSCP.Protocol]::ftp    ; break }
            "s3" { $PSBoundParameters["Protocol"] = [WinSCP.Protocol]::s3     ; break }
            "scp" { $PSBoundParameters["Protocol"] = [WinSCP.Protocol]::scp    ; break }
            "webdav" { $PSBoundParameters["Protocol"] = [WinSCP.Protocol]::webdav ; break }
            default {     
                Write-Host "Unknown protocol specified"
                Write-Host "Exiting..."
                Exit 4
            }
        }
    }

    'SshPrivateKeyPath' {
        $PSBoundParameters["SshPrivateKeyPath"] = $(Resolve-Path -Path $SshPrivateKeyPath | Select-Object -ExpandProperty ProviderPath)
    }

    'SecurePrivateKeyCSEntryName' {
        $PSBoundParameters["SecurePrivateKeyPassphrase"] = Get-Cred $SecurePrivateKeyCSEntryName
    }
}
if ($DebugPreference) {
    $PSBoundParameters
}

# Setup session options
try {
    $sessionOptionObjectProperties = $sessionOptions |
    Get-Member -MemberType Property |
    Select-Object -ExpandProperty Name
    $keys = ($PSBoundParameters.Keys).Where( {
            $_ -in $sessionOptionObjectProperties
        })

    foreach ($key in $keys) {
        Write-Debug -Message ("Adding {0} value {1}" -f $key, $PSBoundParameters[$key])
        $sessionOptions.$key = $PSBoundParameters[$key]
    }
}
catch {
    $PSCmdlet.ThrowTerminatingError(
        $_
    )
}

$returnCode = 0
    
$session = New-Object WinSCP.Session

try {
    # Connect
    $session.Open($sessionOptions)

    if ($command -eq "upload") {
        # Upload files
        $destinationMsg = "$remotePath"
        $transferOptions = New-Object WinSCP.TransferOptions
        $transferOptions.TransferMode = [WinSCP.TransferMode]::Binary
        #$transferResult = $session.PutFiles($localPath,($remotePath + $filemask), $deleteSourceFile , $transferOptions)
        $transferResult = $session.PutFilestoDirectory($localPath, $remotePath, $filemask, $deleteSourceFile)
    }

    if ($command -eq "download") {
        $destinationMsg = " $localPath"
        # Download the file and throw on any error
        #$sessionResult = $session.GetFiles(($remotePath + $fileName),($localPath + $filemask))
        $transferOptions = New-Object WinSCP.TransferOptions
        $transferOptions.TransferMode = [WinSCP.TransferMode]::Binary
        $transferResult = $session.GetFilesToDirectory($remotePath, $localPath, $filemask, $deleteSourceFile)
    }

    # Throw error if found
    $transferResult.Check()
    
    # if ($VerbosePreference) {
        foreach ($transfer in $transferResult.Transfers) {
            Write-Host "[  $command OK  ]   $($transfer.FileName)`t$command succeeded"
        }
        foreach ($failure in $transferResult.Failures) {
            Write-Host "[ $command FAIL ]   $($transfer.FileName)`t$command did NOT succeed"
        }
    # }
    if ($($transferResult.IsSuccess)) {
        Write-Host "$command transfer job ended successfully"
    }
    Write-Host "Files transferred successfully : $($transferResult.Transfers.count)"
    if ($($transferResult.Failures.Count)) {
        Write-Host "Files not transferred : $($transferResult.Failures.Count)"
    }

}
catch [Exception] {
    Write-Host $session.Output
    Write-Host $_.Exception.Message
    $returnCode = 7
}    
finally {
    # Disconnect, clean up
    $session.Dispose()
    exit $returnCode
}    
