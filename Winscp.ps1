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
param (
    [ValidateScript( {
            Test-Path -Path $_ -PathType Leaf
        })]
    [string]
    $winscpPath = "C:\Program Files (x86)\WinSCP\WinSCPnet.dll",

    # [Parameter(ParameterSetName = 'combined')]
    [string]
    $sessionURL, 

    [ValidateScript( {
            Test-Path -Path $_
        })]
    [string]
    $localPath,

    [string]
    $remotePath,

    [string]
    $HostName,
    
    [Alias("Port")]
    [int]
    $PortNumber,

    [Alias("User")]
    [string]
    $UserName,

    [Alias("Password")]
    [SecureString] $SecurePassword,

    [string]
    $CSEntryName,

    [string]
    $Filemask = $null,

    [ValidateSet('download', 'upload')]
    [string] 
    $command,

    [ValidateSet('sftp', 'ftp', 's3', 'scp', 'webdav')]
    [string]
    $Protocol,

    [Alias("serverFingerprint")]
    [string]
    $SshHostKeyFingerprint,


    [ValidateScript( {
            Test-Path -Path $_ -PathType Leaf 
        })]
    [String]
    $SshPrivateKeyPath,

    [string]
    $SecurePrivateKeyCSEntryName,

    [ValidateSet("Active", "Passive")]
    [String]
    $FtpMode,

    [ValidateSet("Implicit", "Explicit", "None")]
    [String]
    $FtpSecure,

    [Switch]
    $IgnoreHostAuthenticityCheck,

    [switch] $deleteSourceFile = $false
)
         
try {
    # Load WinSCP .NET assembly
    Add-Type -Path "$winscpPath"
}
catch [Exception] {
    Write-Host "Error: "$_.Exception.Message
    Exit 1
}




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

            "sftp"  { $PSBoundParameters["Protocol"] = [WinSCP.Protocol]::sftp   ; break }
            "ftp"   { $PSBoundParameters["Protocol"] = [WinSCP.Protocol]::ftp    ; break }
            "s3"    { $PSBoundParameters["Protocol"] = [WinSCP.Protocol]::s3     ; break }
            "scp"   { $PSBoundParameters["Protocol"] = [WinSCP.Protocol]::scp    ; break }
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

    'FtpMode' {
        if ($FtpMode -eq "Active") {
            $PSBoundParameters["FtpMode"] = [WinSCP.FtpMode]::Active
        }
        if ($FtpMode -eq "Passive") {
            $PSBoundParameters["FtpMode"] = [WinSCP.FtpMode]::Passive
        }

    }

    'FtpSecure' {
        if ($SecurityMode -eq "Implicit") {
            $PSBoundParameters["FtpSecure"] = [WinSCP.FtpSecure]::Implicit
        }
        if ($SecurityMode -eq "Explicit") {
            $PSBoundParameters["FtpSecure"] = [WinSCP.FtpSecure]::Explicit
        }
    }

    'IgnoreHostAuthenticityCheck' {
        if ($IgnoreHostAuthenticityCheck) {
            if ($Protocol -in ("sftp", "scp")) {
                $PSBoundParameters["GiveUpSecurityAndAcceptAnySshHostKey"] = $true
            }         
            if ($FtpSecure -eq "Implicit") {
                $PSBoundParameters["GiveUpSecurityAndAcceptAnyTlsHostCertificate"] = $true
            }
            # elseif ($FtpSecure -eq "Explicit") {
            #     $PSBoundParameters["GiveUpSecurityAndAcceptAnySshHostKey"] = $true
            # }
            elseif ($FtpSecure -eq "Explicit") {
                $PSBoundParameters["GiveUpSecurityAndAcceptAnyTlsHostCertificate"] = $true
            }
        }
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
        $operationMessage = $command.ToUpper() + " Job from $LocalPath to $RemotePath Results:"
        $transferOptions = New-Object WinSCP.TransferOptions
        $transferOptions.TransferMode = [WinSCP.TransferMode]::Binary
        #$transferResult = $session.PutFiles($localPath,($remotePath + $filemask), $deleteSourceFile , $transferOptions)
        $transferResult = $session.PutFilestoDirectory($localPath, $remotePath, $filemask, $deleteSourceFile)
    }

    if ($command -eq "download") {
        $operationMessage = $command.ToUpper() + " Job from $RemotePath to $LocalPath Results:"
        # Download the file and throw on any error
        #$sessionResult = $session.GetFiles(($remotePath + $fileName),($localPath + $filemask))
        $transferOptions = New-Object WinSCP.TransferOptions
        $transferOptions.TransferMode = [WinSCP.TransferMode]::Binary
        $transferResult = $session.GetFilesToDirectory($remotePath, $localPath, $filemask, $deleteSourceFile)
    }

    # Throw error if found
    $transferResult.Check()
    
    # if ($VerbosePreference) {
        Write-Host $operationMessage
        foreach ($transfer in $transferResult.Transfers) {
            Write-Host "[  $command OK  ]   $($transfer.FileName)"
        }
        foreach ($failure in $transferResult.Failures) {
            Write-Host "[ $command FAIL ]   $($transfer.FileName)"
        }
    # }
    if ($($transferResult.IsSuccess)) {
        Write-Host "$command job ended successfully"
    }
    Write-Host "Files transferred successfully : $($transferResult.Transfers.count)"
    if ($($transferResult.Failures.Count)) {
        Write-Host "Files not transferred : $($transferResult.Failures.Count)"
    }

}
catch [Exception] {
    Write-Error "Received Error '$_'"
    Write-Host "<---- Session Output ---->"
    $session.Output | Out-String
    Write-Host "^---- Session Output ----^"
    $returnCode = 7
}    
finally {
    # Disconnect, clean up
    $session.Dispose()
    exit $returnCode
}    
