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
    $LocalPath,

    [string]
    $RemotePath,

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
# [System.Collections.Generic.List[string]] $PSBoundKeys = $PSBoundParameters.Keys
switch ($PSBoundParameters.Keys) {
    
    'CSEntryName' {
        $sessionOptions.UserName = Get-UserName $CSEntryName
        $sessionOptions.SecurePassword = Get-Cred $CSEntryName
    }
    
    'sessionURL' {
        
        if (-not $RemotePath) {
            
            [uri] $ParsedURL = $sessionURL
            
            if ($ParsedURL.AbsolutePath -and 
            ( $ParsedURL.AbsolutePath -ne "/" ) -and
            ( -not ($parsedURL.Scheme -in ("s3", "webdav")))
            ) {
                # Remote Path not supported in sessionURL for other protocol than s3 and webdav.
                # Workaround to allow it
                # $Protocol = $ParsedURL.Scheme
                $RemotePath = [uri]::UnescapeDataString($ParsedURL.AbsolutePath)
                $sessionURL = $sessionURL.Substring(0,$sessionURL.LastIndexOf($parsedURL.AbsolutePath))
            }
        }
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

            "sftp"  { $sessionOptions.Protocol = [WinSCP.Protocol]::sftp   ; break }
            "ftp"   { $sessionOptions.Protocol = [WinSCP.Protocol]::ftp    ; break }
            "s3"    { $sessionOptions.Protocol = [WinSCP.Protocol]::s3     ; break }
            "scp"   { $sessionOptions.Protocol = [WinSCP.Protocol]::scp    ; break }
            "webdav" { $sessionOptions.Protocol = [WinSCP.Protocol]::webdav ; break }
            default {     
                Write-Host "Unknown protocol specified"
                Write-Host "Exiting..."
                Exit 4
            }
        }
    }

    'SshPrivateKeyPath' {
        $sessionOptions.SshPrivateKeyPath = $(Resolve-Path -Path $SshPrivateKeyPath | Select-Object -ExpandProperty ProviderPath)
    }

    'SecurePrivateKeyCSEntryName' {
        $sessionOptions.SecurePrivateKeyPassphrase = Get-Cred $SecurePrivateKeyCSEntryName
    }

    'FtpMode' {
        if ($FtpMode -eq "Active") {
            $sessionOptions.FtpMode = [WinSCP.FtpMode]::Active
        }
        if ($FtpMode -eq "Passive") {
            $sessionOptions.FtpMode = [WinSCP.FtpMode]::Passive
        }
    }

    'FtpSecure' {
        if ($FtpSecure -eq "Implicit") {
            $sessionOptions.FtpSecure = [WinSCP.FtpSecure]::Implicit
        }
        if ($FtpSecure -eq "Explicit") {
            $sessionOptions.FtpSecure = [WinSCP.FtpSecure]::Explicit
        }
    }

    'IgnoreHostAuthenticityCheck' {
        if ($IgnoreHostAuthenticityCheck.IsPresent) {
            if ($Protocol -in ("sftp", "scp") -or $sessionOptions.Protocol -in ("sftp", "scp")) {
                $sessionOptions.GiveUpSecurityAndAcceptAnySshHostKey = $true
            }         
            if ($FtpSecure -eq "Implicit") {
                $sessionOptions.GiveUpSecurityAndAcceptAnyTlsHostCertificate = $true
            }
            # elseif ($FtpSecure -eq "Explicit") {
            #     $PSBoundParameters["GiveUpSecurityAndAcceptAnySshHostKey"] = $true
            # }
            elseif ($FtpSecure -eq "Explicit") {
                $sessionOptions.GiveUpSecurityAndAcceptAnyTlsHostCertificate = $true
            }
        }
    }
}
# if ($Debug.IsPresent) {
    $sessionOptions
# }
# Setup session options
# try {
#     $sessionOptionObjectProperties = $sessionOptions |
#     Get-Member -MemberType Property |
#     Select-Object -ExpandProperty Name
#     $keys = ($PSBoundParameters.Keys).Where( {
#             $_ -in $sessionOptionObjectProperties
#         })

#     foreach ($key in $keys) {
#         Write-Debug -Message ("Adding {0} value {1}" -f $key, $PSBoundParameters[$key])
#         $sessionOptions.$key = $PSBoundParameters[$key]
#     }
# }
# catch {
#     $PSCmdlet.ThrowTerminatingError(
#         $_
#     )
# }

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
        $transferResult = $session.PutFilestoDirectory($localPath, $RemotePath, $Filemask, $deleteSourceFile)
    }

    if ($command -eq "download") {
        $operationMessage = $command.ToUpper() + " Job from $RemotePath to $LocalPath Results:"
        # Download the file and throw on any error
        #$sessionResult = $session.GetFiles(($remotePath + $fileName),($localPath + $filemask))
        $transferOptions = New-Object WinSCP.TransferOptions
        $transferOptions.TransferMode = [WinSCP.TransferMode]::Binary
        $transferResult = $session.GetFilesToDirectory($RemotePath, $LocalPath, $Filemask, $deleteSourceFile)
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
    $PSBoundParameters
    $sessionOptions
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
