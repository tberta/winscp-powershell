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
$LocalPath = "C:\To_upload"
$RemotePath = "incoming"
$filemask = "*.txt"
$Command = "upload"
PS> Winscp.ps1 -winscpPath $winscp -sessionURL $sessionURL -localPath $LocalPath -filemask $filemask -remotePath $RemotePath -command $Command -password $pass -user $UserName

.EXAMPLE
$UserName = "mylogin"
$pass = "mypass"
$pass = [uri]::EscapeDataString($pass) # To make it valid in sessionURL
$bucket = "s3-my-bucketname-001"
$sessionURL = "s3://$($UserName):$($pass)@s3.amazonaws.com/$($bucket)/"
$LocalPath = "C:\To_upload"
$RemotePath = "incoming"
$Command = "upload"
PS> Winscp.ps1 -winscpPath $winscp -sessionURL $sessionURL -localPath $LocalPath -filemask $filemask -remotePath $RemotePath -command $Command -password $pass -user $UserName

.INPUTS

None. You cannot pipe objects to Winscp1.ps1.

.OUTPUTS

None. Besides some Console output, Winscp1.ps1 does not generate any output object.

.LINK
https://github.com/tberta/winscp-powershell

#>
[CmdletBinding(SupportsShouldProcess = $False, PositionalBinding = $False)]
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
    $Command,

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

    [ValidateSet("Binary","Ascii", "Automatic")]
    [String]
    $TransferMode = "Automatic",

    [Parameter()]
    [Bool]
    $PreserveTimestamp = $true,

    [Switch]
    $IgnoreHostAuthenticityCheck,

    [switch] $DeleteSourceFile = $false
)
         
try {
    # Load WinSCP .NET assembly
    Add-Type -Path "$winscpPath"
}
catch [Exception] {
    Write-Host "Error: "$_.Exception.Message
    Exit 1
}

#region functions
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



Function FileTransferProgress {
    Param($e)
 
    # New line for every new file
    if (($Null -ne $script:lastFileName) -and
        ($script:lastFileName -ne $e.FileName))
    {
        Write-Host "[  $($Script:Command.PadRight(9)) OK ] `t$($Script:lastFileName)"
        if($PSBoundParameters.ContainsKey("Verbose")) {
            Write-Host
        }
        
    }
 
    # Print transfer progress
    if($PSBoundParameters.ContainsKey("Verbose")) {
        Write-Host -NoNewline ("`r{0} ({1:P0})" -f $e.FileName, $e.FileProgress)
    }
    # Remember a name of the last file reported
    $script:lastFileName = $e.FileName
}


Function FileTransferred {
    Param($e)

    if ($null -eq $e.Error)
    {
        Write-Host "[  $($Script:Command.PadRight(9)) OK ] `t$($e.FileName)"
    }
    else
    {
        Write-Host "[  $($Script:Command.PadRight(9)) FAILED ] `t$($e.FileName)"
    }
}
#endregion functions

$SessionOptions = New-Object WinSCP.SessionOptions
$TransferOptions = New-Object WinSCP.TransferOptions

#region argumentsParsing
[System.Collections.Generic.List[string]] $PSBoundKeys = $PSBoundParameters.Keys
switch ($PSBoundKeys) {
    
    'CSEntryName' {
        $PSBoundParameters["UserName"] = Get-UserName $CSEntryName
        $PSBoundParameters["SecurePassword"] = Get-Cred $CSEntryName
    }
    
    'sessionURL' {
        # Remote Path not supported by WinSCP in sessionURL for  protocol other than s3 and webdav.
        # Workaround to allow it
        # if $RemotePath not defined but defined in sessionURL, we get it and define $RemotePath with it
        if (-not $RemotePath) {
            
            [uri] $ParsedURL = $sessionURL
            
            if ($ParsedURL.AbsolutePath -and 
                ( -not ($parsedURL.Scheme -in ("s3", "webdav")))
            ) {
                $RemotePath = [uri]::UnescapeDataString($ParsedURL.AbsolutePath)
                $sessionURL = $sessionURL.Substring(0, $sessionURL.LastIndexOf($parsedURL.AbsolutePath))
            }
        }
        try {
            $SessionOptions.ParseURL($sessionURL)    
            # Defining $Protocol a posteriori to allow 'IgnoreHostAuthenticityCheck' logic to work
            $Protocol = $SessionOptions.Protocol
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

    'FtpMode' {
        if ($FtpMode -eq "Active") {
            $PSBoundParameters["FtpMode"] = [WinSCP.FtpMode]::Active
        }
        if ($FtpMode -eq "Passive") {
            $PSBoundParameters["FtpMode"] = [WinSCP.FtpMode]::Passive
        }
    }

    'FtpSecure' {
        if ($FtpSecure -eq "Implicit") {
            $PSBoundParameters["FtpSecure"] = [WinSCP.FtpSecure]::Implicit
        }
        if ($FtpSecure -eq "Explicit") {
            $PSBoundParameters["FtpSecure"] = [WinSCP.FtpSecure]::Explicit
        }
    }

    'TransferMode' {
        if($TransferMode -eq "Binary") {
            $PSBoundParameters["TransferMode"] = [WinSCP.TransferMode]::Binary
        } elseif ($TransferMode -eq "Ascii") {
            $PSBoundParameters["TransferMode"] = [WinSCP.TransferMode]::Ascii
        } else {
            $PSBoundParameters["TransferMode"] = [WinSCP.TransferMode]::Automatic
        }
    }

    'IgnoreHostAuthenticityCheck' {
        if ($IgnoreHostAuthenticityCheck.IsPresent) {
            #  $SessionOptions.Protocol =  "sftp" by default at variable initialization
            if ($Protocol -in ("sftp", "scp") -or $PSBoundParameters["Protocol"] -in ("sftp", "scp")) {
                $PSBoundParameters["GiveUpSecurityAndAcceptAnySshHostKey"] = $true
            }         
            if ($FtpSecure -eq "Implicit") {
                $PSBoundParameters["GiveUpSecurityAndAcceptAnyTlsHostCertificate"] = $true
            }
            elseif ($FtpSecure -eq "Explicit") {
                $PSBoundParameters["GiveUpSecurityAndAcceptAnyTlsHostCertificate"] = $true
            }
        }
    }
}

if ($PSBoundParameters.ContainsKey("Debug")) {
    $DebugPreference = "Continue"
    $PSBoundParameters | Out-String
}

# Setup session options
try {
    $sessionOptionObjectProperties = $SessionOptions | 
        Get-Member -MemberType Property | 
        Select-Object -ExpandProperty Name
    $keys = ($PSBoundParameters.Keys).Where( {
            $_ -in $sessionOptionObjectProperties
        })
    foreach ($key in $keys) {
        Write-Debug -Message ("Adding {0} value {1}" -f $key, $PSBoundParameters[$key])
        $SessionOptions.$key = $PSBoundParameters[$key]
    }
}
catch {
    $PSCmdlet.ThrowTerminatingError( $_ )
}

if ($PSBoundParameters.ContainsKey("Debug")) {
    $SessionOptions | Select-Object -ExcludeProperty "Password" -Property "*"
}

try {
    $transferOptionsObjectProperties = $TransferOptions | 
        Get-Member -MemberType Property | 
            Select-Object -ExpandProperty Name
    $keys = ($PSBoundParameters.Keys).Where( {
        $_ -in $transferOptionsObjectProperties
    })
    foreach ($key in $keys) {
        Write-Debug -Message ("Adding {0} value {1}" -f $key, $PSBoundParameters[$key])
        $TransferOptions.$key = $PSBoundParameters[$key]
    }
}
catch {
    $PSCmdlet.ThrowTerminatingError( $_ )
}

if ($PSBoundParameters.ContainsKey("Debug")) {
    $TransferOptions
}

#endregion argumentsParsing

$returnCode = 0
    
$Session = New-Object WinSCP.Session

try {
    # $Session.add_FileTransferProgress({ FileTransferProgress($_) } )
    $Session.add_FileTransferred( { FileTransferred($_) } )
    # Connect
    $Session.Open($SessionOptions)
	if (-not $Session.Opened -eq $true) {
		throw [System.Management.Automation.RuntimeException] "Can't open connection to remote server.`r`n" + $Session.Output
	}
    $Command = $Command.ToUpper()
    if ($Command -eq "UPLOAD") {
        # Upload files
        $operationMessage = $Command + " Job from '{0}' to '{1}' Results:" -f $LocalPath, $RemotePath
        Write-Host $operationMessage
        $TransferResult = $Session.PutFiles($LocalPath, $RemotePath, $DeleteSourceFile, $TransferOptions)
        # $TransferResult = $Session.PutFilestoDirectory($LocalPath, $RemotePath, $Filemask, $DeleteSourceFile)
    }

    if ($Command -eq "DOWNLOAD") {
        $operationMessage = $Command + " Job from '{0}' to '{1}' Results:" -f $RemotePath, $LocalPath
        Write-Host $operationMessage
        # Download the file and throw on any error
        $TransferResult = $Session.GetFiles($RemotePath, $LocalPath, $DeleteSourceFile, $TransferOptions)
        # $TransferResult = $Session.GetFilesToDirectory($RemotePath, $LocalPath, $Filemask, $DeleteSourceFile)
    }

    if ($null -ne $Script:lastFileName) {
        Write-Host "[  $($Script:Command.PadRight(9)) OK ] `t$($Script:lastFileName)"
    }
    # if ($VerbosePreference) {
        #     Write-Host $operationMessage
        # foreach ($transfer in $TransferResult.Transfers) {
            #     Write-Host "[  $Command OK  ]   $($transfer.FileName)"
            # }
            foreach ($failure in $TransferResult.Failures) {
                Write-Host "[ $Command FAIL ]  $($transfer.FileName)"
            }
            # }
            if ($($TransferResult.IsSuccess)) {
        Write-Host "$Command job ended successfully"
    }
    Write-Host "Files transferred successfully : $($TransferResult.Transfers.count)"
    if ($($TransferResult.Failures.Count)) {
        Write-Host "Files not transferred : $($TransferResult.Failures.Count)"
    }
    # Throw error if found
    $TransferResult.Check()
    
    if ($TransferResult.IsSuccess -eq $true ) {
        if ($TransferResult.Transfers.Count -gt 0) {
            $returnCode = 0
        } else {
            $returnCode = 20
        }
    } else {
        $returnCode = 1
    }
}
catch [Exception] {
    $PSBoundParameters | Out-String
    $SessionOptions | Select-Object -ExcludeProperty "Password" -Property "*"
    Write-Error "Received Error '$_'"
    Write-Host "<---- Session Output ---->"
	$Session.Output | Out-String | Write-Host
    Write-Host "^---- Session Output ----^"
    $returnCode = 1
}    
finally {
    # Disconnect, clean up
    $Session.Dispose()
    exit $returnCode
}    
