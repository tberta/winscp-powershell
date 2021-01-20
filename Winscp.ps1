<#
.SYNOPSIS
Upload or download files to or from remote server using WinSCP.

.DESCRIPTION
Upload or download files to or from a remote server. Based on WinSCP .NET library
Protocols supported by WinSCP : SFTP / SCP / S3 / FTP / FTPS / WebDAV

.PARAMETER WinscpPath
Path to WinSCPnet.dll matching the .NET Framework Version

.PARAMETER SessionURL
SessionURL are one string containing several or all information needed for remote site connection
SessionURL Syntax:
<protocol> :// [ <username> [ : <password> ] [ ; <advanced> ] @ ] <host> [ : <port> ] / [ <destination directory> / ]
SessionURL Example 
s3://s3.eu-west-1.amazonaws.com/my-test-bucket/

.PARAMETER LocalPath
Path to folder or to individual file

.PARAMETER RemotePath
Path to remote folder or to 

.PARAMETER HostName
Remote server hostname

.PARAMETER PortNumber
Remote server port number

.PARAMETER UserName
Login for authentication on remote server

.PARAMETER SecurePassword
Password as SecureString for authentication on remote server 

.PARAMETER CSEntryName
CredentialStore Entry Name to get the Login and Password securely

.PARAMETER Include
Filename or wildcard expression to select files.

.PARAMETER Filemask
FileMask is a touchy option allowing to select or exclude files or folder to download or upload.
It needs $Include to include a larger set of files. FileMask will then include or exclude files from this set.

*	Matches any number (including zero) of arbitrary characters.	*.doc; about*.html
?	Matches exactly one arbitrary character.	photo????.jpg
[abc]	Matches one character from the set.	index_[abc].html
[a-z]	Matches one character from the range.	index_[a-z].html

Exclude sub-folders:
-FileMask "|*/"

Include all files from current directory but exclude subfolders:
-FileMask "*|*/"

Other file mask examples and features are documented on https://winscp.net/eng/docs/file_mask

.PARAMETER Command
Operation to executes : upload or download

.PARAMETER Protocol
Protocol to use for remote : sftp, ftp, s3, scp, webdav

.PARAMETER SshHostKeyFingerprint
String containing the remote SSH Public Key fingerprint
"ssh-rsa 2048 e0:a3:0f:1a:04:df:5a:cf:c9:81:84:4e:08:4c:9a:06"

.PARAMETER SshPrivateKeyPath
Path to SSH/SFTP Private Key file (can be passphrase protected)

.PARAMETER SecurePrivateKeyCSEntryName
CredentialStore Entry Name to get the passphrase allowing to read and use the SSH/SFTP Private Key for authentication

.PARAMETER FtpMode
Enable Passive or Active FTP Transfer Mode

.PARAMETER FtpSecure
Enable Explicit or Implicit FTP Secure Mode

.PARAMETER TransferMode
Ascii or Binary, or Automatic (based on extension). Binary by default
Using this option enables an "advanced" behaviour using FileMask (and Include)

.PARAMETER PreserveTimestamp
Add this switch to preserve timestamps on transfered files. $true by default

.PARAMETER IgnoreHostAuthenticityCheck
Add this switch if the remote host authenticity should not be checked.
SshHostKeyFingerprint won't be checked for SSH / SFTP server.
Server certificate won't be checked for FTPS / S3 / WebDAV server.

.PARAMETER DeleteSourceFile
Add this switch if source file should be deleted after transfer successful 

.EXAMPLE
.\Winscp.ps1 -protocol sftp -Hostname remotehost.com -Port 2222 -User mylogin -pass <SecureString> -remotePath "/incoming" -localPath "C:\to_send" -filemask "*"

.EXAMPLE
.\Winscp.ps1 -WinscpPath "C:\Program Files (x86)\WinSCP\WinSCPnet.dll" -SessionURL "s3://s3.amazonaws.com/s3-my-bucketname-001/incoming" -localPath "C:\To_upload" -filemask "*.txt"  -command "upload"  -user "mylogin"

.Example
Using SessionURL with new filename on destination : 
.\Winscp.ps1 -SessionURL "s3://s3.amazonaws.com/s3-my-bucketname-001/incoming/" -LocalPath "C:\To_upload\myfile.tst" -RemotePath "NewName.tst"  -command "upload"  -CsEntry "mylogin"

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

    [Parameter(Mandatory)]
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

    [Alias("Password", "Pass")]
    [SecureString] $SecurePassword,

    [string]
    $CSEntryName,

    [string]
    $Include = $null,

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
    $TransferMode,

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
        throw [System.Management.Automation.RuntimeException] "Module CredentialStore not present.`r`nPlease install it first"
        Exit 3
    }
    try {
        Import-Module -Name CredentialStore -ErrorAction Stop
        [SecureString] $Password = Get-CsPassword -Name $EntryName -ErrorAction Stop
    }
    catch {
        throw [System.Management.Automation.RuntimeException] ("Entry '${_}' does not exist.`r`n" +
        "Please set it first with :`r`n" + 
        "  Import-Module CredentialStore`r`n" +
        "  Set-CsEntry -Name $EntryName`r`n" +
        "Can't continue. Exiting.")
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
        throw [System.Management.Automation.RuntimeException] ("Module CredentialStore not present.`r`n" +
        "Please install it first")
        Exit 3
    }
    try {
        Import-Module -Name CredentialStore -ErrorAction Stop
        $UserName = (Get-CsCredential -Name $EntryName -ErrorAction Stop).UserName
    }
    catch {
        throw ("Error $_`r`n" + $_.Exception.Message)
    }
    
    return $UserName
}



Function FileTransferProgress {
    Param($e)
 
    # New line for every new file
    if (($Null -ne $script:lastFileName) -and
        ($script:lastFileName -ne $e.FileName))
    {
        Write-Host "[  $($Script:Command) OK ] $($Script:lastFileName)"
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
        Write-Host "[  $($Script:Command) OK ] $($e.FileName) to $($e.Destination)"
    }
    else
    {
        Write-Host "[  $($Script:Command) FAILED ] $($e.FileName) to $($e.Destination)"
    }
}
#endregion functions


if (-not (Test-Path -Path $LocalPath -IsValid -ErrorAction Stop)) {
	throw "Specified Path is incorrect : $LocalPath"
	Exit 1
}

$SessionOptions = New-Object WinSCP.SessionOptions
$TransferOptions = New-Object WinSCP.TransferOptions

#region argumentsParsing
[System.Collections.Generic.List[string]] $PSBoundKeys = $PSBoundParameters.Keys
switch ($PSBoundKeys) {
    
    'CSEntryName' {
        $PSBoundParameters["UserName"] = Get-UserName $CSEntryName
        $PSBoundParameters["SecurePassword"] = Get-Cred $CSEntryName
    }
    
	'LocalPath' {
		if (Test-Path -Path $LocalPath -PathType Container) {
			if (-not $LocalPath.EndsWith('\')) {
					$LocalPath += "\"
			}
		}
	}
	
	# 'RemotePath' {
		# if (-not $RemotePath.EndsWith('/')) {
			# $RemotePath += "/"
		# }
	# }
	
    'sessionURL' {
        # Remote Path not supported by WinSCP in sessionURL for protocol other than s3 and webdav.
        # Workaround to allow it
        # if $RemotePath not defined but defined in sessionURL, we get it and define $RemotePath with it
        if (-not $RemotePath) {
            
            [uri] $ParsedURL = $sessionURL
            
            if ($ParsedURL.AbsolutePath) {
				if ($parsedURL.Scheme -in ("s3", "webdav")) {
					$RemotePath = "./" # [uri]::UnescapeDataString($ParsedURL.AbsolutePath)
				} else {
					# Other Protocols => We remove the RemotePath from sessionURL
					$RemotePath = [uri]::UnescapeDataString($ParsedURL.AbsolutePath)
					$sessionURL = $sessionURL.Substring(0, $sessionURL.LastIndexOf($parsedURL.AbsolutePath))
				}
			}
        }
        try {
            $SessionOptions.ParseURL($sessionURL)    
            # Defining $Protocol a posteriori to allow 'IgnoreHostAuthenticityCheck' logic to work
			#  $SessionOptions.Protocol =  "sftp" by default at SessionOptions object initialization
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

}

# Depends on sessionURL Parsing or Protocol argument handling
if ($IgnoreHostAuthenticityCheck.IsPresent) {

	if ($Protocol -in ("sftp", "scp") -or $PSBoundParameters["Protocol"] -in ("sftp", "scp")) {
		$PSBoundParameters["GiveUpSecurityAndAcceptAnySshHostKey"] = $true
	}
	
	if ($FtpSecure -or 
		($Protocol -in ("s3", "webdav")) -or
		($PSBoundParameters["Protocol"] -in ("s3", "webdav"))) {
		$PSBoundParameters["GiveUpSecurityAndAcceptAnyTlsHostCertificate"] = $true
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
    
	if ($FileMask -or $TransferMode) {
		$AdvancedFunction = $true
		Write-Debug "Using Advanced $command function"
	} else {
		$AdvancedFunction = $false
		if (-not $Include) {
			if ($Command -eq "UPLOAD" -and (Test-Path -Path $LocalPath -PathType Container)) {
				Throw "-Include switch is missing from commandline arguments"
			} else {
				Write-Warning "-Include switch is missing from commandline arguments"
			}
		}
	}
	
	if ($Command -eq "UPLOAD") {
		
		if ($AdvancedFunction) {
			if (Test-Path -Path $LocalPath -PathType Container) {
				if ($Include) {
					$LocalPath = $LocalPath + $Include
				}
			}
		
			# Upload files using advanced function
			$operationMessage = $Command + " '{0}' to '{1}' (Advanced FileMask '{2}') Results:" -f $LocalPath, $RemotePath, $FileMask
			Write-Host $operationMessage
			$TransferResult = $Session.PutFiles($LocalPath, $RemotePath, $DeleteSourceFile, $TransferOptions)
        
		} else {
			# Upload files using simpler function
			$operationMessage = $Command + " '{0}' from {1} to '{2}' Results:" -f $Include, $LocalPath, $RemotePath
			Write-Host $operationMessage
			$TransferResult = $Session.PutFilestoDirectory($LocalPath, $RemotePath, $Include, $DeleteSourceFile)
		}
    }

    if ($Command -eq "DOWNLOAD") {
		if ($AdvancedFunction) {
			if ($RemotePath.EndsWith("/")) {
				if ($Include) {
					$RemotePathValue = $RemotePath + $Include
				}
			}
			$operationMessage = $Command + " '{0}' to '{1}' (Advanced FileMask '{2}') Results:" -f $RemotePathValue, $LocalPath, $FileMask
			Write-Host $operationMessage
			# Download the file and throw on any error
			$TransferResult = $Session.GetFiles($RemotePathValue, $LocalPath, $DeleteSourceFile, $TransferOptions)
		} else {
			$RemotePathValue = $RemotePath
			$operationMessage = $Command + " '{0}' from '{1}' to '{2}' Results:" -f $Include, $RemotePathValue, $LocalPath
			Write-Host $operationMessage
			
			$TransferResult = $Session.GetFilesToDirectory($RemotePathValue, $LocalPath, $Include, $DeleteSourceFile)
		}
    }

    if ($null -ne $Script:lastFileName) {
        Write-Host "[  $($Script:Command) OK ] `t$($Script:lastFileName)"
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
	$PSBoundParameters | Out-String | Write-Host
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
