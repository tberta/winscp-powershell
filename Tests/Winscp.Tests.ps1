Param(

	[Parameter()]
	[ValidateSet("S","M", "L")]
	[string] $size = "S",
	
	[Parameter(Mandatory = $false)]
	[ValidateScript( {
			Test-Path -Path $_ -PathType Container 
		})]
	[string] $logFolder = "D:\Tests"
) 

$TestsScriptName = Split-Path -Leaf $PSCommandPath
$ScriptName = $TestsScriptName.Replace(".Tests.", ".")
$ScriptFullName = Join-Path (Resolve-Path "$PSScriptRoot\..") $ScriptName
$LogFileName = $TestsScriptName.Replace(".ps1", ".") + $(Get-Date -Format "yyyyMMdd") + ".log"

if ($LogFolder) {
	Start-Transcript -Path $(Join-Path $LogFolder $LogFileName) -Append
}

$user = "nxy-dev"
$stenv = "QUA"
[PSCustomObject[]] $results = @()


switch($size) {
	"S" { $Filemask = "*small*.tst"}
	"M" { $Filemask = "*medium*.tst"}
	"L" { $Filemask = "*large*.tst" }
	default { }
}

. "$PSScriptRoot\Tests-Helpers.ps1"

# $winSCPPath = 


# .\Winscp.ps1 -sessionURL "ftp://moninvite@localhost:21/" -CSEntryName test -Filemask "*.tst" -LocalPath 'C:\tools\Out' -command download -DEbug
# rc = 0
# .\Winscp.ps1 -sessionURL "sftp://moninvite@localhost:22/C:/Users/moninvite" -CSEntryName test -Filemask "*.tmp" -LocalPath 'C:\tools\Out' -command upload
# rc = 7
# .\Winscp.ps1 -sessionURL "ftp://moninvite@localhost:990/In" -CSEntryName test -Filemask "*" -LocalPath 'C:\tools\Out' -command upload -FtpSecure Implicit
# rc = 7
# .\Winscp.ps1 -sessionURL "ftp://moninvite@localhost:990/In" -CSEntryName test -Filemask "*" -LocalPath 'C:\tools\Out' -command upload -FtpSecure Implicit -IgnoreHostAuthenticityCheck
# rc = 0
# .\Winscp.ps1 -sessionURL "ftp://moninvite@localhost/In" -CSEntryName test -Filemask "*" -LocalPath 'C:\tools\Out' -command download -FtpSecure Explicit -IgnoreHostAuthenticityCheck
# rc = 0
# .\Winscp.ps1 -sessionURL "sftp://moninvite@localhost:/In" -CSEntryName test -Filemask "*" -LocalPath 'C:\tools\Out' -command download -FtpSecure Explicit -IgnoreHostAuthenticityCheck

$testName = "UPLOAD - SFTP - SMALL"
$expected = 0

$sessionURL = "sftp://nxy-dev@stc.qua.nexity.net:8022/TEST/" 
$CSEntryName = "nxy-dev"
$LocalPath = "D:\Tests\1-VariousSize"
$Command = "upload"
$testArgs = @{
	testName = $testName
	argList  = @{
		sessionURL      = $sessionURL
		CSEntryName     = $CSEntryName
		Filemask        = $Filemask
		LocalPath       = $LocalPath
		command			= $command
		IgnoreHostAuthenticityCheck = $true
	}
	expected = $expected
}
$r = runTest $testArgs
$results += [PSCustomObject] $r


$testName = "UPLOAD - SFTP - RemotePath"
$expected = 0

$sessionURL = "sftp://nxy-dev@stc.qua.nexity.net:8022/" 
$RemotePath = "TEST"
$CSEntryName = "nxy-dev"
$Filemask = "*small*.tst"
$LocalPath = "D:\Tests\1-VariousSize"
$Command = "upload"
$testArgs = @{
	testName = $testName
	argList  = @{
		sessionURL      = $sessionURL
		CSEntryName     = $CSEntryName
		Filemask        = $Filemask
		LocalPath       = $LocalPath
		command			= $command
		IgnoreHostAuthenticityCheck = $true
	}
	expected = $expected
}
$r = runTest $testArgs
$results += [PSCustomObject] $r


$testName = "DOWNLOAD - SFTP"
$expected = 0

$sessionURL = "sftp://nxy-dev@stc.qua.nexity.net:8022/TEST/" 
$CSEntryName = "nxy-dev"
$Filemask = "*small*.tst"
$LocalPath = "D:\Tests\Junk"
$Command = "download"
$testArgs = @{
	testName = $testName
	argList  = @{
		sessionURL      = $sessionURL
		CSEntryName     = $CSEntryName
		Filemask        = $Filemask
		LocalPath       = $LocalPath
		command			= $command
		
		IgnoreHostAuthenticityCheck = $true
		DeleteSourceFile = $true
	}
	expected = $expected
}
$r = runTest $testArgs
$results += [PSCustomObject] $r

$testName = "DOWNLOAD - SFTP - No Files"
$expected = 20

$sessionURL = "sftp://nxy-dev@stc.qua.nexity.net:8022/TEST/" 
$CSEntryName = "nxy-dev"
$Filemask = "*nofiles*.tst"
$LocalPath = "D:\Tests\Junk"
$Command = "download"
$IgnoreHostAuthenticityCheck = $true
$testArgs = @{
	testName = $testName
	argList  = @{
		sessionURL      = $sessionURL
		CSEntryName     = $CSEntryName
		Filemask        = $Filemask
		LocalPath       = $LocalPath
		Command			= $command
		
		IgnoreHostAuthenticityCheck = $true
	}
	expected = $expected
}
$r = runTest $testArgs
$results += [PSCustomObject] $r

$testName = "UPLOAD - SFTP - No Files"
$expected = 20

$sessionURL = "sftp://nxy-dev@stc.qua.nexity.net:8022/TEST/" 
$CSEntryName = "nxy-dev"
$Filemask = "*nofiles*.tst"
$LocalPath = "D:\Tests\Junk"
$Command = "upload"
$IgnoreHostAuthenticityCheck = $true
$testArgs = @{
	testName = $testName
	argList  = @{
		sessionURL      = $sessionURL
		CSEntryName     = $CSEntryName
		Filemask        = $Filemask
		LocalPath       = $LocalPath
		Command			= $command
		
		IgnoreHostAuthenticityCheck = $true
	}
	expected = $expected
}
$r = runTest $testArgs
$results += [PSCustomObject] $r


$testName = "UPLOAD - SFTP - Get UserName from Get-Cred"
$expected = 20

$sessionURL = "sftp://stc.qua.nexity.net:8022/TEST/" 
$CSEntryName = "nxy-dev"
$Filemask = "*nofiles*.tst"
$LocalPath = "D:\Tests\Junk"
$Command = "upload"
$IgnoreHostAuthenticityCheck = $true
$testArgs = @{
	testName = $testName
	argList  = @{
		sessionURL      = $sessionURL
		CSEntryName     = $CSEntryName
		Filemask        = $Filemask
		LocalPath       = $LocalPath
		Command			= $command
		
		IgnoreHostAuthenticityCheck = $true
	}
	expected = $expected
}
$r = runTest $testArgs
$results += [PSCustomObject] $r


$testName = "UPLOAD - SFTP - Bad Port"
$expected = 1

$sessionURL = "sftp://stc.qua.nexity.net:8888/TEST/" 
$CSEntryName = "nxy-dev"
$Filemask = "*nofiles*.tst"
$LocalPath = "D:\Tests\Junk"
$Command = "upload"
$IgnoreHostAuthenticityCheck = $true
$testArgs = @{
	testName = $testName
	argList  = @{
		sessionURL      = $sessionURL
		CSEntryName     = $CSEntryName
		Filemask        = $Filemask
		LocalPath       = $LocalPath
		Command			= $command
		
		IgnoreHostAuthenticityCheck = $true
	}
	expected = $expected
}
$r = runTest $testArgs
$results += [PSCustomObject] $r

$testName = "UPLOAD - FTP Unsecure"
$expected = 0

$sessionURL = "ftp://stc.qua.nexity.net:2021/TEST/" 
$CSEntryName = "nxy-dev"
$Filemask = "*small*.tst"
$LocalPath = "D:\Tests\1-VariousSize"
$Command = "upload"
$testArgs = @{
	testName = $testName
	argList  = @{
		sessionURL      = $sessionURL
		CSEntryName     = $CSEntryName
		Filemask        = $Filemask
		LocalPath       = $LocalPath
		Command			= $command
		IgnoreHostAuthenticityCheck = $true	
	}
	expected = $expected
}
$r = runTest $testArgs
$results += [PSCustomObject] $r

## Test Case

$testName = "UPLOAD - FTPS Explicit - Ignore"
$expected = 0

$sessionURL = "ftp://stc.qua.nexity.net:2021/TEST/" 
$CSEntryName = "nxy-dev"
$Filemask = "*small*.tst"
$LocalPath = "D:\Tests\1-VariousSize"
$Command = "upload"
$FtpSecure = "Explicit"
$testArgs = @{
	testName = $testName
	argList  = @{
		sessionURL      = $sessionURL
		CSEntryName     = $CSEntryName
		Filemask        = $Filemask
		LocalPath       = $LocalPath
		Command			= $command
		IgnoreHostAuthenticityCheck = $true	
		FtpSecure		  = $FtpSecure
	}
	expected = $expected
}
$r = runTest $testArgs
$results += [PSCustomObject] $r

# End

$results

if ($LogFolder) {
	Stop-Transcript
}