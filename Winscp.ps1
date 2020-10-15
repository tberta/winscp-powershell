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
$user = "mylogin"
$pass = "mypass"
$bucket = "s3-my-bucketname-001"
$winscp = "C:\Programs Files\WinSCP\WinSCPnet.dll"
$sessionURL = "s3://s3.amazonaws.com/$($bucket)/"
$localPath = "C:\To_upload"
$remotePath = "incoming"
$filemask = "*.txt"
$command = "upload"
PS> Winscp.ps1 -winscpPath $winscp -sessionURL $sessionURL -localPath $localPath -filemask $filemask -remotePath $remotePath -command $command -password $pass -user $user

.EXAMPLE
$user = "mylogin"
$pass = "mypass"
$pass = [uri]::EscapeDataString($pass) # To make it valid in sessionURL
$bucket = "s3-my-bucketname-001"
$sessionURL = "s3://$($user):$($pass)@s3.amazonaws.com/$($bucket)/"
$localPath = "C:\To_upload"
$remotePath = "incoming"
$command = "upload"
PS> Winscp.ps1 -winscpPath $winscp -sessionURL $sessionURL -localPath $localPath -filemask $filemask -remotePath $remotePath -command $command -password $pass -user $user

.INPUTS

None. You cannot pipe objects to Winscp1.ps1.

.OUTPUTS

None. Besides some Console output, Winscp1.ps1 does not generate any output object.

.LINK
https://github.com/tberta/winscp-powershell

#>
[CmdletBinding(DefaultParameterSetName = 'combined')]
param (
    [Parameter(ParameterSetName = 'combined')]
    [Parameter(ParameterSetName = 'splitted')]
    [string] $winscpPath = "C:\Program Files (x86)\WinSCP\WinSCPnet.dll",
    [Parameter(ParameterSetName = 'combined')]
        [string] $sessionURL, 
    [Parameter(ParameterSetName = 'combined', Mandatory=$true)]
    [Parameter(ParameterSetName = 'splitted', Mandatory=$true)]
        [string] $localPath,
    [Parameter(ParameterSetName = 'combined')]
    [Parameter(ParameterSetName = 'splitted')]
        [string] $remotePath,
    [Parameter(ParameterSetName = 'splitted', Mandatory=$true)]
        [string] $hostname,
    [Parameter(ParameterSetName = 'combined')]
    [Parameter(ParameterSetName = 'splitted')]
        [string] $user,
    [Parameter(ParameterSetName = 'splitted')]
        [int]    $port = 22,
    [Parameter(ParameterSetName = 'combined')]
    [Parameter(ParameterSetName = 'splitted')]
        [string] $password,
    [Parameter(ParameterSetName = 'combined')]
    [Parameter(ParameterSetName = 'splitted')]
        [string] $filemask = $null,
    [Parameter(ParameterSetName = 'combined', Mandatory=$true)]
    [Parameter(ParameterSetName = 'splitted', Mandatory=$true)]
    [ValidateSet('download','upload')]
        [string] $command,
    [Parameter(ParameterSetName = 'splitted')]
    [ValidateSet('sftp','ftp','s3','scp','webdav')]
        [string] $protocol = "sftp",
    [Parameter(ParameterSetName = 'splitted')]
        [string] $serverFingerprint,
    [Parameter(ParameterSetName = 'combined')]
    [Parameter(ParameterSetName = 'splitted')]
        [switch] $deleteSourceFile = $false
    # [Parameter(ParameterSetName = 'combined')]
    # [Parameter(ParameterSetName = 'splitted')]
    #     [switch] $verbose = $false
    )
         
try
{
    # Load WinSCP .NET assembly
    Add-Type -Path "$winscpPath"
}
catch [Exception]
{
    Write-Host "Error: "$_.Exception.Message
    Exit 1
}

if(-not (Test-Path -Path $localPath))
{
    Write-Error "Local folder doesn't exist" -Category ObjectNotFound
    Write-Verbose "When running in OpCon, you need an extra *ending* slash for localPath : -localPath=""C:\Test\\"" " -Verbose
    Write-Verbose "Received argument: '$localPath'" -Verbose
    Exit 2
}

if($sessionURL) {
    $sessionOptions = New-Object WinSCP.SessionOptions
    try {
        $sessionOptions.ParseURL($sessionURL)    
    }
    catch [Exception]
    {
        Write-Host "Error while parsing provided sessionURL argument : '$sessionURL'"
        Write-Host $_.Exception.Message
        Exit 3
    }
    # User or Password may need to be HTML encoded in a sessionURL
    # So for convenience, User and Password value can be passed as arguments
    if (-not ($null -eq $password)) { $sessionOptions.Password = $password }
    if (-not ($null -eq $user))     { $sessionOptions.UserName = $user }

} else {
    switch ($protocol)
    {
        "sftp"   {  $winscpProtocol = [WinSCP.Protocol]::sftp   }
        "ftp"    {  $winscpProtocol = [WinSCP.Protocol]::ftp    }
        "s3"     {  $winscpProtocol = [WinSCP.Protocol]::s3     }
        "scp"    {  $winscpProtocol = [WinSCP.Protocol]::scp    }
        "webdav" {  $winscpProtocol = [WinSCP.Protocol]::webdav }
        default  
        {     
            Write-Host "Unknown protocol specified"
            Write-Host "Exiting..."
            Exit 4
        }
    }

    if(($protocol -eq "sftp") -or ($protocol -eq "scp")) {
        if (-not $serverFingerprint)
        {
            
            "Protocol $protocol specified but serverFingerprint is missing. Must be defined "
            'Argument Example : -serverFingerprint "ssh-rsa 2048 e0:a3:0f:1a:04:df:5a:cf:c9:81:84:4e:08:4c:9a:06"'
            Exit 5
        }
    }

    # Setup session options
    $sessionOptions = New-Object WinSCP.SessionOptions -Property @{
        Protocol = $winscpProtocol
        HostName = $hostname
        UserName = $user
        Password = $password
    }
    if (-not ($null -eq $port))                { $sessionOptions.PortNumber = $port }
    if (-not ($null -eq $serverFingerprint))   { $sessionOptions.SshHostKeyFingerprint = $serverFingerprint }
}

$returnCode = 0
$session = New-Object WinSCP.Session

try 
{
    # Connect
    $session.Open($sessionOptions)

    if($command -eq "upload") {
        # Upload files
        #$transferOptions = New-Object WinSCP.TransferOptions
        #$transferOptions.TransferMode = [WinSCP.TransferMode]::Binary
        #$transferResult = $session.PutFiles($localPath,($remotePath + $filemask), $deleteSourceFile , $transferOptions)
        $transferResult = $session.PutFilestoDirectory($localPath,$remotePath, $filemask, $deleteSourceFile)
    }

    if ($command -eq "download") {
        # Download the file and throw on any error
        #$sessionResult = $session.GetFiles(($remotePath + $fileName),($localPath + $filemask))
        $transferResult = $session.GetFilesToDirectory($remotePath, $localPath, $filemask, $deleteSourceFile)
    }

    # Throw error if found
    $transferResult.Check()
    
    if($VerbosePreference) {
        foreach ($transfer in $transferResult.Transfers)
        {
            Write-Verbose "$($transfer.FileName) : $command succeed"
        }
        foreach ($failure in $transferResult.Failures)
        {
            Write-Verbose "$($transfer.FileName): $command did NOT succeed"
        }
    }
    if ($($transferResult.IsSuccess)) {
        Write-Host "$command transfer job ended successfully"
    }
    Write-Host "Files transferred successfully : $($transferResult.Transfers.count)"
    if ($($transferResult.Failures.Count)) {
        Write-Host "Files not transferred : $($transferResult.Failures.Count)"
    }

}
catch [Exception]
{
    Write-Host $session.Output
    Write-Host $_.Exception.Message
    $returnCode = 7
}    
finally
{
    # Disconnect, clean up
    $session.Dispose()
    exit $returnCode
}    
