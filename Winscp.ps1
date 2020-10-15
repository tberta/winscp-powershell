[CmdletBinding(DefaultParameterSetName = 'combined')]
param (
    [Parameter(ParameterSetName = 'combined')]
    [Parameter(ParameterSetName = 'splitted')]
    [string] $winscpPath = "C:\Program Files (x86)\WinSCP\WinSCPnet.dll",
    [Parameter(ParameterSetName = 'combined')]
        [string] $sessionURL, 
    #sessionURL are like:
    # <protocol> :// [ <username> [ : <password> ] [ ; <advanced> ] @ ] <host> [ : <port> ] /
    [Parameter(ParameterSetName = 'combined')]
    [Parameter(ParameterSetName = 'splitted')]
        [string] $localPath,
    [Parameter(ParameterSetName = 'combined')]
    [Parameter(ParameterSetName = 'splitted')]
        [string] $remotePath,
    [Parameter(ParameterSetName = 'splitted')]
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
    [Parameter(ParameterSetName = 'combined')]
    [Parameter(ParameterSetName = 'splitted')]
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
