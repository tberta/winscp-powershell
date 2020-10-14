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
    [Parameter(ParameterSetName = 'splitted')]
        [string] $user,
    [Parameter(ParameterSetName = 'splitted')]
        [int]    $port = 22,
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

if($sessionURL) {
    $sessionOptions = New-Object WinSCP.SessionOptions
    try {
        $sessionOptions.ParseURL($sessionURL)    
    }
    catch [Exception]
    {
        Write-Host "Error while parsing provided sessionURL argument : '$sessionURL'"
        Write-Host $_.Exception.Message
        Exit 1
    }
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
            Exit 2
        }
    }

    if(($protocol -eq "sftp") -or ($protocol -eq "scp")) {
        if (-not $serverFingerprint)
        {
            
            "Protocol $protocol specified but serverFingerprint is missing. Must be defined "
            'Argument Example : -serverFingerprint "ssh-rsa 2048 e0:a3:0f:1a:04:df:5a:cf:c9:81:84:4e:08:4c:9a:06"'
            Exit 3
        }
    }

    # Setup session options
    $sessionOptions = New-Object WinSCP.SessionOptions -Property @{
        Protocol = $winscpProtocol
        HostName = $hostname
        UserName = $user
        Password = $password
    }
    if (-not ($null -eq $port))                 { $sessionOptions.Add("PortNumber"              , $port) }
    if (-not ($null -eq $serverFingerprint))    { $sessionOptions.Add("SshHostKeyFingerprint"   , $serverFingerprint) }
}

#debug: 
$sessionOptions

$returnCode = 0
$session = New-Object WinSCP.Session

if($command -eq "upload") 
{
    try
    {
        # Connect
        $session.Open($sessionOptions)

        # Upload files
        #$transferOptions = New-Object WinSCP.TransferOptions
        #$transferOptions.TransferMode = [WinSCP.TransferMode]::Binary
        #$transferResult = $session.PutFiles($localPath,($remotePath + $filemask), $deleteSourceFile , $transferOptions)
        $transferResult = $session.PutFilestoDirectory($localPath,$remotePath, $filemask, $deleteSourceFile)

        # Throw on any error
        $transferResult.Check()
        $transferResult
        # Print results
        foreach ($transfer in $transferResult.Transfers)
        {
            Write-Host "$($transfer.FileName): Upload succeed"
        }
        foreach ($failure in $transferResult.Failures)
        {
            Write-Host "$($transfer.FileName): Upload did NOT succeed"
        }
    }
    catch [Exception]
    {
        Write-Host $session.Output
        Write-Host $_.Exception.Message
        $returnCode = 5
    }
    finally
    {
        # Disconnect, clean up
        $session.Dispose()
        exit $returnCode
    }
}
elseif($command -eq "download")
{
    try
    {
        # Connect
        $session.Open($sessionOptions)

        # Download Files from remote
        
        # Download the file and throw on any error
        #$sessionResult = $session.GetFiles(($remotePath + $fileName),($localPath + $filemask))
        $sessionResult = $session.GetFilesToDirectory($remotePath, $localPath, $filemask, $deleteSourceFile)
        
        # Throw error if found
        $sessionResult.Check()

        foreach ($transfer in $transferResult.Transfers)
        {
            Write-Host "$($transfer.FileName) : Download succeed"
        }
        foreach ($failure in $transferResult.Failures)
        {
            Write-Host "$($transfer.FileName): Download did NOT succeed"
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
}
else 
{
    Write-Host "Command not specified, must be 'upload' or 'download'"
    Exit 8
}
