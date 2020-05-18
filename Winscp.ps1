param (
    [string] $winscpPath = "C:\Program Files (x86)\WinSCP\WinSCPnet.dll",
    [string] $localPath,
    [string] $remotePath,
    [string] $hostname,
    [string] $user,
    [int]    $port = 22,
    [string] $password,
    [string] $filemask = $null, # $null
    [string] $direction,
    [string] $protocol = "sftp",
    [string] $serverFingerprint,
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

if(($protocol -eq "sftp") -or ($protocol -eq "scp"))
{
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
    PortNumber = $port
    SshHostKeyFingerprint = "ssh-rsa 2048 e0:a3:0f:1a:04:df:5a:cf:c9:81:84:4e:08:4c:9a:06"
}
 
$session = New-Object WinSCP.Session

if($direction -eq "upload")
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
    
        # Print results
        foreach ($transfer in $transferResult.Transfers)
        {
            Write-Host "Upload of $($transfer.FileName) succeeded"
        }
    }
    catch [Exception]
    {
        Write-Host $session.Output
        Write-Host $_.Exception.Message
        Exit 7
    }
    finally
    {
        # Disconnect, clean up
        $session.Dispose()
    }
}
elseif($direction -eq "download")
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
            Write-Host "Upload of $($transfer.FileName) succeeded"
        }
    }
    catch [Exception]
    {
        Write-Host $session.Output
        Write-Host $_.Exception.Message
        Exit 7
    }
    finally
    {
        # Disconnect, clean up
        $session.Dispose()
    }    
}
else 
{
    Write-Host "Direction not specified, must be 'upload' or 'download'"
    Exit 8
}
