# WinSCP
This repo contains an example of running the <a href url="https://winscp.net/eng/index.php">WINSCP</a> ftp program from a PowerShell script.

The intent was to keep this version as generic as possible so it could be used for as many different sites as possible.  If you are running this from OpCon it should fail (with an Exit Code of 1) if there are any issues.  Any other pertinant logging will be in the job output.

# Prerequisites
* WinSCP 5.15.9 : Download [WinSCP .NET / COM Libray](https://winscp.net/eng/downloads.php)
* PowerShell 5.1

# Instructions
This script contains several parameters:<br>
* <b>winscpPath</b> - Path to the WINSCP .Net dll <br>
* <b>localPath</b> - path to the location in your local environment (when running in OpCon you need an extra ending slash) <br>
* <b>remotePath</b> - remote folder on the remote host you connect to <br>
* <b>hostname</b> - remote host you are connecting to <br>
* <b>protocol</b> - protocol to use with remote host<br>
* <b>port</b> - remote port to connect to<br>
* <b>user</b> - username used for authenticating to the remote site <br>
* <b>password</b> - password used for authentication to the remote site (recommend OpCon encrypted global property) <br>
* <b>filemask</b> - name of the file you are trying to upload/download (supports wildcards) <br>
* <b>direction</b> - upload or download<br>
* <b>serverFingerprint</b> - When using SCP or SFTP, specify remote server host key fingerprint in following format with hex value "ssh-rsa 2048 xx:xx:xx:xx:xx:xx:xx:xx:xx:xx:xx:xx:xx:xx:xx:xx"<br>
* <b>deleteSourceFile</b> - to delete source files after successful transmission<br>
<br>
Execution example:<br>

```
powershell.exe -ExecutionPolicy Bypass -File winscp.ps1 -winscpPath "C:\Program Files (x86)\WinSCP\WinSCPnet.dll" -localPath "C:\\" -remotePath "/somedirectory" -filemask "files*.txt" -hostname "someftp.com" -port 822 -user "xman" -password "encrypted" -direction "download" -protocol "sftp" -serverFingerprint "ssh-rsa 2048 xx:xx:xx:xx:xx:xx:xx:xx:xx:xx:xx:xx:xx:xx:xx:xx" -deleteSourceFile
```
# Disclaimer
No Support and No Warranty are provided by SMA Technologies for this project and related material. The use of this project's files is on your own risk.

SMA Technologies assumes no liability for damage caused by the usage of any of the files offered here via this Github repository.

# License
Copyright 2020 SMA Technologies

Licensed under the Apache License, Version 2.0 (the "License");
you may not use this file except in compliance with the License.
You may obtain a copy of the License at [apache.org/licenses/LICENSE-2.0](http://www.apache.org/licenses/LICENSE-2.0)

Unless required by applicable law or agreed to in writing, software
distributed under the License is distributed on an "AS IS" BASIS,
WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
See the License for the specific language governing permissions and
limitations under the License.

# Contributing
We love contributions, please read our [Contribution Guide](CONTRIBUTING.md) to get started!

# Code of Conduct
[![Contributor Covenant](https://img.shields.io/badge/Contributor%20Covenant-v2.0%20adopted-ff69b4.svg)](code-of-conduct.md)
SMA Technologies has adopted the [Contributor Covenant](CODE_OF_CONDUCT.md) as its Code of Conduct, and we expect project participants to adhere to it. Please read the [full text](CODE_OF_CONDUCT.md) so that you can understand what actions will and will not be tolerated.
