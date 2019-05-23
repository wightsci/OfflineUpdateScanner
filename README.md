# OfflineUpdateScanner
Utility to use Microsoft's offline update CAB file to scan a computer for Windows Updates

If you have computers that are not connected to the internet it isn't straightforward to scan them for Windows Updates. This PowerShell script is designed to help.

This script can scan a machine directly using Microsoft's Offline Update CAB, or create a Scheduled Task to do so. Output formats include CSV, XML and HTML.

Tested on Windows 7 (PowerShell 3.0) and Windows Server 2016


See my blog post for more [details](https://carisbrookelabs.wordpress.com/2019/05/23/offline-windows-update-scans-using-powershell/).

Example:
```PowerShell
OfflineUpdateScan.ps1 -Run -Format html -CabSource .\wsusscn2.cab
```

Example:
```PowerShell
OfflineUpdateScan.ps1 -AddTask -Format html
```
