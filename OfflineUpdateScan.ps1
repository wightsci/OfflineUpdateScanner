$orcNotStarted	= 0
$orcInProgress	= 1
$orcSucceeded	= 2
$orcSucceededWithErrors	= 3
$orcFailed	= 4
$orcAborted	= 5


$UpdateSession = New-Object -ComObject "Microsoft.Update.Session"
$UpdateServiceManager = New-Object -ComObject "Microsoft.Update.ServiceManager"
$UpdateService = $UpdateServiceManager.AddScanPackageService("Offline Sync Service", "$env:userprofile\Downloads\wsusscn2.cab", 1)
$UpdateSearcher = $UpdateSession.CreateUpdateSearcher()


Write-Output "Searching for updates..."

$UpdateSearcher.ServerSelection = 3 # ssOthers

$UpdateSearcher.ServiceID = $UpdateService.ServiceID

$SearchResult = $UpdateSearcher.Search("IsInstalled=0")

$Updates = $SearchResult.Updates

If ($searchResult.Updates.Count -eq 0) {
    Write-Output "There are no applicable updates."
    $UpdateServiceManager.Services | Where-Object { $_.Name -eq 'Offline Sync Service' } | ForEach-Object { $UpdateServiceManager.RemoveService($_.ServiceID) }
    Exit
}

Write-Output "List of applicable items on the machine when using wssuscan.cab:"

For ($I = 0;$I -le $searchResult.Updates.Count-1;$I++) {
    $update = $searchResult.Updates.Item($I)
    Write-Output "$($I + 1)> $($update.Title)"
}

$UpdateServiceManager.Services | Where-Object { $_.Name -eq 'Offline Sync Service' } | ForEach-Object { $UpdateServiceManager.RemoveService($_.ServiceID) }