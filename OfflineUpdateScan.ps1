<#
.SYNOPSIS 

Scans a computer for Windows updates using the offline CAB file

.DESCRIPTION

Scans a computer fo Windows updates using the offline CAB file, or creates a Scheduled Task to do so.

.PARAMETER AddTask
Specifies that a Scheduled Task should be created.

.PARAMETER Run
Specifies that a scan should be run interactively

.PARAMETER Format
Specifies the report format: CSV, HTML or XML for files, Console for the ineractive screen.

.PARAMETER Path
Specifies the file path for the report file.

.PARAMETER StartAt
Specifies a Date/Time that the Scheduled Task should start. By default the task will start 45 seconds after it is created.

.PARAMETER CabSource
The location of the Microsoft Offline Update CAB file. Defaults to the user's Documents folder.

.INPUTS
None. You cannot pipe objects to this script.

.OUTPUTS
None to the pipeline.

.EXAMPLE
OfflineUpdateScan.ps1 -AddTask

This example creates a Scheduled Task that will run in 45 seconds time.

.EXAMPLE
OfflineUpdateScan.ps1 -Run

This example starts an offline scan.

.NOTES
#>


Param(
    [Parameter(Mandatory=$true,ParameterSetName='Task')]
    [Switch]
    $AddTask,
    [Parameter(Mandatory=$false,ParameterSetName='Task')]
    [DateTime]
    $StartAt,
    [Parameter(Mandatory=$true,ParameterSetName="Exec")]
    [Switch]
    $Run,
    [Parameter(ParameterSetName="Exec")]
    [Parameter(ParameterSetName="Task")]
    [ValidateSet("csv","xml","console","html")]
    [String]
    $Format="csv",
    [Parameter(Mandatory=$false,ParameterSetName="Exec")]
    [Parameter(Mandatory=$false,ParameterSetName="Task")]
    [String]
    $Path,
    [Parameter(ParameterSetName="Exec")]
    [String]
    $CabSource = "$env:userprofile\documents\wsusscn2.cab"
)
$ScriptGuid = '50bf2b41-ffb4-4381-b693-71a14f5874dd'

## Constant Enums for Schedule Tasks. Derived from taskschd.h
Add-Type -TypeDefinition @" 
public enum TASK_RUN_FLAGS
    {
        TASK_RUN_NO_FLAGS	= 0,
        TASK_RUN_AS_SELF	= 0x1,
        TASK_RUN_IGNORE_CONSTRAINTS	= 0x2,
        TASK_RUN_USE_SESSION_ID	= 0x4,
        TASK_RUN_USER_SID	= 0x8
    }
public enum TASK_ENUM_FLAGS
    {
        TASK_ENUM_HIDDEN	= 0x1
    }
public enum TASK_LOGON_TYPE
    {
        TASK_LOGON_NONE	= 0,
        TASK_LOGON_PASSWORD	= 1,
        TASK_LOGON_S4U	= 2,
        TASK_LOGON_INTERACTIVE_TOKEN	= 3,
        TASK_LOGON_GROUP	= 4,
        TASK_LOGON_SERVICE_ACCOUNT	= 5,
        TASK_LOGON_INTERACTIVE_TOKEN_OR_PASSWORD	= 6
    }
public enum TASK_RUNLEVEL
    {
        TASK_RUNLEVEL_LUA	= 0,
        TASK_RUNLEVEL_HIGHEST	= 1
    }
public enum TASK_PROCESSTOKENSID
    {
        TASK_PROCESSTOKENSID_NONE	= 0,
        TASK_PROCESSTOKENSID_UNRESTRICTED	= 1,
        TASK_PROCESSTOKENSID_DEFAULT	= 2
    }
public enum TASK_STATE
    {
        TASK_STATE_UNKNOWN	= 0,
        TASK_STATE_DISABLED	= 1,
        TASK_STATE_QUEUED	= 2,
        TASK_STATE_READY	= 3,
        TASK_STATE_RUNNING	= 4
    }
public enum TASK_CREATION
    {
        TASK_VALIDATE_ONLY	= 0x1,
        TASK_CREATE	= 0x2,
        TASK_UPDATE	= 0x4,
        TASK_CREATE_OR_UPDATE	= ( TASK_CREATE | TASK_UPDATE ) ,
        TASK_DISABLE	= 0x8,
        TASK_DONT_ADD_PRINCIPAL_ACE	= 0x10,
        TASK_IGNORE_REGISTRATION_TRIGGERS	= 0x20
    }
public enum TASK_TRIGGER_TYPE2
    {
        TASK_TRIGGER_EVENT	= 0,
        TASK_TRIGGER_TIME	= 1,
        TASK_TRIGGER_DAILY	= 2,
        TASK_TRIGGER_WEEKLY	= 3,
        TASK_TRIGGER_MONTHLY	= 4,
        TASK_TRIGGER_MONTHLYDOW	= 5,
        TASK_TRIGGER_IDLE	= 6,
        TASK_TRIGGER_REGISTRATION	= 7,
        TASK_TRIGGER_BOOT	= 8,
        TASK_TRIGGER_LOGON	= 9,
        TASK_TRIGGER_SESSION_STATE_CHANGE	= 11,
        TASK_TRIGGER_CUSTOM_TRIGGER_01	= 12
    }
public enum TASK_SESSION_STATE_CHANGE_TYPE
    {
        TASK_CONSOLE_CONNECT	= 1,
        TASK_CONSOLE_DISCONNECT	= 2,
        TASK_REMOTE_CONNECT	= 3,
        TASK_REMOTE_DISCONNECT	= 4,
        TASK_SESSION_LOCK	= 7,
        TASK_SESSION_UNLOCK	= 8
    }
public enum TASK_ACTION_TYPE
    {
        TASK_ACTION_EXEC	= 0,
        TASK_ACTION_COM_HANDLER	= 5,
        TASK_ACTION_SEND_EMAIL	= 6,
        TASK_ACTION_SHOW_MESSAGE	= 7
    }
public enum TASK_INSTANCES_POLICY
    {
        TASK_INSTANCES_PARALLEL	= 0,
        TASK_INSTANCES_QUEUE	= 1,
        TASK_INSTANCES_IGNORE_NEW	= 2,
        TASK_INSTANCES_STOP_EXISTING	= 3
    }
public enum TASK_COMPATIBILITY
    {
        TASK_COMPATIBILITY_AT	= 0,
        TASK_COMPATIBILITY_V1	= 1,
        TASK_COMPATIBILITY_V2	= 2,
        TASK_COMPATIBILITY_V2_1	= 3,
        TASK_COMPATIBILITY_V2_2	= 4,
        TASK_COMPATIBILITY_V2_3	= 5,
        TASK_COMPATIBILITY_V2_4	= 6
    }
"@
#Constants from wuapi.h
Add-Type -TypeDefinition @"
public enum OperationResultCode
    {
        orcNotStarted	= 0,
        orcInProgress	= 1,
        orcSucceeded	= 2,
        orcSucceededWithErrors	= 3,
        orcFailed	= 4,
        orcAborted	= 5
    }
"@
#Constant from wuapicommon.h
Add-Type -TypeDefinition @"
public enum ServerSelection
    {
        ssDefault	= 0,
        ssManagedServer	= 1,
        ssWindowsUpdate	= 2,
        ssOthers	= 3
    } 
"@
Function Remove-OfflineUpdateScantask {
    $STService = New-Object -ComObject Schedule.Service 
    $STService.Connect()
    $RootFolder = $STService.GetFolder("\")
    try {
        $RootFolder.DeleteTask($Script:ScriptGuid,$Null)
    }
    catch {}
}
Function Add-OfflineUpdateScanTask {

    $STService = New-Object -ComObject Schedule.Service 
    $STService.Connect()

    $RootFolder = $STService.GetFolder("\")

    $NewTaskDef = $STService.NewTask(0)
    $RegInfo = $NewTaskDef.RegistrationInfo
    $RegInfo.Description = "Offline Update Scan"
    $RegInfo.Author = "Stuart Squibb"

    $Principal = $NewTaskDef.Principal
    $Principal.LogonType = [TASK_LOGON_TYPE]::Task_Logon_Service_Account
    $Principal.UserId = 'NT AUTHORITY\SYSTEM'
    $Principal.Id = "System"
    $Principal | Select-Object * | Write-Verbose
    $Settings = $NewTaskDef.Settings
    $Settings.Enabled = $True
    $Settings.DisallowStartIfOnBatteries = $False

    $Trigger = $NewTaskDef.Triggers.Create([TASK_TRIGGER_TYPE2]::TASK_TRIGGER_TIME)

    if ($Script:StartAt) {
        $StartTime = $Script:StartAt
        Write-Verbose $StartTime
    }
    else {
        $StartTime = (Get-Date).AddSeconds(45)
    }
     
    $EndTime = ($StartTime.AddMinutes(5)).ToString("yyyy-MM-ddTHH:mm:ss")
    $StartTime = $StartTime.toString("yyyy-MM-ddTHH:mm:ss")

    Write-Verbose "Time Now  : $((Get-Date).ToString('yyyy-MM-ddTHH:mm:ss'))"
    Write-Verbose "Start Time: $($StartTime)"
    Write-Verbose "End Time  : $($EndTime)"

    $Trigger.StartBoundary = $StartTime
    $Trigger.EndBoundary = $EndTime
    $Trigger.ExecutionTimeLimit = "PT5M"
    $Trigger.Id = "TimeTriggerId"
    $Trigger.Enabled = $True 

    $Action = $NewTaskDef.Actions.Create([TASK_ACTION_TYPE]::TASK_ACTION_EXEC)
    $Action.Path = "C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe"
    $Action.Arguments = "-ExecutionPolicy ByPass -NoProfile -NonInteractive -File C:\OfflineUpdateScan\OfflineUpdateScan.ps1 -Run -Format $Format -Path $Path"
    $Action.WorkingDirectory = "C:\OfflineUpdateScan"

    Write-Verbose "Task Definition created. About to submit Task..."

    [Void]$RootFolder.RegisterTaskDefinition($ScriptGuid, $NewTaskDef,[TASK_CREATION]::TASK_CREATE_OR_UPDATE,$Null,$Null,$Null)

    Write-Verbose "Task $ScriptGuid Submitted"
}

#Microsoft link for CAB file: http://go.microsoft.com/fwlink/?LinkID=74689
#WIN-90CID1J2CS5

$WorkDirectory = 'C:\OfflineUpdateScan'
$CabLocation = "$WorkDirectory\wsusscn2.cab"

if (!($Path)) {
    $Path = "C:\OfflineUpdateScan\OfflineUpdateScan.$($Format)"
}

# get-updatecollection
Function Get-OfflineUpdateCollection {
    $UpdateServiceManager = New-Object -ComObject "Microsoft.Update.ServiceManager"
    $UpdateService = $UpdateServiceManager.AddScanPackageService("Offline Sync Service", $CabLocation, 1)

    $UpdateSession = New-Object -ComObject "Microsoft.Update.Session"
    $UpdateSearcher = $UpdateSession.CreateUpdateSearcher()
    
    Write-Verbose "Searching for updates..."

    $UpdateSearcher.ServerSelection = [ServerSelection]::SSOthers
    $UpdateSearcher.ServiceID = $UpdateService.ServiceID
    $SearchResult = $UpdateSearcher.Search("IsInstalled=0")
    $Updates = $SearchResult.Updates
    
    If ($Updates.Count -eq 0) {
        #This area blank by design
     }
    else {
        $Updates
    }
}

# export-updatecollection
Function Export-OfflineUpdateCollection {
Param (
    [Parameter(Mandatory=$True)]
    [ValidateSet("xml","csv","console","html")]
    [String]
    $Format,
    [Parameter(Mandatory=$False)]
    [String]
    $FileName,
    [Parameter(ValueFromPipeline=$True,Mandatory=$True)]
    [Object]
    $OfflineUpdateCollection
)
$OutPutObject = Select-Object -InputObject $OfflineUpdateCollection -Property MsrcSeverity, Title, MaxDownloadSize, MinDownloadSize, @{Name="KBs";Expression={$_.KBArticleIds -join ';'}}    
switch ($Format) {
        'csv'  { $OutPutObject | Export-Csv -Path $FileName -NoTypeInformation  }
        'xml'  { ($OutPutObject | ConvertTo-Xml -NoTypeInformation -As Document).OuterXML | Out-File -FilePath $FileName } #Export-Clixml -Path $FileName
        'html' { $OutPutObject | ConvertTo-Html -Title "Needed Windows Updates for $env:ComputerName" | Out-File -FilePath $FileName}
        'console' { Format-Table -InputObject $OutPutObject}
    }
}

if ($AddTask.IsPresent) {
    Add-OfflineUpdateScanTask
}

if ($Run.IsPresent) {
    try {
        New-Item -ItemType Directory -Path $WorkDirectory
    }
    catch {}

    try {
        Copy-Item $CabSource -Destination $CabLocation
    }
    catch {}

    Write-Verbose "Exporting $Format format file to $Path"
    Get-OfflineUpdateCollection | Export-OfflineUpdateCollection -Format $Format -FileName  $Path
    Remove-OfflineUpdateScantask
}





