#Header
$header = 
@"
    <style type='text/css'>
    body {font-size:13pt;}
    p {color:lightgrey;font-size:10pt;margin-top:0px;}
    table, th, td, tr {border:1px solid lightgrey}
    table {font-family:Calibri;font-size:11pt;}
    th {background-color:darkSlateBlue;text-transform: uppercase;color:white;text-align:center;padding-top:2px;padding-bottom:2px;padding-left:5px;padding-right:5px;}
    td {padding:2px 5px 2px 5px;}
    h1, h2, h3, h4, h5 {font-family:Calibri;color:DarkBlue;margin-top:5px;margin-bottom:5px;}
    </style>
"@

$ColorTagTable = 
@{
    Stopped = ' style=background-color:OrangeRED;color:Yellow;>Stopped<';
    Running = ' style=background-color:Lime;>Running<';
    Online = ' style=background-color:Lime;>Online<';
    Offline = ' style=background-color:OrangeRED;color:Yellow;>Offline<';
    Failed  = ' style=background-color:OrangeRED;color:Yellow;>Failed<';
    Cancelled = ' style=background-color:Orange;>Cancelled<';
    Up = ' style=background-color:Lime;>Up<';
    Down = ' style=background-color:OrangeRED;color:Yellow;>Down<';
    Restoring = ' style=background-color:Orange;>Restoring<';
	RECOVERING = ' style=background-color:Yellow;>RECOVERING<';
    Success = ' style=background-color:Lime;>Success<';
}
$Red = "style=background-color:OrangeRED;color:Yellow;"
$Amber = "style=background-color:Orange;"
$Green = "style=background-color:Lime;"

#Parameters
$ServerName = ([System.Net.Dns]::GetHostByName(($env:computerName))).Hostname
$ServicesList = Get-Content -Path 'G:\PS_Health_Check_Report\Services.txt'
$AppPoolList = Get-Content -Path 'G:\PS_Health_Check_Report\AppPool.txt'
$SQLscriptsPath = "G:\ServerMonitor_Powershell\SQL_Scripts\"
$SQLInstance = 'GOAAZRAPP1018\EAUDIT'
$SQLService = 'MSSQL$EAUDIT'
$TimeNow = Get-Date

#TimeZone
$TimeZone = Get-TimeZone | Select-Object -ExpandProperty StandardName
$date = Get-Date -UFormat "%m/%d/%Y %r"
$TimeZone = "This report is generated from Time Zone - $TimeZone ($date)"

#Report Header
$ReportHeader = "Health Check Report of $ServerName"

#Server Recycle time
$ServerRecycleTime = Get-CimInstance -ClassName win32_operatingsystem | 
Select-Object @{Name="Server";Expression={$_.csname}},@{Name="LastRestartTime";Expression={[datetime] $_.lastbootuptime -F 'MM/dd/yyyy HH:mm:ss'} }
$SLastRestartTime = [datetime]$ServerRecycleTime.LastRestartTime
$TimeDiff = New-TimeSpan -Start $SLastRestartTime -End $TimeNow
$TimeDiff = $TimeDiff.TotalDays
$ServerRecycleTime = $ServerRecycleTime  | ConvertTo-Html -Head $header -PreContent "<h3>Server Recycle Time</h3>"
    if($TimeDiff -lt 1){
    $ServerRecycleTime = $ServerRecycleTime -replace "<td>$SLastRestartTime","<td $red>$SLastRestartTime" }
    else {
    $ServerRecycleTime = $ServerRecycleTime -replace "<td>$SLastRestartTime","<td $green>$SLastRestartTime" }

#Server UP time
$bootuptime = (Get-CimInstance -ClassName Win32_OperatingSystem).LastBootUpTime
$uptime = $TimeNow - $bootuptime | ConvertTo-Html -Head $header -PreContent "<h3>Server Uptime</h3>" -Property Days, Hours, Minutes 

#Azure Backup Report
$servers = $env:computername
$collection = $()
foreach ($server in $servers)
{
    $status = @{ "ServerName" = $server; "TimeStamp" = (Get-Date -Format yyyy-MM-dd)};
    if (((Get-WinEvent -ComputerName $server -FilterHashtable @{Logname='CloudBackup'; ID = 1; StartTime=(Get-Date).AddHours(-24)}) -and 
(Get-WinEvent -ComputerName $server -FilterHashtable @{Logname='CloudBackup'; ID = 3; StartTime=(Get-Date).AddHours(-24)})))
    { 
    $status["Results"] = "Backup is Successful"
    } 
    elseif (Get-WinEvent -ComputerName $server -FilterHashtable @{Logname='CloudBackup'; ID = 11; StartTime=(Get-Date).AddHours(-24)})
    { 
    $status["Results"] = "Backup is Failed" 
    }
    elseif (Get-WinEvent -ComputerName $server -FilterHashtable @{Logname='CloudBackup'; ID = 1; StartTime=(Get-Date).AddHours(-24)})
    {
    $status["Results"] = "Backup is Running/Errors Occured"
    }
New-Object -TypeName PSObject -Property $status -OutVariable serverStatus
$collection += $serverStatus
} 
$AzureBackupInfo = $collection | ConvertTo-Html -Head $header -PreContent "<h3> Azure backup Status</h3>" -Property Sl.No,ServerName,TimeStamp,Results
$AzureBackupInfo = $AzureBackupInfo -replace "<td>Backup is Successful<","<td $Green>Backup is Successful<"
$AzureBackupInfo = $AzureBackupInfo -replace "<td>Backup is Failed<","<td $Red>Backup is Failed<"
$AzureBackupInfo = $AzureBackupInfo -replace "<td>Backup is Running/Errors Occured<","<td $Amber>Backup is Running/Errors Occured<"

#Instance Recycle Time
$InstanceRecycleTime = invoke-sqlcmd –ServerInstance $SQLInstance -Query "Select @@SERVERNAME as InstanceName, LastRestartTime = create_date from sys.databases where name = 'tempdb'" | 
Select * -ExcludeProperty RowError, RowState, Table, ItemArray, HasErrors| Convertto-html -Head $header -PreContent "<h3>Instance Recycle Time</h3>"

#CPU Usage
$TotalCPUUsedPct = Get-WmiObject Win32_Processor | Measure-Object -Property LoadPercentage -Average | Select-Object -ExpandProperty Average 
$CPUReport = "<h3>CPU Usage</h3>
<table>
        <tr><th>CPUUsage</th></tr>
        <tr><td>$TotalCPUUsedPct %</td></tr>
        </table>"

if($TotalCPUUsedPct -ge 80 -and $TotalCPUUsedPct -lt 85){
    $CPUReport = $CPUReport -replace "<td>$TotalCPUUsedPct %","<td $Amber>$TotalCPUUsedPct %" }
    elseif($TotalCPUUsedPct -ge 85){
	$CPUReport = $CPUReport -replace "<td>$TotalCPUUsedPct %","<td $Red>$TotalCPUUsedPct %" }
    else{
    $CPUReport = $CPUReport -replace "<td>$TotalCPUUsedPct %","<td $Green>$TotalCPUUsedPct %" }

#Memory Usage
$ComputerMemory =  Get-WmiObject -Class WIN32_OperatingSystem | select TotalVisibleMemorySize,FreePhysicalMemory
$ComputerMemoryInGB = ([Math]::Round(($ComputerMemory.TotalVisibleMemorySize/1024/1024)))
$ComputerFreeMemoryInGB = ([Math]::Round(($ComputerMemory.FreePhysicalMemory/1024/1024)))
$MemoryPercent = ((($ComputerMemory.TotalVisibleMemorySize - $ComputerMemory.FreePhysicalMemory)*100)/ $ComputerMemory.TotalVisibleMemorySize)
$MemoryPercent = ([Math]::Round($MemoryPercent, 2))

$MemReport = "<h3>Memory Usage</h3>
    <table>
        <tr><th>Total</th><th>Free</th><th>%Used</th></tr>
        <tr><td>$ComputerMemoryInGB GB</td><td>$ComputerFreeMemoryInGB GB</td><td>$MemoryPercent %</td></tr>
        </table>"

if($MemoryPercent -ge 80 -and $MemoryPercent -lt 85){
    $MemReport = $MemReport -replace "<td>$MemoryPercent %","<td $Amber>$MemoryPercent %" }
    elseif($MemoryPercent -ge 85){
	$MemReport = $MemReport -replace "<td>$MemoryPercent","<td $Red>$MemoryPercent" }
    else{
    $MemReport = $MemReport -replace "<td>$MemoryPercent %","<td $Green>$MemoryPercent %" }

#SQL Instance Memory Usage
$ProcessID = Get-WmiObject -Class Win32_Service | where name -eq ($SQLService) | select -ExpandProperty ProcessId 
$ConfigMaxMemory = Invoke-Sqlcmd -ServerInstance $SQLInstance -Query "SELECT name, value FROM sys.configurations WHERE name like '%max server memory%'" | select -ExpandProperty value
$ConfigMaxMemory = ([Math]::Round(($ConfigMaxMemory/1024),2))
$SQLMem = Get-WmiObject WIN32_PROCESS | where-object -property ProcessId -eq $ProcessID | Select-Object processname, @{Name="MemUsageMB";Expression={[math]::round($_.ws / 1mb)}} 
$SQLMemUsedPct=0
    if($SQLMem.MemUsageMB -gt 0){
		$SQLMemUsage = $SQLMem.MemUsageMB
        $SQLMemUsage = ([Math]::Round(($SQLMemUsage/1024),2))
        $SQLMemUsedPct = ($SQLMemUsage/$ConfigMaxMemory)*100
        $SQLMemUsedPct = ([Math]::Round($SQLMemUsedPct, 2))}
    else{
		$SQLMemUsage = 0 }
$SQLMemReport = "<h3>$SQLInstance</h3>
    <table>
        <tr><th>MAXMem</th><th>UsedMem</th><th>%Used</th></tr>
        <tr><td>$ConfigMaxMemory GB</td><td>$SQLMemUsage GB</td><td>$SQLMemUsedPct %</td></tr>
        </table>"
if($SQLMemUsedPct -ge 80 -and $SQLMemUsedPct -lt 85){
    $SQLMemReport = $SQLMemReport -replace "<td>$SQLMemUsedPct %","<td $Amber>$SQLMemUsedPct %" }
    elseif($SQLMemUsedPct -ge 85){
	$SQLMemReport = $SQLMemReport -replace "<td>$SQLMemUsedPct %","<td $Red>$SQLMemUsedPct %" }
    else{
    $SQLMemReport = $SQLMemReport -replace "<td>$SQLMemUsedPct %","<td $Green>$SQLMemUsedPct %" }

#CLUSTER Information (if any)
    $ModuleCheck = Get-Module -ListAvailable | where name -like 'FailoverClusters' | Select-Object -ExpandProperty Name -ErrorAction SilentlyContinue
    if($ModuleCheck -eq 'FailoverClusters') {
    $IsClustered = 1
    $Cluster = Get-Cluster 
    $ClusterName = $Cluster.Name
    $ClusterInfo = Resolve-DnsName -Name $ClusterName -ErrorAction SilentlyContinue
    #$ClusterInfo
    if(($ClusterInfo.Name) -ne '') {
    $ClusterInfo = $ClusterInfo | select Name,IPAddress | ConvertTo-Html -PreContent "<h3>Cluster Info</h3>" }
    else {
    $ClusterInfo = "<h3>Cluster Info</h3>
		<table><tr><th>Name</th><th>IPAddress</th></tr><tr><td>$ClusterName</td><td style=color:red;>DNS does not exist</td></tr>
		</table>" }
	$ClusterNodes = Get-ClusterNode | select cluster,Name,State | ConvertTo-Html -PreContent "<h3>Cluster Nodes</h3>"
    $ColorTagTable.Keys | foreach { $ClusterNodes = $ClusterNodes -replace ">$_<",($ColorTagTable.$_) }
    $ClusterRoles = Get-ClusterGroup | where Name -like 'AG*'| select Name,OwnerNode,State | ConvertTo-Html -PreContent "<h3>Cluster Roles</h3>"
    $ColorTagTable.Keys | foreach { $ClusterRoles = $ClusterRoles -replace ">$_<",($ColorTagTable.$_) }
    $ClusterResrc = Get-ClusterGroup |where Name -like 'AG*'| Get-ClusterResource | select Name,ResourceType,OwnerGroup,State | ConvertTo-Html -PreContent "<h3>Cluster Resources</h3>"
    $ColorTagTable.Keys | foreach { $ClusterResrc = $ClusterResrc -replace ">$_<",($ColorTagTable.$_) }}

#Disk Report
$DriveFreePct = Get-WmiObject win32_logicaldisk | where DriveType -eq 3 | Select-Object @{Name="FreePct";Expression={[math]::round(($_.freespace/$_.size*100),2)}}
$DriveDetails = Get-WmiObject win32_logicaldisk | where DriveType -eq 3 | 
Select-Object DeviceID,VolumeName,
@{Name="Size(GB)";Expression={[math]::round($_.Size / 1gb,2)}},
@{Name="Used(GB)";Expression={[math]::round(($_.Size - $_.FreeSpace) / 1gb,2)}},
@{Name="%Free";Expression={"{0:n2}" -f ($_.freespace/$_.size*100)}} | ConvertTo-Html -PreContent "<h3>Drive Details</h3>" 
    foreach($Drive in $DriveFreePct){
        if($Drive.FreePct -ge 25.00 -and $Drive.FreePct -lt 30.00){
            $Val = $Drive.FreePct
			$DriveDetails = $DriveDetails -replace "<td>$Val<","<td $Amber>$Val %<"}
		elseif($Drive.FreePct -lt 25.00){
			$Val = $Drive.FreePct
			$DriveDetails = $DriveDetails -replace "<td>$Val<","<td $Red>$Val %<"}
        else{
            $Val = $Drive.FreePct
			$DriveDetails = $DriveDetails -replace "<td>$Val<","<td $Green>$Val %<"}}

#Services Report
$ServicesReport = Get-Service -Name $ServicesList
$ServicesInfo = $ServicesReport | ConvertTo-Html -Property Name, DisplayName,Status -PreContent "<h3>Services Information</h3>"
$ServicesInfo = $ServicesInfo -replace "<td>Running<","<td $Green>Running<"
$ServicesInfo = $ServicesInfo -replace "<td>Stopped<","<td $Red>Stopped<"

#App pool Report
$AppPoolInfo = Get-IISAppPool -Name $AppPoolList
$AppPoolReport = $AppPoolInfo | ConvertTo-Html -Property Name, State -PreContent "<h3>AppPool Information</h3>"
$AppPoolReport = $AppPoolReport -replace "<td>Started<","<td $Green>Started<"
$AppPoolReport= $AppPoolReport -replace "<td>Stopped<","<td $Red>Stopped<"

#SQL Agent Job information Report - Instance
$SQLJobInfo = Get-SqlAgentJob -ServerInstance $SQLInstance | 
Where-Object {$_.LastRunDate -ge ((Get-Date).AddDays(-7))} | ForEach-Object {
    $h = Get-SqlAgentJobHistory -ServerInstance $SQLInstance -JobName $_.Name
    [PSCustomObject]@{
        Name           = $_.Name
        IsEnabled      = $_.IsEnabled
        LastRunDate    = $_.LastRunDate
        LastRunOutcome = $_.LastRunOutcome
        NextRunDate    = $_.NextRunDate
        LastRunStep    = $h[0].StepName
    }
} 
$SQLBackupReport = $SQLJobInfo| ConvertTo-Html -Property Name,LastRunDate,LastRunOutcome,NextRunDate -PreContent "<h3>SqlAgentJobHistory - $SQLInstance </h3>" 
$SQLBackupReport = $SQLBackupReport -replace "<td>Succeeded<","<td $Green>Succeeded<"
$SQLBackupReport = $SQLBackupReport -replace "<td>Failed<","<td $Red>Failed<"

# SQL Scripts Report
$SQLScriptReport = foreach ($filename in get-childitem -path $SQLscriptsPath -filter "*.sql")
{
invoke-sqlcmd –ServerInstance $SQLInstance -InputFile $filename.fullname |Select * -ExcludeProperty RowError, RowState, Table, ItemArray, HasErrors|
ConvertTo-Html -Head $header -PreContent "<h3>$filename</h3>"
}
$ColorTagTable.Keys | foreach { $SQLScriptReport = $SQLScriptReport -replace ">$_<",($ColorTagTable.$_) }

#SQL Error Report
$SQLErrorlogs = Get-SqlErrorLog -ServerInstance $SQLInstance -Since Yesterday | 
Where-object {$_.Text  -match  '( Error | Fail | failure | IO requests taking longer|is full)'  -and  $_.Text  -notmatch  '(Login failed for user |without errors|found 0 errors)' -and $_.Source -ne 'Logon'}
$ErrorCount = $SQLErrorlogs | Measure-Object | Select-Object -ExpandProperty Count
    if($ErrorCount -gt 0 -and $ErrorCount -le 20) {
	$SQLErrorlogDetails = $SQLErrorlogs | ConvertTo-Html -Property Date,Source,Text -PreContent "<h3>Recent Errors from SQL Error Log - $SQLInstance</h3>"  
	$SQLErrorlogDetails = $SQLErrorlogDetails -replace"<tr>","<tr style=color:Red;>"}
	else {
        $SQLErrorlogDetails = "<h3>Recent Errors from SQL Error Log - $SQLInstance </h3><h4 style=color:green;>No errors since yesterday</h4>"}

#Final Report
 $HTMLReport = "<head>
    <style>
		.MainDiv {width:Auto;height:auto;background-color:ghostwhite;}
		.div1 {width:100%;height:55px;background-color:MidnightBlue;}
		.div2 {width:100%;height:90px;}
        .div3 {width:100%;height:auto;}
    </style>
    </head>
    <table class=MainDiv><tr><td style=padding:0px;>
		<table class=div1><tr><td><h2 Style=font-family:tahoma,arial,sans-serif;margin-top:5px;margin-bottom:5px;color:white;>$ReportHeader</h2><p>$TimeZone</p></td></tr>
		</table>
        <table class=div2 style=border:0px><tr style=border:0px><td style=border:0px;>$ServerRecycleTime</td><td style=border:0px;>$uptime</td><td style=border:0px;>$AzureBackupInfo</td></tr>
		</table>
		<table class=div3 style=border:0px;><tr style=border:0px;><td style=border:0px;>$InstanceRecycleTime</td><td style=border:0px;>$CPUReport</td><td style=border:0px;>$MemReport</td><td style=border:0px;>$SQLMemReport</td></tr>
		</table>"
        if($IsClustered -eq 1){
            $HTMLReport = $HTMLReport +
            "<table class=div2 style=border:0px><tr style=border:0px><td style=border:0px;>$ClusterInfo</td><td style=border:0px;>$ClusterNodes</td><td style=border:0px;>$ClusterRoles</td></tr>
             </table>"}
        $HTMLReport = $HTMLReport + 
        "<table class=div3 style=border:0px;><tr style=border:0px;><td style=border:0px;>$ClusterResrc</td></tr>
		</table>
        <table class=div3 style=border:0px;><tr style=border:0px;><td style=border:0px;>$DriveDetails</td><td style=border:0px;>$ServicesInfo</td></tr>
		</table>
        <table style=border:0px><tr style=border:0px><td style=border:0px>$AppPoolReport</td></tr>
		</table>
		<table style=border:0px><tr style=border:0px><td style=border:0px>$SQLBackupReport</td></tr>
		</table>
        <table style=border:0px><tr style=border:0px><td style=border:0px>$SQLScriptReport</td></tr>
		</table>
        <table style=border:0px><tr style=border:0px;width:850px><td style=border:0px;width:850px>$SQLErrorlogDetails</td></tr>
		</table>
		</td></tr></table>" #>
    $line = '*'*118
    $HTMLReport+$line | out-file -FilePath G:\PS_Health_Check_Report\ServerReport.html