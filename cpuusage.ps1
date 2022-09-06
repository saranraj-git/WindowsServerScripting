function get-sysinfo($ser)
{
#RAM:
$ComputerMemory =  Get-WmiObject -class WIN32_OperatingSystem -computername "$ser" | select TotalVisibleMemorySize,FreePhysicalMemory$ComputerMemoryInGB = ([Math]::Round(($ComputerMemory.TotalVisibleMemorySize/1024/1024)))$ComputerFreeMemoryInGB = ([Math]::Round(($ComputerMemory.FreePhysicalMemory/1024/1024)))$MemoryPercent = ((($ComputerMemory.TotalVisibleMemorySize - $ComputerMemory.FreePhysicalMemory)*100)/ $ComputerMemory.TotalVisibleMemorySize)$MemoryPercent = ([Math]::Round($MemoryPercent, 2))

write "RAM - $ComputerMemoryInGB GB `n"

#RAM Usage


#CPU:
$cor=$null;
$cores = Get-WmiObject -class win32_processor -computername $ser | select -ExpandProperty numberofcores
foreach($cr in $cores){$cor = $cor + $cr};
write "CPU CORE : $cor `n"

#CPU Usage
$cpuusage = gwmi Win32_Processor -computername $ser | Measure-Object  -Property LoadPercentage -Average | Select-Object -ExpandProperty Average
write "CPU usage avg percent - $cpuusage"

}
get-sysinfo "localhost"


function get-uptime($ser){
#Server UP time$TimeNow = get-date$bootuptime = (Get-CimInstance -ClassName Win32_OperatingSystem -ComputerName $ser).LastBootUpTime$uptime = $TimeNow - $bootuptime 
write $uptime
}
get-uptime "localhost"




1. Hostname to IPinfo
2. CPU/Memory/drive 
3. Restart time 
4. uptime
5. All software report
6. search specific software report

1. IIS version2. SQL version1. Port connection1. certificateSvc monitoringAzure backup Status


https://docs.microsoft.com/en-us/windows/win32/msi/uninstall-registry-key
HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Uninstall
