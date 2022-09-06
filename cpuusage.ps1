﻿function get-sysinfo($ser)
{
#RAM:
$ComputerMemory =  Get-WmiObject -class WIN32_OperatingSystem -computername "$ser" | select TotalVisibleMemorySize,FreePhysicalMemory

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
#Server UP time
write $uptime
}
get-uptime "localhost"




1. Hostname to IPinfo
2. CPU/Memory/drive 
3. Restart time 
4. uptime
5. All software report
6. search specific software report

1. IIS version


https://docs.microsoft.com/en-us/windows/win32/msi/uninstall-registry-key
HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Uninstall