﻿

function hostname_to_ip($ser){
    $ipadd = [system.net.dns]::resolve($ser) | select -ExpandProperty Addresslist | select -ExpandProperty IPAddressToString
    return "$ipadd"
}


function get-sysinfo($ser){
    #RAM:
    $ComputerMemory =  Get-WmiObject -class WIN32_OperatingSystem -computername "$ser" | select TotalVisibleMemorySize,FreePhysicalMemory

    #CPU:
    $cor=$null;
    $cores = Get-WmiObject -class win32_processor -computername $ser | select -ExpandProperty numberofcores
    foreach($cr in $cores){$cor = $cor + $cr};


    #CPU Usage
    $cpuusage = gwmi Win32_Processor -computername $ser | Measure-Object  -Property LoadPercentage -Average | Select-Object -ExpandProperty Average

    return $ComputerMemoryInGB,$MemoryPercent,$cor,$cpuusage
}

function get-uptime($ser){
    #Server UP time
    write $TimeDiff
}



$outpath = ".\output.csv"

$csvcontents = @()  

$servers = gc .\servers.txt

foreach($ser in $servers)
{
    write "Working on $ser `n"
    $ipadd1 = hostname_to_ip $ser
    $totalmem, $memper, $totalcore, $cpuusage = get-sysinfo $ser


    $obj1 = New-Object PSObject
    $obj1 | Add-Member -MemberType NoteProperty -Name "ComputerName" -Value "$ser"
    $obj1 | Add-Member -MemberType Noteproperty -Name "IPAddress" -value "$ipadd1"
    $obj1 | Add-Member -MemberType Noteproperty -Name "TotalMemory(GB)" -value "$totalmem"
    $obj1 | Add-Member -MemberType Noteproperty -Name "MemoryFreePercent" -value "$memper"
    $obj1 | Add-Member -MemberType Noteproperty -Name "TotalCpuCore" -value "$totalcore"
    $obj1 | Add-Member -MemberType Noteproperty -Name "Avg_CPU_Usage" -value "$cpuusage"
    $csvcontents += $obj1
}

$csvcontents | Export-Csv -Path $outpath



