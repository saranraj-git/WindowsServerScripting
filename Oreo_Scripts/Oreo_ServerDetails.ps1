
function hostname_to_ip($ser){
    $ipadd = [system.net.dns]::resolve($ser) | select -ExpandProperty Addresslist | select -ExpandProperty IPAddressToString
    return "$ipadd"
}


function get-sysinfo($ser){
    #RAM:
    $ComputerMemory =  Get-WmiObject -class WIN32_OperatingSystem -computername "$ser" | select TotalVisibleMemorySize,FreePhysicalMemory
    $ComputerMemoryInGB = ([Math]::Round(($ComputerMemory.TotalVisibleMemorySize/1024/1024)))
    $ComputerFreeMemoryInGB = ([Math]::Round(($ComputerMemory.FreePhysicalMemory/1024/1024)))
    $MemoryPercent = ((($ComputerMemory.TotalVisibleMemorySize - $ComputerMemory.FreePhysicalMemory)*100)/ $ComputerMemory.TotalVisibleMemorySize)
    $MemoryPercent = ([Math]::Round($MemoryPercent, 2))

    #CPU:
    $cor=$null;
    $cores = Get-WmiObject -class win32_processor -computername $ser | select -ExpandProperty numberofcores
    foreach($cr in $cores){$cor = $cor + $cr};


    #CPU Usage
    $cpuusage = Get-WmiObject Win32_Processor -computername $ser | Measure-Object  -Property LoadPercentage -Average | Select-Object -ExpandProperty Average

    return $ComputerMemoryInGB,$MemoryPercent,$cor,$cpuusage
}

function get-uptime($ser){
    #Server UP time
    $TimeNow =  get-date -F 'MM/dd/yyyy HH:mm:ss'
    #$bootuptime = (Get-CimInstance -ClassName Win32_OperatingSystem -ComputerName $ser).LastBootUpTime
    $bootuptime = Get-WmiObject win32_operatingsystem -ComputerName $ser | select @{LABEL='LastBootUpTime';EXPRESSION={$_.ConverttoDateTime($_.lastbootuptime)}} | select LastBootUpTime
    $TimeDiff = New-TimeSpan -Start $bootuptime -End $TimeNow
    #$uptime = $TimeNow - $bootuptime 
    write-output $TimeDiff
}



$outpath = ".\output.csv"

$csvcontents = @()  

$servers = get-content -Path ".\servers.txt"

foreach($ser in $servers)
{
    write-output-output "Working on $ser `n"
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




