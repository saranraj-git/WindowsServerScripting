$servers = gc .\servers.txt

foreach($server in $servers)
{
write "working in $server"
ac "C:\users\ze4kmy9\Desktop\IISoutput.txt" -value "============================" 
ac "C:\users\ze4kmy9\Desktop\IISoutput.txt" -value "Deploying IIS on $server" 
$deployIIS = Invoke-Command -ComputerName $server -sc{
$features = "Web-Server","Web-WebServer","Web-Common-Http","Web-Static-Content","Web-Default-Doc","Web-Dir-Browsing","Web-Http-Errors","Web-Http-Redirect","Web-App-Dev","Web-Asp-Net","Web-Net-Ext","Web-ASP","Web-CGI","Web-ISAPI-Ext","Web-ISAPI-Filter","Web-Health","Web-Http-Logging","Web-Log-Libraries","Web-Request-Monitor","Web-Http-Tracing","Web-Security","Web-Basic-Auth","Web-Windows-Auth","Web-Digest-Auth","Web-Client-Auth","Web-Cert-Auth","Web-Url-Auth","Web-Filtering","Web-IP-Security","Web-Performance","Web-Stat-Compression","Web-Dyn-Compression","Web-Mgmt-Tools","Web-Mgmt-Console","Web-Scripting-Tools","Web-Mgmt-Service","Web-Mgmt-Compat","Web-Metabase","Web-WMI","Web-Lgcy-Scripting","Web-Lgcy-Mgmt-Console"
foreach($fea in $features)
{
$reso = Add-WindowsFeature $fea 
$out = $reso | select -ExpandProperty "Success"
write-output "$fea - $out"
}

}
ac "C:\users\ze4kmy9\Desktop\IISoutput.txt" -value "============================" 
ac "C:\users\ze4kmy9\Desktop\IISoutput.txt" -value "$deployIIS" 
}


==============================================================
