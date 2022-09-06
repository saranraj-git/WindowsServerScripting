function Get-InstalledProducts($ser){

$ser_products = @()
$UninstallKey="SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"
$reg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine',$ser)
$regkey = $reg.OpenSubKey($UninstallKey)
$subkeys=$regkey.GetSubKeyNames()

    foreach($key in $subkeys){
        $thisKey=$UninstallKey+"\\"+$key
        $thisSubKey=$reg.OpenSubKey($thisKey)

        $obj = New-Object PSObject
        $obj | Add-Member -MemberType NoteProperty -Name "ComputerName" -Value $ser
        $obj | Add-Member -MemberType NoteProperty -Name "DisplayName" -Value $($thisSubKey.GetValue("DisplayName"))
        $obj | Add-Member -MemberType NoteProperty -Name "DisplayVersion" -Value $($thisSubKey.GetValue("DisplayVersion")) 
        $obj | Add-Member -MemberType NoteProperty -Name "Publisher" -Value $($thisSubKey.GetValue("Publisher"))
        $obj | Add-Member -MemberType NoteProperty -Name "InstallDate" -Value $($thisSubKey.GetValue("InstallDate"))
        $ser_products += $obj
    }

return $ser_products
}

$outpath = ".\InstalledSoftwares.csv"

$csvcontents = @()  

$servers = get-content .\servers.txt

foreach($ser in $servers)
{
    write-output "Working on $ser `n"
    $InstalledSoft = Get-InstalledProducts $ser
    $csvcontents += $InstalledSoft
}

$csvcontents | Export-Csv -Path $outpath

