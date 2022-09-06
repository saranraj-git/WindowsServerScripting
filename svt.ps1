    [reflection.assembly]::loadwithpartialname("System.Windows.Forms") | Out-Null
    [reflection.assembly]::loadwithpartialname("System.Drawing") | Out-Null

    $pbrTest = New-Object System.Windows.Forms.ProgressBar
    $pbrTest.Maximum = 100
    $pbrTest.Minimum = 0
    $pbrTest.Location = new-object System.Drawing.Size(300,520)
    $pbrTest.size = new-object System.Drawing.Size(290,50)
    $i = 0
    $listBox1 = New-Object System.Windows.Forms.TextBox
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 310
    $System_Drawing_Size.Height = 350
    $listBox1.Size = $System_Drawing_Size
    $listBox1.DataBindings.DefaultDataSourceUpdateMode = 0
    $listBox1.Name = "listBox1"
    $System_Drawing_Point = New-Object System.Drawing.Point
    $System_Drawing_Point.X = 9
    $System_Drawing_Point.Y = 140
    $listBox1.Location = $System_Drawing_Point
    $listBox1.TabIndex = 8
    $listBox1.MultiLine = $True
    $listBox1.ScrollBars = "Vertical"
#======================== writing input for log file

#======================== Converting Excel input sheet to CSV file for powershell processing
    function xlstocsv
    {
    $d = get-date
    $listBox1.appendtext("********************************************************************** `r`n")
    $listBox1.appendtext("                         Script Execution Begins `r`n")
    $listBox1.appendtext("********************************************************************** `r`n")
    $listBox1.appendtext(" `r`n")
    $listBox1.appendtext("$d : Converting xls sheet to CSV sheet `r `n")
    $listBox1.appendtext(" `r`n")
    $tmp = Resolve-Path "serverlist1.xls"
    $objExcel = New-Object -ComObject Excel.Application
        if ((test-path $tmp) -and ($tmp -match ".xl\w*$"))
            {
            $path = (resolve-path -Path $tmp).path
            $savePath = $tmp -replace ".xl\w*$",".csv"
            $i += 2
            $pbrTest.Value = $i
            if(Test-path $savePath)
                {
                Remove-Item -Path $savePath -Force | Out-Null
                }
            $objworkbook=$objExcel.Workbooks.Open($tmp)
            $objworkbook.SaveAs($savePath,6)
            $objworkbook.Close($false)
            $i += 2
            $pbrTest.Value = $i
            }
    $i = 10
    $pbrTest.Value = $i
    $d = get-date
    $listBox1.appendtext("$d : CSV Conversion Successful `r `n")
    $listBox1.appendtext(" `r`n")
    }
#======================== XLSTOCSV ends here

#======================== Creating Output CSV file at C:\Script_Output\Results.csv
    Function CSVcreation
    {

    $d = get-date
    $listBox1.appendtext("$d : Creating Output CSV `r `n")
    $listBox1.appendtext(" `r`n")
    $erroractionpreference = "SilentlyContinue"
    $Excelfile = New-Object -comobject Excel.Application
    $Excelfile.visible = $false
    $Excelfile.DisplayAlerts = $false
    $Excelfile.AskToUpdateLinks = $false
    $Excelfile.AlertBeforeOverwriting = $false
    $wb = $ExcelFile.Workbooks.Add()
    #===============================================
    #===============================================
  <#  $Specific_Reg_Key = $ExcelFile.Worksheets.Add()
    $Specific_Reg_Key.name = "Specific_Reg_Key"
    $Specific_Reg_Key.activate()
    $cellc6 = $wb.Worksheets.Item(1)

    $cellc6.Cells.Item(1,1) = "ServerName"
    $cellc6.Cells.Item(1,2) = "Path"
    $cellc6.Cells.Item(1,3) = "Value" #>
    #===============================================
    #===============================================
    $i = 20
    $pbrTest.Value = $i
    $asd = 1
    $d = get-date
    $listBox1.appendtext("$d : Creating Directory C:\script_output `r `n")
    $listBox1.appendtext(" `r`n")

    md "C:\script_output"
    #remove-item "C:\script_output\*.*" -Force
    $strpath_un = "C:\script_output\results.csv"
    
    $files = "C:\script_output\results.csv"
    
            $Wb.Worksheets.Item("Sheet2").Delete()
            $Wb.Worksheets.Item("Sheet3").Delete()
            #$Wb.Worksheets.Item("Sheet1").Delete()
            
    $wb = $Excelfile.WorkBooks.close($strPath_un)
    $wb.Saveas($files)
    
    $Excelfile.quit()
    $d = get-date
    $listBox1.appendtext("$d : Output CSV created `r `n")
    $listBox1.appendtext(" `r`n")
    }

#=========================== CSVCreation Ends here

#=========================== Temp CSV (serverlist.csv) deletion
    function csvdeletion
    {
    $tmp = Resolve-Path ".\serverlist1.csv"
    remove-item $tmp -force
    copy-item -path "C:\script_output\results.csv" -destination (Resolve-Path ".\" | select -expand Path)
    remove-item "C:\script_output\results.csv" -force

    $d = get-date
    $listBox1.appendtext("$d : Deleting the Temp CSV file `r `n")
    $listBox1.appendtext(" `r`n")
    $listBox1.appendtext("$d : Script execution Completed `r`n")
    $listBox1.appendtext(" `r`n")
    $lo = Resolve-Path ".\" | select -expand path
    $listBox1.appendtext("Results stored in $lo `r`n")
    $listBox1.appendtext(" `r`n")
    $i = 100
    $pbrTest.Value = $i
    $listBox1.appendtext("********************************************************************** `r`n")
    $listBox1.appendtext("                       Script execution Completed `r`n")
    $listBox1.appendtext("********************************************************************** `r`n")
    $listBox1.appendtext(" `r`n")
    }

#============================ CSV deletion ends here =============================

<# ==========================  Module 1 : Product Validation ============================

This Module helps in finding the installed product details on the remote server 

-> It will go ahead and search the following registry path in READ ONLY mode

   1. HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall
   2. HKEY_LOCAL_MACHINE\WoW6432Node\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall
             
#>
function Specific_uninstallkey
{

[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
$objForm = New-Object System.Windows.Forms.Form
$objForm.Text = "Data Entry Form"
$objForm.Size = New-Object System.Drawing.Size(500,200)
$objForm.StartPosition = "CenterScreen"

$objLabel = New-Object System.Windows.Forms.Label
$objLabel.Location = New-Object System.Drawing.Size(10,20)
$objLabel.Size = New-Object System.Drawing.Size(280,20)
$objLabel.Text = "Please enter the Software name"
$objForm.Controls.Add($objLabel)

$info = New-Object System.Windows.Forms.TextBox
$info.Location = New-Object System.Drawing.Size(10,50)
$info.Size = New-Object System.Drawing.Size(460,25)
$objForm.Controls.Add($info)

$OKButton = New-Object System.Windows.Forms.Button
$OKButton.Location = New-Object System.Drawing.Size(10,80)
$OKButton.Size = New-Object System.Drawing.Size(75,23)
$OKButton.Text = "Submit"
$objForm.Controls.Add($OKButton)

$chumma = {
#function begins here
        $softname = $info.text
        $d = get-date
        $array = @()
        $array2 = @()
        $listBox1.appendtext("$d : Started finding specific software Module `r `n")
        $listBox1.appendtext(" `r`n")
        ac .\log.txt -Value " "
        add-content .\log.txt -Value "******** Product Validation Module begins *********"
        ac .\log.txt -Value " "
        $d = get-date
        $listBox1.appendtext("$d : Please enter the software name `r `n")
        $listBox1.appendtext(" `r`n")
        
        ac .\log.txt -Value "User entered product name : $softname "
        ac .\log.txt -Value "Started searching for $softname "
        $a = import-csv ".\serverlist1.csv"

        foreach($pc in $a)
        { 
        $k = $pc.server
        ac .\log.txt -Value " "
        ac .\log.txt -Value "Working on Computer : $k "
        $d = get-date
        $listBox1.appendtext("$d : Working on $k `r `n")
        $listBox1.appendtext(" `r`n")
            if ( Test-Connection -ComputerName $pc.server -Count 1 -erroraction silentlycontinue )
             {
                    ac .\log.txt -Value "$k - Server is in network "
                    
                    $reg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine',$pc.server)
                    if($reg -ne $null)
                    {
                        
                        $flg = 0
                        
                        ac .\log.txt -Value "$k - Checking in HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Uninstall"
                        $reg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine',$pc.server)
                        $regKey= $reg.OpenSubKey("SOFTWARE\Microsoft\Windows NT\CurrentVersion") 
                        $UninstallKey="SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Uninstall"
                        try 
                        { $regkey = $reg.OpenSubKey($UninstallKey)}
                         catch 
                        {
                            $computer = $pc.server
                            $accessDenied = "Access Denied"
                            $obj1 = New-Object PSObject
                            $obj1 | Add-Member -MemberType NoteProperty -Name "ComputerName" -Value $computer
                            $obj1 | Add-Member -MemberType Noteproperty -Name "DisplayName" -value $accessDenied
                            $obj1 | Add-Member -MemberType Noteproperty -Name "DisplayVersion" -value $accessDenied
                            $obj1 | Add-Member -MemberType Noteproperty -Name "DisplayLocation" -value $accessDenied
                            $obj1 | Add-Member -MemberType Noteproperty -Name "Publisher" -value $accessDenied
                            $array += $obj1
                            ac .\log.txt -Value "$computer - Registry access denied"
                            $flg = 2
                            $red = 1
                            
                        }
                        $subkeys=$regkey.GetSubKeyNames()
                        foreach($key in $subkeys)
                        {
                            $thisKey=$UninstallKey+"\\"+$key
                            $thisSubKey=$reg.OpenSubKey($thisKey)
                                if ( $thisSubKey.GetValue("DisplayName") -like "*$softname*" )
                                {
                                $computer = $pc.server
                                $obj = New-Object PSObject
                                $obj | Add-Member -MemberType NoteProperty -Name "ComputerName" -Value $computer
                                $obj | Add-Member -MemberType NoteProperty -Name "DisplayName" -Value $($thisSubKey.GetValue("DisplayName"))
                                $obj | Add-Member -MemberType NoteProperty -Name "DisplayVersion" -Value $($thisSubKey.GetValue("DisplayVersion"))
                                $obj | Add-Member -MemberType NoteProperty -Name "InstallLocation" -Value $($thisSubKey.GetValue("InstallLocation"))
                                $obj | Add-Member -MemberType NoteProperty -Name "Publisher" -Value $($thisSubKey.GetValue("Publisher"))
                                $obj | Add-Member -MemberType NoteProperty -Name "InstallDate" -Value $($thisSubKey.GetValue("InstallDate"))
                                $array += $obj
                                ac .\log.txt -Value "$k - $softname found"
                                $flg = 1
                                }#if this sub key ends here
                            }  # for each ends here
                        
                       

                         #===============
                         $UninstallKey="SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall"
                         ac .\log.txt -Value "$k - Checking in HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\Microsoft\Windows NT\CurrentVersion\Uninstall"
                         try 
                        { 
                        $computer = $pc.server
                        $regkey = $reg.OpenSubKey($UninstallKey)
                        $subkeys=$regkey.GetSubKeyNames()
                        foreach($key in $subkeys)
                            {
                            $thisKey=$UninstallKey+"\\"+$key
                            $thisSubKey=$reg.OpenSubKey($thisKey)
                                if ( $thisSubKey.GetValue("DisplayName") -like "*$softname*" )
                                {
                                $obj = New-Object PSObject
                                $obj | Add-Member -MemberType NoteProperty -Name "ComputerName" -Value $computer
                                $obj | Add-Member -MemberType NoteProperty -Name "DisplayName" -Value $($thisSubKey.GetValue("DisplayName"))
                                $obj | Add-Member -MemberType NoteProperty -Name "DisplayVersion" -Value $($thisSubKey.GetValue("DisplayVersion"))
                                $obj | Add-Member -MemberType NoteProperty -Name "InstallLocation" -Value $($thisSubKey.GetValue("InstallLocation"))
                                $obj | Add-Member -MemberType NoteProperty -Name "Publisher" -Value $($thisSubKey.GetValue("Publisher"))
                                $obj | Add-Member -MemberType NoteProperty -Name "InstallDate" -Value $($thisSubKey.GetValue("InstallDate"))
                                $array += $obj
                                ac .\log.txt -Value "$k - $softname found"
                                $flg = 1
                                }#if this sub key ends here
                            }  # for each ends here
                        } 
                        catch 
                        {
                            if ($red -ne 1)
                            {
                            $computer = $pc.server
                            $accessDenied = "Access Denied"
                            $obj1 = New-Object PSObject
                            $obj1 | Add-Member -MemberType NoteProperty -Name "ComputerName" -Value $computer
                            $obj1 | Add-Member -MemberType Noteproperty -Name "DisplayName" -value $accessDenied
                            $obj1 | Add-Member -MemberType Noteproperty -Name "DisplayVersion" -value $accessDenied
                            $obj1 | Add-Member -MemberType Noteproperty -Name "DisplayLocation" -value $accessDenied
                            $obj1 | Add-Member -MemberType Noteproperty -Name "Publisher" -value $accessDenied
                            $array += $obj1
                            ac .\log.txt -Value "$computer - Registry access denied"
                            $flg = 1
                            }
                        }
                    

                        if ($flg -eq 0)
                        {
                                $computer = $pc.server
                                $prodnotfound = "Product not found"
                                $obj1 = New-Object PSObject
                                $obj1 | Add-Member -MemberType NoteProperty -Name "ComputerName" -Value $computer
                                $obj1 | Add-Member -MemberType Noteproperty -Name "DisplayName" -value $prodnotfound
                                $obj1 | Add-Member -MemberType Noteproperty -Name "DisplayVersion" -value $prodnotfound
                                $obj1 | Add-Member -MemberType Noteproperty -Name "DisplayLocation" -value $prodnotfound
                                $obj1 | Add-Member -MemberType Noteproperty -Name "Publisher" -value $prodnotfound
                                $array += $obj1
                                ac .\log.txt -Value "$computer - $softname not found"
                    } 
                    }  #
             }  #test connection ends here
            else
             {
                        $obj = New-Object PSObject
                        $obj | add-member -membertype noteproperty -name "ComputerName" -value $pc.server
                        $obj | add-member -membertype noteproperty -name "Displayname" -value "Out of network"
                        $array += $obj
                        ac .\log.txt -Value "$k - Ping Check Failed"
                     }
        }
    $array2 = $array | Where-Object { $_.DisplayName } | select ComputerName, DisplayName, DisplayVersion, Publisher, InstallDate
    $array2 | export-csv ".\Specific_Installed_Products.csv"
    $array2 | ogv
    #copy and conversion begins

    #4. perform copy from sec_update.xlsx to results.xlsx (in Security_update Tab )

    $file1 = resolve-path ".\Specific_Installed_Products.csv" # source's fullpath
    $file2 = 'C:\script_output\results.csv' # destination's fullpath
    $xl = new-object -com excel.application
    $xl.displayAlerts = $false   # don't prompt the user
    $xl.AlertBeforeOverwriting = $false
    $wb2 = $xl.workbooks.open($file1) # open source, if u need readonly $wb2 = $xl.workbooks.open($file1, $null, $true)
    $wb1 = $xl.workbooks.open($file2) # open target
    $sh1_wb1 = $wb1.sheets.item("Sheet1") # second sheet in destination workbook
    $sheetToCopy = $wb2.sheets.item(1) # source sheet to copy
    $sheetToCopy.copy($sh1_wb1) # copy source sheet to destination workbook
    $wb2.close($true) # close source workbook and saving
    $wb1.close($true) # close and save destination workbook
    $xl.quit()


        $strpath = "C:\script_output\results.csv"
        $erroractionpreference = "SilentlyContinue"
        $Exf = New-Object -comobject Excel.Application
        $Exf.visible = $false
        $Exf.DisplayAlerts = $false
        $Exf.AskToUpdateLinks = $false
        $Exf.AlertBeforeOverwriting = $false

        $wb1 = $exf.WorkBooks.Open($strPath)
        $ws1 = $wb1.worksheets.item("Specific_Installed_Products")

        $ws1.activate()
        $ws1.cells.item(1,1) =""
    $dd = $ws1.UsedRange
    $dd.EntireColumn.AutoFit()
    $dd.Interior.ColorIndex = 19
    $dd.Font.ColorIndex = 11
    $dd.EntireColumn.AutoFit()
    $aaAutomatic=-4105
    $aaBottom = -4107
    $aaCenter = -4108
    $aaContext = -5002
    $aaContinuous=1
    $aaDiagonalDown=5
    $aaDiagonalUp=6
    $aaEdgeBottom=9
    $aaEdgeLeft=7
    $aaEdgeRight=10
    $aaEdgeTop=8
    $aaInsideHorizontal=12
    $aaInsideVertical=11
    $aaNone=-4142
    $aaAutomatic=-4105
    $aaThin=2
    $aaMedium = -4138
    $aaThick = 4
    [void]$dd.select()
    $dd.Borders.Item($aaEdgeLeft).LineStyle = 1
    $dd.Borders.Item($aaEdgeLeft).ColorIndex = -4105
    $dd.Borders.Item($aaEdgeLeft).Color = 1
    $dd.Borders.Item($aaEdgeLeft).Weight = -4138
    $dd.Borders.Item($aaEdgeTop).LineStyle = 1
    $dd.Borders.Item($aaEdgeBottom).LineStyle = 1
    $dd.Borders.Item($aaEdgeRight).LineStyle = 1
    $dd.Borders.Item($aaInsideVertical).LineStyle = 1
    $dd.Borders.Item($aaInsideHorizontal).LineStyle = 1
    $wb1.save()
    $exf.quit()
    [gc]::collect()
    remove-item $file1

        $d = get-date
        $listBox1.appendtext("$d : Completed Specific software Module `r `n")
        $listBox1.appendtext(" `r`n")
        $objForm.close()
        }

    $OKButton.add_Click($chumma)
    $objForm.ShowDialog() | out-null
    ac .\log.txt -Value " "
    ac .\log.txt -Value "******** Product Validation Module Completed *********"
    ac .\log.txt -Value " "
    } #function ends here

    #=======================================================================================================
function Specific_security_update
    {
    #****************************************
    $objForm444 = New-Object System.Windows.Forms.Form
    $objForm444.Text = "Getting KB info"
    $objForm444.Size = New-Object System.Drawing.Size(200,400)
    $objForm444.StartPosition = "CenterScreen"

    $objLabel = New-Object System.Windows.Forms.Label
    $objLabel.Location = New-Object System.Drawing.Size(10,20)
    $objLabel.Size = New-Object System.Drawing.Size(280,20)
    $objLabel.Text = "Please enter KB info"
    $objForm444.Controls.Add($objLabel)
    #****************************************
    $objLabel1 = New-Object System.Windows.Forms.Label
    $objLabel1.Location = New-Object System.Drawing.Size(10,50)
    $objLabel1.Size = New-Object System.Drawing.Size(30,20)
    $objLabel1.Text = "KB1"
    $objForm444.Controls.Add($objLabel1)

    $info1 = New-Object System.Windows.Forms.TextBox
    $info1.Location = New-Object System.Drawing.Size(50,50)
    $info1.Size = New-Object System.Drawing.Size(120,25)
    $objForm444.Controls.Add($info1)
    #****************************************
    $objLabel2 = New-Object System.Windows.Forms.Label
    $objLabel2.Location = New-Object System.Drawing.Size(10,80)
    $objLabel2.Size = New-Object System.Drawing.Size(30,20)
    $objLabel2.Text = "KB2"
    $objForm444.Controls.Add($objLabel2)

    $info2 = New-Object System.Windows.Forms.TextBox
    $info2.Location = New-Object System.Drawing.Size(50,80)
    $info2.Size = New-Object System.Drawing.Size(120,25)
    $objForm444.Controls.Add($info2)
    #****************************************
    $objLabel3 = New-Object System.Windows.Forms.Label
    $objLabel3.Location = New-Object System.Drawing.Size(10,110)
    $objLabel3.Size = New-Object System.Drawing.Size(30,20)
    $objLabel3.Text = "KB3"
    $objForm444.Controls.Add($objLabel3)

    $info3 = New-Object System.Windows.Forms.TextBox
    $info3.Location = New-Object System.Drawing.Size(50,110)
    $info3.Size = New-Object System.Drawing.Size(120,25)
    $objForm444.Controls.Add($info3)
    #****************************************
    $objLabel4 = New-Object System.Windows.Forms.Label
    $objLabel4.Location = New-Object System.Drawing.Size(10,140)
    $objLabel4.Size = New-Object System.Drawing.Size(30,20)
    $objLabel4.Text = "KB4"
    $objForm444.Controls.Add($objLabel4)

    $info4 = New-Object System.Windows.Forms.TextBox
    $info4.Location = New-Object System.Drawing.Size(50,140)
    $info4.Size = New-Object System.Drawing.Size(120,25)
    $objForm444.Controls.Add($info4)
    #****************************************
    $objLabel5 = New-Object System.Windows.Forms.Label
    $objLabel5.Location = New-Object System.Drawing.Size(10,170)
    $objLabel5.Size = New-Object System.Drawing.Size(30,20)
    $objLabel5.Text = "KB5"
    $objForm444.Controls.Add($objLabel5)

    $info5 = New-Object System.Windows.Forms.TextBox
    $info5.Location = New-Object System.Drawing.Size(50,170)
    $info5.Size = New-Object System.Drawing.Size(120,25)
    $objForm444.Controls.Add($info5)
    #****************************************
    $objLabel6 = New-Object System.Windows.Forms.Label
    $objLabel6.Location = New-Object System.Drawing.Size(10,200)
    $objLabel6.Size = New-Object System.Drawing.Size(30,20)
    $objLabel6.Text = "KB6"
    $objForm444.Controls.Add($objLabel6)

    $info6 = New-Object System.Windows.Forms.TextBox
    $info6.Location = New-Object System.Drawing.Size(50,200)
    $info6.Size = New-Object System.Drawing.Size(120,25)
    $objForm444.Controls.Add($info6)
    #****************************************
    $objLabel7 = New-Object System.Windows.Forms.Label
    $objLabel7.Location = New-Object System.Drawing.Size(10,230)
    $objLabel7.Size = New-Object System.Drawing.Size(30,20)
    $objLabel7.Text = "KB7"
    $objForm444.Controls.Add($objLabel7)

    $info7 = New-Object System.Windows.Forms.TextBox
    $info7.Location = New-Object System.Drawing.Size(50,230)
    $info7.Size = New-Object System.Drawing.Size(120,25)
    $objForm444.Controls.Add($info7)
    #****************************************
    $objLabel8 = New-Object System.Windows.Forms.Label
    $objLabel8.Location = New-Object System.Drawing.Size(10,260)
    $objLabel8.Size = New-Object System.Drawing.Size(30,20)
    $objLabel8.Text = "KB8"
    $objForm444.Controls.Add($objLabel8)

    $info8 = New-Object System.Windows.Forms.TextBox
    $info8.Location = New-Object System.Drawing.Size(50,260)
    $info8.Size = New-Object System.Drawing.Size(120,25)
    $objForm444.Controls.Add($info8)
    #****************************************
    $OKButton = New-Object System.Windows.Forms.Button
    $OKButton.Location = New-Object System.Drawing.Size(60,300)
    $OKButton.Size = New-Object System.Drawing.Size(75,23)
    $OKButton.Text = "Submit"
    $objForm444.Controls.Add($OKButton)
    #****************************************
    $chumma = {
    $objForm444.Close()
    [string]$OS= "ProductName"
    [string]$path= "SOFTWARE\Microsoft\Windows NT\CurrentVersion"
    $row = 2
    $countkb1 = 0
    $countkb2 = 0
    $countkb3 = 0
    $countkb4 = 0
    $countkb5 = 0
    $countkb6 = 0
    $countkb7 = 0
    $countkb8 = 0
    ac .\log.txt -Value "***** Started processing MS patch validation module "
    ac .\log.txt -Value " "
    $array = @()
    $introw = 2
    if ( ( $info1.text -eq $null) -or ($info1.text -eq "" ) ) { $kb1 = $null } else { $kb1 = $info1.Text; ac .\log.txt -Value "User entered : $kb1"; }
    if ( ( $info2.text -eq $null) -or ($info2.text -eq "" ) ) { $kb2 = $null } else { $kb2 = $info2.Text; ac .\log.txt -Value "User entered : $kb2"; }
    if ( ( $info3.text -eq $null) -or ($info3.text -eq "" ) ) { $kb3 = $null } else { $kb3 = $info3.Text; ac .\log.txt -Value "User entered : $kb3"; }
    if ( ( $info4.text -eq $null) -or ($info4.text -eq "" ) ) { $kb4 = $null } else { $kb4 = $info4.Text; ac .\log.txt -Value "User entered : $kb4"; }
    if ( ( $info5.text -eq $null) -or ($info5.text -eq "" ) ) { $kb5 = $null } else { $kb5 = $info5.Text; ac .\log.txt -Value "User entered : $kb5"; }
    if ( ( $info6.text -eq $null) -or ($info6.text -eq "" ) ) { $kb6 = $null } else { $kb6 = $info6.Text; ac .\log.txt -Value "User entered : $kb6"; }
    if ( ( $info7.text -eq $null) -or ($info7.text -eq "" ) ) { $kb7 = $null } else { $kb7 = $info7.Text; ac .\log.txt -Value "User entered : $kb7"; }
    if ( ( $info8.text -eq $null) -or ($info8.text -eq "" ) ) { $kb8 = $null } else { $kb8 = $info8.Text; ac .\log.txt -Value "User entered : $kb8"; }
    $array = @()
    $array2 = @()
    $d = get-date
    $listBox1.appendtext("$d : Started finding specific security update `r `n")
    $listBox1.appendtext(" `r`n")
    $aaa = import-csv ".\serverlist1.csv"

    foreach ( $comp in $aaa )
    {
    $d = get-date
    $ser = $comp.server
    ac .\log.txt -Value " "
    ac .\log.txt -Value "Working on $ser"
    $listBox1.appendtext("$d : Working on $ser `r `n")
    $listBox1.appendtext(" `r`n")
    $a = $null
    $a = Test-Connection -ComputerName $ser -Count 1 -erroraction silentlycontinue
        if($a -ne $null)
         {
             ac .\log.txt -Value "$ser - server is in network"
             $reg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine',$ser)
             $regKey= $reg.OpenSubKey("$path")
             [string]$productName=$regkey.GetValue("$os")
             $obj = $null
             $obj = gwmi Win32_QuickFixEngineering -Computer $ser | Select-Object CSName,description,HotfixID,InstalledBy,Installedon 
             if($obj -ne $null)
             {
                    if ( $productname -like "*2008*" )
                    {
                            foreach ( $hotfix in $obj )
                            {
                            if ( ($kb1 -ne $null) -and ($hotfix.hotfixid -like $kb1) )
                            {
                            $countkb1 = 1
                            $obj12 = New-Object PSObject
                            $obj12 | Add-Member -MemberType NoteProperty -Name "ComputerName" -Value $ser
                            $obj12 | Add-Member -MemberType NoteProperty -Name "Description" -Value $hotfix.description
                            $obj12 | Add-Member -MemberType NoteProperty -Name "hotfixid" -Value $hotfix.hotfixid
                            $obj12 | Add-Member -MemberType NoteProperty -Name "InstalledBy" -Value $hotfix.InstalledBy
                            $obj12 | Add-Member -MemberType NoteProperty -Name "Installedon" -Value $hotfix.Installedon
                            $array += $obj12
                            ac .\log.txt -Value "$ser - $kb1 found"
                            }
                            
                            if ( ($kb2 -ne $null) -and ($hotfix.hotfixid -like $kb2) )
                            {
                            $countkb2 = 1
                            $obj12 = New-Object PSObject
                            $obj12 | Add-Member -MemberType NoteProperty -Name "ComputerName" -Value $ser
                            $obj12 | Add-Member -MemberType NoteProperty -Name "Description" -Value $hotfix.description
                            $obj12 | Add-Member -MemberType NoteProperty -Name "hotfixid" -Value $hotfix.hotfixid
                            $obj12 | Add-Member -MemberType NoteProperty -Name "InstalledBy" -Value $hotfix.InstalledBy
                            $obj12 | Add-Member -MemberType NoteProperty -Name "Installedon" -Value $hotfix.Installedon
                            $array += $obj12
                            ac .\log.txt -Value "$ser - $kb2 found"
                            }
                            if ( ($kb3 -ne $null) -and ($hotfix.hotfixid -like $kb3) )
                            {
                            $countkb3 = 1
                            $obj12 = New-Object PSObject
                            $obj12 | Add-Member -MemberType NoteProperty -Name "ComputerName" -Value $ser
                            $obj12 | Add-Member -MemberType NoteProperty -Name "Description" -Value $hotfix.description
                            $obj12 | Add-Member -MemberType NoteProperty -Name "hotfixid" -Value $hotfix.hotfixid
                            $obj12 | Add-Member -MemberType NoteProperty -Name "InstalledBy" -Value $hotfix.InstalledBy
                            $obj12 | Add-Member -MemberType NoteProperty -Name "Installedon" -Value $hotfix.Installedon
                            $array += $obj12
                            ac .\log.txt -Value "$ser - $kb3 found"
                            }
                            if ( ($kb4 -ne $null) -and ($hotfix.hotfixid -like $kb4) )
                            {
                            $countkb4 = 1
                            $obj12 = New-Object PSObject
                            $obj12 | Add-Member -MemberType NoteProperty -Name "ComputerName" -Value $ser
                            $obj12 | Add-Member -MemberType NoteProperty -Name "Description" -Value $hotfix.description
                            $obj12 | Add-Member -MemberType NoteProperty -Name "hotfixid" -Value $hotfix.hotfixid
                            $obj12 | Add-Member -MemberType NoteProperty -Name "InstalledBy" -Value $hotfix.InstalledBy
                            $obj12 | Add-Member -MemberType NoteProperty -Name "Installedon" -Value $hotfix.Installedon
                            $array += $obj12
                            ac .\log.txt -Value "$ser - $kb4 found"
                            }
                            if ( ($kb5 -ne $null) -and ($hotfix.hotfixid -like $kb5) )
                            {
                            $countkb5 = 1
                            $obj12 = New-Object PSObject
                            $obj12 | Add-Member -MemberType NoteProperty -Name "ComputerName" -Value $ser
                            $obj12 | Add-Member -MemberType NoteProperty -Name "Description" -Value $hotfix.description
                            $obj12 | Add-Member -MemberType NoteProperty -Name "hotfixid" -Value $hotfix.hotfixid
                            $obj12 | Add-Member -MemberType NoteProperty -Name "InstalledBy" -Value $hotfix.InstalledBy
                            $obj12 | Add-Member -MemberType NoteProperty -Name "Installedon" -Value $hotfix.Installedon
                            $array += $obj12
                            ac .\log.txt -Value "$ser - $kb5 found"
                            }
                            if ( ($kb6 -ne $null) -and ($hotfix.hotfixid -like $kb6) )
                            {
                            $countkb6 = 1
                            $obj12 = New-Object PSObject
                            $obj12 | Add-Member -MemberType NoteProperty -Name "ComputerName" -Value $ser
                            $obj12 | Add-Member -MemberType NoteProperty -Name "Description" -Value $hotfix.description
                            $obj12 | Add-Member -MemberType NoteProperty -Name "hotfixid" -Value $hotfix.hotfixid
                            $obj12 | Add-Member -MemberType NoteProperty -Name "InstalledBy" -Value $hotfix.InstalledBy
                            $obj12 | Add-Member -MemberType NoteProperty -Name "Installedon" -Value $hotfix.Installedon
                            $array += $obj12
                            ac .\log.txt -Value "$ser - $kb6 found"
                            }
                            if ( ($kb7 -ne $null) -and ($hotfix.hotfixid -like $kb7) )
                            {
                            $countkb7 = 1
                            $obj12 = New-Object PSObject
                            $obj12 | Add-Member -MemberType NoteProperty -Name "ComputerName" -Value $ser
                            $obj12 | Add-Member -MemberType NoteProperty -Name "Description" -Value $hotfix.description
                            $obj12 | Add-Member -MemberType NoteProperty -Name "hotfixid" -Value $hotfix.hotfixid
                            $obj12 | Add-Member -MemberType NoteProperty -Name "InstalledBy" -Value $hotfix.InstalledBy
                            $obj12 | Add-Member -MemberType NoteProperty -Name "Installedon" -Value $hotfix.Installedon
                            $array += $obj12
                            ac .\log.txt -Value "$ser - $kb7 found"
                            }
                            if ( ($kb8 -ne $null) -and ($hotfix.hotfixid -like $kb8) )
                            {
                            $countkb8 = 1
                            $obj12 = New-Object PSObject
                            $obj12 | Add-Member -MemberType NoteProperty -Name "ComputerName" -Value $ser
                            $obj12 | Add-Member -MemberType NoteProperty -Name "Description" -Value $hotfix.description
                            $obj12 | Add-Member -MemberType NoteProperty -Name "hotfixid" -Value $hotfix.hotfixid
                            $obj12 | Add-Member -MemberType NoteProperty -Name "InstalledBy" -Value $hotfix.InstalledBy
                            $obj12 | Add-Member -MemberType NoteProperty -Name "Installedon" -Value $hotfix.Installedon
                            $array += $obj12
                            ac .\log.txt -Value "$ser - $kb8 found"
                            }
                    }


                            #retrieving uninstall work here
                            $reg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine',$ser)
                            $regKey= $reg.OpenSubKey("SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall")
                            $subkeys=$regkey.GetSubKeyNames()

                            foreach ( $ke in $subkeys )
                            {
                        $key =  "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\$ke"
                        $new = $reg.opensubkey("$key")
                        $out = $new.getvalue("DisplayName")
                        $out1 = $new.getvalue("")
                        if ($kb1 -ne $null)
                        {
                            if ( ( $out1 -like "*$kb1*" ) -or ( $out -like "*$kb1*" ) )
                            {
                                $obj12 = New-Object PSObject
                                $obj12 | Add-Member -MemberType NoteProperty -Name "ComputerName" -Value $ser
                                if ( $out -like "*$kb1*" )
                                {
                                  $obj12 | Add-Member -MemberType NoteProperty -Name "description" -Value $out
                                  $countkb1 = 1
                                }
                                if ( $out1 -like "*$kb1*" )
                                {
                                  $obj12 | Add-Member -MemberType NoteProperty -Name "hotfixid" -Value $out1
                                  $countkb1 = 1
                                }
                                $array += $obj12
                            }
                        }
                        if ($kb2 -ne $null)
                        {
                            if ( ( $out1 -like "*$kb2*" ) -or ( $out -like "*$kb2*" ) )
                            {
                                $obj12 = New-Object PSObject
                                $obj12 | Add-Member -MemberType NoteProperty -Name "ComputerName" -Value $comp.server
                                if ( $out -like "*$kb2*" )
                                {
                                  $obj12 | Add-Member -MemberType NoteProperty -Name "description" -Value $out
                                  $countkb2 = 1
                                }
                                if ( $out1 -like "*$kb2*")
                                {
                                  $obj12 | Add-Member -MemberType NoteProperty -Name "hotfixid" -Value $out1
                                  $countkb2 = 1
                                }
                               $array += $obj12
                            }
                        }
                        if ($kb3 -ne $null)
                        {
                            if ( ( $out1 -like "*$kb3*" ) -or ( $out -like "*$kb3*" ) )
                            {
                                $obj12 = New-Object PSObject
                                $obj12 | Add-Member -MemberType NoteProperty -Name "ComputerName" -Value $comp.server
                                if ( $out -like "*$kb3*" )
                                {
                                  $obj12 | Add-Member -MemberType NoteProperty -Name "description" -Value $out
                                  $countkb3 = 1
                                }
                                if ( $out1 -like "*$kb3*" )
                                {
                                  $obj12 | Add-Member -MemberType NoteProperty -Name "hotfixid" -Value $out1
                                  $countkb3 = 1
                                }
                               $array += $obj12
                         }
                        }
                        if ($kb4 -ne $null)
                        {
                            if ( ( $out1 -like "*$kb4*" ) -or ( $out -like "*$kb4*" ) )
                            {
                                $obj12 = New-Object PSObject
                                $obj12 | Add-Member -MemberType NoteProperty -Name "ComputerName" -Value $comp.server
                                if ( $out -like "*$kb4*" )
                                {
                                  $obj12 | Add-Member -MemberType NoteProperty -Name "description" -Value $out
                                  $countkb4 = 1
                                }
                                if ( $out1 -like "*$kb4*" )
                                {
                                  $obj12 | Add-Member -MemberType NoteProperty -Name "hotfixid" -Value $out1
                                  $countkb4 = 1
                                }
                               $array += $obj12
                         }
                        }
                        if ($kb5 -ne $null)
                        {
                            if ( ( $out1 -like "*$kb5*" ) -or ( $out -like "*$kb5*" ) )
                            {
                                $obj12 = New-Object PSObject
                                $obj12 | Add-Member -MemberType NoteProperty -Name "ComputerName" -Value $comp.server
                                if ( $out -like "*$kb5*" )
                                {
                                  $obj12 | Add-Member -MemberType NoteProperty -Name "description" -Value $out
                                  $countkb5 = 1
                                }
                                if ( $out1 -like "*$kb5*" )
                                {
                                  $obj12 | Add-Member -MemberType NoteProperty -Name "hotfixid" -Value $out1
                                  $countkb5 = 1
                                }
                               $array += $obj12
                         }
                        }
                        if ($kb6 -ne $null)
                        {
                            if ( ( $out1 -like "*$kb6*" ) -or ( $out -like "*$kb6*" ) )
                            {
                                $obj12 = New-Object PSObject
                                $obj12 | Add-Member -MemberType NoteProperty -Name "ComputerName" -Value $comp.server
                                if ( $out -like "*$kb6*" )
                                {
                                  $obj12 | Add-Member -MemberType NoteProperty -Name "description" -Value $out
                                  $countkb6 = 1
                                }
                                if ( $out1 -like "*$kb6*" )
                                {
                                  $obj12 | Add-Member -MemberType NoteProperty -Name "hotfixid" -Value $out1
                                  $countkb6 = 1
                                }
                               $array += $obj12
                         }
                        }
                        if ($kb7 -ne $null)
                        {
                            if ( ( $out1 -like "*$kb7*" ) -or ( $out -like "*$kb7*" ) )
                            {
                                $obj12 = New-Object PSObject
                                $obj12 | Add-Member -MemberType NoteProperty -Name "ComputerName" -Value $comp.server
                                if ( $out -like "*$kb7*" )
                                {
                                  $obj12 | Add-Member -MemberType NoteProperty -Name "description" -Value $out
                                  $countkb7 = 1
                                }
                                if ( $out1 -like "*$kb7*" )
                                {
                                  $obj12 | Add-Member -MemberType NoteProperty -Name "hotfixid" -Value $out1
                                  $countkb7 = 1
                                }
                               $array += $obj12
                         }
                        }
                        if ($kb8 -ne $null)
                        {
                            if ( ( $out1 -like "*$kb8*" ) -or ( $out -like "*$kb8*" ) )
                            {
                                $obj12 = New-Object PSObject
                                $obj12 | Add-Member -MemberType NoteProperty -Name "ComputerName" -Value $comp.server
                                if ( $out -like "*$kb8*" )
                                {
                                  $obj12 | Add-Member -MemberType NoteProperty -Name "description" -Value $out
                                  $countkb8 = 1
                                }
                                if ( $out1 -like "*$kb8*" )
                                {
                                  $obj12 | Add-Member -MemberType NoteProperty -Name "hotfixid" -Value $out1
                                  $countkb8 = 1
                                }
                               $array += $obj12
                            }
                        }

                    }

                            #retrieving WOW6432 node uninstall work here
                            $reg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine',$ser)
                            $regKey= $reg.OpenSubKey("SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall")
                            $subkeys=$regkey.GetSubKeyNames()

                            foreach ( $ke in $subkeys )
                             {
                        $key =  "SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\$ke"
                        $new = $reg.opensubkey("$key")
                        $out = $new.getvalue("DisplayName")
                        $out1 = $new.getvalue("")

                        if ($kb1 -ne $null)
                        {
                            if ( ( $out1 -like "*$kb1*" ) -or ( $out -like "*$kb1*" ) )
                            {
                                $obj12 = New-Object PSObject
                                $obj12 | Add-Member -MemberType NoteProperty -Name "ComputerName" -Value $ser
                                if ( $out -like "*$kb1*" )
                                {
                                  $obj12 | Add-Member -MemberType NoteProperty -Name "description" -Value $out
                                  $countkb1 = 1
                                }
                                if ( $out1 -like "*$kb1*" )
                                {
                                  $obj12 | Add-Member -MemberType NoteProperty -Name "hotfixid" -Value $out1
                                  $countkb1 = 1
                                }
                                $array += $obj12
                        }
                        }
                        if ($kb2 -ne $null)
                        {
                            if ( ( $out1 -like "*$kb2*" ) -or ( $out -like "*$kb2*" ) )
                            {
                                $obj12 = New-Object PSObject
                                $obj12 | Add-Member -MemberType NoteProperty -Name "ComputerName" -Value $comp.server
                                if ( $out -like "*$kb2*" )
                                {
                                  $obj12 | Add-Member -MemberType NoteProperty -Name "description" -Value $out
                                  $countkb2 = 1
                                }
                                if ( $out1 -like "*$kb2*" )
                                {
                                  $obj12 | Add-Member -MemberType NoteProperty -Name "hotfixid" -Value $out1
                                  $countkb2 = 1
                                }
                               $array += $obj12
                            }
                        }
                        if ($kb3 -ne $null)
                        {
                            if ( ( $out1 -like "*$kb3*" ) -or ( $out -like "*$kb3*" ) )
                            {
                                $obj12 = New-Object PSObject
                                $obj12 | Add-Member -MemberType NoteProperty -Name "ComputerName" -Value $comp.server
                                if ( $out -like "*$kb3*" )
                                {
                                  $obj12 | Add-Member -MemberType NoteProperty -Name "description" -Value $out
                                  $countkb3 = 1
                                }
                                if ( $out1 -like "*$kb1*" )
                                {
                                  $obj12 | Add-Member -MemberType NoteProperty -Name "hotfixid" -Value $out1
                                  $countkb3 = 1
                                }
                               $array += $obj12
                         }
                        }
                        if ($kb4 -ne $null)
                        {
                            if ( ( $out1 -like "*$kb4*" ) -or ( $out -like "*$kb4*" ) )
                            {
                                $obj12 = New-Object PSObject
                                $obj12 | Add-Member -MemberType NoteProperty -Name "ComputerName" -Value $comp.server
                                if ( $out -like "*$kb4*" )
                                {
                                  $obj12 | Add-Member -MemberType NoteProperty -Name "description" -Value $out
                                  $countkb4 = 1
                                }
                                if ( $out1 -like "*$kb4*" )
                                {
                                  $obj12 | Add-Member -MemberType NoteProperty -Name "hotfixid" -Value $out1
                                  $countkb4 = 1
                                }
                               $array += $obj12
                         }
                        }
                        if ($kb5 -ne $null)
                        {
                            if ( ( $out1 -like "*$kb5*" ) -or ( $out -like "*$kb5*" ) )
                            {
                                $obj12 = New-Object PSObject
                                $obj12 | Add-Member -MemberType NoteProperty -Name "ComputerName" -Value $comp.server
                                if ( $out -like "*$kb5*" )
                                {
                                  $obj12 | Add-Member -MemberType NoteProperty -Name "description" -Value $out
                                  $countkb5 = 1
                                }
                                if ( $out1 -like "*$kb5*" )
                                {
                                  $obj12 | Add-Member -MemberType NoteProperty -Name "hotfixid" -Value $out1
                                  $countkb5 = 1
                                }
                               $array += $obj12
                         }
                        }
                        if ($kb6 -ne $null)
                        {
                            if ( ( $out1 -like "*$kb6*" ) -or ( $out -like "*$kb6*" ) )
                            {
                                $obj12 = New-Object PSObject
                                $obj12 | Add-Member -MemberType NoteProperty -Name "ComputerName" -Value $comp.server
                                if ( $out -like "*$kb6*" )
                                {
                                  $obj12 | Add-Member -MemberType NoteProperty -Name "description" -Value $out
                                  $countkb6 = 1
                                }
                                if ( $out1 -like "*$kb6*" )
                                {
                                  $obj12 | Add-Member -MemberType NoteProperty -Name "hotfixid" -Value $out1
                                  $countkb6 = 1
                                }
                               $array += $obj12
                         }
                        }
                        if ($kb7 -ne $null)
                        {
                            if ( ( $out1 -like "*$kb7*" ) -or ( $out -like "*$kb7*" ) )
                            {
                                $obj12 = New-Object PSObject
                                $obj12 | Add-Member -MemberType NoteProperty -Name "ComputerName" -Value $comp.server
                                if ( $out -like "*$kb7*" )
                                {
                                  $obj12 | Add-Member -MemberType NoteProperty -Name "description" -Value $out
                                  $countkb7 = 1
                                }
                                if ( $out1 -like "*$kb7*" )
                                {
                                  $obj12 | Add-Member -MemberType NoteProperty -Name "hotfixid" -Value $out1
                                  $countkb7 = 1
                                }
                               $array += $obj12
                         }
                        }
                        if ($kb8 -ne $null)
                        {
                            if ( ( $out1 -like "*$kb8*" ) -or ( $out -like "*$kb8*" ) )
                            {
                                $obj12 = New-Object PSObject
                                $obj12 | Add-Member -MemberType NoteProperty -Name "ComputerName" -Value $comp.server
                                if ( $out -like "*$kb8*" )
                                {
                                  $obj12 | Add-Member -MemberType NoteProperty -Name "description" -Value $out
                                  $countkb8 = 1
                                }
                                if ( $out1 -like "*$kb8*" )
                                {
                                  $obj12 | Add-Member -MemberType NoteProperty -Name "hotfixid" -Value $out1
                                  $countkb8 = 1
                                }
                               $array += $obj12
                             }
                        }

                    }

                            if ( ($kb1 -ne $null) -and ($countkb1 -eq 0) ){
                          $obj12 = New-Object PSObject
                          $nf = "Not Found"
                          $obj12 | Add-Member -MemberType NoteProperty -Name "ComputerName" -Value $comp.server
                          $obj12 | Add-Member -MemberType NoteProperty -Name "hotfixid" -Value $kb1
                          $obj12 | Add-Member -MemberType NoteProperty -Name "description" -Value $nf
                          $array += $obj12
                          ac .\log.txt -Value "$ser - $kb1 Not found"
                          $countkb1 = 0
                            }
                            if ( ($kb2 -ne $null) -and ($countkb2 -eq 0) ){
                          $obj12 = New-Object PSObject
                          $nf = "Not Found"
                          $obj12 | Add-Member -MemberType NoteProperty -Name "ComputerName" -Value $comp.server
                          $obj12 | Add-Member -MemberType NoteProperty -Name "hotfixid" -Value $kb2
                          $obj12 | Add-Member -MemberType NoteProperty -Name "description" -Value $nf
                          $array += $obj12
                          ac .\log.txt -Value "$ser - $kb2 Not found"
                          $countkb2 = 0
                            }
                            if ( ($kb3 -ne $null) -and ($countkb3 -eq 0) ){
                          $obj12 = New-Object PSObject
                          $nf = "Not Found"
                          $obj12 | Add-Member -MemberType NoteProperty -Name "ComputerName" -Value $comp.server
                          $obj12 | Add-Member -MemberType NoteProperty -Name "hotfixid" -Value $kb3
                          $obj12 | Add-Member -MemberType NoteProperty -Name "description" -Value $nf
                          $array += $obj12
                          ac .\log.txt -Value "$ser - $kb3 Not found"
                          $countkb3 = 0
                            }
                            if ( ($kb4 -ne $null) -and ($countkb4 -eq 0) ){
                          $obj12 = New-Object PSObject
                          $nf = "Not Found"
                          $obj12 | Add-Member -MemberType NoteProperty -Name "ComputerName" -Value $comp.server
                          $obj12 | Add-Member -MemberType NoteProperty -Name "hotfixid" -Value $kb4
                          $obj12 | Add-Member -MemberType NoteProperty -Name "description" -Value $nf
                          $array += $obj12
                          ac .\log.txt -Value "$ser - $kb4 Not found"
                          $countkb4 = 0
                             }
                            if ( ($kb5 -ne $null) -and ($countkb5 -eq 0) ){
                          $obj12 = New-Object PSObject
                          $nf = "Not Found"
                          $obj12 | Add-Member -MemberType NoteProperty -Name "ComputerName" -Value $comp.server
                          $obj12 | Add-Member -MemberType NoteProperty -Name "hotfixid" -Value $kb5
                          $obj12 | Add-Member -MemberType NoteProperty -Name "description" -Value $nf
                          $array += $obj12
                          ac .\log.txt -Value "$ser - $kb5 Not found"
                          $countkb5 = 0
                            }
                            if ( ($kb6 -ne $null) -and ($countkb6 -eq 0) ){
                          $obj12 = New-Object PSObject
                          $nf = "Not Found"
                          $obj12 | Add-Member -MemberType NoteProperty -Name "ComputerName" -Value $comp.server
                          $obj12 | Add-Member -MemberType NoteProperty -Name "hotfixid" -Value $kb6
                          $obj12 | Add-Member -MemberType NoteProperty -Name "description" -Value $nf
                          $array += $obj12
                          ac .\log.txt -Value "$ser - $kb6 Not found"
                          $countkb6 = 0
                             }
                            if ( ($kb7 -ne $null) -and ($countkb7 -eq 0) ){
                          $obj12 = New-Object PSObject
                          $nf = "Not Found"
                          $obj12 | Add-Member -MemberType NoteProperty -Name "ComputerName" -Value $comp.server
                          $obj12 | Add-Member -MemberType NoteProperty -Name "hotfixid" -Value $kb7
                          $obj12 | Add-Member -MemberType NoteProperty -Name "description" -Value $nf
                          $array += $obj12
                          ac .\log.txt -Value "$ser - $kb7 Not found"
                          $countkb7 = 0
                            }
                            if ( ($kb8 -ne $null) -and ($countkb8 -eq 0) ){
                          $obj12 = New-Object PSObject
                          $nf = "Not Found"
                          $obj12 | Add-Member -MemberType NoteProperty -Name "ComputerName" -Value $comp.server
                          $obj12 | Add-Member -MemberType NoteProperty -Name "hotfixid" -Value $kb8
                          $obj12 | Add-Member -MemberType NoteProperty -Name "description" -Value $nf
                          $array += $obj12
                          ac .\log.txt -Value "$ser - $kb8 Not found"
                          $countkb8 = 0
                           }
                   
             } #2008 ends here
                    else #2003 begins here
                    {
                        $reg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine',$ser)
                        $regKey= $reg.OpenSubKey("$path")
                        [string]$productName=$regkey.GetValue("$os")
                        $obj = gwmi Win32_QuickFixEngineering -Computer $ser | Select-Object CSName,description,HotfixID,InstalledBy,Installedon 
                        if($obj -ne $null)
                         {
                    
                        foreach ( $hotfix in $obj )
                        {
                                if ( ($kb1 -ne $null ) -and ($hotfix.hotfixid -like $kb1 ) )
                                {
                             $countkb1 = 1
                            $obj12 = New-Object PSObject
                            $obj12 | Add-Member -MemberType NoteProperty -Name "ComputerName" -Value $comp.server
                            $obj12 | Add-Member -MemberType NoteProperty -Name "Description" -Value $hotfix.description
                            $obj12 | Add-Member -MemberType NoteProperty -Name "hotfixid" -Value $hotfix.hotfixid
                            $obj12 | Add-Member -MemberType NoteProperty -Name "InstalledBy" -Value $hotfix.InstalledBy
                            $obj12 | Add-Member -MemberType NoteProperty -Name "Installedon" -Value $hotfix.Installedon
                            ac .\log.txt -Value "$ser - $kb1 found"
                            $array += $obj12
                                 }
                                if ( ($kb2 -ne $null ) -and ($hotfix.hotfixid -like $kb2 ) )
                                {
                             $countkb2 = 1
                            $obj12 = New-Object PSObject
                            $obj12 | Add-Member -MemberType NoteProperty -Name "ComputerName" -Value $comp.server
                            $obj12 | Add-Member -MemberType NoteProperty -Name "Description" -Value $hotfix.description
                            $obj12 | Add-Member -MemberType NoteProperty -Name "hotfixid" -Value $hotfix.hotfixid
                            $obj12 | Add-Member -MemberType NoteProperty -Name "InstalledBy" -Value $hotfix.InstalledBy
                            $obj12 | Add-Member -MemberType NoteProperty -Name "Installedon" -Value $hotfix.Installedon
                            ac .\log.txt -Value "$ser - $kb2 found"
                            $array += $obj12
                    }
                                if ( ($kb3 -ne $null ) -and ($hotfix.hotfixid -like $kb3 ) )
                                {
                             $countkb3 = 1
                            $obj12 = New-Object PSObject
                            $obj12 | Add-Member -MemberType NoteProperty -Name "ComputerName" -Value $comp.server
                            $obj12 | Add-Member -MemberType NoteProperty -Name "Description" -Value $hotfix.description
                            $obj12 | Add-Member -MemberType NoteProperty -Name "hotfixid" -Value $hotfix.hotfixid
                            $obj12 | Add-Member -MemberType NoteProperty -Name "InstalledBy" -Value $hotfix.InstalledBy
                            $obj12 | Add-Member -MemberType NoteProperty -Name "Installedon" -Value $hotfix.Installedon
                            ac .\log.txt -Value "$ser - $kb3 found"
                            $array += $obj12
                    }
                                if ( ($kb4 -ne $null ) -and ($hotfix.hotfixid -like $kb4 ) )
                                {
                             $countkb4 = 1
                            $obj12 = New-Object PSObject
                            $obj12 | Add-Member -MemberType NoteProperty -Name "ComputerName" -Value $comp.server
                            $obj12 | Add-Member -MemberType NoteProperty -Name "Description" -Value $hotfix.description
                            $obj12 | Add-Member -MemberType NoteProperty -Name "hotfixid" -Value $hotfix.hotfixid
                            $obj12 | Add-Member -MemberType NoteProperty -Name "InstalledBy" -Value $hotfix.InstalledBy
                            $obj12 | Add-Member -MemberType NoteProperty -Name "Installedon" -Value $hotfix.Installedon
                            $array += $obj12
                            ac .\log.txt -Value "$ser - $kb4 found"
                    }
                                if ( ($kb5 -ne $null ) -and ($hotfix.hotfixid -like $kb5 ) )
                                {
                             $countkb5 = 1
                            $obj12 = New-Object PSObject
                            $obj12 | Add-Member -MemberType NoteProperty -Name "ComputerName" -Value $comp.server
                            $obj12 | Add-Member -MemberType NoteProperty -Name "Description" -Value $hotfix.description
                            $obj12 | Add-Member -MemberType NoteProperty -Name "hotfixid" -Value $hotfix.hotfixid
                            $obj12 | Add-Member -MemberType NoteProperty -Name "InstalledBy" -Value $hotfix.InstalledBy
                            $obj12 | Add-Member -MemberType NoteProperty -Name "Installedon" -Value $hotfix.Installedon
                            ac .\log.txt -Value "$ser - $kb5 found"
                            $array += $obj12
                    }
                                if ( ($kb6 -ne $null ) -and ($hotfix.hotfixid -like $kb6) )
                                {
                             $countkb6 = 1
                            $obj12 = New-Object PSObject
                            $obj12 | Add-Member -MemberType NoteProperty -Name "ComputerName" -Value $comp.server
                            $obj12 | Add-Member -MemberType NoteProperty -Name "Description" -Value $hotfix.description
                            $obj12 | Add-Member -MemberType NoteProperty -Name "hotfixid" -Value $hotfix.hotfixid
                            $obj12 | Add-Member -MemberType NoteProperty -Name "InstalledBy" -Value $hotfix.InstalledBy
                            $obj12 | Add-Member -MemberType NoteProperty -Name "Installedon" -Value $hotfix.Installedon
                            ac .\log.txt -Value "$ser - $kb6 found"
                            $array += $obj12
                    }
                                if ( ($kb7 -ne $null ) -and ($hotfix.hotfixid -like $kb7 ) )
                                {
                             $countkb7 = 1
                            $obj12 = New-Object PSObject
                            $obj12 | Add-Member -MemberType NoteProperty -Name "ComputerName" -Value $comp.server
                            $obj12 | Add-Member -MemberType NoteProperty -Name "Description" -Value $hotfix.description
                            $obj12 | Add-Member -MemberType NoteProperty -Name "hotfixid" -Value $hotfix.hotfixid
                            $obj12 | Add-Member -MemberType NoteProperty -Name "InstalledBy" -Value $hotfix.InstalledBy
                            $obj12 | Add-Member -MemberType NoteProperty -Name "Installedon" -Value $hotfix.Installedon
                            $array += $obj12
                            ac .\log.txt -Value "$ser - $kb7 found"
                    }
                                if ( ($kb8 -ne $null ) -and ($hotfix.hotfixid -like $kb8 ) )
                                {
                             $countkb8 = 1
                            $obj12 = New-Object PSObject
                            $obj12 | Add-Member -MemberType NoteProperty -Name "ComputerName" -Value $comp.server
                            $obj12 | Add-Member -MemberType NoteProperty -Name "Description" -Value $hotfix.description
                            $obj12 | Add-Member -MemberType NoteProperty -Name "hotfixid" -Value $hotfix.hotfixid
                            $obj12 | Add-Member -MemberType NoteProperty -Name "InstalledBy" -Value $hotfix.InstalledBy
                            $obj12 | Add-Member -MemberType NoteProperty -Name "Installedon" -Value $hotfix.Installedon
                            $array += $obj12
                            ac .\log.txt -Value "$ser - $kb8 found"
                    }
                        }# foreach hotfix loop ends
                        if ( ($kb1 -ne $null) -and ($countkb1 -eq 0) )
                        {
                          $obj12 = New-Object PSObject
                          $nf = "Not Found"
                          $obj12 | Add-Member -MemberType NoteProperty -Name "ComputerName" -Value $comp.server
                          $obj12 | Add-Member -MemberType NoteProperty -Name "description" -Value $nf
                          $obj12 | Add-Member -MemberType NoteProperty -Name "hotfixid" -Value $kb1
                          $array += $obj12
                          ac .\log.txt -Value "$ser - $kb1 Not found"
                          $countkb1 = 0
                    }
                        if ( ($kb2 -ne $null) -and ($countkb2 -eq 0) )
                        {
                          $obj12 = New-Object PSObject
                          $nf = "Not Found"
                          $obj12 | Add-Member -MemberType NoteProperty -Name "ComputerName" -Value $comp.server
                          $obj12 | Add-Member -MemberType NoteProperty -Name "description" -Value $nf
                          $obj12 | Add-Member -MemberType NoteProperty -Name "hotfixid" -Value $kb2
                          $array += $obj12
                          ac .\log.txt -Value "$ser - $kb2 Not found"
                          $countkb2 = 0
                    }
                        if ( ($kb3 -ne $null) -and ($countkb3 -eq 0) )
                        {
                          $obj12 = New-Object PSObject
                          $nf = "Not Found"
                          $obj12 | Add-Member -MemberType NoteProperty -Name "ComputerName" -Value $comp.server
                          $obj12 | Add-Member -MemberType NoteProperty -Name "description" -Value $nf
                          $obj12 | Add-Member -MemberType NoteProperty -Name "hotfixid" -Value $kb3
                          $array += $obj12
                          ac .\log.txt -Value "$ser - $kb3 Not found"
                          $countkb3 = 0
                    }
                        if ( ($kb4 -ne $null) -and ($countkb4 -eq 0) )
                        {
                          $obj12 = New-Object PSObject
                          $nf = "Not Found"
                          $obj12 | Add-Member -MemberType NoteProperty -Name "ComputerName" -Value $comp.server
                          $obj12 | Add-Member -MemberType NoteProperty -Name "description" -Value $nf
                          $obj12 | Add-Member -MemberType NoteProperty -Name "hotfixid" -Value $kb4
                          $array += $obj12
                          ac .\log.txt -Value "$ser - $kb4 Not found"
                          $countkb4 = 0
                    }
                        if ( ($kb5 -ne $null) -and ($countkb5 -eq 0) )
                        {
                          $obj12 = New-Object PSObject
                          $nf = "Not Found"
                          $obj12 | Add-Member -MemberType NoteProperty -Name "ComputerName" -Value $comp.server
                          $obj12 | Add-Member -MemberType NoteProperty -Name "description" -Value $nf
                          $obj12 | Add-Member -MemberType NoteProperty -Name "hotfixid" -Value $kb5
                          $array += $obj12
                          ac .\log.txt -Value "$ser - $kb5 Not found"
                          $countkb5 = 0
                    }
                        if ( ($kb6 -ne $null) -and ($countkb6 -eq 0) )
                        {
                          $obj12 = New-Object PSObject
                          $nf = "Not Found"
                          $obj12 | Add-Member -MemberType NoteProperty -Name "ComputerName" -Value $comp.server
                          $obj12 | Add-Member -MemberType NoteProperty -Name "description" -Value $nf
                          $obj12 | Add-Member -MemberType NoteProperty -Name "hotfixid" -Value $kb6
                          $array += $obj12
                          ac .\log.txt -Value "$ser - $kb6 Not found"
                          $countkb6 = 0
                    }
                        if ( ($kb7 -ne $null) -and ($countkb7 -eq 0) )
                        {
                          $obj12 = New-Object PSObject
                          $nf = "Not Found"
                          $obj12 | Add-Member -MemberType NoteProperty -Name "ComputerName" -Value $comp.server
                          $obj12 | Add-Member -MemberType NoteProperty -Name "description" -Value $nf
                          $obj12 | Add-Member -MemberType NoteProperty -Name "hotfixid" -Value $kb7
                          $array += $obj12
                          ac .\log.txt -Value "$ser - $kb7 Not found"
                          $countkb7 = 0
                    }
                        if ( ($kb8 -ne $null) -and ($countkb8 -eq 0) )
                        {
                          $obj12 = New-Object PSObject
                          $nf = "Not Found"
                          $obj12 | Add-Member -MemberType NoteProperty -Name "ComputerName" -Value $comp.server
                          $obj12 | Add-Member -MemberType NoteProperty -Name "description" -Value $nf
                          $obj12 | Add-Member -MemberType NoteProperty -Name "hotfixid" -Value $kb8
                          $array += $obj12
                          ac .\log.txt -Value "$ser - $kb8 Not found"
                          $countkb8 = 0
                    }
                    } # if obj null ends here
                        
                        else{
                          
                          $nf = "Access Denied"
                          $obj12 = New-Object PSObject
                            $obj12 | Add-Member -MemberType NoteProperty -Name "ComputerName" -Value $ser
                            $obj12 | Add-Member -MemberType NoteProperty -Name "Description" -Value $nf
                            $obj12 | Add-Member -MemberType NoteProperty -Name "hotfixid" -Value $nf
                            $obj12 | Add-Member -MemberType NoteProperty -Name "InstalledBy" -Value $nf
                            $obj12 | Add-Member -MemberType NoteProperty -Name "Installedon" -Value $nf
                          $array += $obj12
                          ac .\log.txt -Value "$ser - Registry access denied" 
                            }#else access denied ends here
                    }#else 2003loop ends here
             } # $obj null ends
             else
             {
                    
                    $nf = "Access Denied"
                    $com = $ser
                    $ser
                    $obj12 = New-Object PSObject
                            $obj12 | Add-Member -MemberType NoteProperty -Name "ComputerName" -Value $ser
                            $obj12 | Add-Member -MemberType NoteProperty -Name "Description" -Value $nf
                            $obj12 | Add-Member -MemberType NoteProperty -Name "hotfixid" -Value $nf
                            $obj12 | Add-Member -MemberType NoteProperty -Name "InstalledBy" -Value $nf
                            $obj12 | Add-Member -MemberType NoteProperty -Name "Installedon" -Value $nf
                    $array += $obj12
                    $obj12
                    ac .\log.txt -Value "$ser - Registry access denied"   
             }#else obj ends here

         } # test ping loop ends
        else
        {
                          $c = $comp.server
                         $nf = "Out of network"
                          $obj12 = New-Object PSObject
                            $obj12 | Add-Member -MemberType NoteProperty -Name "ComputerName" -Value $ser
                            $obj12 | Add-Member -MemberType NoteProperty -Name "Description" -Value $nf
                            $obj12 | Add-Member -MemberType NoteProperty -Name "hotfixid" -Value $nf
                            $obj12 | Add-Member -MemberType NoteProperty -Name "InstalledBy" -Value $nf
                            $obj12 | Add-Member -MemberType NoteProperty -Name "Installedon" -Value $nf
                          $array += $obj12
                          ac .\log.txt -Value "$ser - Server is not in network"
             write-host "$comp.server out of network"
         }

     }# server ping  loop ends here
    $d = get-date
    $listBox1.appendtext("$d : writing results of security update `r `n")
    $listBox1.appendtext(" `r`n")
    $array2 = $array | where-Object { $_.Description } | select ComputerName, Description, hotfixid, InstalledBy, Installedon
    $array2 | export-csv ".\Specific_Security_Update.csv"
    $array2 | ogv

    #copy and conversion begins

    #4. perform copy from sec_update.xlsx to results.xlsx (in Security_update Tab )
    $file1 = resolve-path ".\Specific_Security_Update.csv" # source's fullpath
    $file2 = 'C:\script_output\results.csv' # destination's fullpath
    $xl = new-object -c excel.application
    $xl.displayAlerts = $false # don't prompt the user
    $wb2 = $xl.workbooks.open($file1) # open source, if u need readonly $wb2 = $xl.workbooks.open($file1, $null, $true)
    $wb1 = $xl.workbooks.open($file2) # open target
    $sh1_wb1 = $wb1.sheets.item("Sheet1") # second sheet in destination workbook
    $sheetToCopy = $wb2.sheets.item(1) # source sheet to copy
    $sheetToCopy.copy($sh1_wb1) # copy source sheet to destination workbook
    $wb2.close($true) # close source workbook and saving
    $wb1.close($true) # close and save destination workbook
    $xl.quit()

        $strpath = "C:\script_output\results.csv"
        $erroractionpreference = "SilentlyContinue"
        $Exf = New-Object -comobject Excel.Application
        $Exf.visible = $false
        $Exf.DisplayAlerts = $false
        $Exf.AskToUpdateLinks = $false
        $Exf.AlertBeforeOverwriting = $false

        $wb1 = $exf.WorkBooks.Open($strPath)
        $ws1 = $wb1.worksheets.item("Specific_security_update")

        $ws1.activate()
        $ws1.cells.item(1,1) =""
    $dd = $ws1.UsedRange
    $dd.EntireColumn.AutoFit()
    $dd.Interior.ColorIndex = 19
    $dd.Font.ColorIndex = 11
    $aaAutomatic=-4105
    $aaBottom = -4107
    $aaCenter = -4108
    $aaContext = -5002
    $aaContinuous=1
    $aaDiagonalDown=5
    $aaDiagonalUp=6
    $aaEdgeBottom=9
    $aaEdgeLeft=7
    $aaEdgeRight=10
    $aaEdgeTop=8
    $aaInsideHorizontal=12
    $aaInsideVertical=11
    $aaNone=-4142
    $aaAutomatic=-4105
    $aaThin=2
    $aaMedium = -4138
    $aaThick = 4
    [void]$dd.select()
    $dd.Borders.Item($aaEdgeLeft).LineStyle = 1
    $dd.Borders.Item($aaEdgeLeft).ColorIndex = -4105
    $dd.Borders.Item($aaEdgeLeft).Color = 1
    $dd.Borders.Item($aaEdgeLeft).Weight = -4138
    $dd.Borders.Item($aaEdgeTop).LineStyle = 1
    $dd.Borders.Item($aaEdgeBottom).LineStyle = 1
    $dd.Borders.Item($aaEdgeRight).LineStyle = 1
    $dd.Borders.Item($aaInsideVertical).LineStyle = 1
    $dd.Borders.Item($aaInsideHorizontal).LineStyle = 1
    $wb1.save()
    $exf.quit()
    [gc]::collect()
    remove-item $file1
    $d = get-date
    $listBox1.appendtext("$d : completed writing results `r `n")
    $listBox1.appendtext(" `r`n")
       write-host "Completed Security update module"
       ac .\log.txt -Value " "
        ac .\log.txt -Value " ******* MS patch validation module completed"
        ac .\log.txt -Value " "
       #$f2 = resolve-path ".\b.ps1"
       #start-job -filepath $f2 -arg (,$array2)
       #& .\b.ps1 $array2
    }#chumma ends here

       $OKButton.add_Click($chumma)
       $objForm444.ShowDialog() | out-null
     
       #$objForm444.Close()
       
} # function ends here


function specific_reg_key
{

[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")

$objForm = New-Object System.Windows.Forms.Form
$objForm.Text = "Data Entry Form"
$objForm.Size = New-Object System.Drawing.Size(500,200)
$objForm.StartPosition = "CenterScreen"

$objLabel = New-Object System.Windows.Forms.Label
$objLabel.Location = New-Object System.Drawing.Size(10,20)
$objLabel.Size = New-Object System.Drawing.Size(280,20)
$objLabel.Text = "Please enter the Registry Path for 32bit:"
$objForm.Controls.Add($objLabel)

$regpath32 = New-Object System.Windows.Forms.TextBox
$regpath32.Location = New-Object System.Drawing.Size(10,50)
$regpath32.Size = New-Object System.Drawing.Size(460,25)
$objForm.Controls.Add($regpath32)

$value = New-Object System.Windows.Forms.Label
$value.Location = New-Object System.Drawing.Size(10,90)
$value.Size = New-Object System.Drawing.Size(70,20)
$value.Text = "Key Value"
$objForm.Controls.Add($value)

$val32 = New-Object System.Windows.Forms.TextBox
$val32.Location = New-Object System.Drawing.Size(80,90)
$val32.Size = New-Object System.Drawing.Size(100,25)
$objForm.Controls.Add($val32)

$OKButton = New-Object System.Windows.Forms.Button
$OKButton.Location = New-Object System.Drawing.Size(80,130)
$OKButton.Size = New-Object System.Drawing.Size(75,33)
$OKButton.Text = "Submit"
#$OKButton.Add_Click({$x=$regpath32.Text;$objForm.Close()})
$objForm.Controls.Add($OKButton)
$chumma = {
#function begins here
        ac .\log.txt -Value " ******* Registry validation Module begins ********"
        ac .\log.txt -Value " "
        
        $regpath32t = $regpath32.text
        if($regpath32t -ne $null) {ac .\log.txt -Value "User input for 32bit registry path : $regpath32t"}

        $val32t = $val32.text
        if($val32t -ne $null ) {ac .\log.txt -Value "User input for 32bit registry value : $val32t";ac .\log.txt -Value " ";}

        $regpath64t = $regpath64.text
        if($regpath64t -ne $null) {ac .\log.txt -Value "User input for 64bit registry path : $regpath64t"}

        $val64t = $val64.text
        if($val64t -ne $null){ac .\log.txt -Value "User input for 64bit registry value : $val64t";ac .\log.txt -Value " ";}
        
        $array = @()
        $hkLM = "HKEY_LOCAL_MACHINE"
        $hkcu = "HKEY_CURRENT_USER"

        if ( ($regpath32t -like "$hkLM*") -or ($regpath64t -like "$hkLM*"))
                {
                $hkey = 'LocalMachine'
                if($regpath32t -ne $null) {$path = $regpath32t.replace("HKEY_LOCAL_MACHINE\","")}
                
                }

        if ( ($regpath32t -like "$hkcu*") -or ($regpath64t -like "$hkcu*") )
                {
                 $hkey = 'CurrentUser'
                 if($regpath32t -ne $null) {$path = $regpath32t.replace("HKEY_CURRENT_USER\","")}
                 
                }

         $a = import-csv ".\serverlist1.csv"
         foreach ( $pc in $a )
         {
            $d = get-date
            $k = $pc.server
            $listBox1.appendtext("$d : working on $k `r`n")
            ac .\log.txt -Value "Working on $k"
            $listBox1.appendtext(" `r`n")
            if ( Test-Connection -ComputerName $pc.server -Count 1 -erroraction silentlycontinue )
            {
                 ac .\log.txt -Value "$k - server is in network"
                 $reg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine',$pc.server)
                 $regkey = $null
                 $regKey= $reg.OpenSubKey("SYSTEM\CurrentControlSet\Control\Session Manager\Environment")
                 if($regkey -ne $null)
                 {
                 $arch_ser = $regkey.GetValue("PROCESSOR_ARCHITECTURE")
                 $regkey = $null
                 $regKey= $reg.OpenSubKey("$path")
                    if( $regkey -ne $null )
                    {
                       [string]$val = $regkey.GetValue("$val32t")
                        write-host("Found Registry value on"+" "+$pc.server)
                        ac .\log.txt -Value "Found Registry value on $k"
                        ac .\log.txt -Value " "

                        $obj12 = New-Object PSObject
                        $obj12 | Add-Member -MemberType NoteProperty -Name "ComputerName" -Value $k
                        $obj12 | Add-Member -MemberType NoteProperty -Name "Path" -Value $path
                        $obj12 | Add-Member -MemberType NoteProperty -Name "Value" -Value $val
                        $array += $obj12

                     }

                    else
                    {
                         write-host("Key Doesn't Exist on "+" "+$pc.server)
                         $val = "Key Doesn't Exist"
                         ac .\log.txt -Value "Key Doesn't Exist on $k"
                         ac .\log.txt -Value " "
                         $ser = $pc.server
                         $obj12 = New-Object PSObject
                        $obj12 | Add-Member -MemberType NoteProperty -Name "ComputerName" -Value $ser
                        $obj12 | Add-Member -MemberType NoteProperty -Name "Path" -Value $path
                        $obj12 | Add-Member -MemberType NoteProperty -Name "Value" -Value $val
                        $array += $obj12
                    }

                 }
                 else
                 {
                  $val = "Access Denied"
                  $k = $pc.server
                  ac .\log.txt -Value "$k : $val"
                  ac .\log.txt -Value " "
                  $obj12 = New-Object PSObject
                  $obj12 | Add-Member -MemberType NoteProperty -Name "ComputerName" -Value $k
                  $obj12 | Add-Member -MemberType NoteProperty -Name "Path" -Value $path
                  $obj12 | Add-Member -MemberType NoteProperty -Name "Value" -Value $val
                  $array += $obj12
                 }

             } # test connection ends here
               else
                {
                  $val = "Ping test failed"
                  $k = $pc.server
                  ac .\log.txt -Value "$k : $val"
                  ac .\log.txt -Value " "
                  $obj12 = New-Object PSObject
                  $obj12 | Add-Member -MemberType NoteProperty -Name "ComputerName" -Value $k
                  $obj12 | Add-Member -MemberType NoteProperty -Name "Path" -Value $path
                  $obj12 | Add-Member -MemberType NoteProperty -Name "Value" -Value $val
                  $array += $obj12
                }

        } # Ends For each object
    $array2 = @()
    $array2 = $array | where-Object { $_.Path } | select ComputerName, Path, Value
    $array2 | export-csv ".\Specific_reg_key.csv"
    $array2 | ogv

    $file1 = resolve-path ".\Specific_Reg_Key.csv" # source's fullpath
    $file2 = 'C:\script_output\results.csv' # destination's fullpath

    $xl = new-object -c excel.application
    $xl.displayAlerts = $false # don't prompt the user
    $wb2 = $xl.workbooks.open($file1) # open source, if u need readonly $wb2 = $xl.workbooks.open($file1, $null, $true)
    $wb1 = $xl.workbooks.open($file2) # open target
    $sh1_wb1 = $wb1.sheets.item("Sheet1") # second sheet in destination workbook
    $sheetToCopy = $wb2.sheets.item(1) # source sheet to copy
    $sheetToCopy.copy($sh1_wb1) # copy source sheet to destination workbook
    $wb2.close($true) # close source workbook and saving
    $wb1.close($true) # close and save destination workbook
    $xl.quit()

    #=========================

    $strpath = "C:\script_output\results.csv"
    $erroractionpreference = "SilentlyContinue"
    $Exf = New-Object -comobject Excel.Application
    $Exf.visible = $false
    $Exf.DisplayAlerts = $false
    $Exf.AskToUpdateLinks = $false
    $Exf.AlertBeforeOverwriting = $false

    $wb1 = $exf.WorkBooks.Open($strPath)
    $ws1 = $wb1.worksheets.item("Specific_Reg_Key")

    $ws1.activate()
    $ws1.cells.item(1,1) =""

    $dd = $ws1.UsedRange
    $dd.EntireColumn.AutoFit()
    $dd.Interior.ColorIndex = 19
    $dd.Font.ColorIndex = 11
    $aaAutomatic=-4105
    $aaBottom = -4107
    $aaCenter = -4108
    $aaContext = -5002
    $aaContinuous=1
    $aaDiagonalDown=5
    $aaDiagonalUp=6
    $aaEdgeBottom=9
    $aaEdgeLeft=7
    $aaEdgeRight=10
    $aaEdgeTop=8
    $aaInsideHorizontal=12
    $aaInsideVertical=11
    $aaNone=-4142
    $aaAutomatic=-4105
    $aaThin=2
    $aaMedium = -4138
    $aaThick = 4
    [void]$dd.select()
    $dd.Borders.Item($aaEdgeLeft).LineStyle = 1
    $dd.Borders.Item($aaEdgeLeft).ColorIndex = -4105
    $dd.Borders.Item($aaEdgeLeft).Color = 1
    $dd.Borders.Item($aaEdgeLeft).Weight = -4138
    $dd.Borders.Item($aaEdgeTop).LineStyle = 1
    $dd.Borders.Item($aaEdgeBottom).LineStyle = 1
    $dd.Borders.Item($aaEdgeRight).LineStyle = 1
    $dd.Borders.Item($aaInsideVertical).LineStyle = 1
    $dd.Borders.Item($aaInsideHorizontal).LineStyle = 1
    $wb1.save()
    $exf.quit()
    [gc]::collect()
    remove-item $file1
    #=========================

    $listBox1.appendtext("$d : Reg key module completed `r `n")
    ac .\log.txt -Value " "
    ac .\log.txt -Value "Registry validation module completed"
    ac .\log.txt -Value " "
    $listBox1.appendtext(" `r`n")
    $objForm.Close()
}
$OKButton.add_Click($chumma)
$objForm.ShowDialog() | Out-Null
}
    #specific_reg_key
    #===================================================================================
    function scriptexec
    {
    $listBox1.appendtext("********************************************************************** `r`n")
    $listBox1.appendtext("                         Script Execution Begins `r`n")
    $listBox1.appendtext("********************************************************************** `r`n")
    $listBox1.appendtext(" `r`n")
    }


    #===================================================================================

    #====================================================
    $form1 = New-Object System.Windows.Forms.Form
    $form1.Text = "Server_Validation_Tool"
    $form1.Name = "form1"
    $form1.DataBindings.DefaultDataSourceUpdateMode = 0
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 320
    $System_Drawing_Size.Height = 600
    $form1.ClientSize = $System_Drawing_Size
    $form1.StartPosition = "CenterScreen"
    #=====================================================
    $button1 = New-Object System.Windows.Forms.Button
    $button1.TabIndex = 9
    $button1.Name = "button1"
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 125
    $System_Drawing_Size.Height = 33
    $button1.Size = $System_Drawing_Size
    $button1.UseVisualStyleBackColor = $True
    $button1.Text = "Run Script"
    $System_Drawing_Point = New-Object System.Drawing.Point
    $System_Drawing_Point.X = 27
    $System_Drawing_Point.Y = 535
    $button1.Location = $System_Drawing_Point
    $button1.DataBindings.DefaultDataSourceUpdateMode = 0
    #=====================================================
    $button2 = New-Object System.Windows.Forms.Button
    $button2.TabIndex = 9
    $button2.Name = "button2"
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 125
    $System_Drawing_Size.Height = 33
    $button2.Size = $System_Drawing_Size
    $button2.UseVisualStyleBackColor = $True
    $button2.Text = "Quit"
    $System_Drawing_Point = New-Object System.Drawing.Point
    $System_Drawing_Point.X = 160
    $System_Drawing_Point.Y = 535
    $button2.Location = $System_Drawing_Point
    $button2.DataBindings.DefaultDataSourceUpdateMode = 0
    $form1.Controls.Add($button2)
    #=====================================================
    $buttonPress_Click={
        $this.Enabled = $False

         # Simulate work
         $dg=$null
         $dstart = get-date
         ac .\log.txt -Value "***** Script Execution begins at $dstart *******"
         ac .\log.txt -Value " "
    
        xlstocsv;CSVcreation;
        if ($checkbox1.checked ) { Specific_uninstallkey }
        if ($checkbox2.checked ) { Specific_security_update }
        if ($checkbox3.checked ) { specific_reg_key }

        # Process the pending messages before enabling the button
        csvdeletion
        $dend = get-date
        ac .\log.txt -Value "Script Execution completed at $dend"
        ac .\log.txt -Value " "
        
        $exec_time = NEW-TIMESPAN –Start $dstart –End $dend 
        $Tmin = $exec_time.minutes
        $Tsec = $exec_time.seconds
        if($Tmin -gt 1)
        {
            if($Tsec -gt 1)
            {
            ac .\log.txt -Value "***** Script execution Time : $Tmin Minutes $Tsec seconds"
            ac .\log.txt -Value " "
            }
            else
            {
            ac .\log.txt -Value "***** Script execution Time : $Tmin Minutes $Tsec second"
            ac .\log.txt -Value " "
            }
        }
        else
        {
            if($Tsec -gt 1)
            {
            ac .\log.txt -Value "***** Script execution Time : $Tmin Minute $Tsec seconds"
            ac .\log.txt -Value " "
            }
            else
            {
            ac .\log.txt -Value "***** Script execution Time : $Tmin Minutes $Tsec second"
            ac .\log.txt -Value " "
            }
        }
        $this.Enabled = $True

    }
    $buttonPress_Click_2 = { $form1.close() }
    #=====================================================
    $button1.add_Click($buttonPress_Click)
    $button2.add_Click($buttonPress_Click_2)
    $form1.Controls.Add($button1)
    $Form1.MinimizeBox = $False
    $Form1.MaximizeBox = $False
    #=====================================================
    $lbl=New-Object System.Windows.Forms.Label
    $lblfont = New-Object System.Drawing.Font("Times New Roman",14,[System.Drawing.FontStyle]::BOLD)
     $lbl.Location=New-Object System.Drawing.Point( 60, 20 )
     $lbl.font = $lblfont
         $lbl.Name = "label1"
         $lbl.Size=New-Object System.Drawing.Size( 325, 20 )
         $lbl.TabIndex=0
         $lbl.Text="Server Validation Tool"
    $form1.controls.add($lbl)
    #=====================================================

        $checkBox1 = New-Object System.Windows.Forms.checkbox
        $checkBox1.UseVisualStyleBackColor = $True
        $System_Drawing_Size = New-Object System.Drawing.Size
        $System_Drawing_Size.Width = 210
        $System_Drawing_Size.Height = 29
        $checkBox1.Size = $System_Drawing_Size
        $checkBox1.TabIndex = 1
        $checkBox1.Text = "To find the Specific Software status"
        $System_Drawing_Point = New-Object System.Drawing.Point
        $System_Drawing_Point.X = 20
        $System_Drawing_Point.Y = 50
        $checkBox1.Location = $System_Drawing_Point
        $checkBox1.Name = "ServerInfo1"
        $form1.Controls.Add($checkBox1)
    #==========================================================
    $checkBox2 = New-Object System.Windows.Forms.checkbox
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 210
    $System_Drawing_Size.Height = 29
    $checkBox2.Size = $System_Drawing_Size
    $checkBox2.TabIndex = 2
    $checkBox2.Text = "To Find Specific MS security update"
    $System_Drawing_Point = New-Object System.Drawing.Point
    $System_Drawing_Point.X = 20
    $System_Drawing_Point.Y = 75
    $checkBox2.Location = $System_Drawing_Point
    #$checkBox2.DataBindings.DefaultDataSourceUpdateMode = 0
    $checkBox2.Name = "Driveinfo"
    $form1.Controls.Add($checkBox2)
    #==========================================================
    $checkBox3 = New-Object System.Windows.Forms.checkbox
    $checkBox3.UseVisualStyleBackColor = $True
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 260
    $System_Drawing_Size.Height = 29
    $checkBox3.Size = $System_Drawing_Size
    $checkBox3.TabIndex = 3
    $checkBox3.Text = "To Find the registry key value"
    $System_Drawing_Point = New-Object System.Drawing.Point
    $System_Drawing_Point.X = 20
    $System_Drawing_Point.Y = 100
    $checkBox3.Location = $System_Drawing_Point
    #$checkBox3.DataBindings.DefaultDataSourceUpdateMode = 0
    $checkBox3.Name = "UninstallKey"
    $form1.Controls.Add($checkBox3)
    #==========================================================
    $form1.Controls.Add($listBox1)
    #==========================================================
    # $Form1.Controls.Add($pbrTest)
    cls
    #==========================================================
    $i = 0
    $pbrTest.Value = $i
    cls
    #==========================================================
    $form1.ShowDialog()| Out-Null

    #=====================================================
