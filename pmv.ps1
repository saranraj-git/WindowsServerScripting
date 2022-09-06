    $form1 = New-Object System.Windows.Forms.Form
    $form1.Text = "Server Migration Validator"
    $form1.Name = "form1"
    $form1.DataBindings.DefaultDataSourceUpdateMode = 0
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 385
    $System_Drawing_Size.Height = 600
    $form1.ClientSize = $System_Drawing_Size
    $form1.StartPosition = "CenterScreen"  

    $lbl=New-Object System.Windows.Forms.Label
    $lblfont = New-Object System.Drawing.Font("Times New Roman",14,[System.Drawing.FontStyle]::BOLD)
    $lbl.Location=New-Object System.Drawing.Point( 80, 20 )
    $lbl.font = $lblfont
    $lbl.Name = "label1"
    $lbl.Size=New-Object System.Drawing.Size( 325, 20 )
    $lbl.TabIndex=0
    $lbl.Text="Server Migration Validator"
    $form1.controls.add($lbl)

    $lbl1=New-Object System.Windows.Forms.Label
    $lblfont1 = New-Object System.Drawing.Font("Times New Roman",11,[System.Drawing.FontStyle]::BOLD)
    $lbl1.Location=New-Object System.Drawing.Point( 40, 60 )
    $lbl1.font = $lblfont1
    $lbl1.Name = "label11"
    $lbl1.Size=New-Object System.Drawing.Size( 100, 20 )
    $lbl1.TabIndex=0
    $lbl1.Text="ServerNames"
    $form1.controls.add($lbl1)

    $lbl2=New-Object System.Windows.Forms.Label
    $lblfont2 = New-Object System.Drawing.Font("Times New Roman",11,[System.Drawing.FontStyle]::BOLD)
    $lbl2.Location=New-Object System.Drawing.Point( 200, 60 )
    $lbl2.font = $lblfont2
    $lbl2.Name = "l"
    $lbl2.Size=New-Object System.Drawing.Size( 325, 20 )
    $lbl2.TabIndex=0
    $lbl2.Text="Features to Validate:"
    $form1.controls.add($lbl2)

    $flbl1=New-Object System.Windows.Forms.Label
    $flblfont1 = New-Object System.Drawing.Font("Times New Roman",9)
    $flbl1.Location=New-Object System.Drawing.Point( 230, 90 )
    $flbl1.font = $flblfont1
    $flbl1.Name = "l"
    $flbl1.Size=New-Object System.Drawing.Size( 325, 20 )
    $flbl1.TabIndex=0
    $flbl1.Text="Server FQDN / IP Check"
    $form1.controls.add($flbl1)

    $checkBox1 = New-Object System.Windows.Forms.checkbox
    $checkBox1.UseVisualStyleBackColor = $True
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 20
    $System_Drawing_Size.Height = 20
    $checkBox1.Size = $System_Drawing_Size
    $checkBox1.TabIndex = 1
    $checkBox1.Text = "To find the Specific Software status"
    $System_Drawing_Point = New-Object System.Drawing.Point
    $System_Drawing_Point.X = 210
    $System_Drawing_Point.Y = 90
    $checkBox1.Location = $System_Drawing_Point
    $checkBox1.enabled = $false
    $checkBox1.checked = $true
    $form1.Controls.Add($checkBox1)

    $flbl2=New-Object System.Windows.Forms.Label
    $flblfont2 = New-Object System.Drawing.Font("Times New Roman",9)
    $flbl2.Location=New-Object System.Drawing.Point( 230, 110 )
    $flbl2.font = $flblfont2
    $flbl2.Name = "l"
    $flbl2.Size=New-Object System.Drawing.Size( 325, 20 )
    $flbl2.TabIndex=0
    $flbl2.Text="RAM / CPU / DRIVES"
    $form1.controls.add($flbl2)

    $checkBox2 = New-Object System.Windows.Forms.checkbox
    $checkBox2.UseVisualStyleBackColor = $True
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 20
    $System_Drawing_Size.Height = 20
    $checkBox2.Size = $System_Drawing_Size
    $checkBox2.TabIndex = 1
    $checkBox2.Text = "To find the Specific Software status"
    $System_Drawing_Point = New-Object System.Drawing.Point
    $System_Drawing_Point.X = 210
    $System_Drawing_Point.Y = 110
    $checkBox2.Location = $System_Drawing_Point
    $checkBox2.Name = "ServerInfo1"
    $form1.Controls.Add($checkBox2)

    $flbl3=New-Object System.Windows.Forms.Label
    $flblfont3 = New-Object System.Drawing.Font("Times New Roman",9)
    $flbl3.Location=New-Object System.Drawing.Point( 230, 130 )
    $flbl3.font = $flblfont3
    $flbl3.Name = "l"
    $flbl3.Size=New-Object System.Drawing.Size( 325, 20 )
    $flbl3.TabIndex=0
    $flbl3.Text="IIS Details"
    $form1.controls.add($flbl3)

    $checkBox3 = New-Object System.Windows.Forms.checkbox
    $checkBox3.UseVisualStyleBackColor = $True
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 20
    $System_Drawing_Size.Height = 20
    $checkBox3.Size = $System_Drawing_Size
    $checkBox3.TabIndex = 1
    $checkBox3.Text = "To find the Specific Software status"
    $System_Drawing_Point = New-Object System.Drawing.Point
    $System_Drawing_Point.X = 210
    $System_Drawing_Point.Y = 130
    $checkBox3.Location = $System_Drawing_Point
    $checkBox3.Name = "ServerInfo1"
    $form1.Controls.Add($checkBox3)

    $flbl4=New-Object System.Windows.Forms.Label
    $flblfont4 = New-Object System.Drawing.Font("Times New Roman",9)
    $flbl4.Location=New-Object System.Drawing.Point( 230, 150)
    $flbl4.font = $flblfont1
    $flbl4.Name = "l"
    $flbl4.Size=New-Object System.Drawing.Size( 325, 20 )
    $flbl4.TabIndex=0
    $flbl4.Text="Cluster Details"
    $form1.controls.add($flbl4)

    $checkBox4 = New-Object System.Windows.Forms.checkbox
    $checkBox4.UseVisualStyleBackColor = $True
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 20
    $System_Drawing_Size.Height = 20
    $checkBox4.Size = $System_Drawing_Size
    $checkBox4.TabIndex = 1
    $checkBox4.Text = "To find the Specific Software status"
    $System_Drawing_Point = New-Object System.Drawing.Point
    $System_Drawing_Point.X = 210
    $System_Drawing_Point.Y = 150
    $checkBox4.Location = $System_Drawing_Point
    $checkBox4.Name = "ServerInfo1"
    $form1.Controls.Add($checkBox4)

    $flbl5=New-Object System.Windows.Forms.Label
    $flblfont5 = New-Object System.Drawing.Font("Times New Roman",9)
    $flbl5.Location=New-Object System.Drawing.Point( 230, 170 )
    $flbl5.font = $flblfont1
    $flbl5.Name = "l"
    $flbl5.Size=New-Object System.Drawing.Size( 325, 20 )
    $flbl5.TabIndex=0
    $flbl5.Text="HostFile Content"
    $form1.controls.add($flbl5)

    $checkBox5 = New-Object System.Windows.Forms.checkbox
    $checkBox5.UseVisualStyleBackColor = $True
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 20
    $System_Drawing_Size.Height = 20
    $checkBox5.Size = $System_Drawing_Size
    $checkBox5.TabIndex = 1
    $checkBox5.Text = "To find the Specific Software status"
    $System_Drawing_Point = New-Object System.Drawing.Point
    $System_Drawing_Point.X = 210
    $System_Drawing_Point.Y = 170
    $checkBox5.Location = $System_Drawing_Point
    $checkBox5.Name = "ServerInfo1"
    $form1.Controls.Add($checkBox5)

    $flbl6=New-Object System.Windows.Forms.Label
    $flblfont6 = New-Object System.Drawing.Font("Times New Roman",9)
    $flbl6.Location=New-Object System.Drawing.Point( 230, 190 )
    $flbl6.font = $flblfont1
    $flbl6.Name = "l"
    $flbl6.Size=New-Object System.Drawing.Size( 325, 20 )
    $flbl6.TabIndex=0
    $flbl6.Text="Env Var:PATH backup"
    $form1.controls.add($flbl6)

    $checkBox6 = New-Object System.Windows.Forms.checkbox
    $checkBox6.UseVisualStyleBackColor = $True
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 20
    $System_Drawing_Size.Height = 20
    $checkBox6.Size = $System_Drawing_Size
    $checkBox6.TabIndex = 1
    $checkBox6.Text = "To find the Specific Software status"
    $System_Drawing_Point = New-Object System.Drawing.Point
    $System_Drawing_Point.X = 210
    $System_Drawing_Point.Y = 190
    $checkBox6.Location = $System_Drawing_Point
    $form1.Controls.Add($checkBox6)

    $flbl7=New-Object System.Windows.Forms.Label
    $flblfont7 = New-Object System.Drawing.Font("Times New Roman",9)
    $flbl7.Location=New-Object System.Drawing.Point( 230, 210 )
    $flbl7.font = $flblfont7
    $flbl7.Name = "l"
    $flbl7.Size=New-Object System.Drawing.Size( 325, 20 )
    $flbl7.TabIndex=0
    $flbl7.Text="Listening Ports check"
    $form1.controls.add($flbl7)

    $checkBox7 = New-Object System.Windows.Forms.checkbox
    $checkBox7.UseVisualStyleBackColor = $True
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 20
    $System_Drawing_Size.Height = 20
    $checkBox7.Size = $System_Drawing_Size
    $checkBox7.TabIndex = 1
    $checkBox7.Text = "To find the Specific Software status"
    $System_Drawing_Point = New-Object System.Drawing.Point
    $System_Drawing_Point.X = 210
    $System_Drawing_Point.Y = 210
    $checkBox7.Location = $System_Drawing_Point
    $checkBox7.Name = "ServerInfo1"
    $form1.Controls.Add($checkBox7)

    $flbl8=New-Object System.Windows.Forms.Label
    $flblfont8 = New-Object System.Drawing.Font("Times New Roman",9)
    $flbl8.Location=New-Object System.Drawing.Point( 230, 230 )
    $flbl8.font = $flblfont8
    $flbl8.Name = "l"
    $flbl8.Size=New-Object System.Drawing.Size( 325, 20 )
    $flbl8.TabIndex=0
    $flbl8.Text="Installed Products check"
    $form1.controls.add($flbl8)

    $checkBox8 = New-Object System.Windows.Forms.checkbox
    $checkBox8.UseVisualStyleBackColor = $True
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 20
    $System_Drawing_Size.Height = 20
    $checkBox8.Size = $System_Drawing_Size
    $checkBox8.TabIndex = 1
    $checkBox8.Text = "To find the Specific Software status"
    $System_Drawing_Point = New-Object System.Drawing.Point
    $System_Drawing_Point.X = 210
    $System_Drawing_Point.Y = 230
    $checkBox8.Location = $System_Drawing_Point
    $form1.Controls.Add($checkBox8)

    $flbl9=New-Object System.Windows.Forms.Label
    $flblfont9 = New-Object System.Drawing.Font("Times New Roman",9)
    $flbl9.Location=New-Object System.Drawing.Point( 230, 250 )
    $flbl9.font = $flblfont9
    $flbl9.Name = "l"
    $flbl9.Size=New-Object System.Drawing.Size( 325, 20 )
    $flbl9.TabIndex=0
    $flbl9.Text="Server Band Build Info"
    $form1.controls.add($flbl9)

    $checkBox9 = New-Object System.Windows.Forms.checkbox
    $checkBox9.UseVisualStyleBackColor = $True
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 20
    $System_Drawing_Size.Height = 20
    $checkBox9.Size = $System_Drawing_Size
    $checkBox9.TabIndex = 1
    $checkBox9.Text = "To find the Specific Software status"
    $System_Drawing_Point = New-Object System.Drawing.Point
    $System_Drawing_Point.X = 210
    $System_Drawing_Point.Y = 250
    $checkBox9.Location = $System_Drawing_Point
    $checkBox9.Name = "ServerInfo1"
    $form1.Controls.Add($checkBox9)

    $flbl10=New-Object System.Windows.Forms.Label
    $flblfont10 = New-Object System.Drawing.Font("Times New Roman",9)
    $flbl10.Location=New-Object System.Drawing.Point( 230, 270 )
    $flbl10.font = $flblfont10
    $flbl10.Name = "l"
    $flbl10.Size=New-Object System.Drawing.Size( 325, 20 )
    $flbl10.TabIndex=0
    $flbl10.Text="CA scalability server"
    $form1.controls.add($flbl10)

    $checkBox10 = New-Object System.Windows.Forms.checkbox
    $checkBox10.UseVisualStyleBackColor = $True
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 20
    $System_Drawing_Size.Height = 20
    $checkBox10.Size = $System_Drawing_Size
    $checkBox10.TabIndex = 1
    $checkBox10.Text = "To find the Specific Software status"
    $System_Drawing_Point = New-Object System.Drawing.Point
    $System_Drawing_Point.X = 210
    $System_Drawing_Point.Y = 270
    $checkBox10.Location = $System_Drawing_Point
    $checkBox10.Name = "ServerInfo1"
    $form1.Controls.Add($checkBox10)
    
    $listBox1 = New-Object System.Windows.Forms.TextBox
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 170
    $System_Drawing_Size.Height = 195
    $listBox1.Size = $System_Drawing_Size
    $listBox1.DataBindings.DefaultDataSourceUpdateMode = 0
    $listBox1.Name = "listBox1"
    $System_Drawing_Point = New-Object System.Drawing.Point
    $System_Drawing_Point.X = 9
    $System_Drawing_Point.Y = 90
    $listBox1.Location = $System_Drawing_Point
    $listBox1.TabIndex = 8
    $listBox1.MultiLine = $True
    $listBox1.ScrollBars = "Vertical"
    $form1.controls.add($listbox1)

    $lbl3=New-Object System.Windows.Forms.Label
    $lblfont3 = New-Object System.Drawing.Font("Times New Roman",11,[System.Drawing.FontStyle]::BOLD)
    $lbl3.Location=New-Object System.Drawing.Point( 140, 300 )
    $lbl3.font = $lblfont1
    $lbl3.Name = "label11"
    $lbl3.Size=New-Object System.Drawing.Size( 120, 20 )
    $lbl3.TabIndex=0
    $lbl3.Text="Output Console"
    $form1.controls.add($lbl3)

    $Outbox = New-Object System.Windows.Forms.TextBox
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 370
    $System_Drawing_Size.Height = 195
    $Outbox.Size = $System_Drawing_Size
    $Outbox.DataBindings.DefaultDataSourceUpdateMode = 0
    $Outbox.Name = "listBox1"
    $System_Drawing_Point = New-Object System.Drawing.Point
    $System_Drawing_Point.X = 9
    $System_Drawing_Point.Y = 330
    $Outbox.Location = $System_Drawing_Point
    $Outbox.TabIndex = 8
    $Outbox.MultiLine = $True
    $Outbox.ScrollBars = "Vertical"
    $Outbox.readonly = $true
    $form1.controls.add($Outbox)
    
    $copyoutput = New-Object System.Windows.Forms.Button
    $copyoutput.Location = New-Object System.Drawing.Size(10,550)
    $copyoutput.Size = New-Object System.Drawing.Size(85,23)
    $copyoutput.Text = "Copy Output"
    $form1.Controls.Add($copyoutput)

    $OKButton = New-Object System.Windows.Forms.Button
    $OKButton.Location = New-Object System.Drawing.Size(160,550)
    $OKButton.Size = New-Object System.Drawing.Size(75,23)
    $OKButton.Text = "Run Script"
    $form1.Controls.Add($OKButton)
    
    $quitbutton = New-Object System.Windows.Forms.Button
    $quitbutton.Location = New-Object System.Drawing.Size(295,550)
    $quitbutton.Size = New-Object System.Drawing.Size(85,23)
    $quitbutton.Text = "Quit"
    $form1.Controls.Add($quitbutton)
    
    $copyoutputevent = {$Outbox.Text.Split("`n") | clip}
    $copyoutput.add_Click($copyoutputevent)

    $runscriptevent={
    $OKButton.Text = "Please wait"
    $quitbutton.enabled = $false
    $copyoutput.enabled = $false
    $this.enabled =$false
    
    $allserver = $listBox1.text.split("`n")
    
    Import-module servermanager 
    $aryDNSSuffixes = "pnn-p01.chp.bankofamerica.com","sdi.corp.bankofamerica.com","corp.bankofamerica.com","amrs.win.ml.com","bankofamerica.com"
    invoke-wmimethod -Class win32_networkadapterconfiguration -Name setDNSSuffixSearchOrder -ComputerName localhost -ArgumentList @($aryDNSSuffixes), $null
    
        if($allserver -ne "")
        {
        $cred = (Get-Credential)
            foreach($ser in $allserver)
            {
                $ser = $ser.Trim()
                #NSLookup started
                $ser = [system.net.dns]::resolve($ser) | select -ExpandProperty Hostname
                $outbox.AppendText("------------Script Execution started----------- `n")
                If(Test-Connection -ComputerName $ser -Count 1) 
                {
                    $outbox.AppendText("Server Name : $ser `n")
                    $IPaddr = [system.net.dns]::resolve($ser) | select -ExpandProperty Addresslist | select -ExpandProperty IPAddressToString
                    $outbox.AppendText("IPAddress : $IPAddr `n")
        
                    #access check
                    try 
                    {
                    $reg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine',$ser)
                    $regKey= $reg.OpenSubKey("SOFTWARE\Microsoft\Windows NT\CurrentVersion") 
                    $UninstallKey="SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Uninstall"
                    $regkey = $reg.OpenSubKey($UninstallKey)
                    $outbox.AppendText("$ser - Access allowed `n")
                    $flag = 1
                    }
                    catch 
                    {
                        $outbox.AppendText("$ser - Access denied `n")
                        $flag = 0
                    }
                    If($flag -eq 1)
                    {
                        
                        if($checkbox2.checked)
                        {
                            #RAM:
                            $outbox.AppendText("`n#RAM/CPU/Drive`n`n")
	                        $ram = (Get-WmiObject Win32_PhysicalMemory -computername $ser | select -ExpandProperty Capacity)/(1073741824)
                            $outbox.AppendText("RAM config : $ram GB`n")
	        
                            #CPU:
                            $cor=$null;
	                        $cores = Get-WmiObject -class win32_processor -computername $ser | select -ExpandProperty numberofcores
	                        foreach($cr in $cores){$cor = $cor + $cr};
                            $outbox.AppendText("`nNo. of CPU CORE : $cor`n")
                            #DRIVE:
	                        $d = GET-WMIOBJECT  -computername $ser –query “SELECT * from win32_logicaldisk where DriveType = 3” | select DeviceID,Size,FreeSpace
                            $dcount = $d.count	
                            $outbox.AppendText("`nDrive count : $dcount`n")
                            foreach ($drive in $d)
	                        {
	                        $drivename = $drive.DeviceID
	                        $size = "{0:N2}" -f (($drive.Size)/1073741824)
	                        $free = "{0:N2}" -f (($drive.FreeSpace)/1073741824)
    	                    $outbox.AppendText("`n$drivename drive - Total size $size GB with Free space of $free GB `n")
	                        }
                        }
                        if($checkbox3.checked)
                        {
                            #Windows Feature Installed
                            $outbox.AppendText("`n#IIS Details`n")
                            $winfeatures = invoke-command -ComputerName $ser -Credential $cred -sc {
                            $winF = (Get-WindowsFeature | Where Installed) | select -ExpandProperty Name;
                            if($winf.Contains("Web-Server"))
                            {
                
                                $outbox.AppendText("`nIIS configured : YES`n")
                                #IIS module begins
                                import-module webadministration
                                #get-childitem iis:AppPools
                                $sites = get-childitem iis:sites
                                $allsitenames = $sites.name
                                $htt="http"
                                
                                foreach ($site in $sites)
                                {
                                $sname = $site.name
                                $sstatus = $site.state
                                $sapppool = $site.applicationPool
                                
                                write-output "Site Name : $sname`n"
                                write-output "Site Status : $sstatus`n"
                                write-output "App Pool Name : $sapppool`n"
                                #$port = $site | select -ExpandProperty Bindings
                                
                                $pc = $site.Bindings.collection.protocol.Count
                                if($pc -eq 1)
                                {
                                            if($site.Bindings.collection.protocol -like "$htt*")
                                            {
                                                #$csite = $site.Bindings.collection.protocol
                                                $port = $site.Bindings.collection.bindingInformation
                                                $port = $port.replace("*","")
                                                $port = $port.replace(":","")
                                                write-output "`nPort used : $port"
                                                
                                            }

                                 }
                                 elseif($pc -gt 1)
                                 {
                                        for($i=0;$i-lt$pc;$i++)
                                        {
                                            if($site.Bindings.collection.protocol[$i] -like "$htt*")
                                            {
                                                #$csite = $site.Bindings.collection.protocol
                                                $port = $site.Bindings.collection.bindingInformation[$i]
                                                $port = $port.replace("*","")
                                                $port = $port.replace(":","")
                                                write-output "`nPort used : $port"
                                                
                                            }
                                         }
                                 }
                                write-output "----------------------------"
                                #Cert associated check
                                $localcert = dir cert:\localmachine\my 
                                $localcerttp = $localcert | select -ExpandProperty Thumbprint
                                $localcertname = $localcert | select  -ExpandProperty subject 

                                foreach($t in $localcert)
                                    {
                                    $certbinding = get-childitem iis:sslbindings
    
                                    foreach ($c in $certbinding)
                                        {
                                        if($c.Thumbprint -like $t.Thumbprint)
                                        { 
                                            $sitname = $c.Sites.value
                                            $scertname = $t.Subject -split ","
                                            $scertname = $scertname[0].replace("CN=","")
                                            if($sitname -ne $null)
                                            {
                                                if($sitname -like $sname)
                                                {
                                                write-output "Cert Associated  : $scertname`n"
                                                $cexp = $t.NotAfter
                                                write-output "Cert Expiry Date : $cexp`n"
                                                write-output "----------------------------"
                                                }
                                            }
                                        }
                                    }

                                    }
                
                                }#IIS module ends
                            }
                            else{write-output "IIS configured `t : NO `n" }
                            }
                            foreach($line in $winfeatures)
                            {
                            $outbox.AppendText("$line `n")
                            }
                        }
                        if($checkbox4.checked)
                        {
                             #CLUSTER CONFIG
                            $winfeatures = invoke-command -ComputerName $ser -Credential $cred -sc {
                            $winF = (Get-WindowsFeature | Where Installed) | select -ExpandProperty Name;
   		                    if($winf.Contains("RSAT-Clustering"))
                            {
                                write-output "#Failover Cluster Configured `n"
                                $g = Get-ClusterGroup | Get-ClusterResource 
                                 write-output "OWNERGROUP `t `t `t STATE `t `t NAME `t `n"
                                foreach($info in $g)
                                {
                                       $Clname = $info.name
                                       $ClOwnGRP = $info.OwnerGroup.name
                                       $clState = $info.state
                                       $clType = $info.Resourcetype.name
                                       write-output "$clowngrp`t `t $clstate `t `t $CLNAME `n"
                                }
                                write-output "----------------------------------"
                            }
                            else{write-output "Cluster not Configured `n `n";write-output "----------------------------------";}
                            }
                            foreach($line in $winfeatures)
                            {
                            $outbox.AppendText("$line `n")
                            }
                        }
                        if($checkbox5.checked)
                        {
                            #host file contents
                            $winfeatures = invoke-command -ComputerName $ser -Credential $cred -sc {
 	                        $hst = gc C:\Windows\System32\drivers\etc\hosts
                            write-output "----------------------------------"
                            write-output "#HostFile Contents: `n"
                            foreach($line in $hst)
                            {write-output "$line `n"}
                            }
                            foreach($line in $winfeatures)
                            {
                            $outbox.AppendText("$line `n")
                            }
                            $outbox.AppendText("HostFile-EOF-----: `n")
                        }
                        if($checkbox6.checked)
                        {
                            #Env variable
                            $winfeatures = invoke-command -ComputerName $ser -Credential $cred -sc {
                            $envar = $Env:Path -split ";"
                            $ecount =$envar.count
                            write-output "----------------------------------"
                            write-output "#Env var as follows (Total vars = $ecount) `n"
                            foreach($line in $envar){write-output "$line `n"}
                             }
                            foreach($line in $winfeatures)
                            {
                            $outbox.AppendText("$line `n")
                            }
                        }
                        if($checkbox7.checked)
                        {
                            #listening port
                            $winfeatures = invoke-command -ComputerName $ser -Credential $cred -sc {
                            Function Get-Listeningport 
                                                                                                                                                                                                                                                                                                                                                                                                            {            
	            [cmdletbinding()]            
	            param()           
	            
	            try 
                {            
	                $TCPProperties = [System.Net.NetworkInformation.IPGlobalProperties]::GetIPGlobalProperties()            
	                $Connections = $TCPProperties.GetActiveTcpListeners()            
	                foreach($Connection in $Connections) 
                    {            
	                    if($Connection.address.AddressFamily -eq "InterNetwork" ) { $IPType = "IPv4" }  else { $IPType = "IPv6" }           
	                    $OutputObj = New-Object -TypeName PSobject    
	                    $OutputObj | Add-Member -MemberType NoteProperty -Name "IPV4Or6" -Value $IPType          
	                    $OutputObj | Add-Member -MemberType NoteProperty -Name "ListeningPort" -Value $Connection.Port            
	                    if($IPType -notlike "IPv6"){$OutputObj | select -ExpandProperty ListeningPort}
	                }            
	            
	            } 
                catch 
                {            
	                $outbox.AppendText("Failed to get listening connections. $_ `n")
	            }           
	        }
                            $ports = Get-Listeningport
                            write-output "----------------------------------"
                            write-output "#Listening ports are: `n"
                            foreach($line in $ports){write-output "$line `n"}
                            }
                            foreach($line in $winfeatures)
                            {
                            $outbox.AppendText("$line `n")
                            }
                        }
                        if($checkbox8.checked)
                        {
                             #Installed products
                            $winfeatures = invoke-command -ComputerName $ser -Credential $cred -sc {
                            $Allproducts =	Get-WmiObject -class win32_product | select Name,Version | Sort-Object Name
                            $allc = $Allproducts.count
            
                            write-output "Total no of products installed : $allc `n"
            
                            foreach($prod in $allproducts){
                            $name = $prod | select -ExpandProperty name
                            $ver = $prod| select -ExpandProperty version
                            write-output "$name - $ver `n"
                            }
                            }#winfeatures completed
                            foreach($line in $winfeatures)
                            {
                            $outbox.AppendText("$line `n")
                            }

                        }
                        if($checkbox9.checked)
                        {
                            #Band and build type
                            $reg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine',$ser)
	                        $regKey= $reg.OpenSubKey("SOFTWARE\Bank of America\AutoInstall")
	                        $Bandinfo= $regkey.GetValue("Version")
	                        $BuildType = $regkey.GetValue("Build Type")
                            $outbox.AppendText("--------------------------------- `n")
                            $outbox.AppendText("Band Build : $Bandinfo `n")
                            $outbox.AppendText("Server buildtype : $buildtype `n")
                            $outbox.AppendText("--------------------------------- `n")
                        }
                        if($checkbox10.checked)
                        {
                            #CA REG info
                            $winfeatures = invoke-command -ComputerName $ser -Credential $cred -sc {
   	                        $ca = caf setserveraddress
	                        $cas = $ca[0].Replace("Caf currently registers with the scalability server at ","")
                            write-output "`n CA SERVER : $cas `n"
                            }
                            $outbox.AppendText("$winfeatures `n")
                            
                        }
                    
        
        
   	    $outbox.AppendText("--------------------------------------------- `n")
     }

                }#test connectionif loop ends here
                else
                {
                    $outbox.AppendText("$ser is out of network `n")
                }
              
            }#foreach $ser ends here
        }
        else
        {
            $outbox.AppendText("Please enter server name `n")
        }
        $OKButton.Text = "Run Script"
        $quitbutton.enabled = $true
    $copyoutput.enabled = $true
    $this.enabled =$true
    }#runscript event ends here

    $OKButton.add_Click($runscriptevent)
    $quitevent = {$form1.close()}
    $quitbutton.add_Click($quitevent)

    $form1.ShowDialog()| Out-Null 
