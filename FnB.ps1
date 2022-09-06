    [reflection.assembly]::loadwithpartialname("System.Windows.Forms") | Out-Null
    [reflection.assembly]::loadwithpartialname("System.Drawing") | Out-Null
    $form1 = New-Object System.Windows.Forms.Form
    $form1.Text = "CA Fast Deployer"
    $form1.Name = "form1"
    $form1.DataBindings.DefaultDataSourceUpdateMode = 0
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 750
    $System_Drawing_Size.Height = 600
    $form1.ClientSize = $System_Drawing_Size
    $form1.StartPosition = "CenterScreen"

    $lbl=New-Object System.Windows.Forms.Label
    $lblfont = New-Object System.Drawing.Font("Times New Roman",16,[System.Drawing.FontStyle]::BOLD)
    $lbl.Location=New-Object System.Drawing.Point( 270, 15 )
    $lbl.font = $lblfont
    $lbl.Name = "label1"
    $lbl.Size=New-Object System.Drawing.Size( 325, 30 )
    $lbl.TabIndex=0
    $lbl.Text="Fast Deployer"
    $form1.controls.add($lbl)

    $lbl1=New-Object System.Windows.Forms.Label
    $lblfont1 = New-Object System.Drawing.Font("Times New Roman",11)
    $lbl1.Location=New-Object System.Drawing.Point( 240, 50 )
    $lbl1.font = $lblfont1
    $lbl1.Name = "label1"
    $lbl1.Size=New-Object System.Drawing.Size( 150, 20 )
    $lbl1.TabIndex=0
    $lbl1.Text="Enter the Server names"
    $form1.controls.add($lbl1)
    


    $Outbox = New-Object System.Windows.Forms.TextBox
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 150
    $System_Drawing_Size.Height = 200
    $Outbox.Size = $System_Drawing_Size
    $Outbox.DataBindings.DefaultDataSourceUpdateMode = 0
    $Outbox.Name = "listBox1"
    $System_Drawing_Point = New-Object System.Drawing.Point(240,70)
    $Outbox.Location = $System_Drawing_Point
    $Outbox.TabIndex = 8
    $Outbox.MultiLine = $True
    $Outbox.ScrollBars = "Vertical"
    #$Outbox.readonly = $true
    $form1.controls.add($Outbox)

    $lbl2=New-Object System.Windows.Forms.Label
    $lblfont2 = New-Object System.Drawing.Font("Times New Roman",11)
    $lbl2.Location=New-Object System.Drawing.Point( 15, 50 )
    $lbl2.font = $lblfont1
    $lbl2.Name = "label1"
    $lbl2.Size=New-Object System.Drawing.Size( 325, 20 )
    $lbl2.TabIndex=0
    $lbl2.Text="CA ServerName"
    $form1.controls.add($lbl2)

    $Outbox1 = New-Object System.Windows.Forms.TextBox
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 170
    $System_Drawing_Size.Height = 10
    $Outbox1.Size = $System_Drawing_Size
    $Outbox1.DataBindings.DefaultDataSourceUpdateMode = 0
    $Outbox1.Name = "listBox1"
    $System_Drawing_Point = New-Object System.Drawing.Point(15,70)
    $Outbox1.Location = $System_Drawing_Point
    $Outbox1.TabIndex = 8
    $Outbox1.MultiLine = $False
    $form1.controls.add($Outbox1)

    $lbl3=New-Object System.Windows.Forms.Label
    $lblfont3 = New-Object System.Drawing.Font("Times New Roman",11)
    $lbl3.Location=New-Object System.Drawing.Point( 15, 110 )
    $lbl3.font = $lblfont3
    $lbl3.Name = "label1"
    $lbl3.Size=New-Object System.Drawing.Size( 325, 20 )
    $lbl3.TabIndex=0
    $lbl3.Text= "Stack Name"
    $form1.controls.add($lbl3)

    $Outbox3 = New-Object System.Windows.Forms.TextBox
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 170
    $System_Drawing_Size.Height = 10
    $Outbox3.Size = $System_Drawing_Size
    $Outbox3.DataBindings.DefaultDataSourceUpdateMode = 0
    $Outbox3.Name = "listBox1"
    $System_Drawing_Point = New-Object System.Drawing.Point(15, 130)
    $Outbox3.Location = $System_Drawing_Point
    $Outbox3.TabIndex = 8
    $Outbox3.MultiLine = $False
    $form1.controls.add($Outbox3)

    $lbl4=New-Object System.Windows.Forms.Label
    $lblfont4 = New-Object System.Drawing.Font("Times New Roman",11)
    $lbl4.Location=New-Object System.Drawing.Point( 400, 48 )
    $lbl4.font = $lblfont3
    $lbl4.Name = "label1"
    $lbl4.Size=New-Object System.Drawing.Size( 103, 20 )
    $lbl4.TabIndex=0
    $lbl4.Text="Output Console"
    $form1.controls.add($lbl4)

   
    $Outbox4 = New-Object System.Windows.Forms.TextBox
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 340
    $System_Drawing_Size.Height = 200
    $Outbox4.Size = $System_Drawing_Size
    $Outbox4.DataBindings.DefaultDataSourceUpdateMode = 0
    $Outbox4.Name = "listBox1"
    $System_Drawing_Point = New-Object System.Drawing.Point(400, 68)
    $Outbox4.Location = $System_Drawing_Point
    $Outbox4.TabIndex = 8
    $Outbox4.MultiLine = $True
    $Outbox4.ScrollBars = "Vertical"
    $Outbox4.readonly = $true
    $form1.controls.add($Outbox4)

       
    Function Get-FileName($initialDirectory)
    {
    [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null
    
    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $OpenFileDialog.initialDirectory = $initialDirectory
    $OpenFileDialog.filter = "CSV (*.csv)| *.csv"
    $OpenFileDialog.ShowDialog() | Out-Null
    $OpenFileDialog.filename
    }

    $Addservers = New-Object System.Windows.Forms.Button
    $Addservers.Location = New-Object System.Drawing.Size(10,160)
    $Addservers.Size = New-Object System.Drawing.Size(100,35)
    $Addservers.Text = "Add Servers"
    $form1.Controls.Add($Addservers)

    $listcomp = New-Object System.Windows.Forms.Button
    $listcomp.Location = New-Object System.Drawing.Size(120,160)
    $listcomp.Size = New-Object System.Drawing.Size(100,35)
    $listcomp.Text = "List the Servers in Stack"
    $form1.Controls.Add($listcomp)

    $setupjob = New-Object System.Windows.Forms.Button
    $setupjob.Location = New-Object System.Drawing.Size(10,200)
    $setupjob.Size = New-Object System.Drawing.Size(100,35)
    $setupjob.Text = "Evaluate Stack"
    $form1.Controls.Add($setupjob)
    
    $removescomp = New-Object System.Windows.Forms.Button
    $removescomp.Location = New-Object System.Drawing.Size(120,200)
    $removescomp.Size = New-Object System.Drawing.Size(100,35)
    $removescomp.Text = "Remove Specific Servers"
    $form1.Controls.Add($removescomp)

    $sealstack = New-Object System.Windows.Forms.Button
    $sealstack.Location = New-Object System.Drawing.Size(10,240)
    $sealstack.Size = New-Object System.Drawing.Size(100,35)
    $sealstack.Text = "Activate Stack"
    $form1.Controls.Add($sealstack)

    $removecomp = New-Object System.Windows.Forms.Button
    $removecomp.Location = New-Object System.Drawing.Size(120,240)
    $removecomp.Size = New-Object System.Drawing.Size(100,35)
    $removecomp.Text = "Remove all servers"
    $form1.Controls.Add($removecomp)

    $AddserverEvt = {
    $outbox4.AppendText("<-----------Script execution started---------------->`n")
    $addservers.text = "Executing"
    $caserver = $Outbox1.Text.Trim()
    $quts = '"'
    $stkname = $Outbox3.Text.Trim()
    $srvrs = $Outbox.Text.Trim().Split("`n")
    $flag = 0
    if(($caserver -ne "") -and ($stkname -ne "") -and ($srvrs -ne ""))
    {
    foreach ($srv in $srvrs)
    {
    $srv = $srv.trim()
    $adcmd = ""
    $param =""
        if($flag -ne 1)
        {
        $o = cadsmcmd local $caserver templategroup action=addcomp name="$stkname" computer=$srv
            foreach ($ln in $o)
            {
                $sd = "SDCMD"
                if($ln -like "$sd*")
                {
                    $objlinked = "Object is already linked"
                    $stacknotavl = "Computer group does not exist"
                    $compnotavbl = "At least one of multiple operations failed"
                    $CAnotavbl = "Session establishment failed"
                    $OK = "SDCMD<A000000>: OK"
                    if($ln -like "*$objlinked*") { $outbox4.AppendText("$srv - Already added to $stkname`n");write all}
                    if($ln -like "*$compnotavbl*") { $outbox4.AppendText("$srv - Invalid computer name`n")}
                    if($ln -like "*$OK*") { $outbox4.AppendText("$srv - added to $stkname`n")}
                    if($ln -like "*$stacknotavl*") { $outbox4.AppendText("$stkname - Invalid stack`n");$flag =1}
                    if($ln -like "*$CAnotavbl*") { $outbox4.AppendText("$caserver - Invalid CA server/Access Denied`n");$flag =1}

                }
            }
        }
    }
    }
    else
    {
    $outbox4.AppendText("Please enter CA servername, StackName and TargetComputer `n")
    }
    $outbox4.AppendText("<-----------Script execution Completed---------->`n")
    $Addservers.Text = "Add Servers"
    }
    $Addservers.add_Click($AddserverEvt)

    
    $setupjobevt = {
    $outbox4.AppendText("<-----------Script execution started---------------->`n")
    $setupjob.text = "Executing"

    $caserver = $Outbox1.Text.Trim()
    $quts = '"'
    $stkname = $Outbox3.Text.Trim()

    if(($caserver -ne "") -and ($stkname -ne ""))
    {
            $unsl = cadsmcmd local $caserver templategroup action=unseal name="$stkname"
            foreach($ln in $unsl)
            {
                $sd = "SDCMD"
                 if($ln -like "$sd*")
                 {
                        $objlinked = "Object is already linked"
                        $stacknotavl = "Computer group does not exist"
                        $compnotavbl = "At least one of multiple operations failed"
                        $CAnotavbl = "Session establishment failed"
                        $emptystack = "SDCMD<A001251>"
                        $OK = "SDCMD<A000000>: OK"
                         if($ln -like "*$objlinked*") { $outbox4.AppendText("$srv - Already added to stack`n")}
                        elseif($ln -like "*$compnotavbl*") { $outbox4.AppendText("$srv - Invalid computer name`n")}
                        
                        elseif($ln -like "*$emptystack*") {$outbox4.AppendText("$stkname - No servers in this stack`n"); }
                        elseif($ln -like "*$stacknotavl*") { $outbox4.AppendText("$stkname - Invalid stack`n");$flag =1}
                        elseif($ln -like "*$CAnotavbl*") { $outbox4.AppendText("$caserver - Invalid CA server/Access Denied`n");$flag =1}
                        elseif($ln -like "*$OK*") {
                        $chpolicy=$null
                        $chpolicy=$stkname+' [$date $time]'
                        cadsmcmd local $caserver swPolicy action=modify name=$chpolicy setup_jobs
                        $o = cadsmcmd local $caserver templategroup action=seal name="$stkname"
                            foreach ($ln in $o)
                            {
                                $sd = "SDCMD"
                                 if($ln -like "$sd*")
                                    {
                                        $objlinked = "Object is already linked"
                                        $stacknotavl = "Computer group does not exist"
                                        $compnotavbl = "At least one of multiple operations failed"
                                        $CAnotavbl = "Session establishment failed"
                                        $emptystack = "SDCMD<A001251>"
                                        $OK = "SDCMD<A000000>: OK"
                                        if($ln -like "*$objlinked*") { $outbox4.AppendText("$srv - Already added to stack`n")}
                                        elseif($ln -like "*$compnotavbl*") { $outbox4.AppendText("$srv - Invalid computer name`n")}
                        
                                        elseif($ln -like "*$emptystack*") {$outbox4.AppendText("$stkname - No servers in this stack`n"); }
                                        elseif($ln -like "*$stacknotavl*") { $outbox4.AppendText("$stkname - Invalid stack`n");$flag =1}
                                        elseif($ln -like "*$CAnotavbl*") { $outbox4.AppendText("$caserver - Invalid CA server/Access Denied`n");$flag =1}
                                        elseif($ln -like "*$OK*") {
                                        $cassserver = $Outbox1.Text.Trim()
                                        $stknamess = $Outbox3.Text.Trim()
                                        listcomp $cassserver $stknamess;
                                        $outbox4.AppendText("$stkname - Job Evaluation submitted Successfully`n");}
                                        else{ $outbox4.AppendText("$ln`n")}
                                    }
                            }
                        }
                        else{ $outbox4.AppendText("$ln`n")}
                 }
            }
    }
    else
    {
    $outbox4.AppendText("Please enter CA servername and StackName`n")
    }

    $outbox4.AppendText("<-----------Script execution Completed---------->`n")
    $setupjob.text = "Evaluate Stack"
    }

    $setupjob.add_click($setupjobevt)

    $sealstackevt = {
    $outbox4.AppendText("<-----------Stack Activate script started---------------->`n")
    $sealstack.text = "Executing"

    $caserver = $Outbox1.Text.Trim()
    $quts = '"'
    $stkname = $Outbox3.Text.Trim()

    if(($caserver -ne "") -and ($stkname -ne ""))
    {
     
    $quts = '"'
    $jofil = "Job container name=$stkname"
    $job = "$quts$jofil"
    $adcmd = "cadsmcmd"
    $param ="local $caserver jobcontainer action=list  filter=$job"
    #start-process -filepath $adcmd -argumentlist $param* -RedirectStandardOutput C:\temp\SortOut.txt -WindowStyle Hidden -wait
    $jobcontainer = cadsmcmd local $caserver jobcontainer action=list filter="Job container name=$stkname*"
        $noofjob="Number of job containers shown:"
        
        foreach($lin in $jobcontainer)
        {
            if($lin -like "$noofjob*")
            {
                $jobcount=$lin.Replace("Number of job containers shown:	","")
                $inc=16
                if(!$lin.endswith(0))
                {
                $inc =16
                    for($i=0;$i -lt $jobcount;$i++)
                    {
                        $jcname = $jobcontainer[$inc]
                        if($jcname.EndsWith("(successfully built)"))
                        {
                            $index = $jcname.IndexOf("M] (")
                            $jobname = $jcname.Substring(0,$index+2)
                            #$param ="local $caserver jobcontainer action=activate name=$quts$jobname$quts"
                            #if(Test-Path -Path "C:\temp\SortOut.txt"){remove-item "C:\temp\SortOut.txt"}
                            #start-process -filepath $adcmd -argumentlist $param -RedirectStandardOutput C:\temp\SortOut.txt -WindowStyle Hidden -wait
                            $o = cadsmcmd local $caserver jobcontainer action=activate name="$jobname"
                            #write "im in"
                            foreach($ln in $o)
                            {
                                $ok = "SDCMD<A000000>: OK"
                                if ($ln -like $ok)
                                {
                                    $outbox4.AppendText("$jobname - Activated Successfully`n");
                                }
                            }
                        }
                        $inc = $inc + 1
                    }
                } 
                else
                {
                $outbox4.AppendText("Evaluated $stkname not found in the Job container`n")
                }             
            }
            $unkn = "SDCMD<CMD000035>: Unknown parameter "
            $jcn = "SDCMD<A001666>: Job Container not found."
            if(($lin -like "$unkn*") -or ($lin -like "$jcn*") )
            {
                $outbox4.AppendText("$stkname not found in the Job container`n")
            }
            

        }
   
    }
    else
    {
    $outbox4.AppendText("Please enter CA servername and StackName`n")
    }
    $outbox4.AppendText("<-----------Stack Activate script Completed---------->`n")
    $sealstack.text = "Activate Stack"
    }
    $sealstack.add_click($sealstackevt)
    function Listcomp($caspar,$stkpar)
    {
    
        #
        
        $caserver = $caspar
        
        $stkname = $stkpar
        #$srvrs = $Outbox.Text.Trim().Split("`n")
        $flag = 0
        if(($caserver -ne "") -and ($stkname -ne ""))
        {
            if($flag -ne 1)
            {
                $a = cadsmcmd local $caserver templategroup action=listcomp name="$stkname"
                foreach ($ln in $a)
                {
                    $sd = "SDCMD"
                    if($ln -like "$sd*")
                    {
                        $objlinked = "Object is already linked"
                        $stacknotavl = "Computer group does not exist"
                        $compnotavbl = "At least one of multiple operations failed"
                        $CAnotavbl = "Session establishment failed"
                        $OK = "SDCMD<A000000>: OK"
                        if($ln -like "*$objlinked*") { $outbox4.AppendText("$srv - Already added to stack`n")}
                        if($ln -like "*$compnotavbl*") { $outbox4.AppendText("$srv - Invalid computer name`n")}
                        if($ln -like "*$OK*") { 
                        $noft = "Number of target computer read: "
                        $totser = 0
                        $sflag = 0
                        foreach ($line in $a)
                        {
                            if($line -like "*$noft*")
                            {
                                $totser = $line.replace("Number of target computer read: ","")
                                $sflag = 1
                            }
                        }
                        $inc =16
                        if(($sflag -eq 1) -and ($totser -gt 0))
                        {
                            $outbox4.AppendText("$stkname contains $totser servers`n")
                            for($i = 0;$i -lt $totser;$i++)
                            {
                                $ser = $a[$inc]
                                $inc=$inc+1
                                $outbox4.AppendText("$ser`n")
                            }
                        }
                        else
                        {
                            $outbox4.AppendText("$stkname doesn't have any servers`n")
                        }
                        }
                        if($ln -like "*$stacknotavl*") { $outbox4.AppendText("$stkname - Invalid stack`n");$flag =1}
                        if($ln -like "*$CAnotavbl*") { $outbox4.AppendText("$caserver - Invalid CA server/Access Denied`n");$flag =1}
                    }
                    
                }
                
            }
        
        }
        else
        {
        $outbox4.AppendText("Please enter CA servername, StackName and TargetComputer `n")
        }
        
        

    
    }
    $listcompevt = {
    $listcomp.text = "Executing"
    $outbox4.AppendText("<-----------Script execution started---------------->`n");
    $cassserver = $Outbox1.Text.Trim()
    $stknamess = $Outbox3.Text.Trim()
    listcomp $cassserver $stknamess;
    $outbox4.AppendText("<-----------Script execution Completed---------->`n");
    $listcomp.Text = "List Servers from Stack"
    }
    $listcomp.add_click($listcompevt)

    $removecompevt = {
        $outbox4.AppendText("<-----------Script execution started---------------->`n")
        $removecomp.text = "Executing"
        $caserver = $Outbox1.Text.Trim()
        $quts = '"'
        $stkname = $Outbox3.Text.Trim()
        #$srvrs = $Outbox.Text.Trim().Split("`n")
        $flag = 0
        $adseraray = @()
        if(($caserver -ne "") -and ($stkname -ne ""))
        {
            if($flag -ne 1)
            {
                $a = cadsmcmd local $caserver templategroup action=listcomp name="$stkname"
                foreach ($ln in $a)
                {
                    $sd = "SDCMD"
                    if($ln -like "$sd*")
                    {
                        #$objlinked = "Object is already linked"
                        $stacknotavl = "Computer group does not exist"
                        #$compnotavbl = "At least one of multiple operations failed"
                        $CAnotavbl = "Session establishment failed"
                        $OK = "SDCMD<A000000>: OK"
                        #if($ln -like "*$objlinked*") { $outbox4.AppendText("$srv - Already added to stack`n")}
                        #if($ln -like "*$compnotavbl*") { $outbox4.AppendText("$srv - Invalid computer name`n")}
                        if($ln -like "*$OK*") { 
                        $noft = "Number of target computer read: "
                        $totser = 0
                        $sflag = 0
                        foreach ($line in $a)
                        {
                            if($line -like "*$noft*")
                            {
                                $totser = $line.replace("Number of target computer read: ","")
                                $sflag = 1
                            }
                        }
                        $inc =16
                        if(($sflag -eq 1) -and ($totser -gt 0))
                        {
                            $outbox4.AppendText("$stkname contains $totser servers`n")
                            for($i = 0;$i -lt $totser;$i++)
                            {
                                $ser = $a[$inc]
                                $inc=$inc+1
                                $rmser = cadsmcmd local $caserver templategroup action=removecomp name="$stkname" computer=$ser
                                $stacknotavl = "Computer group does not exist"
                                $compnotavbl = "At least one of multiple operations failed"
                                foreach($ltm in $rmser)
                                {
                                    if($ltm -like "SDCMD<A000000>: OK"){ $outbox4.AppendText("$ser - Removed`n")}
                                    
                                }
                            }
                        }
                        else
                        {
                            $outbox4.AppendText("$stkname doesn't have any servers`n")
                        }
                        }
                        if($ln -like "*$stacknotavl*") { $outbox4.AppendText("$stkname - Invalid stack`n");$flag =1}
                        if($ln -like "*$CAnotavbl*") { $outbox4.AppendText("$caserver - Invalid CA server/Access Denied`n");$flag =1} 
                    }
                    
                }
                
            }
        
        }
        else
        {
        $outbox4.AppendText("Please enter CA servername, StackName and TargetComputer `n")
        }
        
        $outbox4.AppendText("<-----------Script execution Completed---------->`n")
        $removecomp.Text = "Remove all servers"

    }
    $removecomp.add_click($removecompevt)

    $removescompevt = {
    $outbox4.AppendText("<-----------Script execution started---------------->`n")
    $removescomp.text = "Executing"
    $caserver = $Outbox1.Text.Trim()
    $quts = '"'
    $stkname = $Outbox3.Text.Trim()
    $srvrs = $Outbox.Text.Trim().Split("`n")
    $flag = 0
    if(($caserver -ne "") -and ($stkname -ne "") -and ($srvrs -ne ""))
    {
    foreach ($srv in $srvrs)
    {
    $srv = $srv.trim()
        if($flag -ne 1)
        {
        $o = cadsmcmd local $caserver templategroup action=removecomp name="$stkname" computer=$srv
            foreach ($ln in $o)
            {
                $sd = "SDCMD"
                if($ln -like "$sd*")
                {
                    $objlinked = "Object is not linked to group"
                    $stacknotavl = "Computer group does not exist"
                    $compnotavbl = "At least one of multiple operations failed"
                    $CAnotavbl = "Session establishment failed"
                    $OK = "SDCMD<A000000>: OK"
                    if($ln -like "*$objlinked*") { $outbox4.AppendText("$srv - Stack doesn't have this server`n");write all}
                    if($ln -like "*$compnotavbl*") { $outbox4.AppendText("$srv - Invalid computer name`n")}
                    if($ln -like "*$OK*") { $outbox4.AppendText("$srv - Removed from $stkname`n")}
                    if($ln -like "*$stacknotavl*") { $outbox4.AppendText("$stkname - Invalid stack `n");$flag =1}
                    if($ln -like "*$CAnotavbl*") { $outbox4.AppendText("$caserver - Invalid CA server/Access Denied`n");$flag =1}
                }
            }
        }
    }
    }
    else
    {
    $outbox4.AppendText("Please enter CA servername, StackName and TargetComputer `n")
    }
    $outbox4.AppendText("<-----------Script execution Completed---------->`n")
    $removescomp.Text = "Remove Specific Servers"
    }
    $removescomp.add_click($removescompevt)

    $clrconsole = New-Object System.Windows.Forms.Button
    $clrconsole.Location = New-Object System.Drawing.Size(670,45)
    $clrconsole.Size = New-Object System.Drawing.Size(70,20)
    $clrconsole.Text = "Clear Log"
    $form1.Controls.Add($clrconsole)

    $clrconsoleevt = {$outbox4.Text = "";}
    $clrconsole.add_click($clrconsoleevt)

    $blkclrconsole = New-Object System.Windows.Forms.Button
    $blkclrconsole.Location = New-Object System.Drawing.Size(670,335)
    $blkclrconsole.Size = New-Object System.Drawing.Size(70,20)
    $blkclrconsole.Text = "Clear Log"
    $form1.Controls.Add($blkclrconsole)

    $blkclrconsoleevt = {$bulkoutbox.Text = "";}
    $blkclrconsole.add_click($blkclrconsoleevt)

    $dtemplate = New-Object System.Windows.Forms.Button
    $dtemplate.Location = New-Object System.Drawing.Size(520,335)
    $dtemplate.Size = New-Object System.Drawing.Size(120,20)
    $dtemplate.Text = "Download Template"
    $form1.Controls.Add($dtemplate)

    $blkclrconsoleevt = {$bulkoutbox.Text = "";}
    $blkclrconsole.add_click($blkclrconsoleevt)

    #$dlg=New-Object System.Windows.Forms.SaveFileDialog
    #$dlg.filter = "CSV (*.csv)| *.csv"
    #$dlg.initialDirectory = "c:\temp"

    $dtemplateevt = {
    
     $obj12 = New-Object PSObject
     $obj12 | Add-Member -MemberType NoteProperty -Name "CA server name" -value ""
     $obj12 | Add-Member -MemberType NoteProperty -Name "Stack name" -value ""
     $obj12 | Add-Member -MemberType NoteProperty -Name "Server Name" -value ""
     
    if(test-path -path "c:\temp\servers.csv"){Remove-Item -Path "c:\temp\servers.csv" -Force;export-csv -path "c:\temp\servers.csv" -Force -InputObject $obj12;} 
    else{export-csv -path "c:\temp\servers.csv" -Force -InputObject $obj12}
    $bulkoutbox.AppendText("Template downloaded in c:\temp\servers.csv`n")

    
    }
    $dtemplate.add_click($dtemplateevt)

    $lbl7=New-Object System.Windows.Forms.Label
    $lbl7.Location=New-Object System.Drawing.Point( 240, 300 )
    $lbl7.Size=New-Object System.Drawing.Size( 220, 30 )
    $lblfont = New-Object System.Drawing.Font("Times New Roman",16,[System.Drawing.FontStyle]::BOLD)
    $lbl7.font = $lblfont
    $lbl7.TabIndex=0
    $lbl7.Text="Bulk Deployment Tool"
    $form1.controls.add($lbl7)

    $bulkoutbox = New-Object System.Windows.Forms.TextBox
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 340
    $System_Drawing_Size.Height = 200
    $bulkoutbox.Size = $System_Drawing_Size
    $bulkoutbox.DataBindings.DefaultDataSourceUpdateMode = 0
    $bulkoutbox.Name = "listBox1"
    $System_Drawing_Point = New-Object System.Drawing.Point(400, 360)
    $bulkoutbox.Location = $System_Drawing_Point
    $bulkoutbox.TabIndex = 8
    $bulkoutbox.MultiLine = $True
    $bulkoutbox.ScrollBars = "Vertical"
    $bulkoutbox.readonly = $true
    $form1.controls.add($bulkoutbox)

    $lbl9=New-Object System.Windows.Forms.Label
    $lblfont4 = New-Object System.Drawing.Font("Times New Roman",11)
    $lbl9.Location=New-Object System.Drawing.Point( 400, 335 )
    $lbl9.font = $lblfont4
    $lbl9.Name = "label1"
    $lbl9.Size=New-Object System.Drawing.Size( 325, 20 )
    $lbl9.TabIndex=0
    $lbl9.Text="Output Console"
    $form1.controls.add($lbl9)

    $ll9=New-Object System.Windows.Forms.Label
    $llfont4 = New-Object System.Drawing.Font("Times New Roman",11)
    $ll9.Location=New-Object System.Drawing.Point( 15, 335 )
    $ll9.font = $llfont4
    $ll9.Name = "label1"
    $ll9.Size=New-Object System.Drawing.Size( 325, 20 )
    $ll9.TabIndex=0
    $ll9.Text="Enter the CSV path"
    $form1.controls.add($ll9)
    
    $browseOutbox5 = New-Object System.Windows.Forms.TextBox
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 210
    $System_Drawing_Size.Height = 10
    $browseOutbox5.Size = $System_Drawing_Size
    $browseOutbox5.DataBindings.DefaultDataSourceUpdateMode = 0
    $Icon = New-Object system.drawing.icon ("C:\Program Files (x86)\CA\DSM\Bin\Images\CAFSTarted.ico")
    $Form1.Icon = $Icon
    $System_Drawing_Point = New-Object System.Drawing.Point(15, 365)
    $browseOutbox5.Location = $System_Drawing_Point
    $browseOutbox5.TabIndex = 8
    $browseOutbox5.MultiLine = $False
    $browseOutbox5.ReadOnly =$false
    $form1.controls.add($browseOutbox5)

    $browsebutton = New-Object System.Windows.Forms.Button
    $System_Drawing_Size = New-Object System.Drawing.Size(70,20)
    $browsebutton.size = $System_Drawing_Size
    $browsebutton.text = "Browse"
    $System_Drawing_Point = New-Object System.Drawing.Point(230,365)
    $browsebutton.Location = $System_Drawing_Point
    #$form1.controls.add($browsebutton)

    $browseevt = {
    $inputfile = Get-FileName "C:\temp"
    $inputdata = import-csv -path $inputfile
    $browseOutbox5.Text = $inputfile
    }
    #$browsebutton.add_click($browseevt)

    $clrbutton = New-Object System.Windows.Forms.Button
    $System_Drawing_Size = New-Object System.Drawing.Size(70,20)
    $clrbutton.size = $System_Drawing_Size
    $clrbutton.text = "Clear"
    $System_Drawing_Point = New-Object System.Drawing.Point(310,365)
    $clrbutton.Location = $System_Drawing_Point
    #$form1.controls.add($clrbutton)

    $clrbuttonevt = {$browseOutbox5.Text = "";}
    $clrbutton.add_click($clrbuttonevt)
    $bulkAddservers = New-Object System.Windows.Forms.Button
    $bulkAddservers.Location = New-Object System.Drawing.Size(30,395)
    $bulkAddservers.Size = New-Object System.Drawing.Size(125,45)
    $bulkAddservers.Text = "1. Add Servers (Multiple CA servers)"
    $form1.Controls.Add($bulkAddservers)

    $bulkaddevt={
        $bulkoutbox.AppendText("<-----------Bulk server addition started---------------->`n")
        $input = $browseOutbox5.Text.Trim()
        $cs = ".CSV"
        if(($input -ne "") -and (test-path -path $input) -and ($input -like "*$cs"))
        {
            $inputdata = import-csv -Path $input
            foreach ($ln in $inputdata)
            {
            $cas = $ln.'CA server name'
            $stk = $ln.'Stack name'
            $srv = $ln.'Server Name'
            if(($cas -ne "") -and ($stk -ne "") -and ($srv -ne ""))
            {
                
                $bulkAddservers.text = "Executing"
                $caserver = $cas
                $quts = '"'
                $stkname = $stk
                $srvrs = $srv
                $flag = 0
                
                foreach ($srv in $srvrs)
                {
                $srv = $srv.trim()
                $adcmd = "cadsmcmd"
                $param ="local $caserver templategroup action=addcomp name=$quts$stkname$quts computer=$srv"
                    if($flag -ne 1)
                    {
                    $o = cadsmcmd local $caserver templategroup action=addcomp name="$stkname" computer=$srv
                        foreach ($ln in $o)
                        {
                            $sd = "SDCMD"
                            if($ln -like "$sd*")
                            {
                                $objlinked = "Object is already linked"
                                $stacknotavl = "Computer group does not exist"
                                $compnotavbl = "At least one of multiple operations failed"
                                $CAnotavbl = "Session establishment failed"
                                $OK = "SDCMD<A000000>: OK"
                                $bulkoutbox.AppendText("-----------------------------------------------------------------------`n")
                                $bulkoutbox.AppendText("CA Server - $caserver`n")
                                $bulkoutbox.AppendText("Stack Name - $stkname`n")
                                $bulkoutbox.AppendText("ServerName - $srv`n")
                                if($ln -like "*$objlinked*") { $bulkoutbox.AppendText("Status - Already added to stack`n");}
                                if($ln -like "*$compnotavbl*") { $bulkoutbox.AppendText("Status - Invalid computer name`n")}
                                if($ln -like "*$OK*") { $bulkoutbox.AppendText("Status - Server added successfully`n")}
                                if($ln -like "*$stacknotavl*") { $bulkoutbox.AppendText("Status - Stack not avail/Invalid Stack `n");}
                                if($ln -like "*$CAnotavbl*") { $bulkoutbox.AppendText("Status - Invalid CA server/Access Denied`n");}
                    
                            }
                        }
                    }
                }
          
            }

        }
        }
        else
        {
        $bulkoutbox.AppendText("Please enter valid CSV path`n")
        }
        $bulkoutbox.AppendText("<-----------Bulk server addition Completed---------->`n")
        $bulkAddservers.Text = "1. Add Servers (Multiple CA servers)"

    }
    $bulkAddservers.add_click($bulkaddevt) 

    $bulkevaluate = New-Object System.Windows.Forms.Button
    $bulkevaluate.Location = New-Object System.Drawing.Size(30,455)
    $bulkevaluate.Size = New-Object System.Drawing.Size(125,45)
    $bulkevaluate.Text = "2. Evaluate Stacks (Multiple CA servers)"
    $form1.Controls.Add($bulkevaluate)

    $bulkevalevt ={
        $bulkoutbox.AppendText("<-----------Bulk Stacks Evaluate started---------------->`n")
        $input = $browseOutbox5.Text.Trim()
        $cs = ".CSV"
        if(($input -ne "") -and (test-path -path $input) -and ($input -like "*$cs"))
        {
            $inputdata = import-csv -Path $input
            $alcasstk=@()
            foreach ($ln in $inputdata)
            {
                $cas = $ln.'CA server name'
                $stk = $ln.'Stack name'
                $casstk = "$cas*$stk"
                $alcasstk += $casstk
            }
            $alcasstk = $alcasstk | select -Unique

            foreach ($ln in $alcasstk)
            {
            $ln=$ln.split("*")
            $cas = $ln[0]
            $stk = $ln[1]
            #$srv = $ln.'Server Name'
            if(($cas -ne "") -and ($stk -ne ""))
            {
                
                $bulkevaluate.text = "Executing"
                $caserver = $cas
                $quts = '"'
                $stkname = $stk
                #$srvrs = $srv
                #$srv = $srv.trim()
                $chpolicy=$null
                        $chpolicy=$stkname+' [$date $time]'
                        cadsmcmd local $caserver swPolicy action=modify name=$chpolicy setup_jobs
                    $o = cadsmcmd local $caserver templategroup action=seal name="$stkname" 
                        foreach ($ln in $o)
                        {
                            $sd = "SDCMD"
                            if($ln -like "$sd*")
                            {
                                #$objlinked = "Object is already linked"
                                $stacknotavl = "Computer group does not exist"
                                #$compnotavbl = "At least one of multiple operations failed"
                                $CAnotavbl = "Session establishment failed"
                                $OK = "SDCMD<A000000>: OK"
                                $bulkoutbox.AppendText("-----------------------------------------------------------------------`n")
                                $bulkoutbox.AppendText("CA Server - $caserver`n")
                                $bulkoutbox.AppendText("Stack Name - $stkname`n")
                                #$bulkoutbox.AppendText("ServerName - $srv`n")
                                #if($ln -like "*$objlinked*") { $bulkoutbox.AppendText("Status - Already added to stack`n");}
                                #if($ln -like "*$compnotavbl*") { $bulkoutbox.AppendText("Status - Invalid computer name`n")}
                                if($ln -like "*$OK*") { $bulkoutbox.AppendText("Status - Evaluate submitted succesfully`n")}
                                elseif($ln -like "*$stacknotavl*") { $bulkoutbox.AppendText("Status - Stack not avail/Invalid Stack `n");}
                                elseif($ln -like "*$CAnotavbl*") { $bulkoutbox.AppendText("Status - Invalid CA server/Access Denied`n");}
                                else{$bulkoutbox.AppendText("Status - $ln`n")}
                            }
                        }
            }

        }
        }
        else
        {
        $bulkoutbox.AppendText("Please enter valid CSV path`n")
        }
        $bulkoutbox.AppendText("<-----------Bulk Stacks evaluate Completed---------->`n")
        $bulkevaluate.Text = "2. Evaluate Stacks (Multiple CA servers)"

    }
    $bulkevaluate.add_click($bulkevalevt)

    $bulkrmbutton = New-Object System.Windows.Forms.Button
    $bulkrmbutton.Location = New-Object System.Drawing.Size(180,520)
    $bulkrmbutton.Size = New-Object System.Drawing.Size(125,45)
    $bulkrmbutton.Text = "Remove servers from all stacks"
    $form1.Controls.Add($bulkrmbutton)

    $bulkremoveevt ={
    
        $input = $browseOutbox5.Text.Trim()
        $cs = ".CSV"
        if(($input -ne "") -and (test-path -path $input) -and ($input -like "*$cs"))
        {
            $bulkoutbox.AppendText("<-----------Bulk Stacks-server removal started---------------->`n")
            $inputdata = import-csv -Path $input
            $alcasstk=@()
            foreach ($ln in $inputdata)
            {
                $cas = $ln.'CA server name'
                $stk = $ln.'Stack name'
                $casstk = "$cas*$stk"
                $alcasstk += $casstk
            }
            $alcasstk = $alcasstk | select -Unique
            foreach ($ln in $alcasstk)
            {
            $ln=$ln.split("*")
            $cas = $ln[0]
            $stk = $ln[1]
            #$srv = $ln.'Server Name'
                if(($cas -ne "") -and ($stk -ne ""))
                {
                
                $bulkrmbutton.text = "Executing"
                $caserver = $cas
                $quts = '"'
                $stkname = $stk
                #$srvrs = $srv
                #$srv = $srv.trim()
                if(($caserver -ne "") -and ($stkname -ne ""))
                {
                    if($flag -ne 1)
                    {
                        $a = cadsmcmd local $caserver templategroup action=listcomp name="$stkname"
                        $bulkoutbox.AppendText("-----------------------------------------------------------------------`n")
                        $bulkoutbox.AppendText("CA Server - $caserver`n")
                        $bulkoutbox.AppendText("Stack Name - $stkname`n")
                        foreach ($ln in $a)
                        {
                            $sd = "SDCMD"
                            if($ln -like "$sd*")
                            {
                                $objlinked = "Object is already linked"
                                $stacknotavl = "Computer group does not exist"
                                $compnotavbl = "At least one of multiple operations failed"
                                $CAnotavbl = "Session establishment failed"
                                $OK = "SDCMD<A000000>: OK"
                                if($ln -like "*$objlinked*") {$bulkoutbox.AppendText("$srv - Already added to stack`n")}
                                if($ln -like "*$compnotavbl*") {$bulkoutbox.AppendText("$srv - Invalid computer name`n")}
                                if($ln -like "*$OK*") 
                                { 
                                $noft = "Number of target computer read: "
                                $totser = 0
                                $sflag = 0
                                    foreach ($line in $a)
                                    {
                                        if($line -like "*$noft*")
                                        {
                                            $totser = $line.replace("Number of target computer read: ","")
                                            $sflag = 1
                                        }
                                    }
                                    $inc =16
                                    $incc=16
                                    if(($sflag -eq 1) -and ($totser -gt 0))
                                    {
                                        $bulkoutbox.AppendText("$stkname contains $totser servers`n")
                                        for($i = 0;$i -lt $totser;$i++)
                                        {
                                            $ser = $a[$inc]
                                            $inc=$inc+1
                                            $bulkoutbox.AppendText("$ser`n")
                                        }
                                        for($j = 0;$j -lt $totser;$j++)
                                        {
                                            $ser = $a[$incc]
                                            $incc=$incc+1
                                            $o = cadsmcmd local $caserver templategroup action=removecomp name="$stkname" computer=$ser
                                            foreach ($ln in $o)
                                            {
                                                $sd = "SDCMD"
                                                if($ln -like "$sd*")
                                                {
                                                    $objlinked = "Object is not linked to group"
                                                    $stacknotavl = "Computer group does not exist"
                                                    $compnotavbl = "At least one of multiple operations failed"
                                                    $CAnotavbl = "Session establishment failed"
                                                    $OK = "SDCMD<A000000>: OK"
                                                    if($ln -like "*$objlinked*") { $bulkoutbox.AppendText("$ser - Stack doesn't have this server`n");write all}
                                                    if($ln -like "*$compnotavbl*") { $bulkoutbox.AppendText("$ser - Invalid computer name`n")}
                                                    if($ln -like "*$OK*") { $bulkoutbox.AppendText("$ser - Removed from $stkname`n")}
                                                    if($ln -like "*$stacknotavl*") { $bulkoutbox.AppendText("$stkname - Invalid stack `n");$flag =1}
                                                    if($ln -like "*$CAnotavbl*") { $bulkoutbox.AppendText("$caserver - Invalid CA server/Access Denied`n");$flag =1}
                                                }
                                            }
                                        }
                                    }
                                    else
                                    {
                                        $bulkoutbox.AppendText("$stkname doesn't have any servers`n")
                                    }
                                }
                                if($ln -like "*$stacknotavl*") {$bulkoutbox.AppendText("$stkname - Invalid stack`n")} #;$flag =1}
                                if($ln -like "*$CAnotavbl*") { $bulkoutbox.AppendText("$caserver - Invalid CA server/Access Denied`n")}#;$flag =1}
                            }
                        }
                
                    }
        
                }
                $bulkrmbutton.text = "Remove servers from all stacks"
            }#if ends
            #$bulkoutbox.AppendText("<alcasstk `n")
            }
            $bulkoutbox.AppendText("<-----------Bulk Stacks-servers removal Completed---------->`n")
        }
        else
        {
        $bulkoutbox.AppendText("Please enter valid CSV path`n")
        }
        
    }
    $bulkrmbutton.add_click($bulkremoveevt)

    $bulkseal = New-Object System.Windows.Forms.Button
    $bulkseal.Location = New-Object System.Drawing.Size(30,520)
    $bulkseal.Size = New-Object System.Drawing.Size(125,45)
    $bulkseal.Text = "3. Activate Stacks (Multiple CA servers)"
    $form1.Controls.Add($bulkseal)

    $bulksealevt = {
        $bulkoutbox.AppendText("<-----------Bulk Stacks Seal started---------------->`n")
        $input = $browseOutbox5.Text.Trim()
        $cs = ".CSV"
        if(($input -ne "") -and (test-path -path $input) -and ($input -like "*$cs"))
        {
            $inputdata = import-csv -Path $input
            $alcasstk=@()
            foreach ($ln in $inputdata)
            {
                $cas = $ln.'CA server name'
                $stk = $ln.'Stack name'
                $casstk = "$cas*$stk"
                $alcasstk += $casstk
            }
            $alcasstk = $alcasstk | select -Unique

            foreach ($ln in $alcasstk)
            {
            $ln=$ln.split("*")
            $cas = $ln[0]
            $stk = $ln[1]
            $bulkoutbox.AppendText("$cas server")
            #$srv = $ln.'Server Name'
            if(($cas -ne "") -and ($stk -ne ""))
            {
                
                $bulkseal.text = "Executing"
                $caserver = $cas
                $quts = '"'
                $stkname = $stk
                    #$quts = '"'
                    #$jofil = "Job container name=$stkname"
                    #$job = "$quts$jofil"
                    #$adcmd = "cadsmcmd"
                    #$param ="local $caserver jobcontainer action=list  filter=$job"
                    #start-process -filepath $adcmd -argumentlist $param* -RedirectStandardOutput C:\temp\SortOut.txt -WindowStyle Hidden -wait
                    $jobcontainer = cadsmcmd local $caserver jobcontainer action=list filter="Job container name=$stkname*"
                    $noofjob="Number of job containers shown:"
                        foreach($lin in $jobcontainer)
                        {
                            if($lin -like "$noofjob*")
                            {
                                $jobcount=$lin.Replace("Number of job containers shown:	","")
                                $inc=16
                                if(!$lin.endswith(0))
                                {
                                $inc =16
                                    for($i=0;$i -lt $jobcount;$i++)
                                    {
                                        $jcname = $jobcontainer[$inc]
                                        if($jcname.EndsWith("(successfully built)"))
                                        {
                                            $index = $jcname.IndexOf("M] (")
                                            $jobname = $jcname.Substring(0,$index+2)
                                            #$param ="local $caserver jobcontainer action=activate name=$quts$jobname$quts"
                                            #if(Test-Path -Path "C:\temp\SortOut.txt"){remove-item "C:\temp\SortOut.txt"}
                                            #start-process -filepath $adcmd -argumentlist $param -RedirectStandardOutput C:\temp\SortOut.txt -WindowStyle Hidden -wait
                                            $o = cadsmcmd local $caserver jobcontainer action=activate name="$jobname"
                                            foreach($ln in $o)
                                            {
                                                $ok = "SDCMD<A000000>: OK"
                                                if ($ln -like $ok)
                                                {
                                                    $bulkoutbox.AppendText("$jobname sealed Successfully`n");
                                                }
                                            }
                                        }
                                        $inc = $inc + 1
                                    }
                                }                
                            }
                            $unkn = "SDCMD<CMD000035>: Unknown parameter "
                            $jcn = "SDCMD<A001666>: Job Container not found."
                            $sdcmd = "SDCMD"
                            if(($lin -like "$unkn*") -or ($lin -like "$jcn*") )
                            {
                                $bulkoutbox.AppendText("$stkname not found in the Job container`n")
                            }
                            if($lin -like "$sdcmd*")
                            {
                                $bulkoutbox.AppendText("Status - $lin`n")
                            }
                            

                        }
            }

        }
        }
        else
        {
        $bulkoutbox.AppendText("Please enter valid CSV path`n")
        }
        $bulkoutbox.AppendText("<-----------Bulk Stacks Seal Completed---------->`n")
        $bulkseal.Text = "3. Activate Stacks (Multiple CA servers)"

    }
    $bulkseal.add_click($bulksealevt)

    $bulklistservers = New-Object System.Windows.Forms.Button
    $bulklistservers.Location = New-Object System.Drawing.Size(180,395)
    $bulklistservers.Size = New-Object System.Drawing.Size(125,45)
    $bulklistservers.Text = "List Servers (Multiple Stacks)"
    $form1.Controls.Add($bulklistservers)

    $bulklistserversevt = {
        $bulkoutbox.AppendText("<-----------Listing servers in multiple stacks---------------->`n")
        
        $bulklistservers.Text = "Executing"
        $input = $browseOutbox5.Text.Trim()
        $cs = ".CSV"
        if(($input -ne "") -and (test-path -path $input) -and ($input -like "*$cs"))
        {
            $inputdata = import-csv -Path $input
            $alcasstk=@()
            foreach ($ln in $inputdata)
            {
                $cas = $ln.'CA server name'
                $stk = $ln.'Stack name'
                $casstk = "$cas*$stk"
                $alcasstk += $casstk
            }
            $alcasstk = $alcasstk | select -Unique

            foreach ($ln in $alcasstk)
            {
            $ln=$ln.split("*")
            $cas = $ln[0]
            $stk = $ln[1]
            #$srv = $ln.'Server Name'
            if(($cas -ne "") -and ($stk -ne ""))
            {
                
                
                $caserver = $cas
                $quts = '"'
                $stkname = $stk
                    $bulkoutbox.AppendText("-----------------------------------------------------------------------`n")
                    $bulkoutbox.AppendText("`n")
                    $bulkoutbox.AppendText("CA Server - $caserver`n")
                    $bulkoutbox.AppendText("Stack Name - $stkname`n")
                    $flag = 0
                    if(($caserver -ne "") -and ($stkname -ne ""))
                    {
                        if($flag -ne 1)
                        {
                            $a = cadsmcmd local $caserver templategroup action=listcomp name="$stkname"
                            foreach ($ln in $a)
                            {
                                $sd = "SDCMD"
                                if($ln -like "$sd*")
                                {
                                    $objlinked = "Object is already linked"
                                    $stacknotavl = "Computer group does not exist"
                                    $compnotavbl = "At least one of multiple operations failed"
                                    $CAnotavbl = "Session establishment failed"
                                    $OK = "SDCMD<A000000>: OK"
                                    if($ln -like "*$objlinked*") { $bulkoutbox.AppendText("$srv - Already added to stack`n")}
                                    if($ln -like "*$compnotavbl*") { $bulkoutbox.AppendText("$srv - Invalid computer name`n")}
                                    if($ln -like "*$OK*") { 
                                    $noft = "Number of target computer read: "
                                    $totser = 0
                                    $sflag = 0
                                    foreach ($line in $a)
                                    {
                                        if($line -like "*$noft*")
                                        {
                                            $totser = $line.replace("Number of target computer read: ","")
                                            $sflag = 1
                                        }
                                    }
                                    $inc =16
                                    if(($sflag -eq 1) -and ($totser -gt 0))
                                    {
                                        $bulkoutbox.AppendText("Stack contains $totser servers`n`n")
                                        $bulkoutbox.AppendText("`n")
                                        for($i = 0;$i -lt $totser;$i++)
                                        {
                                            $ser = $a[$inc]
                                            $inc=$inc+1
                                            $bulkoutbox.AppendText("$ser`n")
                                        }
                                    }
                                    else
                                    {
                                        $bulkoutbox.AppendText("$stkname doesn't have any servers`n")
                                    }
                                    }
                                    if($ln -like "*$stacknotavl*") { $bulkoutbox.AppendText("$stkname - Invalid stack`n");$flag =1}
                                    if($ln -like "*$CAnotavbl*") { $bulkoutbox.AppendText("$caserver - Invalid CA server/Access Denied`n");$flag =1}
                                }
                    
                            }
                
                        }
        
                    }
                    else
                    {
                    $bulkoutbox.AppendText("Please enter CA servername, StackName and TargetComputer `n")
                    }
                $bulkoutbox.AppendText("`n")
            }

        }
        }
        else
        {
        $bulkoutbox.AppendText("Please enter valid CSV path`n")
        }
        
        $bulkoutbox.AppendText("<-----------Listing servers in multiple stacks Completed---------->`n")
        $bulklistservers.Text = "List Servers (Multiple Stacks)"

    }
    $bulklistservers.add_click($bulklistserversevt)

    $quitbutton = New-Object System.Windows.Forms.Button
    $quitbutton.Location = New-Object System.Drawing.Size(190,460)
    $quitbutton.Size = New-Object System.Drawing.Size(95,23)
    $quitbutton.Text = "Quit"
    $form1.Controls.Add($quitbutton)
    $quitevent = {$form1.close()}
    $quitbutton.add_Click($quitevent)
    $form1.ShowDialog()| Out-Null
    
