    [reflection.assembly]::loadwithpartialname("System.Windows.Forms") | Out-Null
    [reflection.assembly]::loadwithpartialname("System.Drawing") | Out-Null

    $form1 = New-Object System.Windows.Forms.Form
    $form1.Text = "Linux products Retriever "
    $form1.Name = "form1"
    $form1.DataBindings.DefaultDataSourceUpdateMode = 0
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 345
    $System_Drawing_Size.Height = 500
    $form1.ClientSize = $System_Drawing_Size
    $form1.StartPosition = "CenterScreen"  

    $lbl=New-Object System.Windows.Forms.Label
    $lblfont = New-Object System.Drawing.Font("Times New Roman",14,[System.Drawing.FontStyle]::BOLD)
    $lbl.Location=New-Object System.Drawing.Point( 60, 20 )
    $lbl.font = $lblfont
    $lbl.Name = "label1"
    $lbl.Size=New-Object System.Drawing.Size( 325, 20 )
    $lbl.TabIndex=0
    $lbl.Text="Linux Products Retriever"
    $form1.controls.add($lbl)

    $lbl1=New-Object System.Windows.Forms.Label
    $lblfont1 = New-Object System.Drawing.Font("Times New Roman",11)
    $lbl1.Location=New-Object System.Drawing.Point( 40, 60 )
    $lbl1.font = $lblfont1
    $lbl1.Name = "ASIA\NBK_ID"
    $lbl1.Size=New-Object System.Drawing.Size( 100, 20 )
    $lbl1.TabIndex=0
    $lbl1.Text="Domain\User"
    $form1.controls.add($lbl1)

    $lbl2=New-Object System.Windows.Forms.Label
    $lblfont2 = New-Object System.Drawing.Font("Times New Roman",11)
    $lbl2.Location=New-Object System.Drawing.Point( 40, 100 )
    $lbl2.font = $lblfont2
    $lbl2.Name = "l"
    $lbl2.Size=New-Object System.Drawing.Size( 65, 20 )
    $lbl2.TabIndex=0
    $lbl2.Text="Password"
    $form1.controls.add($lbl2)

    $lbl3=New-Object System.Windows.Forms.Label
    $lblfont1 = New-Object System.Drawing.Font("Times New Roman",11)
    $lbl3.Location=New-Object System.Drawing.Point( 40, 140 )
    $lbl3.font = $lblfont1
    $lbl3.Name = "ASIA\NBK_ID"
    $lbl3.Size=New-Object System.Drawing.Size( 100, 20 )
    $lbl3.TabIndex=0
    $lbl3.Text="Email address"
    $form1.controls.add($lbl3)
    
    $nbkid = New-Object System.Windows.Forms.TextBox
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 140
    $System_Drawing_Size.Height = 10
    $nbkid.Size = $System_Drawing_Size
    $nbkid.DataBindings.DefaultDataSourceUpdateMode = 0
    $System_Drawing_Point = New-Object System.Drawing.Point(150, 60)
    $nbkid.Location = $System_Drawing_Point
    $nbkid.TabIndex = 1
    $nbkid.MultiLine = $False
    $form1.controls.add($nbkid)

    $emailid = New-Object System.Windows.Forms.TextBox
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 180
    $System_Drawing_Size.Height = 10
    $emailid.Size = $System_Drawing_Size
    $emailid.DataBindings.DefaultDataSourceUpdateMode = 0
    $System_Drawing_Point = New-Object System.Drawing.Point(150, 140)
    $emailid.Location = $System_Drawing_Point
    $emailid.TabIndex = 3
    $emailid.MultiLine = $False
    $form1.controls.add($emailid)

    $rellbl3=New-Object System.Windows.Forms.Label
    $lblfont1 = New-Object System.Drawing.Font("Times New Roman",11)
    $rellbl3.Location=New-Object System.Drawing.Point( 40, 180 )
    $rellbl3.font = $lblfont1
    $rellbl3.Name = "x"
    $rellbl3.Size=New-Object System.Drawing.Size( 100, 20 )
    $rellbl3.TabIndex=0
    $rellbl3.Text="Release Name"
    $form1.controls.add($rellbl3)

    $release1 = New-Object System.Windows.Forms.TextBox
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 180
    $System_Drawing_Size.Height = 200
    $release1.Size = $System_Drawing_Size
    $release1.DataBindings.DefaultDataSourceUpdateMode = 0
    $System_Drawing_Point = New-Object System.Drawing.Point(150, 180)
    $release1.Location = $System_Drawing_Point
    $release1.TabIndex = 4
    $release1.MultiLine = $true
    $release1.ScrollBars = "Vertical"
    $form1.controls.add($release1)

    $paswd = New-Object System.Windows.Forms.MaskedTextBox
    $paswd.PasswordChar = '*'
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 140
    $System_Drawing_Size.Height = 10
    $paswd.Size = $System_Drawing_Size
    $paswd.DataBindings.DefaultDataSourceUpdateMode = 0
    $System_Drawing_Point = New-Object System.Drawing.Point(150,100 )
    $paswd.Location = $System_Drawing_Point
    $paswd.TabIndex = 2
    $paswd.MultiLine = $False
    $form1.controls.add($paswd)

    $runbutton = New-Object System.Windows.Forms.Button
    $runbutton.Location = New-Object System.Drawing.Size(150,390)
    $runbutton.Size = New-Object System.Drawing.Size(130,30)
    $runbutton.Text = "Run Script"
    $form1.Controls.Add($runbutton)

    $Outbox = New-Object System.Windows.Forms.TextBox
    $System_Drawing_Size = New-Object System.Drawing.Size
    $System_Drawing_Size.Width = 330
    $System_Drawing_Size.Height = 195
    $Outbox.Size = $System_Drawing_Size
    $Outbox.DataBindings.DefaultDataSourceUpdateMode = 0
    $Outbox.Name = "listBox1"
    $System_Drawing_Point = New-Object System.Drawing.Point
    $System_Drawing_Point.X = 9
    $System_Drawing_Point.Y = 260
    $Outbox.Location = $System_Drawing_Point
    $Outbox.TabIndex = 8
    $Outbox.MultiLine = $True
    $Outbox.ScrollBars = "Vertical"
    $Outbox.readonly = $true
    #$form1.controls.add($Outbox)

    $dataSource = "techmdb3.services.us.ml.com\INST3"
    $user = "scmwebapp"
    $pwd = "wearescum"
    $database = "dblc_db"
    $connectionString = "Server=$dataSource;uid=$user; pwd=$pwd;Database=$database;Integrated Security=False;"
function get-relmanid ([string]$rname)
{
$connection = New-Object System.Data.SqlClient.SqlConnection;
$connection.ConnectionString = $connectionString;
$connection.Open()
$command = $connection.CreateCommand()
$getquery = "select * from rel_man_table where rel_man_name='$rname'"
$command.CommandText = $getquery
$result = $command.ExecuteReader()
$relmanid = new-object “System.Data.DataTable”
$relmanid.Load($result)
return $relmanid.rel_man_id
$connection.Close()
}

function pull_prod_id ([int]$rel_id)
{
$storeproc = "get_rel_product_list @release_id='$rel_id'"
$connection = New-Object System.Data.SqlClient.SqlConnection;
$connection.ConnectionString = $connectionString;
$connection.Open()
$command = $connection.CreateCommand()
$command.CommandText = $storeproc
$result = $command.ExecuteReader()
$table = new-object “System.Data.DataTable”
$table.Load($result)
$prodlist = $Table | select Product_name,product_id | Format-Table -AutoSize
$prodid = $Table | select -ExpandProperty product_id 
return $prodid
$connection.Close()
}

function prodtype ([int]$prodid,$revi)
{
$q2 = "get_product_attr @product_id=$prodid"
$connection = New-Object System.Data.SqlClient.SqlConnection;
$connection.ConnectionString = $connectionString;
$connection.Open()
$command = $connection.CreateCommand()
$command.CommandText = $q2
$result = $command.ExecuteReader()
$prodtype = new-object “System.Data.DataTable”
$prodtype.Load($result)
$ptype = $null
$ptype = $prodtype | select -ExpandProperty attr_value
$x = 'Linux,'
    if($ptype -like "$x")
    {
    $q3 = "select product_name from product_table where product_id=$prodid"
    $command1 = $connection.CreateCommand()
    $command1.CommandText = $q3
    $result3 = $command1.ExecuteReader()
    $prodname = new-object “System.Data.DataTable”
    $prodname.Load($result3)
    $pname = $prodname | select -ExpandProperty product_name

    $stp = "get_product_search_details @product_name='$pname'"
    $command = $connection.CreateCommand()
    $command.CommandText = $stp
    $result = $command.ExecuteReader()

    $table = new-object “System.Data.DataTable”
    $table.Load($result)
    $pver=$null
    [string]$s = $revi
    foreach($v in $table)
    {
    $f = $v.release_name
    if($f -like "$s")
    {
    $pverii = $v.LatestIntegVersion
     return "$pname*$pverii"
    }
    }
   
    }
    
$connection.Close()
}
function sendmail($mailser,$mlid,$unam,$pwdd)
{

$SenderMailID=$RecipientList=$mlid
$date = Get-Date
$c = Get-Date –f yyyyMMddHHmmss;$out = "C:\temp\Linuxproducts_$c.csv";
Rename-Item -Path "C:\temp\Linuxproducts.csv" -NewName $out;
$secpasswd = ConvertTo-SecureString $pwdd -AsPlainText -Force
$smtpcred = New-Object System.Management.Automation.PSCredential ($unam,$secpasswd)
Send-MailMessage -SmtpServer $smtpServer -From $SenderMailID -To $RecipientList -Subject "Linux products $Date" -Attachments $out -Body "Linux products $Date" -credential $smtpcred -ErrorAction Stop
}
    $runevent ={
    
    $runbutton.Text = "Please wait";$this.enabled =$false;
    #########333
    
    $releases = $release1.text.trim().Split("`n")
    $prodarr=@()
    foreach($release in $releases)
    {
        $release = $release.trim()
        $pf = $release.split(" ") 
    $platform = $pf[0]
    $outbox.AppendText("$release`n")
    $outbox.AppendText("`n")
    $relid = get-relmanid $release
    $plist = pull_prod_id $relid 
    
    if($plist -ne $null)
    {
        foreach ($pod in $plist)
        {
        $final = $f1 = $prname = $prver = $null
        $final = prodtype $pod $release
        $f1 = $final.split("*")
        $prname = $f1[0]
        $prver =$f1[1]
        [int]$pid=$pod
        #===================================================
        $getservertype="get_servertypes_for_product @product_id=$pid"
        $connection = New-Object System.Data.SqlClient.SqlConnection;
        $connection.ConnectionString = $connectionString;
        $connection.Open()
        $command = $connection.CreateCommand()
        $command.CommandText = $getservertype
        $result = $command.ExecuteReader()
        
        $table = new-object “System.Data.DataTable”
        $table.Load($result)
        
        $out = $Table 
        $s = $out | select -ExpandProperty Column1
        $s | out-file c:\temp\stype.xml
        [xml]$s2 = gc c:\temp\stype.xml
        $allstypes = $s2.product.releaseplatforms.releaseplatform.servertypes.servertype.servertypename 

        #===================================================
        foreach($sty in $allstypes)
        {
            if(($prname -ne $null) -and ($prver -ne $null))
            {
            $obj12 = New-Object PSObject
            $obj12 | Add-Member -MemberType NoteProperty -Name "Platform" -Value $platform
            $obj12 | Add-Member -MemberType NoteProperty -Name "Release_Name" -Value $release
            $obj12 | Add-Member -MemberType NoteProperty -Name "ServerType" -Value $sty
            $obj12 | Add-Member -MemberType NoteProperty -Name "ProductName" -Value $prname
            if($prver -eq "")
            {
            $prver = "Not Integrated in PL"
            $obj12 | Add-Member -MemberType NoteProperty -Name "IntegratedVersion" -Value $prver
            }
            else
            {
            $obj12 | Add-Member -MemberType NoteProperty -Name "IntegratedVersion" -Value $prver
            }
            $prodarr += $obj12
            $outbox.AppendText("$prname - $prver`n")
            }
        }
        }
    }
    ############
    

    } #foreach ends here
    $prodarr | export-csv -Path "C:\temp\Linuxproducts.csv"

    $smtpServer="emeacas.bankofamerica.com"
    $mailid = $emailid.Text.trim()
    $pswd = $paswd.text.trim()
    $usname = $nbkid.text.trim()

    sendmail $smtpServer $mailid $usname $pswd 
    
    $runbutton.Text = "Run Script";$this.enabled =$true;
    }
    $runbutton.add_click($runevent)
    clear
    $form1.ShowDialog()| Out-Null

 
 