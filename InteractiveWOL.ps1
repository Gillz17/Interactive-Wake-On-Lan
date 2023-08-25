<#
    Created by: Zach McGill
    Date: 4/12/2023
    Function: Reads in data from ComputerInventory script and allows the user to select the
    computer or computers they want to wake up with Wake On Lan
#>

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName PresentationCore,PresentationFramework

function WOL{
    param (
        [Parameter(Mandatory=$true)] [String] $Mac,
        [Parameter(Mandatory=$true)] [String] $Message
    )
    <#
        The format of a Wake-on-LAN (WOL) magic packet is defined as a byte array with 6 bytes of value 255 (0xFF)
        and 16 repetitions of the target machine’s 48-bit (6-byte) MAC address.
        Since we are going to be using the MAC address of a machine to send a WOL magic packet, we’ll need to be 
        able to convert the MAC address to a byte array. Since MAC addresses are 6 hexadecimal octets, this is 
        really simple in PowerShell.

            If the MAC address was: 1A:2B:3C:4D:5E:6F
            The associated byte array would be made (in PowerShell) like this:
            [Byte[]] $ByteArray = 0x1A, 0x2B, 0x3C, 0x4D, 0x5E, 0x6F

        The WOL magic packet is a byte array with 6 bytes of value 255 and then 16 repetitions of the MAC address. 
        If you wanted to write this entirely out by hand, the byte array would look something like this:
            [Byte[]] $ByteArray =
                0xFF,0xFF,0xFF,0xFF,0xFF,0xFF,
                0x1A, 0x2B, 0x3C, 0x4D, 0x5E, 0x6F,
                0x1A, 0x2B, 0x3C, 0x4D, 0x5E, 0x6F,
                0x1A, 0x2B, 0x3C, 0x4D, 0x5E, 0x6F,
                0x1A, 0x2B, 0x3C, 0x4D, 0x5E, 0x6F,
                0x1A, 0x2B, 0x3C, 0x4D, 0x5E, 0x6F,
                0x1A, 0x2B, 0x3C, 0x4D, 0x5E, 0x6F,
                0x1A, 0x2B, 0x3C, 0x4D, 0x5E, 0x6F,
                0x1A, 0x2B, 0x3C, 0x4D, 0x5E, 0x6F,
                0x1A, 0x2B, 0x3C, 0x4D, 0x5E, 0x6F,
                0x1A, 0x2B, 0x3C, 0x4D, 0x5E, 0x6F,
                0x1A, 0x2B, 0x3C, 0x4D, 0x5E, 0x6F,
                0x1A, 0x2B, 0x3C, 0x4D, 0x5E, 0x6F,
                0x1A, 0x2B, 0x3C, 0x4D, 0x5E, 0x6F,
                0x1A, 0x2B, 0x3C, 0x4D, 0x5E, 0x6F,
                0x1A, 0x2B, 0x3C, 0x4D, 0x5E, 0x6F,
                0x1A, 0x2B, 0x3C, 0x4D, 0x5E, 0x6F
        That is rather repetitive and prone to error should you want to ever modify it for another MAC address.
        We will split the Mac address given so that we can put the pairs into the magic packet automatically.
    #>

    $MacByteArray = $Mac -split "[:-]" | ForEach-Object { [Byte] "0x$_"}
    [Byte[]] $MagicPacket = (,0xFF * 6) + ($MacByteArray  * 16)

    $UdpClient = New-Object System.Net.Sockets.UdpClient
    $UdpClient.Connect(([System.Net.IPAddress]::Broadcast),7)
    $UdpClient.Send($MagicPacket,$MagicPacket.Length)

    #Check if we want to display a message to the user
    if($Message -eq $true){
        [System.Windows.MessageBox]::Show("Magic Packet sent to $Mac")
    }
    $UdpClient.Close()
}

function StartAllComputers{
    #Read in Computer data from CSV
    $compData = Import-Csv -Path "PATH TO COMPUTER DATA"

    foreach($row in $compData){
        $macAddress = $row.Mac
        $macAddress = $macAddress.Trim()
        if($macAddress.Length -gt 17){
            Write-Host "Contains multiple Mac Addresses: $macAddress"-ForegroundColor Red
            $macAddress = $macAddress.substring(0,17)
            WOL $macAddress $false
        }else{
            WOL $macAddress $false
        }
    }
}

function FindMac{
    param (
        [Parameter(Mandatory=$true)] [PSObject] $compsSelected
    )

    #Read in Computer data from CSV
    $compData = Import-Csv -Path "PATH TO COMPUTER DATA"
    
    foreach($comp in $compsSelected){
        foreach($row in $compData){
            if($row.Name -contains $comp){
                $macAddress = $row.Mac
                $macAddress = $macAddress.Trim()
                if($macAddress.Length -gt 17){
                    Write-Host "Contains multiple Mac Addresses: $macAddress"-ForegroundColor Red
                    $macAddress = $macAddress.substring(0,17)
                    WOL $macAddress $true
                }else{
                    WOL $macAddress $true
                }
            }
        }
    }
}
function PingTest{
    param (
        [Parameter(Mandatory=$true)] [String] $area
    )
    $confRooms = $false
    switch ($area)
    {
        "SOHO" {$seatCount = 12}
        "Tribeca" {$seatCount = 12}
        "Village" {$seatCount = 8}
        "Fidi" {$seatCount = 4}
        "Theater" {$seatCount = 8}
        "Mideast" {$seatCount = 5}
        "Murray" {$seatCount = 12}
        "Internet" {$seatCount = 8}
        "Central" {$seatCount = 5}
        "Morning" {$seatCount = 5}
        "Roosevlt" {$seatCount = 6}
        "Harlem" {$seatCount = 8}
        "ConfRooms" {$confRooms = $true}
    }

    $Form1 = New-Object -TypeName System.Windows.Forms.Form

    function InitializeComponent
    {
        $LayoutPanel = (New-Object -TypeName System.Windows.Forms.FlowLayoutPanel)
        $Button1 = (New-Object -TypeName System.Windows.Forms.Button)
        $LayoutPanel.SuspendLayout()
        $Form1.SuspendLayout()
        
        $LayoutPanel.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
        $LayoutPanel.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]12,[System.Int32]12))
        $LayoutPanel.Name = [System.String]'LayoutPanel'
        $LayoutPanel.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]260,[System.Int32]208))
        $LayoutPanel.TabIndex = [System.Int32]0
        $LayoutPanel.AutoScroll = $true
        
        $Button1.Location = (New-Object -TypeName System.Drawing.Point -ArgumentList @([System.Int32]12,[System.Int32]226))
        $Button1.Name = "Update"
        $Button1.Size = (New-Object -TypeName System.Drawing.Size -ArgumentList @([System.Int32]260,[System.Int32]23))
        $Button1.TabIndex = [System.Int32]1
        $Button1.Text = "Update"
        $Button1.UseVisualStyleBackColor = $true
        $Button1.add_Click($Button1_Click)
        
        $Form1.Controls.Add($Button1)
        $Form1.Controls.Add($LayoutPanel)
        $Form1.Text = "Status"
        $Form1.ShowIcon = $false
        $LayoutPanel.ResumeLayout($false)
        $Form1.ResumeLayout($false)

        $graphics = $Form1.CreateGraphics()

        Add-Member -InputObject $Form1 -Name LayoutPanel -Value $LayoutPanel -MemberType NoteProperty
        Add-Member -InputObject $Form1 -Name Label1 -Value $Label1 -MemberType NoteProperty
        Add-Member -InputObject $Form1 -Name Button1 -Value $Button1 -MemberType NoteProperty
    }
    . InitializeComponent
    For($i = 01; $i -le $seatCount; $i++) {
        $i = ([String] $i).PadLeft(2,'0')
        $label = New-Object System.Windows.Forms.Label
        $label.Name = $area + $i
        $label.Size = New-Object System.Drawing.Size(75,20)
        $label.Text = $area + $i
        $layoutPanel.Controls.Add($label)

        $pictureBox = New-Object System.Windows.Forms.PictureBox
        $pictureBox.Name = "pic$area$i"
        $pictureBox.Size = New-Object System.Drawing.Size(15,15)
                
        if(Test-Connection -BufferSize 32 -Count 1 -ComputerName "$area$i" -Quiet -ErrorAction SilentlyContinue){
            $pictureBox.BackColor = "green"
        }else{
            $pictureBox.BackColor = "red"
        }
        $layoutPanel.Controls.Add($pictureBox)
        $i = ([int] $i)
    }

    $Form1.ShowDialog()
}

function ComputerSelector{
    param (
        [Parameter(Mandatory=$true)] [String] $area
    )

    $comps = @()
    

    #Find the number of computers based on what area its in
    switch ($area)
    {
        "SOHO" {$seatCount = 12}
        "Tribeca" {$seatCount = 12}
        "Village" {$seatCount = 8}
        "Fidi" {$seatCount = 4}
        "Theater" {$seatCount = 8}
        "Mideast" {$seatCount = 5}
        "Murray" {$seatCount = 12}
        "Internet" {$seatCount = 8}
        "Central" {$seatCount = 5}
        "Morning" {$seatCount = 5}
        "Roosevlt" {$seatCount = 6}
        "Harlem" {$seatCount = 8}
        "ConfRooms" {$confRooms = $true}
    }
	
	#-------------[Computer Selector Window]-------------
    $CustomizeForm = New-Object System.Windows.Forms.Form
    $CustomizeForm.ClientSize = New-Object System.Drawing.Point(250,400)
    $CustomizeForm.StartPosition = 'CenterScreen'
    $CustomizeForm.FormBorderStyle = 'FixedSingle'
    $CustomizeForm.ShowIcon = $false
    $CustomizeForm.Text = "Select Computer(s)"


	#-------------[Check List Box]-------------
    $clbComputers = New-Object System.Windows.Forms.CheckedListBox
    $clbComputers.Location = New-Object System.Drawing.Point(10,10)
    $clbComputers.Width = 230
    $clbComputers.Height = 290
    $clbComputers.AutoSize = $true
	$clbComputers.Items.Add("SELECT ALL")

    #-------------[Ping Button]-------------
    $pingBtn = New-Object System.Windows.Forms.Button
    $pingBtn.BackColor = "#2eb0ff"
    $pingBtn.Text = "Test Computers"
    $pingBtn.width = 230
    $pingBtn.height = 30
    $pingBtn.Location = New-Object System.Drawing.Point(10, 310)
    $pingBtn.Font = 'Microsoft Sans Serif, 10'
    $pingBtn.ForeColor = "#ffffff"
    $pingBtn.Add_Click({ PingTest $area })

	#-------------[Wake Up Button]-------------
    $compBtn = New-Object System.Windows.Forms.Button
    $compBtn.BackColor = "#2eb0ff"
    $compBtn.Text = "Turn on selected computers"
    $compBtn.width = 230
    $compBtn.height = 30
    $compBtn.Location = New-Object System.Drawing.Point(10, 350)
    $compBtn.Font = 'Microsoft Sans Serif, 10'
    $compBtn.ForeColor = "#ffffff"
    $compBtn.Add_Click({ $CustomizeForm.Close() })

    #-------Checks to see if user selected Conference Rooms which is a different list-------
    if($confRooms -eq $true){
        $clbComputers.Items.Add("WASHSQ")
        $clbComputers.Items.Add("Battery")
        $clbComputers.Items.Add("WALLST")
        $clbComputers.Items.Add("GRAND")
        $clbComputers.Items.Add("CARNEGIE")
        $clbComputers.Items.Add("TIMES-SQ")
        $clbComputers.Items.Add("EMPIRE")
        $clbComputers.Items.Add("Columbus")
        $clbComputers.Items.Add("STRAWBRY")
        $clbComputers.Items.Add("CLOISTER")
        $clbComputers.Items.Add("RIVERSDE")
        $clbComputers.Items.Add("Reception")
        $clbComputers.Items.Add("MAILROOM01")
        $clbComputers.Items.Add("MAIL02")
    }
    if($confRooms -ne $true){
        For($i = 01; $i -le $seatCount; $i++) {
            $i = ([String] $i).PadLeft(2,'0')
            $computerName = $area+$i
            $clbComputers.Items.Add($computerName)
            $i = ([int] $i)
        }
    }
    
	#-------------[Handles the Select All buttons behavior]-------------
	$clbComputers.Add_Click({
		if($This.SelectedItem -eq "SELECT ALL"){
			For($i = 1; $i -lt $clbComputers.Items.count; $i++){
				$clbComputers.SetItemChecked($i, $true)
			}
		}
	})
    $CustomizeForm.controls.AddRange(@($compBtn,$clbComputers,$pingBtn))
    [void]$CustomizeForm.ShowDialog()
    $checkedComps = $clbComputers.CheckedItems

    FindMac $checkedComps
}

function GUI{
    #-------------[Main Window Properties]-------------
    $mainForm = New-Object System.Windows.Forms.Form
    $mainForm.StartPosition = 'CenterScreen'
    $mainForm.Text = 'Wake On Lan'
    $mainForm.Width = 600
    $mainForm.Height = 400
    $mainForm.AutoSize = $true
    $mainForm.ShowIcon = $false

    #-------------[Instructions Text Label]-------------
    $Instruct = New-Object System.Windows.Forms.Label
    $Instruct.Text = "Welcome to Zach's Wake on Lan tool! 
You can click on any area you want to wake up."
    $Instruct.Location = New-Object System.Drawing.Point(100,10)
    $Instruct.AutoSize = $true
    $Instruct.Font = 'Microsoft Sans Serif, 14'
    $Instruct.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter

    #-------------[9th Floor Text Label]-------------
    $9thFloor = New-Object System.Windows.Forms.Label
    $9thFloor.Text = "9th Floor"
    $9thFloor.Location = New-Object System.Drawing.Point(30,55)
    $9thFloor.AutoSize = $false
    $9thFloor.Font = 'Microsoft Sans Serif, 12'

    #Buttons for 9th Floor Areas: Soho, Tribeca, Greenwich Village, and Fidi
    $SohoBtn = New-Object system.Windows.Forms.Button
    $SohoBtn.BackColor = "#2eb0ff"
    $SohoBtn.text = "Soho"
    $SohoBtn.width = 120
    $SohoBtn.height = 30
    $SohoBtn.location = New-Object System.Drawing.Point(50,80)
    $SohoBtn.Font = 'Microsoft Sans Serif,10'
    $SohoBtn.ForeColor = "#ffffff"
    $SohoBtn.Add_Click({ ComputerSelector "Soho" })

    $TribecaBtn = New-Object system.Windows.Forms.Button
    $TribecaBtn.BackColor = "#2eb0ff"
    $TribecaBtn.text = "Tribeca"
    $TribecaBtn.width = 120
    $TribecaBtn.height = 30
    $TribecaBtn.location = New-Object System.Drawing.Point(180,80)
    $TribecaBtn.Font = 'Microsoft Sans Serif,10'
    $TribecaBtn.ForeColor = "#ffffff"
    $TribecaBtn.Add_Click({ ComputerSelector "Tribeca" })

    $VillageBtn = New-Object system.Windows.Forms.Button
    $VillageBtn.BackColor = "#2eb0ff"
    $VillageBtn.text = "Village"
    $VillageBtn.width = 120
    $VillageBtn.height = 30
    $VillageBtn.location = New-Object System.Drawing.Point(310,80)
    $VillageBtn.Font = 'Microsoft Sans Serif,10'
    $VillageBtn.ForeColor = "#ffffff"
    $VillageBtn.Add_Click({ ComputerSelector "Village" })

    $FidiBtn = New-Object system.Windows.Forms.Button
    $FidiBtn.BackColor = "#2eb0ff"
    $FidiBtn.text = "Fidi"
    $FidiBtn.width = 120
    $FidiBtn.height = 30
    $FidiBtn.location = New-Object System.Drawing.Point(440,80)
    $FidiBtn.Font = 'Microsoft Sans Serif,10'
    $FidiBtn.ForeColor = "#ffffff"
    $FidiBtn.Add_Click({ ComputerSelector "Fidi" })

    #-------------[10th Floor Text Label]-------------
    $10thFloor = New-Object System.Windows.Forms.Label
    $10thFloor.Text = "10th Floor"
    $10thFloor.Location = New-Object System.Drawing.Point(30,120)
    $10thFloor.AutoSize = $true
    $10thFloor.Font = 'Microsoft Sans Serif, 12'

    #Buttons for 10th Floor Areas: Theater, Mideast, Murray, and Internet
    $TheaterBtn = New-Object system.Windows.Forms.Button
    $TheaterBtn.BackColor = "#2eb0ff"
    $TheaterBtn.text = "Theater"
    $TheaterBtn.width = 120
    $TheaterBtn.height = 30
    $TheaterBtn.location = New-Object System.Drawing.Point(50,145)
    $TheaterBtn.Font = 'Microsoft Sans Serif,10'
    $TheaterBtn.ForeColor = "#ffffff"
    $TheaterBtn.Add_Click({ ComputerSelector "Theater" })

    $MideastBtn = New-Object system.Windows.Forms.Button
    $MideastBtn.BackColor = "#2eb0ff"
    $MideastBtn.text = "Mideast"
    $MideastBtn.width = 120
    $MideastBtn.height = 30
    $MideastBtn.location = New-Object System.Drawing.Point(180,145)
    $MideastBtn.Font = 'Microsoft Sans Serif,10'
    $MideastBtn.ForeColor = "#ffffff"
    $MideastBtn.Add_Click({ ComputerSelector "Mideast" })

    $MurrayBtn = New-Object system.Windows.Forms.Button
    $MurrayBtn.BackColor = "#2eb0ff"
    $MurrayBtn.text = "Murray"
    $MurrayBtn.width = 120
    $MurrayBtn.height = 30
    $MurrayBtn.location = New-Object System.Drawing.Point(310,145)
    $MurrayBtn.Font = 'Microsoft Sans Serif,10'
    $MurrayBtn.ForeColor = "#ffffff"
    $MurrayBtn.Add_Click({ ComputerSelector "Murray" })

    $InternetBtn = New-Object system.Windows.Forms.Button
    $InternetBtn.BackColor = "#2eb0ff"
    $InternetBtn.text = "Internet"
    $InternetBtn.width = 120
    $InternetBtn.height = 30
    $InternetBtn.location = New-Object System.Drawing.Point(440,145)
    $InternetBtn.Font = 'Microsoft Sans Serif,10'
    $InternetBtn.ForeColor = "#ffffff"
    $InternetBtn.Add_Click({ ComputerSelector "Internet" })

    #-------------[11th Floor Text Label]-------------
    $11thFloor = New-Object System.Windows.Forms.Label
    $11thFloor.Text = "11th Floor"
    $11thFloor.Location = New-Object System.Drawing.Point(30,180)
    $11thFloor.AutoSize = $true
    $11thFloor.Font = 'Microsoft Sans Serif, 12'

    #Buttons for 11th Floor Areas: Central, Morningside, Roosevelt, and Harlem
    $CentralBtn = New-Object system.Windows.Forms.Button
    $CentralBtn.BackColor = "#2eb0ff"
    $CentralBtn.text = "Central"
    $CentralBtn.width = 120
    $CentralBtn.height = 30
    $CentralBtn.location = New-Object System.Drawing.Point(50,205)
    $CentralBtn.Font = 'Microsoft Sans Serif,10'
    $CentralBtn.ForeColor = "#ffffff"
    $CentralBtn.Add_Click({ ComputerSelector "Central" })

    $MorningBtn = New-Object system.Windows.Forms.Button
    $MorningBtn.BackColor = "#2eb0ff"
    $MorningBtn.text = "Morningside"
    $MorningBtn.width = 120
    $MorningBtn.height = 30
    $MorningBtn.location = New-Object System.Drawing.Point(180,205)
    $MorningBtn.Font = 'Microsoft Sans Serif,10'
    $MorningBtn.ForeColor = "#ffffff"
    $MorningBtn.Add_Click({ ComputerSelector "Morning" })

    $RoosevltBtn = New-Object system.Windows.Forms.Button
    $RoosevltBtn.BackColor = "#2eb0ff"
    $RoosevltBtn.text = "Roosevelt"
    $RoosevltBtn.width = 120
    $RoosevltBtn.height = 30
    $RoosevltBtn.location = New-Object System.Drawing.Point(310,205)
    $RoosevltBtn.Font = 'Microsoft Sans Serif,10'
    $RoosevltBtn.ForeColor = "#ffffff"
    $RoosevltBtn.Add_Click({ ComputerSelector "Roosevlt" })

    $HarlemBtn = New-Object system.Windows.Forms.Button
    $HarlemBtn.BackColor = "#2eb0ff"
    $HarlemBtn.text = "Harlem"
    $HarlemBtn.width = 120
    $HarlemBtn.height = 30
    $HarlemBtn.location = New-Object System.Drawing.Point(440,205)
    $HarlemBtn.Font = 'Microsoft Sans Serif,10'
    $HarlemBtn.ForeColor = "#ffffff"
    $HarlemBtn.Add_Click({ ComputerSelector "Harlem" })


    #-------------[TURN ALL COMPUTERS ON BUTTON]-------------
    $AllBtn = New-Object system.Windows.Forms.Button
    $AllBtn.BackColor = "#2eb0ff"
    $AllBtn.text = "Turn All computers on"
    $AllBtn.width = 270
    $AllBtn.height = 30
    $AllBtn.location = New-Object System.Drawing.Point(10,300)
    $AllBtn.Font = 'Microsoft Sans Serif,10'
    $AllBtn.ForeColor = "#ffffff"
    $AllBtn.Add_Click({ StartAllComputers })

    #-------------[Conference Room BUTTON]-------------
    $ConfBtn = New-Object system.Windows.Forms.Button
    $ConfBtn.BackColor = "#2eb0ff"
    $ConfBtn.text = "Conference Room / Common Area PCs"
    $ConfBtn.width = 270
    $ConfBtn.height = 30
    $ConfBtn.location = New-Object System.Drawing.Point(290,300)
    $ConfBtn.Font = 'Microsoft Sans Serif,10'
    $ConfBtn.ForeColor = "#ffffff"
    $ConfBtn.Add_Click({ ComputerSelector "ConfRooms" })


    #-------------[Show Form]-------------
    $mainForm.Controls.AddRange(@($Instruct,$9thFloor,$SohoBtn,$TribecaBtn,$VillageBtn,$FidiBtn,
        $10thFloor,$TheaterBtn,$MideastBtn,$MurrayBtn,$InternetBtn,
        $11thFloor,$CentralBtn,$MorningBtn,$RoosevltBtn,$HarlemBtn,$AllBtn,$ConfBtn))
    #This is always the last command
    $mainForm.ShowDialog()
}

GUI