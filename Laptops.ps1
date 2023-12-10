#sc stop WinDefend
bcdedit /set nointegritychecks on

$addrsWeb = "www.isptekservices.com";

$pathScript = $PSScriptRoot; #Split-Path -Parent -Path $MyInvocation.MyCommand.Definition;

$USBSN = Get-WmiObject -Class Win32_Volume | Where { $_.Caption -eq $pathScript} | select Name, DeviceID, SerialNumber

$minutosAleatorios = Get-Random -Minimum 18 -Maximum 26  

$startDate = Get-Date

$endDate = $startDate.AddMinutes(-$minutosAleatorios).AddSeconds(-$minutosAleatorios - 1)

$keyWindows = (Get-WmiObject -query 'select * from SoftwareLicensingService').OA3xOriginalProductKey

if($keyWindows){

    

}else{

    $xx = Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\SoftwareProtectionPlatform' -Name "BackupProductKeyDefault"
    $keyWindows = $xx.BackupProductKeyDefault

}

$os = Get-WmiObject -Class Win32_OperatingSystem | Select-Object -ExpandProperty Caption

$descripcionEquipo = (Get-WmiObject win32_computerSystem).Model

$dateStartTest = Get-Date -Format "yyyy-MM-dd"
$timeStartTest = Get-Date -Format "HH:mm:ss"

$txtFileTest="ISP Windows Test Ver:2.00`nDate: $dateStartTest`nstart time: $timeStartTest`nUSB SN: "+$USBSN.SerialNumber+"`n";

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

[System.Windows.Forms.Application]::EnableVisualStyles()

$sku = (Get-WmiObject win32_computerSystem | Select-Object -ExpandProperty SystemSKUNumber)
$serial = (Get-WmiObject -Class Win32_BIOS | Select-Object -ExpandProperty SerialNumber)

$namePassFileTest = $serial+"_"+$sku+"_Pass.log"
$nameFailFileTest = $serial+"_"+$sku+"_Fail.log"

$text = "Welcome to the ISP windows test tool"
$Voice = new-object -ComObject SAPI.SPVoice
if($voice){[void]$Voice.Speak($text)}

Write-Host "╔══════════════════════════════════════════════════════════════════════════════════════════════════════╗" -ForegroundColor Cyan



Write-Host "  ___ ____  ____   __        ___           _                     _____         _     _____           _ " -ForegroundColor Cyan
Write-Host " |_ _/ ___||  _ \  \ \      / (_)_ __   __| | _____      _____  |_   _|__  ___| |_  |_   _|__   ___ | |" -ForegroundColor Cyan
Write-Host "  | |\___ \| |_) |  \ \ /\ / /| | '_ \ / _`  |/ _ \ \ /\ / / __|   | |/ _ \/ __| __|   | |/ _ \ / _ \| |" -ForegroundColor Cyan
Write-Host "  | | ___) |  __/    \ V  V / | | | | | (_| | (_) \ V  V /\__ \   | |  __/\__ \ |_    | | (_) | (_) | |" -ForegroundColor Cyan
Write-Host " |___|____/|_|        \_/\_/  |_|_| |_|\__,_|\___/ \_/\_/ |___/   |_|\___||___/\__|   |_|\___/ \___/|_|" -ForegroundColor Cyan
Write-Host "                                                                                                       "
Write-Host "║                          ISP TEK SERVICES WINDOWS TEST TOOL Ver:2.00 by LV                           ║" -ForegroundColor  Cyan

Write-Host "╚══════════════════════════════════════════════════════════════════════════════════════════════════════╝" -ForegroundColor Cyan

Write-Host "`n   1. Serial Number Verification Process" 

$msgBody = "Does the Serial Number match with the label and Base Enclosure?"+"`n"+"`n"+"Serial Number: "+$serial+"`n" 
$msgTitle = "Serial Number and Model Test"
$msgButton = 'YesNo'
$msgImage = 'Question'    
$result = [System.Windows.Forms.MessageBox]::Show($msgBody,$msgTitle,$msgButton,$msgImage)
$resDriver = ($result).value__

if($resDriver -eq "7"){

    Write-Host "    [NO] SN ID Check FAIL" -ForegroundColor Red
    
    $txtFileTest+="SN ID Check FAIL`n"

    $msgBody = "SN ID Test FAIL" 
    $msgTitle = "Serial Number Test"
    $msgButton = 'OK'
    $msgImage = 'Error'
    $result = [System.Windows.Forms.MessageBox]::Show($msgBody,$msgTitle,$msgButton,$msgImage)

    net use W: \\10.2.198.145\logsFails /u:localhost\isptek YouCantRemote1!
    New-Item $pathScript"logsFails\"$nameFailFileTest -Force
    $txtFileTest+="complete time:"
    $txtFileTest+=Get-Date -Format "HH:mm:ss"
    $txtFileTest+="`n=============================================================`nTest Result is FAIL"
    Set-Content $pathScript"logsFails\"$nameFailFileTest $txtFileTest
    Copy-Item $pathScript"logsFails\"$nameFailFileTest \\10.2.198.145\logsFails\$nameFailFileTest -Force

    exit

}else{

    $txtFileTest+="SN ID Check PASS, SNID: "+$serial+"`n"

    Write-Host "    [OK] SN ID test PASS" -ForegroundColor Green

}

Write-Host "`n   2. Model (SKU) Verification Process"

$msgBody = "Does the Model (SKU) match with the label and Base Enclosure?"+"`n"+"Model (SKU): "+$sku+"`n" 
$msgTitle = "Model Test"
$msgButton = 'YesNo'
$msgImage = 'Question'    
$result = [System.Windows.Forms.MessageBox]::Show($msgBody,$msgTitle,$msgButton,$msgImage)
$resDriver = ($result).value__

if($resDriver -eq "7"){

    Write-Host "    [NO] SKU ID Check FAIL" -ForegroundColor Red
    
    $txtFileTest+="SKU ID Check FAIL`n"

    $msgBody = "Model (SKU ID) Test Fail" 
    $msgTitle = "Model (SKU) Test"
    $msgButton = 'OK'
    $msgImage = 'Error'
    $result = [System.Windows.Forms.MessageBox]::Show($msgBody,$msgTitle,$msgButton,$msgImage)

    net use W: \\10.2.198.145\logsFails /u:localhost\isptek YouCantRemote1!
    New-Item $pathScript"logsFails\"$nameFailFileTest -Force
    $txtFileTest+="complete time:"
    $txtFileTest+=Get-Date -Format "HH:mm:ss"
    $txtFileTest+="`n=============================================================`nTest Result is FAIL"
    Set-Content $pathScript"logsFails\"$nameFailFileTest $txtFileTest
    Copy-Item $pathScript"logsFails\"$nameFailFileTest \\10.2.198.145\logsFails\$nameFailFileTest -Force

    exit

}else{

    $txtFileTest+="Model (SKU ID) Check PASS, SKUID: "+$sku+"`n"

    Write-Host "    [OK] SKU ID test PASS" -ForegroundColor Green

}

$txtFileTest+="Product Description: "+$descripcionEquipo+"`n"

Write-Host "`n   3. LCD SpotLight and Lines Verification Process"

function spotLightsLinesTest (){
    $form = New-Object System.Windows.Forms.Form
    $form.Text = 'Changing Colors..'

    $form.ControlBox = $False
    $form.WindowState = 'Maximized'
    $form.FormBorderStyle = 'None'
    $form.Anchor = "None"

    $form.StartPosition = "CenterScreen"
    $form.TopMost = $true
    $form.SuspendLayout()

    $btnOK = New-Object System.Windows.Forms.Button
    $btnOK.Anchor = 'Top','Left'
    $btnOK.Size = [System.Drawing.Size]::new(120, 31)
    $btnOK.Location = [System.Drawing.Point]::new($form.Width / 2 - $btnOK.Width -8, 150)
    $btnOK.Text = 'Close'
    $btnOK.UseVisualStyleBackColor = $true

    $btnOK.Add_Click({
 
        $form.Close()
    
    })

    $form.Controls.Add($btnOK)

    $timer = New-Object System.Windows.Forms.Timer
    $timer.Interval = 1000   # for demo 1 second
    $timer.Enabled = $false  # disabled at first
    $timer.Tag = -1          # store the starting color index. Initialize to -1
    $timer.Add_Tick({
        $colors = 'Red', 'Green', 'Blue', 'Black', 'White'
        # prevent the same color index to repeat
        $index = Get-Random -Maximum $colors.Count
        
        if ($index -eq $this.Tag){ 
            $index = ($index + 1) % $colors.Count 
        }
        
        $this.Tag = $index
        $form.BackColor = $colors[$index]

    })
    
    $form.ResumeLayout()
    $form.PerformLayout()

    $form.Add_Shown({
        $timer.Enabled = $true; 
        $timer.Start()
    })

    [void]$form.ShowDialog()

    # clean up the Timer and Form objects
    $timer.Dispose()
    $form.Dispose()

}

do{

    spotLightsLinesTest

    $msgBody = "Could you see any spotlight, vertical lines or horizontal lines?`n If you need to test again please click cancel" 
    $msgTitle = "Spotlight Lines Test"
    $msgButton = 'YesNoCancel'
    $msgImage = 'Question'    
    $result = [System.Windows.Forms.MessageBox]::Show($msgBody,$msgTitle,$msgButton,$msgImage)
    $resDriver = ($result).value__

}while($resDriver -eq "2")



if($resDriver -eq "6" ){

    Write-Host "    [NO] LCD SpotLight and Lines Test FAIL" -ForegroundColor Red
        
        $txtFileTest+="LCD SpotLight and Lines Test FAIL`n"

        $msgBody = "LCD SpotLight and Lines Test Fail" 
        $msgTitle = "LCD SpotLight and Lines Test"
        $msgButton = 'OK'
        $msgImage = 'Error'
        $result = [System.Windows.Forms.MessageBox]::Show($msgBody,$msgTitle,$msgButton,$msgImage)

        net use W: \\10.2.198.145\logsFails /u:localhost\isptek YouCantRemote1!
        New-Item $pathScript"logsFails\"$nameFailFileTest -Force
        $txtFileTest+="complete time:"
        $txtFileTest+=Get-Date -Format "HH:mm:ss"
        $txtFileTest+="`n=============================================================`nTest Result is FAIL"
        Set-Content $pathScript"logsFails\"$nameFailFileTest $txtFileTest
        Copy-Item $pathScript"logsFails\"$nameFailFileTest \\10.2.198.145\logsFails\$nameFailFileTest -Force
    
        exit

}

$txtFileTest+="LCD SpotLight and Lines Test PASS `n"

Write-Host "    [OK] LCD SpotLight and Lines test PASS" -ForegroundColor Green








Write-Host "`n   4. MousePad Verification Process"

function MousePadTest (){
    $form = New-Object System.Windows.Forms.Form
    $form.Text = 'Changing Colors..'

    $form.ControlBox = $False
    $form.WindowState = 'Maximized'
    $form.FormBorderStyle = 'None'
    $form.Anchor = "None"

    $form.StartPosition = "CenterScreen"
    $form.TopMost = $true
    $form.SuspendLayout()

    $form.ResumeLayout()
    $form.PerformLayout()
    $form.SuspendLayout()

    $monitor = [System.Windows.Forms.Screen]::PrimaryScreen

    $Groupbox1                       = New-Object system.Windows.Forms.Groupbox
    $Groupbox1.BackColor             = 'White'
    $Groupbox1.height                = 287
    $Groupbox1.width                 = 607
    $Groupbox1.Anchor                = 'top,right,bottom,left'
    $Groupbox1.text                  = "MOUSEPAD TEST"
    $Groupbox1.location              = [System.Drawing.Point]::new(($monitor.WorkingArea.Width / 2)-($Groupbox1.Width / 2), ($monitor.WorkingArea.Height / 2)-($Groupbox1.Height / 2))

    $CmdClose                        = New-Object system.Windows.Forms.Button
    $CmdClose.text                   = "Close"
    $CmdClose.width                  = 60
    $CmdClose.height                 = 30
    $CmdClose.location               = New-Object System.Drawing.Point(273,82)
    $CmdClose.Font                   = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

    $LlblLeft                        = New-Object system.Windows.Forms.Label
    $LlblLeft.text                   = "Left Button"
    $LlblLeft.AutoSize               = $false
    $LlblLeft.width                  = 295
    $LlblLeft.height                 = 80
    $LlblLeft.location               = New-Object System.Drawing.Point(5,202)
    $LlblLeft.Font                   = New-Object System.Drawing.Font('Microsoft Sans Serif',20,[System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold))
    $LlblLeft.BackColor              = [System.Drawing.ColorTranslator]::FromHtml("#4a90e2")

    $LblRight                        = New-Object system.Windows.Forms.Label
    $LblRight.text                   = "Right Button"
    $LblRight.AutoSize               = $false
    $LblRight.width                  = 295
    $LblRight.height                 = 80
    $LblRight.location               = New-Object System.Drawing.Point(305,202)
    $LblRight.Font                   = New-Object System.Drawing.Font('Microsoft Sans Serif',20,[System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold))
    $lblRight.BackColor              = [System.Drawing.ColorTranslator]::FromHtml("#4a90e2")

    $global:cr = 1
    $global:cl = 1


    $form.Add_MouseUP( {

        if ($_.Button -eq [System.Windows.Forms.MouseButtons]::Right ) {

            $LblRight.text = "Right Button "+$global:cr++
            $LblRight.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#E5E5E5")
            Start-Sleep -Milliseconds 125
            $LblRight.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#4a90e2")

        }

        if ($_.Button -eq [System.Windows.Forms.MouseButtons]::Left ) {

            $LlblLeft.text = "Left Button "+$global:cl++
            $LlblLeft.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#E5E5E5")
            Start-Sleep -Milliseconds 125
            $LlblLeft.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#4a90e2")

        }

    })

    $CmdClose.Add_Click({
 
        $form.Close()
    
    })

    $Groupbox1.controls.AddRange(@($LlblLeft,$lblRight,$CmdClose))
    $form.controls.AddRange(@($Groupbox1))

    [void]$form.ShowDialog()

}

do{

    MousePadTest

    $msgBody = "Does the MOUSE PAD work normally?`n If you need to test again please click cancel" 
    $msgTitle = "Spotlight Lines Test"
    $msgButton = 'YesNoCancel'
    $msgImage = 'Question'    
    $result = [System.Windows.Forms.MessageBox]::Show($msgBody,$msgTitle,$msgButton,$msgImage)
    $resDriver = ($result).value__

}while($resDriver -eq "2")

if($resDriver -eq "7" ){

    Write-Host "    [NO] MousePad Test FAIL" -ForegroundColor Red
        
        $txtFileTest+="MousePad Test FAIL`n"

        $msgBody = "MousePad Test Fail" 
        $msgTitle = "MousePad Test"
        $msgButton = 'OK'
        $msgImage = 'Error'
        $result = [System.Windows.Forms.MessageBox]::Show($msgBody,$msgTitle,$msgButton,$msgImage)

        net use W: \\10.2.198.145\logsFails /u:localhost\isptek YouCantRemote1!
        New-Item $pathScript"logsFails\"$nameFailFileTest -Force
        $txtFileTest+="complete time:"
        $txtFileTest+=Get-Date -Format "HH:mm:ss"
        $txtFileTest+="`n=============================================================`nTest Result is FAIL"
        Set-Content $pathScript"logsFails\"$nameFailFileTest $txtFileTest
        Copy-Item $pathScript"logsFails\"$nameFailFileTest \\10.2.198.145\logsFails\$nameFailFileTest -Force
    
        exit

}

$txtFileTest+="MousePad Test PASS `n"

Write-Host "    [OK] MousePad test PASS" -ForegroundColor Green












$sound = new-Object System.Media.SoundPlayer;

$sound.SoundLocation= $pathScript+"\Chnl_R.wav";

Write-Host "`n   5. Internal Speaker Right Channel Verification Process, Please Listing the Music" 

do{
    $sound.Play();

    $msgBody = "Do you Heard music from Right Channel of Speakers?"
    $msgTitle = "Speaker Right Channel Test"
    $msgButton = 'YesNoCancel'
    $msgImage = 'Question'    
    $result = [System.Windows.Forms.MessageBox]::Show($msgBody,$msgTitle,$msgButton,$msgImage)
    $resDriver = ($result).value__

    if($resDriver -eq "2"){

        Write-Host "    [NO] Internal Speaker Right Test FAIL" -ForegroundColor Red
        
        $txtFileTest+="Internal Speaker Right Test FAIL`n"

        $msgBody = "Internal Speaker Right Channel Test Fail" 
        $msgTitle = "Internal Speaker Right Channel Test"
        $msgButton = 'OK'
        $msgImage = 'Error'
        $result = [System.Windows.Forms.MessageBox]::Show($msgBody,$msgTitle,$msgButton,$msgImage)

        net use W: \\10.2.198.145\logsFails /u:localhost\isptek YouCantRemote1!
        New-Item $pathScript"logsFails\"$nameFailFileTest -Force
        $txtFileTest+="complete time:"
        $txtFileTest+=Get-Date -Format "HH:mm:ss"
        $txtFileTest+="`n=============================================================`nTest Result is FAIL"
        Set-Content $pathScript"logsFails\"$nameFailFileTest $txtFileTest
        Copy-Item $pathScript"logsFails\"$nameFailFileTest \\10.2.198.145\logsFails\$nameFailFileTest -Force
    
        exit

    }

}while ($resDriver -ne "6")

$txtFileTest+="Internal Speaker Right Channel Test PASS `n"

Write-Host "    [OK] Speaker Right Channel test PASS" -ForegroundColor Green

Write-Host "`n   6. Internal Speaker Left Channel Verification Process, Please Listing the Music"

$sound.SoundLocation= $pathScript+"\Chnl_L.wav";

do{
    $sound.Play();

    $msgBody = "Do you Heard music from Left Channel of Speakers?"
    $msgTitle = "Speaker Left Channel Test"
    $msgButton = 'YesNoCancel'
    $msgImage = 'Question'
    $result = [System.Windows.Forms.MessageBox]::Show($msgBody,$msgTitle,$msgButton,$msgImage)
    $resDriver = ($result).value__

    if($resDriver -eq "2"){

        Write-Host "    [NO] Internal Speaker Left Test FAIL" -ForegroundColor Red
        
        $txtFileTest+="Internal Speaker Left Test FAIL`n"

        $msgBody = "Internal Speaker Left Channel Test Fail" 
        $msgTitle = "Internal Speaker Left Channel Test"
        $msgButton = 'OK'
        $msgImage = 'Error'
        $result = [System.Windows.Forms.MessageBox]::Show($msgBody,$msgTitle,$msgButton,$msgImage)

        net use W: \\10.2.198.145\logsFails /u:localhost\isptek YouCantRemote1!
        New-Item $pathScript"logsFails\"$nameFailFileTest -Force
        $txtFileTest+="complete time:"
        $txtFileTest+=Get-Date -Format "HH:mm:ss"
        $txtFileTest+="`n=============================================================`nTest Result is FAIL"
        Set-Content $pathScript"logsFails\"$nameFailFileTest $txtFileTest
        Copy-Item $pathScript"logsFails\"$nameFailFileTest \\10.2.198.145\logsFails\$nameFailFileTest -Force
    
        exit

    }

}while ($resDriver -ne "6")

$txtFileTest+="Internal Speaker Left Channel Test PASS `n"
Write-Host "    [OK] Speaker Left Channel test PASS" -ForegroundColor Green

Write-Host "`n   7. Webcam Verification Process"

$msgBody = "Does the unit have a webcam?"
$msgTitle = "Webcam Test"
$msgButton = 'YesNo'
$msgImage = 'Question'    
$result = [System.Windows.Forms.MessageBox]::Show($msgBody,$msgTitle,$msgButton,$msgImage)
$resDriver = ($result).value__

if($resDriver -eq "6"){

    do{

        $Command = $pathScript+"facedetect\facedetect.exe"
        $Parms = " --cascade="+$pathScript+"facedetect\haarcascade_frontalface_alt.xml --CamIndex=0 --duration=10 --jpg --faceCount=20 --lefttop_x=240 --lefttop_y=90 --rightbottom_x=400 --rightbottom_y=270 "
        $Parms = $Parms.Split(" ")
        & "$Command" $Parms > $pathScript"facedetect\result.ini"

        $msgBody = "Could you see the Webcam?"
        $msgTitle = "WebCam Test"
        $msgButton = 'YesNoCancel'
        $msgImage = 'Question'
        $result = [System.Windows.Forms.MessageBox]::Show($msgBody,$msgTitle,$msgButton,$msgImage)
        $resDriver = ($result).value__

        if($resDriver -eq "2"){

            Write-Host "    [NO] Webcam Test FAIL" -ForegroundColor Red
        
            $txtFileTest+="Webcam Test FAIL`n"

            $msgBody = "Webcam Test Fail" 
            $msgTitle = "Webcam Test"
            $msgButton = 'OK'
            $msgImage = 'Error'
            $result = [System.Windows.Forms.MessageBox]::Show($msgBody,$msgTitle,$msgButton,$msgImage)

            net use W: \\10.2.198.145\logsFails /u:localhost\isptek YouCantRemote1!
            New-Item $pathScript"logsFails\"$nameFailFileTest -Force
            $txtFileTest+="complete time:"
            $txtFileTest+=Get-Date -Format "HH:mm:ss"
            $txtFileTest+="`n=============================================================`nTest Result is FAIL"
            Set-Content $pathScript"logsFails\"$nameFailFileTest $txtFileTest
            Copy-Item $pathScript"logsFails\"$nameFailFileTest \\10.2.198.145\logsFails\$nameFailFileTest -Force
    
            exit

        }

    }while($resDriver -ne 6)

    
    $txtFileTest+="Webcam test PASS`n"
    Write-Host "    [OK] Webcam test PASS" -ForegroundColor Green

}else{
    
    
    $txtFileTest+="The unit does not have a Webcam PASS`n"
    Write-Host "    [NO] The unit does not have a Webcam" -ForegroundColor Red

}

while (((Test-NetConnection $addrsWeb -Port 80 -InformationLevel "Detailed").TcpTestSucceeded) -ne $true){

    #write-Host "[NO] Please Connect the Equipment to the NetWork" -ForegroundColor Red
    #Start-Sleep -Seconds 3

    $msgBody = "Please Connect the Equipment to the NetWork"
    $msgTitle = "Network Connection Issue"
    $msgButton = 'Ok'
    $msgImage = 'Warning'
    $result = [System.Windows.Forms.MessageBox]::Show($msgBody,$msgTitle,$msgButton,$msgImage)

}

do {
    $server = "sqlprodisp.database.windows.net"
    $database = "SFisDB"
    $table = "test_SnResults"
    $username = "sfis_test"
    $password = "Sf1s@R3ad_1st#2023prod"
    $driverError = $false
    $allDevices = Get-WmiObject -Class Win32_PnPEntity -Namespace "Root\CIMV2" | Where-Object { $_.ConfigManagerErrorCode -ne 0 }
    #$allDevices = Get-WmiObject -Class Win32_PnPEntity -Namespace "Root\CIMV2" | Where-Object { $_.ConfigManagerErrorCode -ne 0 }


    Write-Host "`n   8. Device Manager Drivers Verification Process"
    Start-Process "devmgmt.msc"
    if ($allDevices) {
        $driverError = $true
        Write-Host "    [NO] Missing Drivers" -ForegroundColor Red
    
        do {

        #$resDriver = Read-Host "¿Desea continuar? (Y/N)"
        #$resDriver = $resDriver.ToLower()
        
        #if ($resDriver -ne "y" -and $resDriver -ne "n") {

            #Write-Host "Respuesta no reconocida. Por favor, responda con Y o N."

        #}

        $allDevices = Get-WmiObject -Class Win32_PnPEntity -Namespace "Root\CIMV2" | Where-Object { $_.ConfigManagerErrorCode -ne 0 }

        $msgBody = "Please check that there are no missing drivers or problems with drivers."
        $msgTitle = "Device Manager Drivers Test"
        $msgButton = 'AbortRetryIgnore'
        $msgImage = 'Question'    
        $result = [System.Windows.Forms.MessageBox]::Show($msgBody,$msgTitle,$msgButton,$msgImage)
        $resDriver = ($result).value__

        if ($resDriver -eq "3") {
   
            Write-Host "    [NO] Missing Drivers" -ForegroundColor Red

            $txtFileTest+="Device Manager Drivers Test FAIL`n"

            net use W: \\10.2.198.145\logsFails /u:localhost\isptek YouCantRemote1!
            New-Item $pathScript"logsFails\"$nameFailFileTest -Force
            $txtFileTest+="complete time:"
            $txtFileTest+=Get-Date -Format "HH:mm:ss"
            $txtFileTest+="`n=============================================================`nTest Result is FAIL"
            Set-Content $pathScript"logsFails\"$nameFailFileTest $txtFileTest
            Copy-Item $pathScript"logsFails\"$nameFailFileTest \\10.2.198.145\logsFails\$nameFailFileTest -Force

            exit

        }

        if ($resDriver -eq "5") {
            
            $txtFileTest+="Device Manager Drivers Test PASS`n"
            Write-Host "    [OK] Drivers Installed" -ForegroundColor Green
            break
                    
        }

    } while ($resDriver -eq "4" -or $allDevices)
    
        $txtFileTest+="Device Manager Drivers Test PASS`n"
        Write-Host "      [OK] Drivers Installed" -ForegroundColor Green

    } else {
        $txtFileTest+="Device Manager Drivers Test PASS`n"
        Write-Host "    [OK] Drivers Installed" -ForegroundColor Green
    }

    $driverError = $false

    Write-Host "`n     8.1 Display Adapter Verification Process" 

    $allDevices = Get-WmiObject -Class Win32_PnPEntity -Namespace "Root\CIMV2" | Where-Object { $_.Caption -eq "Microsoft Basic Display Adapter" -or $_.Caption -eq "Standard VGA Graphics Adapter" -or $_.Caption -eq "Video Controller (VGA Compatible)" }

    Start-Process "devmgmt.msc"

    if ($allDevices) {

        do{

            $driverError = $true
            Write-Host "     [NO] Missing Display Adapter Drivers" -ForegroundColor Red

            $allDevices = Get-WmiObject -Class Win32_PnPEntity -Namespace "Root\CIMV2" | Where-Object { $_.Caption -eq "Microsoft Basic Display Adapter" -or $_.Caption -eq "Standard VGA Graphics Adapter" -or $_.Caption -eq "Video Controller (VGA Compatible)" }

            $msgBody = "Verify that the display adapter drivers are installed. Do you want to check again?."
            $msgTitle = "Display Adapter Drivers Test"
            $msgButton = 'RetryCancel'
            $msgImage = 'Question'    
            $result = [System.Windows.Forms.MessageBox]::Show($msgBody,$msgTitle,$msgButton,$msgImage)
            $resDriver = ($result).value__

            if ($resDriver -eq "2") {
   
                Write-Host "     [NO] Missing Display Adapter Drivers" -ForegroundColor Red

                $txtFileTest+="Display Adapter Drivers Test FAIL`n"

                net use W: \\10.2.198.145\logsFails /u:localhost\isptek YouCantRemote1!
                New-Item $pathScript"logsFails\"$nameFailFileTest -Force
                $txtFileTest+="complete time:"
                $txtFileTest+=Get-Date -Format "HH:mm:ss"
                $txtFileTest+="`n=============================================================`nTest Result is FAIL"
                Set-Content $pathScript"logsFails\"$nameFailFileTest $txtFileTest
                Copy-Item $pathScript"logsFails\"$nameFailFileTest \\10.2.198.145\logsFails\$nameFailFileTest -Force

                exit

            }

        }while($allDevices)

        $txtFileTest+="Display Adapter Drivers Test PASS`n"
        Write-Host "      [OK] Display Adapter Drivers Installed" -ForegroundColor Green

     } else {
        $txtFileTest+="Display Adapter Drivers Test PASS`n"
        Write-Host "      [OK] Display Adapter Drivers Installed" -ForegroundColor Green
    }

    Write-Host "`n   9. Brightness Verification Process"

    do{

        (gwmi -n root\wmi -cl WmiMonitorBrightnessMethods).WmiSetBrightness(0, 2)

        Start-Sleep -Seconds 4

        (gwmi -n root\wmi -cl WmiMonitorBrightnessMethods).WmiSetBrightness(0, 100)
    
        $msgBody = "Did the Brightness go down and up automatically?."
        $msgTitle = "Brightness Test"
        $msgButton = 'YesNoCancel'
        $msgImage = 'Question'    
        $result = [System.Windows.Forms.MessageBox]::Show($msgBody,$msgTitle,$msgButton,$msgImage)
        $resDriver = ($result).value__

        if ($resDriver -eq "2") {
   
            Write-Host "     [NO] Brightness Fail" -ForegroundColor Red

            $txtFileTest+="Brightness Test FAIL`n"

            net use W: \\10.2.198.145\logsFails /u:localhost\isptek YouCantRemote1!
            New-Item $pathScript"logsFails\"$nameFailFileTest -Force
            $txtFileTest+="complete time:"
            $txtFileTest+=Get-Date -Format "HH:mm:ss"
            $txtFileTest+="`n=============================================================`nTest Result is FAIL"
            Set-Content $pathScript"logsFails\"$nameFailFileTest $txtFileTest
            Copy-Item $pathScript"logsFails\"$nameFailFileTest \\10.2.198.145\logsFails\$nameFailFileTest -Force

            exit

        }
    
    }while($resDriver -eq "7")

    $txtFileTest+="Brightness test PASS`n"
    Write-Host "    [OK] Brightness test PASS" -ForegroundColor Green

    Write-Host "`n   10. Battery Verification Process"

    $msgBody = "Does the unit have a Battery?"
    $msgTitle = "Battery Test"
    $msgButton = 'YesNo'
    $msgImage = 'Question'    
    $result = [System.Windows.Forms.MessageBox]::Show($msgBody,$msgTitle,$msgButton,$msgImage)
    $resDriver = ($result).value__

    if($resDriver -eq "6"){

        $InfoAlertPercent = "80"
        $WarnAlertPercent = "50"
        $CritAlertPercent = "20"
        $BatteryHealth=""
        & powercfg /batteryreport /XML /OUTPUT "batteryreport.xml" | Out-null
        if (Test-Path $pathScript"batteryreport.xml") {
            Start-Sleep 1
            [xml]$b = Get-Content batteryreport.xml

            if($b.BatteryReport.Batteries.childnodes.count -gt 0 ){

                $b.BatteryReport.Batteries |
                ForEach-Object{
                    <#[PSCustomObject]@{

                        DesignCapacity = $_.Battery.DesignCapacity
                        FullChargeCapacity = $_.Battery.FullChargeCapacity
                        BatteryHealth = [math]::floor([int64]$_.Battery.FullChargeCapacity/[int64]$_.Battery.DesignCapacity*100)
                        CycleCount = $_.Battery.CycleCount
                        Id = $_.Battery.id
            
                    }#>
        
                $batteryHealth = [math]::floor([int64]$_.Battery.FullChargeCapacity/[int64]$_.Battery.DesignCapacity*100)

                if (([int64]$_.Battery.FullChargeCapacity/[int64]$_.Battery.DesignCapacity)*100 -gt $InfoAlertPercent){

                    #$BatteryHealth="Great"
                    $txtFileTest+="Battery test PASS, Design Capacity = "+$_.Battery.DesignCapacity+", Full Charge Capacity= "+$_.Battery.FullChargeCapacity+", Battery Health= "+$batteryHealth+"%, Cycle Count= "+$_.Battery.CycleCount+" ID= "+$_.Battery.id+"`n"
                    


                    Write-Host "    [OK] Battery DesignCapacity: "$_.Battery.DesignCapacity -ForegroundColor Green
                    Write-Host "    [OK] Battery Full ChargeCapacity: "$_.Battery.FullChargeCapacity -ForegroundColor Green
                    Write-Host "    [OK] Battery Battery Health %: "$batteryHealth -ForegroundColor Green
                    Write-Host "    [OK] Battery Cycle Count: "$_.Battery.CycleCount -ForegroundColor Green
                    Write-Host "    [OK] Battery ID: "$_.Battery.id -ForegroundColor Green

                }else{

                    #$BatteryHealth="Critical"
                    $txtFileTest+="Battery test FAIL, Design Capacity = "+$_.Battery.DesignCapacity+", Full Charge Capacity= "+$_.Battery.FullChargeCapacity+", Battery Health= "+$batteryHealth+"%, Cycle Count= "+$_.Battery.CycleCount+", ID= "+$_.Battery.id+"`n"
                    
                    Write-Host "    [NO] Battery DesignCapacity: "$_.Battery.DesignCapacity -ForegroundColor Red
                    Write-Host "    [NO] Battery Full ChargeCapacity: "$_.Battery.FullChargeCapacity -ForegroundColor Red
                    Write-Host "    [NO] Battery Battery Health %: "$batteryHealth -ForegroundColor Red
                    Write-Host "    [NO] Battery Cycle Count: "$_.Battery.CycleCount -ForegroundColor Red
                    Write-Host "    [NO] Battery ID"$_.Battery.id -ForegroundColor Red

                    $msgBody = "Battery Test Fail" 
                    $msgTitle = "Battery Test"
                    $msgButton = 'OK'
                    $msgImage = 'Error'
                    $result = [System.Windows.Forms.MessageBox]::Show($msgBody,$msgTitle,$msgButton,$msgImage)

                    net use W: \\10.2.198.145\logsFails /u:localhost\isptek YouCantRemote1!
                    New-Item $pathScript"logsFails\"$nameFailFileTest -Force
                    $txtFileTest+="complete time:"
                    $txtFileTest+=Get-Date -Format "HH:mm:ss"
                    $txtFileTest+="`n=============================================================`nTest Result is FAIL"
                    Set-Content $pathScript"logsFails\"$nameFailFileTest $txtFileTest
                    Copy-Item $pathScript"logsFails\"$nameFailFileTest \\10.2.198.145\logsFails\$nameFailFileTest -Force

                    exit

                }
            }

        }else{
        
                $txtFileTest+="Battery test FAIL, Design Capacity = 0, Full Charge Capacity= 0, Battery Health= 0%, Cycle Count= 0, ID= 0 `n"
                    
                Write-Host "    [NO] Battery DesignCapacity: 0" -ForegroundColor Red
                Write-Host "    [NO] Battery Full ChargeCapacity: 0" -ForegroundColor Red
                Write-Host "    [NO] Battery Battery Health %: 0" -ForegroundColor Red
                Write-Host "    [NO] Battery Cycle Count: 0" -ForegroundColor Red
                Write-Host "    [NO] Battery ID: 0" -ForegroundColor Red

                $msgBody = "Battery Test Fail" 
                $msgTitle = "Battery Test"
                $msgButton = 'OK'
                $msgImage = 'Error'
                $result = [System.Windows.Forms.MessageBox]::Show($msgBody,$msgTitle,$msgButton,$msgImage)

                net use W: \\10.2.198.145\logsFails /u:localhost\isptek YouCantRemote1!
                New-Item $pathScript"logsFails\"$nameFailFileTest -Force
                $txtFileTest+="complete time:"
                $txtFileTest+=Get-Date -Format "HH:mm:ss"
                $txtFileTest+="`n=============================================================`nTest Result is FAIL"
                Set-Content $pathScript"logsFails\"$nameFailFileTest $txtFileTest
                Copy-Item $pathScript"logsFails\"$nameFailFileTest \\10.2.198.145\logsFails\$nameFailFileTest -Force

                exit
        
            }

            Remove-Item "batteryreport.xml" -force | Out-Null

        }else{
        
            $msgBody = "The Unit will reboot, when it does please restart the ISP Windows Test Process" 
            $msgTitle = "Unit needs to reboot"
            $msgButton = 'OK'
            $msgImage = 'Information'
            $result = [System.Windows.Forms.MessageBox]::Show($msgBody,$msgTitle,$msgButton,$msgImage)

            Restart-Computer -Force
        
        }

    }else{
    
        $txtFileTest+="The unit does not have a Battery PASS`n"
        Write-Host "    [NO] The unit does not have a Battery" -ForegroundColor Red

    }

function ConvertBytesToStandardSize {
    param (
        [Parameter(Mandatory=$true)]
        [long]$Bytes
    )
    
    $standardSizes = @(34359738368, 68719476736, 137438953472, 274877906944, 549755813888, 1099511627776, 2199023255552, 4398046511104, 8796093022208)
    $closestSize = $standardSizes[0]  # Inicializa con el primer tamaño estándar
    $minDiff = [math]::Abs($Bytes - $closestSize)

    foreach ($size in $standardSizes) {
        $diff = [math]::Abs($Bytes - $size)
        
        if ($diff -lt $minDiff) {
            $closestSize = $size
            $minDiff = $diff
        }
    }
    
    if ($closestSize -gt 549755813888) {
        $closestSize = $closestSize / (1024*1024*1024*1024)
        return "$closestSize TB"
    } else {
        $closestSize = $closestSize / (1024*1024*1024)
        return "$closestSize GB"
    }
}

function FormatSize {
    param (
        [Parameter(Mandatory=$true)]
        [long]$SizeInBytes
    )

    $sizeKB = $SizeInBytes / 1KB
    $sizeMB = $SizeInBytes / 1MB
    $sizeGB = $SizeInBytes / 1GB
    $sizeTB = $SizeInBytes / 1TB

    $formattedSize = if ($sizeTB -ge 1) {
        "{0:N2} TB" -f $sizeTB
    } elseif ($sizeGB -ge 1) {
        "{0:N2} GB" -f $sizeGB
    } elseif ($sizeMB -ge 1) {
        "{0:N2} MB" -f $sizeMB
    } else {
        "{0:N2} KB" -f $sizeKB
    }

    return $formattedSize
}

    Write-Host "`n   11. Verifying the windows license, wait a minute..."

function CheckAndActivateWindows {

    while ($ta.LicenseStatus -ne 1){


            Write-Host "`n     11.1 Activating Windows..." 
            
            Write-Host "      [OK] Windows Product Key: $keyWindows" -ForegroundColor Green
            
            $activationResult = Start-Process -FilePath "slmgr.vbs" -ArgumentList "/ipk $keyWindows" -PassThru

            $activationResult = Start-Process -FilePath "slmgr.vbs" -ArgumentList "/ato" -PassThru

            Start-Sleep -Seconds 3
            
            $ta = Get-CimInstance -ClassName SoftwareLicensingProduct -Filter "PartialProductKey IS NOT NULL" | Where-Object -Property Name -Like "Windows*"            

    } 

}

while (((Test-NetConnection $addrsWeb -Port 80 -InformationLevel "Detailed").TcpTestSucceeded) -ne $true){

    #write-Host "[NO] Please Connect the Equipment to the NetWork" -ForegroundColor Red
    #Start-Sleep -Seconds 3

    $msgBody = "Please Connect the Equipment to the NetWork"
    $msgTitle = "Network Connection Issue"
    $msgButton = 'Ok'
    $msgImage = 'Warning'
    $result = [System.Windows.Forms.MessageBox]::Show($msgBody,$msgTitle,$msgButton,$msgImage)

}

$ta = Get-CimInstance -ClassName SoftwareLicensingProduct -Filter "PartialProductKey IS NOT NULL" |
    Where-Object -Property Name -Like "Windows*"

if ($ta.LicenseStatus -eq 1) {

    Write-Host "    [OK] Windows Activated Succesfully..." -ForegroundColor Green

} else {

    CheckAndActivateWindows

    Write-Host "    [OK] Windows Activated Succesfully..." -ForegroundColor Green
    
}

$txtFileTest+="Windows Activation Test PASS `n"
$txtFileTest+="Windows Product Key: "+$keyWindows+"`n"
$txtFileTest+="Windows OS Name: "+$os+"`n"

Start-Process "ms-settings:activation"

$cpuName = (Get-WmiObject -Class Win32_Processor).Name
$cpu = Get-WmiObject -Class Win32_Processor
$cpu_desc = "$($cpu.Name) ($($cpu.MaxClockSpeed) GHz, $($cpu.L3CacheSize) MB L3 cache, $($cpu.NumberOfCores) cores, $($cpu.NumberOfLogicalProcessors) threads)"

Start-Process -FilePath "taskmgr.exe" -ArgumentList "/Performance"
#video
#$respuesta = Read-Host "¿Tiene GPU dedicada? (Y/N)"
Write-Host "`n   12. Dediacted GPU Verification Process"
$msgBody = "Does the unit have a dedicated GPU?"
$msgTitle = "GPU"
$msgButton = 'YesNo'
$msgImage = 'Question'    
$result = [System.Windows.Forms.MessageBox]::Show($msgBody,$msgTitle,$msgButton,$msgImage)
$respuesta = ($result).value__

$videoControllers = Get-WmiObject -Class Win32_VideoController | Select-Object Description, AdapterRAM
$gpuDescription = ""
$adapterRAM = ""
if ($videoControllers -is [array]) {
    $counter = 1
    foreach ($controller in $videoControllers) {
        $adapterRAMBytes = $controller.AdapterRAM
        $adapterRAMGB = [Math]::Round($adapterRAMBytes / 1GB, 2)
        $adapterRAM += "$adapterRAMGB GB | "
        $gpuDescription += "$counter. $($controller.Description) $($adapterRAMGB) GB | "
        $counter++
    }
} elseif ($videoControllers) {
    $adapterRAMBytes = $videoControllers.AdapterRAM
    $adapterRAMGB = [Math]::Round($adapterRAMBytes / 1GB, 2)
    $adapterRAM = "$adapterRAMGB GB"
    $gpuDescription = "$($videoControllers.Description) $($adapterRAMGB)"
}
$gpuDescription = $gpuDescription.TrimEnd(' | ')
$adapterRAM = $adapterRAM.TrimEnd(' | ')
#if ($respuesta -eq "Y" -or $respuesta -eq "y") {
if ($respuesta -eq "6") {
#Start-Process -FilePath "taskmgr.exe" -ArgumentList "/Performance"
    $gpuDescription = ""
    $adapterRAM = ""
    if ($videoControllers -is [array]) {
        $counter = 1
        #Mas de un chip de video
        foreach ($controller in $videoControllers) {
            $adapterRAMBytes = $controller.AdapterRAM
            $adapterRAMGB = [Math]::Round($adapterRAMBytes / 1GB, 2)
            
                #$vManual = Read-Host "$desea ingresar Manualmente el valor de $($controller.Description) (y/n)"
                $msgBody = "you want to manually enter the value of $($controller.Description)?"
                $msgTitle = "GPU Value"
                $msgButton = 'YesNo'
                $msgImage = 'Question'    
                $result = [System.Windows.Forms.MessageBox]::Show($msgBody,$msgTitle,$msgButton,$msgImage)
                $vManual = ($result).value__

                #if($vManual.ToLower() -eq 'y') {
                if($vManual -eq '6') {
                    $txtFileTest+="GPU Verification Test PASS"+$controller.Description+" `n"
                    Write-Host "    [OK]" $controller.Description -ForegroundColor Green
                    $adapterRAMGB = Read-Host "¿cual es el valor de la Memoria dedicada de la GPU en GB? Solo numero"
                    $adapterRAM += " $($adapterRAMGB) GB | "
                }
            if($adapterRAMGB -ge 4 -and $vManual -ne 'y') {
                $txtFileTest+="GPU Verification Test PASS"+$controller.Description+" `n"
                Write-Host "    [OK]" $controller.Description -ForegroundColor Green
                $adapterRAMGB = Read-Host "¿cual es el valor de la Memoria dedicada de la GPU en GB? Solo numero"
                $adapterRAM += " $($adapterRAMGB) GB | "
            }
            elseif($vManual -ne '6') {
                $adapterRAM += "$adapterRAMGB GB | "
            }
            $gpuDescription += "$counter. $($controller.Description) $($adapterRAMGB) GB | "
            $counter++
        }
    } elseif ($videoControllers) {
        $adapterRAMBytes = $videoControllers.AdapterRAM
        $adapterRAMGB = [Math]::Round($adapterRAMBytes / 1GB, 2)
        if($adapterRAMGB -ge 4) {
            $txtFileTest+="GPU Verification Test PASS"+$videoControllers.Description+" `n"
            Write-Host "    [OK]" $videoControllers.Description -ForegroundColor Green
            $s1= Read-Host -Prompt "Enter your subject 1 name" -AsSecureString
            #$adapterRAMGB = Read-Host "¿cual es el valor de la Memoria dedicada de la GPU en GB? Solo numero"
            $adapterRAMGB = Read-Host -Prompt "    What is the value of the GPU Dedicated Memory in GB? Only number"
            $adapterRAM = "$($adapterRAMGB) GB"
        }
        else {
            $adapterRAM = "$($adapterRAMGB) GB"
        }
        
    $gpuDescription = "$($videoControllers.Description) $($adapterRAMGB) GB"
    }
    $gpuDescription = $gpuDescription.TrimEnd(' | ')
    $adapterRAM = $adapterRAM.TrimEnd(' | ')
}else{
    $txtFileTest+="GPU Verification Test PASS, "+$gpuDescription+"`n"
    Write-Host "    [OK] Non-dedicated GPU" $gpuDescription -ForegroundColor Green
}

#Hard Drive
$internalDisks = Get-PhysicalDisk | Where-Object {$_.MediaType -match "HDD|SSD" -and $_.Usage -eq "Auto-Select"}
$hddInformation = ""
foreach ($disk in $internalDisks) {
    $sizeBytes = $disk.Size
    $convertedSize = ConvertBytesToStandardSize -Bytes $sizeBytes
    $friendlyName = $disk.FriendlyName
    $mediaType = $disk.MediaType    
    $hddInformation +=  "$convertedSize $friendlyName $mediaType | "
    $SIZE_HDD += "$convertedSize | "
    $MODEL_HDD += "$friendlyName | "
    $SERIAL_HDD += "$($disk.SerialNumber) | "
}

$SIZE_HDD = $SIZE_HDD.TrimEnd(' | ')
$MODEL_HDD = $MODEL_HDD.TrimEnd(' | ')
$SERIAL_HDD = $SERIAL_HDD.TrimEnd(' | ')
$hddInformation = $hddInformation.TrimEnd(' | ')
# Memory RAM
$tieneRGB = "n"

# Convertir la respuesta en minúsculas para hacer coincidencia
$tieneRGB = $tieneRGB.ToLower()

# Verificar si la respuesta es "sí" y asignar el valor a $ram_rgb
if ($tieneRGB -eq "y") {
    $ram_rgb = "RGB"
} else {
    $ram_rgb = ""
}
$ramModules = Get-WmiObject -Class Win32_PhysicalMemory
$ram_desc = ""
$ramName = 0
foreach ($ram in $ramModules) {
    $manufacturer = $ram.Manufacturer
    $capacityGB = [math]::Round($ram.Capacity / 1GB)
    $ramName = $ramName + $capacityGB
    $speed = $ram.Speed
    $memoryType = $ram.MemoryType
    $ddrVersion = ""
    if ($speed -gt 1000) {
        $ddrVersion = "DDR4"
    } elseif ($speed -gt 667) {
        $ddrVersion = "DDR3"
    } else {
        $ddrVersion = "DDR2 o anterior"
    }
    
    if($tieneRGB -eq "y") {
        $ram_desc += "$manufacturer, $capacityGB GB, $speed MHz, $ddrVersion $ram_rgb | "
    } else {
    $ram_desc += "$manufacturer, $capacityGB GB, $speed MHz, $ddrVersion | "
    }
}
$ramName = "$($ramName) GB"
$ram_desc = $ram_desc.TrimEnd(' | ')
$description = "$descripcionEquipo`r`n$os`r`n$cpu_desc;`r`n$hddInformation`r`n$ramName GB ($ram_desc)`r`n$gpuDescription"
$operator = Read-Host "`n    Operator:"
$operator = $operator.ToUpper()
Write-Host "$($serial) - $($sku)"
Write-Host $description

while (((Test-NetConnection $addrsWeb -Port 80 -InformationLevel "Detailed").TcpTestSucceeded) -ne $true){

    #write-Host "Please Connect the Equipment to the NetWork";
    #Start-Sleep -Seconds 3

    $msgBody = "Please Connect the Equipment to the NetWork"
    $msgTitle = "Network Connection Issue"
    $msgButton = 'Ok'
    $msgImage = 'Warning'
    $result = [System.Windows.Forms.MessageBox]::Show($msgBody,$msgTitle,$msgButton,$msgImage)

}


$connectionString = "Server=$server;Database=$database;User ID=$username;Password=$password;"
$insertQuery = @"
BEGIN TRANSACTION;
IF EXISTS (SELECT 1 FROM $table WHERE Serial = '$serial')
BEGIN
    DELETE FROM $table WHERE Serial = '$serial';
    INSERT INTO $table (test_SnResultsID, RAM, GPU, GPU_RAM, Model_HDD, Model, Serial, date, status, Serial_HDD, HDD_CAPACITY, OPERAtOR, os, CPU, DateEnd, DateStart, Description)
    VALUES (NEWID(), '$ramName', '$gpuDescription', '$adapterRAM', '$MODEL_HDD', '$sku', '$serial', GETDATE(), 'true', '$SERIAL_HDD', '$SIZE_HDD', '$operator','$os', '$cpu_desc', '$startDate', '$endDate', '$description');
END
ELSE
BEGIN
    INSERT INTO $table (test_SnResultsID, RAM, GPU, GPU_RAM, Model_HDD, Model, Serial, date, status, Serial_HDD, HDD_CAPACITY, OPERAtOR, os, CPU, DateEnd, DateStart, Description)
    VALUES (NEWID(), '$ramName', '$gpuDescription', '$adapterRAM', '$MODEL_HDD', '$sku', '$serial', GETDATE(), 'true', '$SERIAL_HDD', '$SIZE_HDD', '$operator','$os', '$cpu_desc', '$startDate', '$endDate', '$description');
END
COMMIT;
"@
try {
    $connection = New-Object System.Data.SqlClient.SqlConnection($connectionString)
    $connection.Open()
    Write-Host "Conexión establecida correctamente."
    $command = New-Object System.Data.SqlClient.SqlCommand($insertQuery, $connection)
    $command.ExecuteNonQuery()
    Write-Host "Datos insertados o actualizados correctamente."

    $txtFileTest+="CPU Verification Test PASS, "+$cpu_desc+"`n"
    $txtFileTest+="HDD Capacity Test PASS, "+$SIZE_HDD+"`n"
    $txtFileTest+="Memory RAM Capacity Test PASS, "+$ramName+"`n"

    while (((Test-NetConnection 10.2.198.145 -Port 80 -InformationLevel "Detailed").TcpTestSucceeded) -ne $true){

        $msgBody = "Please Connect the Equipment to the NetWork and Verify the Connection to Imaging Server"
        $msgTitle = "Network Connection and Imaging Server Issue"
        $msgButton = 'Ok'
        $msgImage = 'Warning'
        $result = [System.Windows.Forms.MessageBox]::Show($msgBody,$msgTitle,$msgButton,$msgImage)

    }

    net use W: \\10.2.198.145\logsPass /u:localhost\isptek YouCantRemote1!
    New-Item $pathScript"logsPass\"$namePassFileTest -Force
    $txtFileTest+="Operator: "+$operator+"`n"
    $txtFileTest+="complete time: "
    $txtFileTest+=Get-Date -Format "HH:mm:ss"
    $txtFileTest+="`n=============================================================`nTest Result is PASS"
    Set-Content $pathScript"logsPass\"$namePassFileTest $txtFileTest
    Copy-Item $pathScript"logsPass\"$namePassFileTest \\10.2.198.145\logsPass\$namePassFileTest -Force

    if (Test-Path $pathScript"Face.jpg") {

        Move-Item $pathScript"Face.jpg" $pathScript"logsPics\"$serial"_"$sku"_"$operator".jpg" -Force

        Copy-Item $pathScript"logsPics\"$serial"_"$sku"_"$operator".jpg" \\10.2.198.145\logsPics\$serial"_"$sku"_"$operator".jpg" -Force
        #Move-Item $pathScript"Face.jpg" \\10.2.198.145\logsPics\$serial"_"$sku"_"$operator".jpg" -Force

    }

    #Start-Sleep -Seconds 5
    #Exit
    Write-Host "1. Apagar la computadora"
    Write-Host "2. Repetir el proceso"
    Write-Host "3. Salir"
    $opcion = Read-Host "Por favor, selecciona una opción"
    switch ($opcion) {
        1 {
            Write-Host "Apagando la computadora en 1 segundo..."
            Start-Sleep -Seconds 1
            Stop-Computer -Force
        }
        2 {
            Write-Host "Reiniciando el proceso..."
        }
        3 {
            Exit
        }
        default {
            Write-Host "Opción inválida. Por favor, selecciona 1, 2 o 3."
        }
    }

} catch {
    Write-Host "Error al establecer la conexión a la base de datos: $_"
    Write-Host "Presiona cualquier tecla para salir..."
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")

} finally {
    if ($connection.State -eq 'Open') {
        $connection.Close()
        $connection.Dispose()
    }
}

} while ($opcion -eq '2')

bcdedit /set nointegritychecks off

# SIG # Begin signature block
# MIIGAAYJKoZIhvcNAQcCoIIF8TCCBe0CAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUuZWzOwrSQGXtFnp7z3HTxAS0
# qfKgggNqMIIDZjCCAk6gAwIBAgIQHKKdSd0fuahHmfjqNfspmzANBgkqhkiG9w0B
# AQsFADBLMRYwFAYDVQQDDA1JU1AgVGVzdCBUb29sMTEwLwYJKoZIhvcNAQkBFiJs
# b3JlbnpvLnZpbnVlemFAaXNwdGVrc2VydmljZXMuY29tMB4XDTIzMTEwNjIxNDk0
# OFoXDTMxMDEwMTA2MDAwMFowSzEWMBQGA1UEAwwNSVNQIFRlc3QgVG9vbDExMC8G
# CSqGSIb3DQEJARYibG9yZW56by52aW51ZXphQGlzcHRla3NlcnZpY2VzLmNvbTCC
# ASIwDQYJKoZIhvcNAQEBBQADggEPADCCAQoCggEBAM6OeNTYrSS1S4RIBQf85ODZ
# TAxiD6BkT7+x436oxRv6TsFpEFcy1Lhrt8oswDixysWZ352D6Bkao4Hqg7hue1ny
# PbGRDoUPFiwmp/nCDr5JVb5UFKFgb4wC0qMZHjAS6AbUWj7GK/Rc3z71ABkENyFf
# oAdffyDA4/nkPDJcON31gAnCiKXfAqhwoj40gSLBmKuXKsCNVGYpTX8ZEzmRqsm5
# OMj8b0CxkfijuD0Acvi+pCZqoVuL/fIu9j5M4z5xwz7qq7FYpvPn9M3NJacyoZ7X
# WCLs/ev0oh9Nl7jq/xRPx6jflIFpnFbLPy1EDTkChrHljTHf6XIjh+zLKFXAIy0C
# AwEAAaNGMEQwDgYDVR0PAQH/BAQDAgWgMBMGA1UdJQQMMAoGCCsGAQUFBwMDMB0G
# A1UdDgQWBBTfyeVgwI0FJt2SI6vioo4gZmAkizANBgkqhkiG9w0BAQsFAAOCAQEA
# HJZtteUkQDBPATYo8UeGTUjhBuQPa7ZhQnlCoSDDzoGEl/FLxxh8ob1hEI+S++qX
# 3fBSrKECaIqE043T4EWFrtsD3kbARgUH5L4Xb3iBtg8PZSMkahRuZ3pO94Df/eV8
# FZhHnn0KqVGEZzgEEwdmUbOl9ZnF3XJaPz3TNdiXpt+OZBS5CCTtsKMwo5fGmzhS
# teyghqw88cw7uHUaRhO62pwcRUv1z/5sAC38z1WhfsgeeX3NQvIcq3R3lAEqmcY5
# Z9QMoGW93xJa3s8yrImOXDHDt01zq2GV8ku6dQe6Oyt7a8Is7OLo1M056QNbvjWc
# twGhwG4FxfP30LiW5ksXBzGCAgAwggH8AgEBMF8wSzEWMBQGA1UEAwwNSVNQIFRl
# c3QgVG9vbDExMC8GCSqGSIb3DQEJARYibG9yZW56by52aW51ZXphQGlzcHRla3Nl
# cnZpY2VzLmNvbQIQHKKdSd0fuahHmfjqNfspmzAJBgUrDgMCGgUAoHgwGAYKKwYB
# BAGCNwIBDDEKMAigAoAAoQKAADAZBgkqhkiG9w0BCQMxDAYKKwYBBAGCNwIBBDAc
# BgorBgEEAYI3AgELMQ4wDAYKKwYBBAGCNwIBFTAjBgkqhkiG9w0BCQQxFgQUflIq
# BTF2nvdfG/JiyJH7UdSIcawwDQYJKoZIhvcNAQEBBQAEggEALB/KvbfML98e3hD7
# fV7h26vwF0C92fiKPSrVXvP3uhsaw0g265gZr7N2Awv/95qf6mtMt7fPtw7uFSw6
# SwQvMh9XW6PYyzEhgc2mQvRUmnxvb3yEGDUkaJTXy5bL0Q59jHsc8H+7L6N4idOI
# m0JPQWFwOvV7TgtRqcvro+gAH+kQPsjUN+LNGebhio866dtOBbdmttV/lVH1/Uci
# qCkI7iIUZR+3Zf8By74XlVZ1ig8g4EXl6wxsQXpPiCsRHqN0qvl+HrJP9yJhR36F
# WYkR2EoM54MHQcvnB75MVW0LlhazZHm/7WW4/8czi4guMvLHbG+Vs2LrLdQI9gBJ
# vJjRXQ==
# SIG # End signature block
