<#
    Automated Software Instller for HCPF
    Author: Kyle King
    Date: 9/7/2021

    Creates a Windows Form with a list of checkable items. Any checked item will be queued for installation.

    7/5/2022 - ALL ADDRESSES AND KEY HAVE BEEN REMOVED FOR PRIVACY CONCERNS
#>

# import required packages
Add-Type -AssemblyName System.Windows.Forms

# Create GUI for installation checkbox
$Form                           = New-Object System.Windows.Forms.Form
$Form.ClientSize                = New-Object System.Drawing.Point(400,300)
$Form.Text                      = "Automated Software Install"

$ListBox                        = New-Object System.Windows.Forms.CheckedListBox
$ListBox.Width                  = 380
$ListBox.Height                 = 250
$ListBox.Location               = New-Object System.Drawing.Point(10,10)
$ListBox.Font                   = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$InstallButton                  = New-Object System.Windows.Forms.Button
$InstallButton.Text             = "Start"
$InstallButton.Location         = New-Object System.Drawing.Point(300,260)
$InstallButton.Font             = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$InstallingLabel                = New-Object system.Windows.Forms.Label
$InstallingLabel.text           = "Installing in Progess..."
$InstallingLabel.AutoSize       = $True
$InstallingLabel.Visible        = $False
$InstallingLabel.width          = 25
$InstallingLabel.height         = 12
$InstallingLabel.location       = New-Object System.Drawing.Point(25,260)
$InstallingLabel.Font           = 'Microsoft Sans Serif,12'

$CompleteLabel                  = New-Object system.Windows.Forms.Label
$CompleteLabel.text             = "Installs Complete"
$CompleteLabel.AutoSize         = $True
$CompleteLabel.Visible          = $False
$CompleteLabel.width            = 25
$CompleteLabel.height           = 12
$CompleteLabel.location         = New-Object System.Drawing.Point(25,260)
$CompleteLabel.Font             = 'Microsoft Sans Serif,12'
  
# Add elements to Form
$Form.Controls.AddRange(@($InstallButton,$ListBox,$InstallingLabel,$CompleteLabel))

# Add list of software options for install
[void]$ListBox.Items.Add("7Zip")
[void]$ListBox.Items.Add("Add Printer")
[void]$ListBox.Items.Add("Adobe Acrobat Pro 2017")
[void]$ListBox.Items.Add("Adobe Acrobat Pro 2020")
[void]$ListBox.Items.Add("ALM & Explorer")
[void]$ListBox.Items.Add("ArcGIS Desktop (Requires Activation)")
[void]$ListBox.Items.Add("ArcGIS Business Analyst (Requires Activation)")
[void]$ListBox.Items.Add("ArcGIS Street View (Pro) (Requires Activation)")
[void]$ListBox.Items.Add("Articulate 360")
[void]$ListBox.Items.Add("Bomgar Jump Client")
[void]$ListBox.Items.Add("Call Center")
[void]$ListBox.Items.Add("Camtasia (Requires Activation)")
[void]$ListBox.Items.Add("Cisco Jabber")
[void]$ListBox.Items.Add("Cisco Webex")
[void]$ListBox.Items.Add("Citrix Workspace App")
[void]$ListBox.Items.Add("FileZilla")
[void]$ListBox.Items.Add("Google Drive for Desktop")
[void]$ListBox.Items.Add("HPIA (Manual)")
[void]$ListBox.Items.Add("Microsoft Office 365 x64")
[void]$ListBox.Items.Add("Microsoft Project x64")
[void]$ListBox.Items.Add("Microsoft Visio x64")
[void]$ListBox.Items.Add("Microsoft PowerBI x64")
[void]$ListBox.Items.Add("NVivo")
[void]$ListBox.Items.Add("Onbase Unity Client and VPD")
[void]$ListBox.Items.Add("Oracle 12c x64 and file transfer")
[void]$ListBox.Items.Add("Oracle 19c x64 and file transfer")
[void]$ListBox.Items.Add("R and R Stuido (Includes Oracle)")
[void]$ListBox.Items.Add("RSA Token")
[void]$ListBox.Items.Add("SAS 9.4")
[void]$ListBox.Items.Add("Snagit (Requires Activation)")
[void]$ListBox.Items.Add("Tableau Reader")
[void]$ListBox.Items.Add("Tableau Desktop")
[void]$ListBox.Items.Add("Toad 5.3 (Includes Oracle and activation)")
[void]$ListBox.Items.Add("Ultra Edit")
[void]$ListBox.Items.Add("VPN PDF Guide")

# Set home directory to location where ps1 script is saved
$homeDir = "" # address to where the file is saved.
# Location where local files will be saved for some installs
$destinationFolder = "C:\ProgramData\Temp\"

# CopyMSI is used to copy the MSI installer locally to a hidden folder on the machine for better install procedure
function CopyMSI($MsiFIle) {
    if (!(Test-Path -path $destinationFolder)) {
		New-Item $destinationFolder -Type Directory
	} 
    Copy-Item -Path $MsiFIle -Destination $destinationFolder
}

# Applies Adobe fix allowing users to login to Adobe products
function AdobeFix {
    if ((Test-Path -path "C:\Program Files (x86)\Common Files\Adobe\OOBE\PDApp\P7")) {
		Copy-Item -Path "$homeDir\" <# address to where the file is saved #> -Destination "C:\Program Files (x86)\Common Files\Adobe\OOBE\PDApp\P7"
	} 
    Start-Process -FilePath "$homeDir\" <# address to where the file is saved #> -Wait -NoNewWindow
}

function AppDataInstaller($Files) {
    $user = Get-WmiObject win32_computersystem | select username
    $user = $user.username.split('\')[1]
    Write-Host "Installing into $user's Appdata"
    Copy-Item -Path $Files -Destination "C:\Users\$user\AppData\Local\" -Force -Recurse
    Write-Host "Installing into Default Appdata"
    Copy-Item -Path $Files -Destination "C:\Users\Default\AppData\Local\" -Force -Recurse
}

# Install Map data for ArcGIS
function ArcGISMapData {
    $arcFolder = "C:\ArcGIS\Business Analyst\US_2015" # "C:\ProgramData\ArcMaps"
    if (!(Test-Path -path $arcFolder)) {
        New-Item $arcFolder -Type Directory
        Write-Host "ArcGIS Street Map Data"
        Invoke-Item -path "$homeDir\" <# address to where the file is saved #>
        Start-Process -FilePath "$homeDir\" <# address to where the file is saved #> -ArgumentList "/S /v/qr" -Wait -NoNewWindow
    }
}

# Removes leftover installs from harddrive
function CleanUp {
    if (Test-Path -path $destinationFolder) {
        Get-ChildItem -Path $destinationFolder -Include *.* -File -Recurse | Remove-Item
        Remove-Item -Path $destinationFolder -Recurse -Force
    }
}


# Install fuctions for each software added in the list above
function Install7Zip { 
    Start-Process -FilePath "$homeDir\" <# address to where the file is saved #> -ArgumentList "/S" -Wait -NoNewWindow 
} # working

function InstallPrinter {
    Start-Process -FilePath "$homeDir\" <# address to where the file is saved #> -Wait -NoNewWindow
} # working

function InstallAdobeAcrobat2017 {
    Start-Process -FilePath "$homeDir\" <# address to where the file is saved #> -ArgumentList "/sAll /rs" -Wait -NoNewWindow 
    AdobeFix
} # working

function InstallAdobeAcrobat2020 {
    Start-Process -FilePath "$homeDir\" <# address to where the file is saved #> -ArgumentList "/sAll /rs" -Wait -NoNewWindow
    AdobeFix
} # working

function InstallAdobeCC {
    CopyMSI "$homeDir\" <# address to where the file is saved #>
    Start-Process msiexec.exe -ArgumentList "/I C:\ProgramData\Temp\filename.exe" -Wait -NoNewWindow 
    AdobeFix
} # not working

function InstallALM {
    CopyMSI "$homeDir\" <# address to where the file is saved #>
    Start-Process msiexec.exe -ArgumentList "/I C:\ProgramData\Temp\filename.msi /passive" -Wait -NoNewWindow 
    Start-Process -FilePath "$homeDir\" <# address to where the file is saved #> -ArgumentList "/Q" -Wait -NoNewWindow
    Start-Process Powershell.exe -Argument "$homeDir\" <# address to where the file is saved #> -Wait -NoNewWindow
    
    $openIE = {"$IE = new-object -com internetexplorer.application; $IE.navigate2('http://usolgwkxco002.coad.xco.dcs-usps.com:8080/qcbin/start_a.jsp'); $IE.visible = $true;"}
    $action = New-ScheduledTaskAction -Execute Powershell.exe -Argument $openIE
    $trigger = New-ScheduledTaskTrigger -AtLogOn
    $principal = New-ScheduledTaskPrincipal -UserId (Get-CimInstance -ClassName Win32_ComputerSystem | Select-Object -expand UserName)
    $task = New-ScheduledTask -Action $action -Trigger $trigger -Principal $principal
    Register-ScheduledTask StartALM -InputObject $task
    Start-ScheduledTask -TaskName StartALM
    Start-Sleep -Seconds 5
    Unregister-ScheduledTask -TaskName StartALM -Confirm:$false
} # who knows

function InstallArcGISDesktop {
    Start-Process -FilePath "$homeDir\" <# address to where the file is saved #> -ArgumentList "/passive /norestart" -Wait -NoNewWindow 
    ArcGISMapData
} # working

function InstallArcGISBA {
    CopyMSI "$homeDir\" <# address to where the file is saved #>
    Start-Process msiexec.exe -ArgumentList "/I C:\ProgramData\Temp\filename.msi /passive" -Wait -NoNewWindow
    ArcGISMapData
} # working

function InstallArcGISStreet {
    #labeld as ArcGIS pro
    CopyMSI "$homeDir\" <# address to where the file is saved #>
    Start-Process msiexec.exe -ArgumentList "/I C:\ProgramData\Temp\filename.exe /passive" -Wait -NoNewWindow 
    ArcGISMapData
} # working

function InstallArticulate {
    Start-Process -FilePath "$homeDir\" <# address to where the file is saved #> -ArgumentList "/passive" -Wait -NoNewWindow 
} # working

function InstallBomgar {
    CopyMSI "$homeDir\" <# address to where the file is saved #>
    Start-Process msiexec.exe -ArgumentList "/I C:\ProgramData\Temp\filename.msi" -Wait -NoNewWindow 
} # working

function InstallCallCenter {
    CopyMSI "$homeDir\" <# address to where the file is saved #>
    CopyMSI "$homeDir\" <# address to where the file is saved #>

    Start-Process msiexec.exe -ArgumentList "/I C:\ProgramData\Temp\filename.msi MSIINSTALLPERUSER=1" -Wait -NoNewWindow
    Start-Process msiexec.exe -ArgumentList "/I C:\ProgramData\Temp\filename.msi /q" -Wait -NoNewWindow
}  # working

function InstallCamtasia {
    Start-Process -FilePath "$homeDir\" <# address to where the file is saved #> -ArgumentList "/passive /norestart" -Wait -NoNewWindow 
} # working

function InstallJabber {
    CopyMSI "$homeDir\" <# address to where the file is saved #>
    Start-Process msiexec.exe -ArgumentList "/I C:\ProgramData\Temp\filename.msi /q" -Wait -NoNewWindow
}  # working

function InstallCiscoWebex {
    CopyMSI "$homeDir\" <# address to where the file is saved #>
    Start-Process msiexec.exe -ArgumentList "/I C:\ProgramData\Temp\filename.msi /passive" -Wait -NoNewWindow 
}  # working

function InstallCitrixWorkspace {
    Start-Process -FilePath "$homeDir\" <# address to where the file is saved #> -ArgumentList "/silent /noreboot /forceinstall" -Wait -NoNewWindow
}  # working

function InstallFileZilla {
    Start-Process -FilePath "$homeDir\" <# address to where the file is saved #> -ArgumentList "/S /user=all" -Wait -NoNewWindow 
} # working

function InstallGoogleDrive {
    Start-Process -FilePath "$homeDir\" <# address to where the file is saved #> -ArgumentList "--silent --desktop_shortcut" -Wait -NoNewWindow 
} # working

function InstallHPIA {
    Start-Process -FilePath "$homeDir\" <# address to where the file is saved #> -Wait -NoNewWindow
}  # working


function InstallProject {
    Start-Process -FilePath "$homeDir\" <# address to where the file is saved #> -Wait -NoNewWindow 
} # working

function InstallOffice365 {
    Start-Process -FilePath "$homeDir\" <# address to where the file is saved #> -Wait -NoNewWindow 
} # working

function InstallVisio {
    Start-Process -FilePath "$homeDir\" <# address to where the file is saved #> -Wait -NoNewWindow 
} # working

function InstallPowerBI {
    Start-Process -FilePath "$homeDir\" <# address to where the file is saved #> -ArgumentList "-q -norestart ACCEPT_EULA=1" -Wait -NoNewWindow 
} # working

function InstallNVivo {
    Start-Process -FilePath "$homeDir\" <# address to where the file is saved #> -ArgumentList "/S /v/qn" -Wait -NoNewWindow 
} # working

function InstallOnbase {
    CopyMSI "$homeDir\" <# address to where the file is saved #>
    CopyMSI "$homeDir\" <# address to where the file is saved #>
    Start-Process msiexec.exe -ArgumentList "/I C:\ProgramData\Temp\filename.msi /passive" -Wait -NoNewWindow 
    Start-Process msiexec.exe -ArgumentList "/I C:\ProgramData\Temp\filename.msi /passive" -Wait -NoNewWindow 
    Copy-Item -Path "$homeDir\" <# address to where the file is saved #> -Destination "C:\Program Files (x86)\Hyland\Unity Client"
} # working

function InstallOracle12_64 {
    if (!(Test-Path "C:\Oracle64")) {
        Write-Host "Oracle 12c x64 and file transfer"
        Start-Process Powershell.exe -Argument "$homeDir\" <# address to where the file is saved #> -Wait -NoNewWindow
        Copy-Item -Path "$homeDir\" <# address to where the file is saved #> -Destination "C:\Oracle64\product\12.1.0\client_1\network\admin"
    }
} # working

function InstallOracle19_64 {
    if (!(Test-Path "C:\Oracle19_x64")) {
        Write-Host "Oracle 19c x64 and file transfer"
        Start-Process Powershell.exe -Argument "$homeDir\" <# address to where the file is saved #> -Wait -NoNewWindow
        Copy-Item -Path "$homeDir\" <# address to where the file is saved #> -Destination "C:\Oracle19_x64\product\19.0.0\client_1\network\admin"
    }
}  # working

function InstallR {
    InstallOracle12_64
    Start-Process -FilePath "$homeDir\" <# address to where the file is saved #> -ArgumentList "/SILENT /NORESTART" -Wait -NoNewWindow 
    Start-Process -FilePath "$homeDir\" <# address to where the file is saved #> -ArgumentList "/S" -Wait -NoNewWindow 
} # working

function InstallRSA {
    CopyMSI "$homeDir\" <# address to where the file is saved #>
    Start-Process msiexec.exe -ArgumentList "/I C:\ProgramData\Temp\filename.msi /passive" -Wait -NoNewWindow 
} # working

function InstallSAS {
    Copy-Item -Path "$homeDir\" <# address to where the file is saved #> -Destination "C:\ProgramData"
    Start-Process -FilePath "$homeDir\" <# address to where the file is saved #> -ArgumentList "-quiet -wait -responsefile C:\ProgramData\response.properties" -Wait -NoNewWindow 
} # working

function InstallSnagit {
    Start-Process -FilePath "$homeDir\" <# address to where the file is saved #> -ArgumentList "/quiet" -Wait -NoNewWindow 
} # working

function InstallTableauReader {
    Start-Process -FilePath "$homeDir\" <# address to where the file is saved #> -ArgumentList "/silent /norestart ACCEPTEULA=1 REMOVEINSTALLEDAPP=1" -Wait -NoNewWindow 
} # working

function InstallTableauDesktop {
    Start-Process -FilePath "$homeDir\" <# address to where the file is saved #> -ArgumentList "/silent /norestart ACCEPTEULA=1 REMOVEINSTALLEDAPP=1" -Wait -NoNewWindow 
} # working

function InstallToad {
    InstallOracle12_64
    CopyMSI "$homeDir\" <# address to where the file is saved #>
    Start-Process -FilePath "$homeDir\" <# address to where the file is saved #> -ArgumentList "/passive" -Wait -NoNewWindow 
    Start-Process -FilePath "$homeDir\" <# address to where the file is saved #> -ArgumentList "/passive" -Wait -NoNewWindow 
    Start-Process msiexec.exe -ArgumentList "/I C:\ProgramData\Temp\filename.msi /passive" -Wait -NoNewWindow 
    Copy-Item -Path "$homeDir\" <# address to where the file is saved #> -Destination "C:\Program Files\Quest Software\Toad Data Point 5.3"
} # working

function InstallUltraEdit {
    CopyMSI "$homeDir\" <# address to where the file is saved #>
    Start-Process msiexec.exe -ArgumentList "/I C:\ProgramData\Temp\filename.msi /passive" -Wait -NoNewWindow 
} # working

function InstallVPNGuide {
    $Source = "$homeDir\" <# address to where the file is saved #>
    $Destination = 'C:\users\*\Desktop\'
    Get-ChildItem $Destination | ForEach-Object {Copy-Item -Path $Source -Destination $_ -Force}
    Copy-Item -Path $Source -Destination "C:\Users\Default\Desktop"
}  # working


# Once the Start button in clicked, loop through each checked item and call the instller. Set Installing label to visible when starting and then ocmplete label when done.
$InstallButton.Add_Click(
    {
        $Count = 1
        $InstallingLabel.Visible = $True
        Write-Host "`n`n`n`n`n`n"
        for($i = 0; $i -lt $ListBox.CheckedItems.Count; $i++) {
            Write-Progress -Id 0 -Activity "Installing Software..." -Status "$Count of $($ListBox.CheckedItems.Count)" -PercentComplete (($Count / $ListBox.CheckedItems.Count) * 100)
            $Count++
            Write-Host $ListBox.CheckedItems[$i]
            switch($ListBox.CheckedItems[$i]){
                "7Zip"                                              {Install7Zip; Break}
                "Add Printer"                                       {InstallPrinter; Break}
                "Adobe Acrobat Pro 2017"                            {InstallAdobeAcrobat2017; Break}
                "Adobe Acrobat Pro 2020"                            {InstallAdobeAcrobat2020; Break}
                "ALM & Explorer"                                    {InstallALM; Break}
                "ArcGIS Desktop (Requires Activation)"              {InstallArcGISDesktop; Break}
                "ArcGIS Business Analyst (Requires Activation)"     {InstallArcGISBA; Break}
                "ArcGIS Street View (Pro) (Requires Activation)"    {InstallArcGISStreet; Break}
                "Articulate 360"                                    {InstallArticulate; Break}
                "Bomgar Jump Client"                                {InstallBomgar; Break}
                "Call Center"                                       {InstallCallCenter; Break}
                "Camtasia (Requires Activation)"                    {InstallCamtasia; Break}
                "Cisco Jabber"                                      {InstallJabber; Break}
                "Cisco Webex"                                       {InstallCiscoWebex; Break}
                "Citrix Workspace App"                              {InstallCitrixWorkspace; Break}
                "FileZilla"                                         {InstallFileZilla; Break}
                "Google Drive for Desktop"                          {InstallGoogleDrive; Break}
                "HPIA (Manual)"                                     {InstallHPIA; Break}
                "Microsoft Office 365 x64"                          {InstallOffice365; Break}
                "Microsoft Project x64"                             {InstallProject; Break}
                "Microsoft Visio x64"                               {InstallVisio; Break}
                "Microsoft PowerBI x64"                             {InstallPowerBI; Break}
                "NVivo"                                             {InstallNVivo; Break}
                "Onbase Unity Client and VPD"                       {InstallOnbase; Break}
                "Oracle 12c x64 and file transfer"                  {InstallOracle12_64; Break}
                "Oracle 19c x64 and file transfer"                  {InstallOracle19_64; Break}
                "R and R Stuido (Includes Oracle)"                  {InstallR; Break}
                "RSA Token"                                         {InstallRSA; Break}
                "SAS 9.4"                                           {InstallSAS; Break}
                "Snagit (Requires Activation)"                      {InstallSnagit; Break}
                "Tableau Reader"                                    {InstallTableauReader; Break}
                "Tableau Desktop"                                   {InstallTableauDesktop; Break}
                "Toad 5.3 (Includes Oracle and activation)"         {InstallToad; Break}
                "Ultra Edit"                                        {InstallUltraEdit; Break}
                "VPN PDF Guide"                                     {InstallVPNGuide; Break}
            }
        }
        CleanUp
        $InstallingLabel.Visible = $False
        $CompleteLabel.Visible = $True
        Write-Host ""
        Write-Host "=================="
        Write-Host "Installs Complete."
        Write-Host "=================="
    }
)

# Bulid the form
$Form.Add_Shown({$Form.Activate()})
$Form.ShowDialog()