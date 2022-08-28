Set-Location "C:\temp"
$CurrentDate = (get-date).tostring("yyyyMMdd")
$transcriptfile = "snipe-it_client-" + $CurrentDate + ".txt"
Start-Transcript -Path $transcriptfile

#####################################################
# Scripted by: Christian Gätcke 					
# Created on: 02/07/2020  							
#####################################################


###################################
# automatically get asset tag?     
# yes: $getTag=0                   
# no : $getTag=1                   
###################################
$getTag = 0

########################################
# enable the GUI?                       
# yes: $enableGUI=1                     
# powershell only : $enableGUI= 0       
# only inventorize asset: $enableGUI= 2 
# This will not create a user           
########################################
$enableGUI = 0

#########################################################
# Global variables					 
# $apiKey = '' your SnipeIt API Key			 
# example:						 
# $apiKey = '.......iIsImp0aSI6IjJmMDkyNDg5MzZk....'     
# 							 
# $baseUrl = '' the Base-Url of your SnipeIT installation
# example:						 
# $baseUrl = 'https://snipeit.example.com'		 
#########################################################

$apiKey = "eyJ0eXAiOiJKV1QiLCJhbGciOiJSUzI1NiJ9.eyJhdWQiOiIxIiwianRpIjoiMTBlZGIzZThjNTVkZDZhNDdhYjFmMzVhMDA1OTBhMmFjNDkyZTc3MzYwMGI1NWI0OTBjMTk4MWI5YWQ1MmQwMTA2YTgxMTgxZjk1ZjMxNTciLCJpYXQiOjE2NjE0OTgwOTksIm5iZiI6MTY2MTQ5ODA5OSwiZXhwIjoyMjkyNjUzNjk4LCJzdWIiOiIyIiwic2NvcGVzIjpbXX0.e4wZMSfe3QG4Znv0PFTi4VoS9vMqCmXmkpn12b6NJiKxrqvoGa3EhDM8DLDLaoGE03mDOCaQoOkqU1kpvvUscpvuvgBM9JhL7xESN_aH0PwPMB98JntlYsuzbBkCSFiLm2UX_j8fXGugdkcMK7hfJOhiAEbOuPuZZqh_FuEXsByHuePW_ZVp6UKLq_KRgR-T-odHFMM0eRbdG-0AGeY2QKbUIkyOOLKlnQocFt6cb-JvolvMU7WIQQbksEPpLjB2O0ipXHPxM3oWHXkhN14R_0mJwIvFtoV--XqjMlQe2QjE4U6nD8fwp4XdAnVSb9r4Tvh421h3NyvMPz5NCf4RZj2A6J_IDMfpvoNCCA3BmodxCDhkopDSuX_VY2p7Cu75tZXyt89KYYftRjzz1sP3iCBVYnwj401n-3204_mTtLc9TlKnFF2uQG_4UK6dRGNdwR3XnjeYnBWa1t-WbVqE6MhEK9G6_o1smY89UPfOxJiKuJ9-hNaAUvAqzGnG51LExcYVvO-wRwQbIZTsKiM9mujA24my2CR9AzucpNJ2DWEI6nEBR5hG4HD4wm-aActFekLIrt2Wd0_LxVAAX67syiQTwXRoeM6XIkJm-YE17Lep-bT6B61_Putb6gPInQzH7zjVAataXh31pT-l4Z-XF_0lR9NDK5LBrY9ERi7i3Cc
"
$baseUrl = "https://inventory.bluvacanze.it"
$header = @{'authorization' = 'Bearer ' + $apiKey ; 'accept' = 'application/json' ; 'content-type' = 'application/json' }



#########################################################################################
# Snipe specific fields:								
# To store values like Ram, CPU, Mac etc, a fieldset with 				
# the corresponding fields needs to be created in snipe.  				
# See: https://snipe-it.readme.io/docs/custom-fields#common-custom-fields-regexes	
# Fieldset-ID:										
# example: $fsField = "2"							
# CPU-Field:										
# example: $cpuField = "_snipeit_cpu_4"				
# RAM-Field:										
# example: $ramField = "_snipeit_ram_2"				
# Macaddress-Field:									
# example: $macField = "_snipeit_mac_address_1"		
# Disk-Field:										
# example: $diskField = "_snipeit_disksize_3"		
# Operatingsystem-Field								
# example: $osField = "_snipeit_operating_system_6"	
# The examples won't work out of the box and need to be generated first			
# 											#
# $statusID is the status checked out assets will be transferred to			
# example: $statusID = "2"								
#########################################################################################

$fsField = "3"
$cpuField = "_snipeit_cpu_7"
$ramField = "_snipeit_ram_8"
$macField = "_snipeit_mac_address_1"
$diskField = "_snipeit_disk_9"
$osField = "_snipeit_operating_system_10"
$ipField = "_snipeit_ip_11"
$statusID = "2"


####################
#generate Random PW#
#of length 12	   #
####################
$ranPass = ( -join ((65..90) + (97..122) | Get-Random -Count 12 | ForEach-Object { [char]$_ })).Trim()

#####################################################################################################################################################################################
#                                                                                   GetSystemInfo                                                                                   #
#####################################################################################################################################################################################

##################
#get Computername
#will be used as 
#Asset Name	 
##################
$computername = $env:computername


###################################
#get System Model and Manufacturer
###################################
$sysModel = (Get-CimInstance -ClassName Win32_ComputerSystem | Select-Object -expand Model) -replace '\s', '-'
$sysManufacturer = (Get-CimInstance -ClassName Win32_ComputerSystem | Select-Object -expand Manufacturer) -replace '\s', '-'
$serialNumber = ( Get-WmiObject -Class Win32_BIOS).SerialNumber

########
#get OS
########

$operatingSystem = (Get-WMIObject win32_operatingsystem | Select-Object -expand name).split('|')[0]
Select-Object -Property SystemType

#####
#CPU
#####
$CPU = Get-CimInstance -ClassName Win32_Processor | Select-Object -ExcludeProperty "CIM*" -expand Name

##################
#Disksize in GB   
##################
$disksizeinbytes = Get-CimInstance -ClassName Win32_LogicalDisk -Filter "DriveType=3" |
Measure-Object -Property Size -Sum |
Select-Object -Property Property, Sum -expand Sum
$disksize = ([math]::Round($disksizeinbytes / 1073741824))
$disksize = [string]$disksize + ' GB'

###############
#Ramsize in GB
###############
$ramsizeinbytes = Get-CimInstance win32_physicalmemory | Select-Object capacity -expand capacity | Measure-Object -Property capacity -Sum | Select-Object -Property Property, Sum -expand Sum 
$ramsize = [math]::Round($ramsizeinbytes / 1073741824) 
$ramsize = [string]$ramsize + ' GB'

####################
#primary MacAddress
####################
$primaryMacaddress = Get-WmiObject win32_networkadapterconfiguration `
| Select-Object -Property @{name = 'IPAddress'; Expression = { ($_.IPAddress[0]) } }, MacAddress, Description `
| Where-Object Description -notlike "Hyper-V*"  |  Where-Object IPAddress -NE $null  | Select-Object -expand MacAddress

if ($primaryMacaddress.Length -lt 17) {
    $primaryMacaddress = $primaryMacaddress[0]
}
else 
{ $primaryMacaddress = $primaryMacaddress }

####################
#Ipv4 ips
####################
$ips = [System.Net.Dns]::GetHostAddresses($computername) | Where-Object {$_.AddressFamily -eq "InterNetwork"} `
| Select-Object -ExpandProperty IPAddressToString

if ($ips.Length -lt 16) {
    $primaryIP = $ips
} else
{ $primaryIP = $ip[0] }

#################################################################################################
# Credits go to:										
# Scripted by: Adam Bacon  									
# Created on: 15/03/2011  									
# Scripted in: Powershell will work in V1 & V2 of Powershell					
# See: https://gallery.technet.microsoft.com/scriptcenter/c0bc039d-5bbf-4c8b-8307-e44da22a42b5	
#################################################################################################
function check-chassis {  
    BEGIN {}  
    PROCESS {  
              
        $chassis = Get-WmiObject win32_systemenclosure -computer "localhost" | Select-Object chassistypes  
        if ($chassis.chassistypes -contains '3') { Write-Output "Desktop" }  
        elseif ($chassis.chassistypes -contains '4') { Write-Output "Low Profile Desktop" }  
        elseif ($chassis.chassistypes -contains '5') { Write-Output "Pizza Box" }  
        elseif ($chassis.chassistypes -contains '6') { Write-Output "Mini Tower" }  
        elseif ($chassis.chassistypes -contains '7') { Write-Output "Tower" }  
        elseif ($chassis.chassistypes -contains '8') { Write-Output "Portable" }  
        elseif ($chassis.chassistypes -contains '9') { Write-Output "Laptop" }  
        elseif ($chassis.chassistypes -contains '10') { Write-Output "Notebook" }  
        elseif ($chassis.chassistypes -contains '11') { Write-Output "Hand Held" }  
        elseif ($chassis.chassistypes -contains '12') { Write-Output "Docking Station" }  
        elseif ($chassis.chassistypes -contains '13') { Write-Output "All in One" }  
        elseif ($chassis.chassistypes -contains '14') { Write-Output "Sub Notebook" }  
        elseif ($chassis.chassistypes -contains '15') { Write-Output "Space-Saving" }   
        elseif ($chassis.chassistypes -contains '16') { Write-Output "Lunch Box" }  
        elseif ($chassis.chassistypes -contains '17') { Write-Output "Main System Chassis" }  
        elseif ($chassis.chassistypes -contains '18') { Write-Output "Expansion Chassis" }  
        elseif ($chassis.chassistypes -contains '19') { Write-Output "Sub Chassis" }  
        elseif ($chassis.chassistypes -contains '20') { Write-Output "Bus Expansion Chassis" }  
        elseif ($chassis.chassistypes -contains '21') { Write-Output "Peripheral Chassis" }  
        elseif ($chassis.chassistypes -contains '22') { Write-Output "Storage Chassis" }  
        elseif ($chassis.chassistypes -contains '23') { Write-Output "Rack Mount Chassis" }  
        elseif ($chassis.chassistypes -contains '24') { Write-Output "Sealed-Case PC" }  
        else { { Write-Output "unknown" } }
    }
}

$categoryName = $( check-chassis ).Trim()


function AddManufacturer {
    param (
        #todo add params
    )
    $manufactJson = Invoke-RestMethod -Uri "$baseUrl/api/v1/manufacturers?search=$sysManufacturer" -Headers $header -Method GET
    $manufactJson = $manufactJson | Select-Object -ExpandProperty rows
    $manufact = ( $manufactJson | Select-Object -expand "name" ) 

    if ($sysManufacturer -eq $manufact)
    { $manId = $manufactJson.id }
    else {
        $JSON = @{"name" = $sysManufacturer } | ConvertTo-Json
        Invoke-RestMethod -Uri $baseUrl/api/v1/manufacturers -Headers $header -Method POST -Body $JSON
        $manId = Invoke-RestMethod -Uri "$baseUrl/api/v1/manufacturers?search=$sysManufacturer" -Headers $header -Method GET | Select-Object -ExpandProperty rows | Select-Object -expand id
    }
}

function AddCategory {
    param (
        #todo add params
    )
    $catJson = Invoke-RestMethod -Uri "$baseUrl/api/v1/categories?search=$categoryName" -Headers $header -Method GET
    $catJson = $catJson | Select-Object -ExpandProperty rows
    $cat = ( $catJson | Select-Object -expand "name" )
    if ($categoryName -eq $cat) { 
        $categoryId = $catJson.id 
    }
    else {
        $JSON = @{"name" = $categoryName ; "category_type" = "asset" ; "checkin_email" = 'false' } | ConvertTo-Json
        Invoke-RestMethod -Uri $baseUrl/api/v1/categories -Headers $header -Method POST -Body $JSON
        $categoryId = Invoke-RestMethod -Uri "$baseUrl/api/v1/categories?search=$categoryName" -Headers $header -Method GET | Select-Object -ExpandProperty rows | Select-Object -expand id
    }

}

function AddModel {
    param (
        #todo add params
    )
    $modelJson = Invoke-RestMethod -Uri "$baseUrl/api/v1/models?limit=50&offset=0&search=$sysModel" -Headers $header -Method GET
    $modelJson = $modelJson | Select-Object -ExpandProperty rows
    $model = ( $modelJson | Select-Object -expand "name" )
    if ($sysModel -eq $model)
    { $modelId = $modelJson.id } else {
        $JSON = @{"name" = $sysModel ; "manufacturer_id" = $manId ; "category_id" = $categoryId ; "fieldset_id" = $fsField } | ConvertTo-Json
        Invoke-RestMethod -Uri "$baseUrl/api/v1/models" -Headers $header -Method POST -Body $JSON
        $modelId = Invoke-RestMethod -Uri $baseUrl/api/v1/models?search=$sysModel -Headers $header -Method GET | Select-Object -ExpandProperty rows | Select-Object -expand id
    }
}

function AddAsset {
    param (
        #todo add params
    )
    $assetJson = Invoke-RestMethod -Uri "$baseUrl/api/v1/hardware?search=$serialNumber" -Headers $header -Method GET
    $assetJson = $assetJson | Select-Object -ExpandProperty rows
    $assets = ( $assetJson | Select-Object -expand "serial" )
    if ($serialNumber -eq $assets) {
        Write-Output "Asset already exists"
    }
    else {
        $JSON = @{"name" = $computername ; "status_id" = $statusID ; "model_id" = $modelId ; $ramField = $ramsize; $cpuField = $CPU ; $diskField = $disksize ; "serial" = $serialNumber ; $macField = $primaryMacaddress ; $osField = $operatingSystem } | ConvertTo-Json
        Invoke-RestMethod -Uri "$baseUrl/api/v1/hardware" -Headers $header -Method POST -Body $JSON
        Write-Output "Asset created"
    }
}

function BuildGUI {
    param (
        #todo add params
    )
    Add-Type -AssemblyName System.Windows.Forms
    [System.Windows.Forms.Application]::EnableVisualStyles()
    Add-Type -Name Window -Namespace Console -MemberDefinition '
[DllImport("Kernel32.dll")]
public static extern IntPtr GetConsoleWindow();

[DllImport("user32.dll")]
public static extern bool ShowWindow(IntPtr hWnd, Int32 nCmdShow);
'
    $iconKey = 'iVBORw0KGgoAAAANSUhEUgAAACAAAAAgCAYAAABzenr0AAAABmJLR0
QA/wD/AP+gvaeTAAAAc0lEQVRYhe2XXQrAIAyDM/H+p5rn6p5k4gLK/Flh+UBQUJqGYhUQNwmALR
5nHfQo5jY/J0oZE7G1YSI0wbAoWDcSIAESIAESIAGsG+5qywCcOvCv9wBzIPO2FrKDVq0prh0YrY
Wu8y4d+OweSBviPT4mQlxtAyI4cVGDVAAAAABJRU5ErkJggg=='

    $Form = New-Object system.Windows.Forms.Form
    $Form.ClientSize = '379,221'
    $Form.text = "Inventorization"
    $Form.TopMost = $false

    $iconBase64 = "$iconKey"
    $iconBytes = [Convert]::FromBase64String($iconBase64)
    $stream = New-Object IO.MemoryStream($iconBytes, 0, $iconBytes.Length)
    $stream.Write($iconBytes, 0, $iconBytes.Length);
    $Form.Icon = [System.Drawing.Icon]::FromHandle((New-Object System.Drawing.Bitmap -Argument $stream).GetHIcon())

    $emailAddress = New-Object system.Windows.Forms.TextBox
    $emailAddress.multiline = $false
    $emailAddress.width = 249
    $emailAddress.height = 20
    $emailAddress.location = New-Object System.Drawing.Point(62, 34)
    $emailAddress.Font = 'Microsoft Sans Serif,10'

    $Label1 = New-Object system.Windows.Forms.Label
    $Label1.text = "Please enter your emailaddress"
    $Label1.AutoSize = $true
    $Label1.width = 25
    $Label1.height = 10
    $Label1.location = New-Object System.Drawing.Point(62, 63)
    $Label1.Font = 'Microsoft Sans Serif,10'

    if ($getTag -eq "1") {
        $compName = New-Object system.Windows.Forms.TextBox
        $compName.multiline = $false
        $compName.width = 249
        $compName.height = 20
        $compName.location = New-Object System.Drawing.Point(62, 94)
        $compName.Font = 'Microsoft Sans Serif,10'

        $Label2 = New-Object system.Windows.Forms.Label
        $Label2.text = "Please enter your computername `naccording to the sticker at the bottom"
        $Label2.AutoSize = $true
        $Label2.width = 25
        $Label2.height = 10
        $Label2.location = New-Object System.Drawing.Point(62, 123)
        $Label2.Font = 'Microsoft Sans Serif,10'
    }

    $AddDeviceBtn = New-Object system.Windows.Forms.Button
    $AddDeviceBtn.text = "Add"
    $AddDeviceBtn.width = 60
    $AddDeviceBtn.height = 30
    $AddDeviceBtn.location = New-Object System.Drawing.Point(63, 170)
    $AddDeviceBtn.Font = 'Microsoft Sans Serif,10'

    $cancelBtn = New-Object system.Windows.Forms.Button
    $cancelBtn.text = "Cancel"
    $cancelBtn.width = 60
    $cancelBtn.height = 30
    $cancelBtn.location = New-Object System.Drawing.Point(251, 170)
    $cancelBtn.Font = 'Microsoft Sans Serif,10'
    $Form.CancelButton = $cancelBtn
    $Form.Controls.Add($cancelBtn)
    if ($getTag -eq "0") {
        $Form.controls.AddRange(@($emailAddress, $Label1, $compName, $Label2, $AddDeviceBtn, $cancelBtn))
    }
    else {
        $Form.controls.AddRange(@($emailAddress, $Label1, $compName, $Label2, $AddDeviceBtn, $cancelBtn))
    }

}
function Hide-Console {
    $consolePtr = [Console.Window]::GetConsoleWindow()
    #0 hide
    [Console.Window]::ShowWindow($consolePtr, 0)
}
#####################################################################################################################################################################################
#                                                                                   Inventorization                                                                                 #
#####################################################################################################################################################################################

if ($enableGUI -eq "2") {
    $manufactJson = Invoke-RestMethod -Uri "$baseUrl/api/v1/manufacturers?search=$sysManufacturer" -Headers $header -Method GET
    $manufactJson = $manufactJson | Select-Object -ExpandProperty rows
    $manufact = ( $manufactJson | Select-Object -expand "name" )
    if ($sysManufacturer -eq $manufact)
    { $manId = $manufactJson.id } else {
        $JSON = @{"name" = $sysManufacturer } | ConvertTo-Json
        Invoke-RestMethod -Uri $baseUrl/api/v1/manufacturers -Headers $header -Method POST -Body $JSON
        $manId = Invoke-RestMethod -Uri "$baseUrl/api/v1/manufacturers?search=$sysManufacturer" -Headers $header -Method GET | Select-Object -ExpandProperty rows | Select-Object -expand id
    }
    $catJson = Invoke-RestMethod -Uri "$baseUrl/api/v1/categories?search=$categoryName" -Headers $header -Method GET
    $catJson = $catJson | Select-Object -ExpandProperty rows
    $cat = ( $catJson | Select-Object -expand "name" )
    if ($categoryName -eq $cat)
    { $categoryId = $catJson.id } else {
        $JSON = @{"name" = $categoryName ; "category_type" = "asset" ; "checkin_email" = 'false' } | ConvertTo-Json
        Invoke-RestMethod -Uri $baseUrl/api/v1/categories -Headers $header -Method POST -Body $JSON
        $categoryId = Invoke-RestMethod -Uri "$baseUrl/api/v1/categories?search=$categoryName" -Headers $header -Method GET | Select-Object -ExpandProperty rows | Select-Object -expand id
    }
    $modelJson = Invoke-RestMethod -Uri "$baseUrl/api/v1/models?limit=50&offset=0&search=$sysModel" -Headers $header -Method GET
    $modelJson = $modelJson | Select-Object -ExpandProperty rows
    $model = ( $modelJson | Select-Object -expand "name" )
    if ($sysModel -eq $model)
    { $modelId = $modelJson.id } else {
        $JSON = @{"name" = $sysModel ; "manufacturer_id" = $manId ; "category_id" = $categoryId ; "fieldset_id" = $fsField } | ConvertTo-Json
        Invoke-RestMethod -Uri "$baseUrl/api/v1/models" -Headers $header -Method POST -Body $JSON
        $modelId = Invoke-RestMethod -Uri $baseUrl/api/v1/models?search=$sysModel -Headers $header -Method GET | Select-Object -ExpandProperty rows | Select-Object -expand id
    }
    $assetJson = Invoke-RestMethod -Uri "$baseUrl/api/v1/hardware?search=$serialNumber" -Headers $header -Method GET
    $assetJson = $assetJson | Select-Object -ExpandProperty rows
    $assets = ( $assetJson | Select-Object -expand "serial" )
    if ($serialNumber -eq $assets) {
        Write-Output "Asset already exists"
    }
    else {
        $JSON = @{"name" = $computername ; "status_id" = $statusID ; "model_id" = $modelId ; $ramField = $ramsize; $cpuField = $CPU ; $diskField = $disksize ; "serial" = $serialNumber ; $macField = $primaryMacaddress ; $osField = $operatingSystem } | ConvertTo-Json
        Invoke-RestMethod -Uri "$baseUrl/api/v1/hardware" -Headers $header -Method POST -Body $JSON
        Write-Output "Asset created"
        exit
    }
}
else {
    if ($enableGUI -eq "1") {
        Add-Type -AssemblyName System.Windows.Forms
        [System.Windows.Forms.Application]::EnableVisualStyles()
        Add-Type -Name Window -Namespace Console -MemberDefinition '
[DllImport("Kernel32.dll")]
public static extern IntPtr GetConsoleWindow();

[DllImport("user32.dll")]
public static extern bool ShowWindow(IntPtr hWnd, Int32 nCmdShow);
'
        $iconKey = 'iVBORw0KGgoAAAANSUhEUgAAACAAAAAgCAYAAABzenr0AAAABmJLR0
QA/wD/AP+gvaeTAAAAc0lEQVRYhe2XXQrAIAyDM/H+p5rn6p5k4gLK/Flh+UBQUJqGYhUQNwmALR
5nHfQo5jY/J0oZE7G1YSI0wbAoWDcSIAESIAESIAGsG+5qywCcOvCv9wBzIPO2FrKDVq0prh0YrY
Wu8y4d+OweSBviPT4mQlxtAyI4cVGDVAAAAABJRU5ErkJggg=='

        $Form = New-Object system.Windows.Forms.Form
        $Form.ClientSize = '379,221'
        $Form.text = "Inventorization"
        $Form.TopMost = $false

        $iconBase64 = "$iconKey"
        $iconBytes = [Convert]::FromBase64String($iconBase64)
        $stream = New-Object IO.MemoryStream($iconBytes, 0, $iconBytes.Length)
        $stream.Write($iconBytes, 0, $iconBytes.Length);
        $Form.Icon = [System.Drawing.Icon]::FromHandle((New-Object System.Drawing.Bitmap -Argument $stream).GetHIcon())

        $emailAddress = New-Object system.Windows.Forms.TextBox
        $emailAddress.multiline = $false
        $emailAddress.width = 249
        $emailAddress.height = 20
        $emailAddress.location = New-Object System.Drawing.Point(62, 34)
        $emailAddress.Font = 'Microsoft Sans Serif,10'

        $Label1 = New-Object system.Windows.Forms.Label
        $Label1.text = "Please enter your emailaddress"
        $Label1.AutoSize = $true
        $Label1.width = 25
        $Label1.height = 10
        $Label1.location = New-Object System.Drawing.Point(62, 63)
        $Label1.Font = 'Microsoft Sans Serif,10'

        if ($getTag -eq "1") {
            $compName = New-Object system.Windows.Forms.TextBox
            $compName.multiline = $false
            $compName.width = 249
            $compName.height = 20
            $compName.location = New-Object System.Drawing.Point(62, 94)
            $compName.Font = 'Microsoft Sans Serif,10'

            $Label2 = New-Object system.Windows.Forms.Label
            $Label2.text = "Please enter your computername `naccording to the sticker at the bottom"
            $Label2.AutoSize = $true
            $Label2.width = 25
            $Label2.height = 10
            $Label2.location = New-Object System.Drawing.Point(62, 123)
            $Label2.Font = 'Microsoft Sans Serif,10'
        }

        $AddDeviceBtn = New-Object system.Windows.Forms.Button
        $AddDeviceBtn.text = "Add"
        $AddDeviceBtn.width = 60
        $AddDeviceBtn.height = 30
        $AddDeviceBtn.location = New-Object System.Drawing.Point(63, 170)
        $AddDeviceBtn.Font = 'Microsoft Sans Serif,10'

        $cancelBtn = New-Object system.Windows.Forms.Button
        $cancelBtn.text = "Cancel"
        $cancelBtn.width = 60
        $cancelBtn.height = 30
        $cancelBtn.location = New-Object System.Drawing.Point(251, 170)
        $cancelBtn.Font = 'Microsoft Sans Serif,10'
        $Form.CancelButton = $cancelBtn
        $Form.Controls.Add($cancelBtn)
        if ($getTag -eq "0") {
            $Form.controls.AddRange(@($emailAddress, $Label1, $compName, $Label2, $AddDeviceBtn, $cancelBtn))
        }
        else {
            $Form.controls.AddRange(@($emailAddress, $Label1, $compName, $Label2, $AddDeviceBtn, $cancelBtn))
        }





        function addDevice {

            $script:unameInput = $emailAddress.Text.Trim()
            $script:assetTag = $compName.Text.Trim()


            ################################################################################################################################################
            #To test this output, run:														       
            #echo $computername ,  $sysModel ,  $sysManufacturer ,  $CPU ,  $disksize ,  $ramsize ,  $primaryMacaddress ,  $operatingSystem , $serialNumber
            ################################################################################################################################################

            ###########################################################
            #call to your Snipe-instance: Does the manufacturer exist?
            #if not, create the manufacturer; 			  
            #else get the manufacturer's ID      			  
            ###########################################################

            BuildGUI
            AddManufacturer
            AddCategory
            AddManufacturer
            AddAsset


            $manufactJson = Invoke-RestMethod -Uri "$baseUrl/api/v1/manufacturers?search=$sysManufacturer" -Headers $header -Method GET
            $manufactJson = $manufactJson | Select-Object -ExpandProperty rows
            $manufact = ( $manufactJson | Select-Object -expand "name" )

            if ($sysManufacturer -eq $manufact) {
                $manId = $manufactJson.id 
            }
            else {
                $JSON = @{"name" = $sysManufacturer } | ConvertTo-Json
                Invoke-RestMethod -Uri $baseUrl/api/v1/manufacturers -Headers $header -Method POST -Body $JSON
                $manId = Invoke-RestMethod -Uri "$baseUrl/api/v1/manufacturers?search=$sysManufacturer" -Headers $header -Method GET | Select-Object -ExpandProperty rows | Select-Object -expand id
            }

            ###########################################################
            #call to your Snipe-instance: Does the category exist?	  
            #if not, create the category; 				  
            #else get the category's ID		      		  
            ###########################################################
            $catJson = Invoke-RestMethod -Uri "$baseUrl/api/v1/categories?search=$categoryName" -Headers $header -Method GET
            $catJson = $catJson | Select-Object -ExpandProperty rows
            $cat = ( $catJson | Select-Object -expand "name" )
            if ($categoryName -eq $cat)
            { $categoryId = $catJson.id } else {
                $JSON = @{"name" = $categoryName ; "category_type" = "asset" ; "checkin_email" = 'false' } | ConvertTo-Json
                Invoke-RestMethod -Uri $baseUrl/api/v1/categories -Headers $header -Method POST -Body $JSON
                $categoryId = Invoke-RestMethod -Uri "$baseUrl/api/v1/categories?search=$categoryName" -Headers $header -Method GET | Select-Object -ExpandProperty rows | Select-Object -expand id
            }

            ###########################################################
            #call to your Snipe-instance: Does the model exist?	  
            #if not, create the model; 				  
            #else get the model's ID		      		  
            ###########################################################

            $modelJson = Invoke-RestMethod -Uri "$baseUrl/api/v1/models?limit=50&offset=0&search=$sysModel" -Headers $header -Method GET
            $modelJson = $modelJson | Select-Object -ExpandProperty rows
            $model = ( $modelJson | Select-Object -expand "name" )


            if ($sysModel -eq $model) {
                $modelId = $modelJson.id 
            }
            else {
                $JSON = @{"name" = $sysModel ; "manufacturer_id" = $manId ; "category_id" = $categoryId ; "fieldset_id" = $fsField } | ConvertTo-Json
                Invoke-RestMethod -Uri "$baseUrl/api/v1/models" -Headers $header -Method POST -Body $JSON
                $modelId = Invoke-RestMethod -Uri $baseUrl/api/v1/models?search=$sysModel -Headers $header -Method GET | Select-Object -ExpandProperty rows | Select-Object -expand id
            }



            ###########################################################
            #call to your Snipe-instance: Does the asset exist?	  
            #if not, create the asset; 				  
            #else get the asset's ID		      		  
            #IMPORTANT: The serial number of the asset will be used   
            #to look the asset up. Make sure, your devices have a SN  
            #as older Boards sometimes don't propagate it		  
            ###########################################################

            $assetJson = Invoke-RestMethod -Uri "$baseUrl/api/v1/hardware?search=$serialNumber" -Headers $header -Method GET
            $assetJson = $assetJson | Select-Object -ExpandProperty rows
            $assets = ( $assetJson | Select-Object -expand "serial" )

            if ($serialNumber -eq $assets) {
                Write-Output "Asset already exists"
                $assetId = Invoke-RestMethod -Uri "$baseUrl/api/v1/hardware?search=$serialNumber" -Headers $header -Method GET | Select-Object -ExpandProperty rows | Select-Object -expand id
            }
            else {
                if ($getTag -eq "1") {
                    $JSON = @{"name" = $computername ; "asset_tag" = $assetTag  ; "status_id" = $statusID ; "model_id" = $modelId ; $ramField = $ramsize; $cpuField = $CPU ; $diskField = $disksize ; "serial" = $serialNumber ; $macField = $primaryMacaddress ; $osField = $operatingSystem } | ConvertTo-Json
                }
                else {
                    $JSON = @{"name" = $computername ; "status_id" = $statusID ; "model_id" = $modelId ; $ramField = $ramsize; $cpuField = $CPU ; $diskField = $disksize ; "serial" = $serialNumber ; $macField = $primaryMacaddress ; $osField = $operatingSystem } | ConvertTo-Json    
                }
                Invoke-RestMethod -Uri "$baseUrl/api/v1/hardware" -Headers $header -Method POST -Body $JSON
                $assetId = Invoke-RestMethod -Uri "$baseUrl/api/v1/hardware?search=$serialNumber" -Headers $header -Method GET | Select-Object -ExpandProperty rows | Select-Object -expand id
            }
	

            ############################################################
            #call to your Snipe-instance: Does the user exist?	   
            #if not, create the user and disallow login; 		   
            # >>"activated" = 'false'<<				   
            #else get the user's ID 		      		   
            #IMPORTANT: This script expects either name@example.com or 
            #firstname.lastname@example.com as username.		   
            #I highly suggest to take the firstname.lastname approach  
            ############################################################

            $unameSplit = ($unameInput).split('@')[0]
            if (($unameSplit.ToCharArray() | Where-Object { $_ -eq '.' } | Measure-Object).Count -gt 0) {
                $firstName = (($unameSplit).split(".")[0]).Trim()
                $lastName = (($unameSplit).split(".")[1]).Trim()
            }
            else {
                $firstName = $unameSplit.Trim() 
            }

            $userJson = Invoke-RestMethod -Uri "$baseUrl/api/v1/users?search=$unameInput" -Headers $header -Method GET
            $userJson = $userJson | Select-Object -ExpandProperty rows
            $users = ( $userJson | Select-Object -expand "username")
            if ($unameInput -eq $users) {
                Write-Output "user already exists"
                $userId = Invoke-RestMethod -Uri "$baseUrl/api/v1/users?search=$unameInput" -Headers $header -Method GET | Select-Object -ExpandProperty rows | Select-Object -expand id
            }
            else {
                if ($null -eq $lastname ) {
                    $JSON = @{"first_name" = $firstName ; "username" = $unameSplit ; "email" = $unameInput ; "password" = $ranPass ; "password_confirmation" = $ranPass ; "activated" = 'false' } | ConvertTo-Json
                    Invoke-RestMethod -Uri "$baseUrl/api/v1/users" -Headers $header -Method POST -Body $JSON
                    $userId = Invoke-RestMethod -Uri "$baseUrl/api/v1/users?search=$unameInput" -Headers $header -Method GET | Select-Object -ExpandProperty rows | Select-Object -expand id
                }
                else {
                    $JSON = @{"first_name" = $firstName ; "last_name" = $lastName ; "username" = $unameSplit ; "email" = $unameInput ; "password" = $ranPass ; "password_confirmation" = $ranPass ; "activated" = 'false' } | ConvertTo-Json
                    Invoke-RestMethod -Uri "$baseUrl/api/v1/users" -Headers $header -Method POST -Body $JSON
                    $userId = Invoke-RestMethod -Uri "$baseUrl/api/v1/users?search=$unameInput" -Headers $header -Method GET | Select-Object -ExpandProperty rows | Select-Object -expand id    
                }
            }

            ##########################################
            # check out sasset to user        	 
            # and update the status and asset tag	 
            # if necessary	(e.g. if you change the  
            # sticker on the device			 
            ##########################################

            $JSON = @{ "checkout_to_type" = "user" ; "assigned_user" = $userId } | ConvertTo-Json
            Invoke-RestMethod -Uri "$baseUrl/api/v1/hardware/$assetId/checkout" -Headers $header -Method POST -Body $JSON
            $JSON = @{ "status_id" = $statusID ; "asset_tag" = $assetTag } | ConvertTo-Json
            Invoke-RestMethod -Uri "$baseUrl/api/v1/hardware/$assetId" -Headers $header -Method PATCH -Body $JSON


            Add-Type -AssemblyName PresentationCore, PresentationFramework
            $ButtonType = [System.Windows.MessageBoxButton]::OK
            $MessageIcon = [System.Windows.MessageBoxImage]::Question
            $MessageBody = "Your device has been inventorized!"
            $MessageTitle = "success"
 
            [System.Windows.MessageBox]::Show($MessageBody, $MessageTitle, $ButtonType, $MessageIcon)
 
            $Form.Close()

        }
        Hide-Console
        $AddDeviceBtn.Add_Click({ addDevice })
        [void]$Form.ShowDialog()
    }
    else {
        if ($getTag -eq "1") {
            $unameInput = Read-Host -Prompt 'Please type in your emailaddress'
            $assetTag = Read-Host -Prompt 'Please type in the name of your device/whats written on the sticker' 
        }
        else {
            $uname = whoami
            if ($uname -ilike "cisadom*") {
                if ($unameInput -inotlike "adm_*") {
                    $unameInput = $uname.Split("\")[1]
                }
                else {
                    $unameInput = ""
                }
            }
            else {
                $unameInput = ""
            }
        }

        ###########################################################
        #call to your Snipe-instance: Does the manufacturer exist?
        #if not, create the manufacturer; 			  
        #else get the manufacturer's ID      			  
        ###########################################################
        $manufactJson = Invoke-RestMethod -Uri "$baseUrl/api/v1/manufacturers?search=$sysManufacturer" -Headers $header -Method GET
        $manufactJson = $manufactJson | Select-Object -ExpandProperty rows
        $manufact = ( $manufactJson | Select-Object -expand "name" )
        if ($sysManufacturer -eq $manufact) {
            $manId = $manufactJson.id 
        }
        else {
            $JSON = @{"name" = $sysManufacturer } | ConvertTo-Json
            Invoke-RestMethod -Uri $baseUrl/api/v1/manufacturers -Headers $header -Method POST -Body $JSON
            $manId = Invoke-RestMethod -Uri "$baseUrl/api/v1/manufacturers?search=$sysManufacturer" -Headers $header -Method GET | Select-Object -ExpandProperty rows | Select-Object -expand id
        }

        ###########################################################
        #call to your Snipe-instance: Does the device or the category exist?
        #if not, create the category; 				  
        #else get the category's ID      			  
        ###########################################################
        $assetJson = Invoke-RestMethod -Uri "$baseUrl/api/v1/hardware?search=$serialNumber" -Headers $header -Method GET
        $assetJson = $assetJson | Select-Object -ExpandProperty rows
        $assets = ( $assetJson | Select-Object -expand "serial" )
        if ($serialNumber -eq $assets) {
            ## serial number exists
            $categoryId = Invoke-RestMethod -Uri "$baseUrl/api/v1/hardware?search=$serialNumber" -Headers $header -Method GET | Select-Object -ExpandProperty rows | Select-Object -ExpandProperty category | Select-Object -ExpandProperty id
        }
        else {
            ## no serial found
            $categoryName = $categoryName + "_found"
            $catJson = Invoke-RestMethod -Uri "$baseUrl/api/v1/categories?search=$categoryName" -Headers $header -Method GET
            $catJson = $catJson | Select-Object -ExpandProperty rows
            $cat = ( $catJson | Select-Object -expand "name" )

            if ($categoryName -eq $cat) {
                $categoryId = ($catJson | Select-Object -First 1).id
            }
            else {
                $JSON = @{"name" = $categoryName ; "category_type" = "asset" ; "checkin_email" = 'false' } | ConvertTo-Json
                Invoke-RestMethod -Uri $baseUrl/api/v1/categories -Headers $header -Method POST -Body $JSON
                $categoryId = Invoke-RestMethod -Uri "$baseUrl/api/v1/categories?search=$categoryName" -Headers $header -Method GET | Select-Object -ExpandProperty rows | Select-Object -expand id
            }
        }

        ###########################################################
        #call to your Snipe-instance: Does the model exist?	  
        #if not, create the model; 				  
        #else get the model's ID		      		  
        ###########################################################
        $modelJson = Invoke-RestMethod -Uri "$baseUrl/api/v1/models?limit=100&offset=0&search=$sysModel" -Headers $header -Method GET
        $modelJson = $modelJson | Select-Object -ExpandProperty rows
        $model = ( $modelJson | Select-Object -expand "name" )
        if ($sysModel -eq $model) {
            $modelId = $modelJson.id 
        }
        else {
            $JSON = @{"name" = $sysModel ; "manufacturer_id" = $manId ; "category_id" = $categoryId ; "fieldset_id" = $fsField } | ConvertTo-Json
            Invoke-RestMethod -Uri "$baseUrl/api/v1/models" -Headers $header -Method POST -Body $JSON
            $modelId = Invoke-RestMethod -Uri $baseUrl/api/v1/models?search=$sysModel -Headers $header -Method GET | Select-Object -ExpandProperty rows | Select-Object -expand id
            if (!($modelId)) {
                Write-Output "ERROR. Unable to create model: $($sysModel)"
            }
        }

        ###########################################################
        #call to your Snipe-instance: Does the asset exist?
        #if not, create the asset;
        #else get the asset's ID		      		  
        #IMPORTANT: The serial number of the asset will be used   
        #to look the asset up. Make sure, your devices have a SN  
        #as older Boards sometimes don't propagate it
        # NEW -- Update IP address
        ###########################################################
        $assetJson = Invoke-RestMethod -Uri "$baseUrl/api/v1/hardware?search=$serialNumber" -Headers $header -Method GET
        $assetJson = $assetJson | Select-Object -ExpandProperty rows
        $assets = ( $assetJson | Select-Object -expand "serial" )
        if ($serialNumber -eq $assets) {
            Write-Output "Asset already exists"
            $assetId = Invoke-RestMethod -Uri "$baseUrl/api/v1/hardware?search=$serialNumber" -Headers $header -Method GET | Select-Object -ExpandProperty rows | Select-Object -expand id
            $assetTag = Invoke-RestMethod -Uri "$baseUrl/api/v1/hardware?search=$serialNumber" -Headers $header -Method GET | Select-Object -ExpandProperty rows | Select-Object -expand asset_tag
            $assetStatus = Invoke-RestMethod -Uri "$baseUrl/api/v1/hardware?search=$serialNumber" -Headers $header -Method GET | Select-Object -ExpandProperty rows | Select-Object -ExpandProperty status_label | Select-Object -ExpandProperty id

            # NEW -- Update IP address
            $JSON = @{$ramField = $ramsize; $cpuField = $CPU ; $diskField = $disksize ; $macField = $primaryMacaddress ; $osField = $operatingSystem ; $ipField = $primaryIP } | ConvertTo-Json
            Invoke-RestMethod -Uri "$baseUrl/api/v1/hardware/$($assetId)" -Headers $header -Method Patch -Body $JSON
        }
        else {
            if ($getTag -eq "1") {
                $JSON = @{"name" = $computername ; "asset_tag" = $assetTag  ; "status_id" = $statusID ; "model_id" = $modelId ; $ramField = $ramsize; $cpuField = $CPU ; $diskField = $disksize ; "serial" = $serialNumber ; $macField = $primaryMacaddress ; $osField = $operatingSystem ; $ipField = $primaryIP  } | ConvertTo-Json
            }
            else {
                $JSON = @{"name" = $computername ; "status_id" = $statusID ; "model_id" = $modelId ; $ramField = $ramsize; $cpuField = $CPU ; $diskField = $disksize ; "serial" = $serialNumber ; $macField = $primaryMacaddress ; $osField = $operatingSystem ; $ipField = $primaryIP } | ConvertTo-Json    
            }
            Invoke-RestMethod -Uri "$baseUrl/api/v1/hardware" -Headers $header -Method POST -Body $JSON
            $assetId = Invoke-RestMethod -Uri "$baseUrl/api/v1/hardware?search=$serialNumber" -Headers $header -Method GET | Select-Object -ExpandProperty rows | Select-Object -expand id
            $assetTag = Invoke-RestMethod -Uri "$baseUrl/api/v1/hardware?search=$serialNumber" -Headers $header -Method GET | Select-Object -ExpandProperty rows | Select-Object -expand asset_tag
            $assetStatus = Invoke-RestMethod -Uri "$baseUrl/api/v1/hardware?search=$serialNumber" -Headers $header -Method GET | Select-Object -ExpandProperty rows | Select-Object -ExpandProperty status_label | Select-Object -ExpandProperty id
        }


        ###########################################################
        #call to your Snipe-instance: Does the user exist?
        #if not, create the user and checkout the asset to him/her;
        #else get the asset's ID and create the user and checkout 
        # the asset to him/her		      		  
        ###########################################################
        if ($unameInput) {

            #$unameSplit = ($unameInput).split('@')[0]
            #if (($unameSplit.ToCharArray() | Where-Object { $_ -eq '.' } | Measure-Object).Count -gt 0) {
            #    $firstName = (($unameSplit).split(".")[0]).Trim()
            #    $lastName = (($unameSplit).split(".")[1]).Trim()
            #}
            #else {
            #    $firstName = $unameSplit.Trim() 
            #}
    
            $userJson = Invoke-RestMethod -Uri "$baseUrl/api/v1/users?limit=10&offset=0&username=$($unameInput)&deleted=false&all=false" -Headers $header -Method GET
            $userJson = $userJson | Select-Object -ExpandProperty rows
            $users = ( $userJson | Select-Object -expand "username")
            if ($unameInput -eq $users) {
                Write-Output "user already exists"
                $userId = Invoke-RestMethod -Uri "$baseUrl/api/v1/users?limit=10&offset=0&username=$($unameInput)&deleted=false&all=false" -Headers $header -Method GET | Select-Object -ExpandProperty rows | Select-Object -expand id
            }
            else {
                Write-Output "user does not exists"
                $userId = 'not trying to checkout asset, user does not exists.'
            }

            #else {
            #    if ($null -eq $lastname ) {
            #        $JSON = @{"first_name" = $firstName ; "username" = $unameSplit ; "email" = $unameInput ; "password" = $ranPass ; "password_confirmation" = $ranPass ; "activated" = 'false' } | ConvertTo-Json
            #        Invoke-RestMethod -Uri "$baseUrl/api/v1/users" -Headers $header -Method POST -Body $JSON
            #        $userId = Invoke-RestMethod -Uri "$baseUrl/api/(v1/users?search=$unameInput" -Headers $header -Method GET | Select-Object -ExpandProperty rows | Select-Object -expand id
            #    }
            #    else {
            #        $JSON = @{"first_name" = $firstName ; "last_name" = $lastName ; "username" = $unameSplit ; "email" = $unameInput ; "password" = $ranPass ; "password_confirmation" = $ranPass ; "activated" = 'false' } | ConvertTo-Json
            #        Invoke-RestMethod -Uri "$baseUrl/api/v1/users" -Headers $header -Method POST -Body $JSON
            #        $userId = Invoke-RestMethod -Uri "$baseUrl/api/v1/users?search=$unameInput" -Headers $header -Method GET | Select-Object -ExpandProperty rows | Select-Object -expand id    
            #    }
            #}

            if (($assetStatus -eq '2' -or $assetStatus -eq '4' -or $assetStatus -eq '8') -and ($userId)) {
                #check if already assigned
                $assignedTo = Invoke-RestMethod -Uri "$baseUrl/api/v1/hardware?search=$serialNumber" -Headers $header -Method GET | Select-Object -ExpandProperty rows | Select-Object -ExpandProperty assigned_to | Select-Object -ExpandProperty username 

                if (!($assignedTo)) {
                    #not assigned
                    $JSON = @{ "checkout_to_type" = "user" ; "assigned_user" = $userId } | ConvertTo-Json
                    Invoke-RestMethod -Uri "$baseUrl/api/v1/hardware/$assetId/checkout" -Headers $header -Method POST -Body $JSON
                    $JSON = @{ "status_id" = $assetStatus ; "asset_tag" = $assetTag } | ConvertTo-Json
                    Invoke-RestMethod -Uri "$baseUrl/api/v1/hardware/$assetId" -Headers $header -Method PATCH -Body $JSON
                }
                else {
                    if ($assignedTo -ne $unameInput) {
                        $JSON = @{ "status_id" = $assetStatus } | ConvertTo-Json
                        Invoke-RestMethod -Uri "$baseUrl/api/v1/hardware/$assetId/checkin" -Headers $header -Method POST -Body $JSON
                        $JSON = @{ "checkout_to_type" = "user" ; "assigned_user" = $userId } | ConvertTo-Json
                        Invoke-RestMethod -Uri "$baseUrl/api/v1/hardware/$assetId/checkout" -Headers $header -Method POST -Body $JSON
                        $JSON = @{ "status_id" = $assetStatus ; "asset_tag" = $assetTag } | ConvertTo-Json
                        Invoke-RestMethod -Uri "$baseUrl/api/v1/hardware/$assetId" -Headers $header -Method PATCH -Body $JSON
                    }
                    else {
                        Write-Output "asset correctly assigned, nothing to do"
                    }
                }
            }
            elseif ($userId) {
                $JSON = @{ "checkout_to_type" = "user" ; "assigned_user" = $userId } | ConvertTo-Json
                Invoke-RestMethod -Uri "$baseUrl/api/v1/hardware/$assetId/checkout" -Headers $header -Method POST -Body $JSON
                $JSON = @{ "status_id" = $statusID ; "asset_tag" = $assetTag } | ConvertTo-Json
                Invoke-RestMethod -Uri "$baseUrl/api/v1/hardware/$assetId" -Headers $header -Method PATCH -Body $JSON
            }
        }
        else {
            Write-Output "not trying to checkout asset, run the script as domain user"
        }

        ###########################################################
        #print variables and 
        #stop transcript
        ###########################################################
        Get-Variable | Out-String
        Stop-Transcript
        exit
    }
}
# SIG # Begin signature block
# MIIP0QYJKoZIhvcNAQcCoIIPwjCCD74CAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUC3YdrJ76xhC08Tv89T13dx07
# 6/6ggg1CMIIGjDCCBHSgAwIBAgITYgAAAAUq+g6wdbTX6AAAAAAABTANBgkqhkiG
# 9w0BAQ0FADAXMRUwEwYDVQQDEwxDaXNhUm9vdENhMDEwHhcNMTgxMDI0MTA1NjU2
# WhcNMjgxMDI0MTEwNjU2WjBBMRMwEQYKCZImiZPyLGQBGRYDUFJJMRcwFQYKCZIm
# iZPyLGQBGRYHQ0lTQURPTTERMA8GA1UEAxMIQ1NJU1NVMDMwggEiMA0GCSqGSIb3
# DQEBAQUAA4IBDwAwggEKAoIBAQDMvnfZM1feRD0X2srd5X0g/eQUoe3pvNNEQEBR
# bYSKHF94sbq12eLzKVEoTlSbUEVhXR3rXgX2cdYPJt0/oQAYKSIgFkp1dlaw6mzI
# F0MCIP6OehwpfueCCHBu5i3VLBq14shUfi7PJhhPDbSHuv3PZ6E456dbUP3jN9HW
# 6xxOcghTyQEQ1BgN4kkTrC4Tn5v51o4wpK0rkCsJMBo+etUrtx2QfifZbCk78mYv
# omcw5ltEuS+6EvByM0VnddSp9wW/SRsxp8gTmfybjHt/ChWdZqW4K7PVh/JnrJP5
# 6Bw6+1cGgVbkgGRyVzz8my6xHhsqdRGXnOEGO3HG42lc850dAgMBAAGjggKlMIIC
# oTAQBgkrBgEEAYI3FQEEAwIBADAdBgNVHQ4EFgQUdr6RfaLj/KqktE3hKwz+C/Z6
# CmIwGQYJKwYBBAGCNxQCBAweCgBTAHUAYgBDAEEwCwYDVR0PBAQDAgGGMA8GA1Ud
# EwEB/wQFMAMBAf8wHwYDVR0jBBgwFoAUD4pZ/Az4Jw8fAOI/a1Gj9dYqT24wggED
# BgNVHR8EgfswgfgwgfWggfKgge+GgbZsZGFwOi8vL0NOPUNpc2FSb290Q2EwMSxD
# Tj1DU1BXVlJDQTAxLENOPUNEUCxDTj1QdWJsaWMlMjBLZXklMjBTZXJ2aWNlcyxD
# Tj1TZXJ2aWNlcyxDTj1Db25maWd1cmF0aW9uLERDPUNJU0FET00sREM9UFJJP2Nl
# cnRpZmljYXRlUmV2b2NhdGlvbkxpc3Q/YmFzZT9vYmplY3RDbGFzcz1jUkxEaXN0
# cmlidXRpb25Qb2ludIY0aHR0cDovL2NzcHd2Y2F3ZWIwMS5jaXNhZG9tLnByaS9w
# a2kvQ2lzYVJvb3RDYTAxLmNybDCCAQsGCCsGAQUFBwEBBIH+MIH7MIGrBggrBgEF
# BQcwAoaBnmxkYXA6Ly8vQ049Q2lzYVJvb3RDYTAxLENOPUFJQSxDTj1QdWJsaWMl
# MjBLZXklMjBTZXJ2aWNlcyxDTj1TZXJ2aWNlcyxDTj1Db25maWd1cmF0aW9uLERD
# PUNJU0FET00sREM9UFJJP2NBQ2VydGlmaWNhdGU/YmFzZT9vYmplY3RDbGFzcz1j
# ZXJ0aWZpY2F0aW9uQXV0aG9yaXR5MEsGCCsGAQUFBzAChj9odHRwOi8vY3Nwd3Zj
# YXdlYjAxLmNpc2Fkb20ucHJpL3BraS9DU1BXVlJDQTAxX0Npc2FSb290Q2EwMS5j
# cnQwDQYJKoZIhvcNAQENBQADggIBAL4bfl2R9dqZTB881NgqtTlPhqRwEQZXrgPg
# IT3+3fsmlzQ+t00mVBrb4/zniJH67zuKAMfqO0ZJiL1ujbB2STCl1qGd1GfP8Uyt
# vsb2fo8FJVTq9o6MOEzQH3NRcs3MDDsW5EBcALPbj+Q3uhjUEWE9J2xoHVlNCakP
# +YAVlZa+xuzZH2pVnzf7DFniy++pvAhb1xAA43e2OGSWofCe0vbc4wSExIgG7YOf
# iKhplbaJwc7mUO//Y1QrU+sLLJjFtijCkx4OpRr4JRBBmVK8fSHBg7V6RbicIt8j
# IJJtymCSCV4HO8Ev31LeQtv+HIjbDXhuSFCG0zwF+AjMdnkR3Sbox8mtOH/tip9X
# PfCJ5sWvs6SXfWRaHIg9UYdc8r+vssRL1MAdK2lVCnQuBlwmKhRRSWAt5DWi9LyE
# qALtLqKAeOW+jHKqxd7ueqpOraosu0jYAfGlwBHOPOv+26ZfYs0HZLLdZN76ImYD
# asViGyVE1VTtXNWEZpC79zfrAap8em37vWw0NoiYWDQVop8tfxhdKvx4Ti8rns6e
# K/BzDbbuTcIJ67Vhxu52SN0wIQ5vGgYe3BUiqF1ZQ2tnZka5qwfo51UAF3hn/nJQ
# caEUtrgDoyjvGJgi1DgTUEF1mgnBmFO6roaMiMjAQzSS1YXSLtGaGztUEPJ2II1d
# FVzcmkDAMIIGrjCCBZagAwIBAgITUwAAks59rC73hPU00QAAAACSzjANBgkqhkiG
# 9w0BAQsFADBBMRMwEQYKCZImiZPyLGQBGRYDUFJJMRcwFQYKCZImiZPyLGQBGRYH
# Q0lTQURPTTERMA8GA1UEAxMIQ1NJU1NVMDMwHhcNMjIwODI3MTc1MjE4WhcNMjUw
# ODI3MTgwMjE4WjBUMRMwEQYKCZImiZPyLGQBGRYDUFJJMRcwFQYKCZImiZPyLGQB
# GRYHQ0lTQURPTTEMMAoGA1UECxMDQURNMRYwFAYDVQQDEw1BRE0gSXZvIEJhY2Nv
# MIIBIjANBgkqhkiG9w0BAQEFAAOCAQ8AMIIBCgKCAQEA0f7Chsp2o+HLXeL0AL1B
# PFE86kiqlNBs/MVECCzDsnWJFg2zQylvX3OB+em/7kuUkuapBL9r/Bi+dyi+V0JA
# CNYkMCWB9nzmzfYVUhbkrXUBuPWxnweQWDhW88Gxjrj5aVv1i1GOvcIOUJXvHIrh
# 5BOu57HH6bhWRrdIQ2IA9XDEsRq9APhyctjG7Gy5WmE42vrLDXsQkXlIxg0vOQ9Z
# AVZuLq6CWahtDf5BJdAe3dR/mNC4I1T1rOUzU6zVJ/kmdbWTMsXRZvafeyYElWN0
# 5dRMSLI4PAZDvV1iNm2IguC9VbisnJTF099ovUqB9Tw4fWYBfAbDWhf//EL/aMj7
# TQIDAQABo4IDijCCA4YwCwYDVR0PBAQDAgeAMDsGCSsGAQQBgjcVBwQuMCwGJCsG
# AQQBgjcVCIeGrACfm1qClZULhauDPoWNjnFiorEmhcbpOwIBZAIBAzAdBgNVHQ4E
# FgQUnUs1boySRAwB06fBm+Yk9Zx+dlMwHwYDVR0jBBgwFoAUdr6RfaLj/KqktE3h
# Kwz+C/Z6CmIwgfsGA1UdHwSB8zCB8DCB7aCB6qCB54aBsmxkYXA6Ly8vQ049Q1NJ
# U1NVMDMsQ049Q1NQV1ZDQUkwMyxDTj1DRFAsQ049UHVibGljJTIwS2V5JTIwU2Vy
# dmljZXMsQ049U2VydmljZXMsQ049Q29uZmlndXJhdGlvbixEQz1DSVNBRE9NLERD
# PVBSST9jZXJ0aWZpY2F0ZVJldm9jYXRpb25MaXN0P2Jhc2U/b2JqZWN0Q2xhc3M9
# Y1JMRGlzdHJpYnV0aW9uUG9pbnSGMGh0dHA6Ly9jc3B3dmNhd2ViMDEuY2lzYWRv
# bS5wcmkvcGtpL0NTSVNTVTAzLmNybDCCAUMGCCsGAQUFBwEBBIIBNTCCATEwgacG
# CCsGAQUFBzAChoGabGRhcDovLy9DTj1DU0lTU1UwMyxDTj1BSUEsQ049UHVibGlj
# JTIwS2V5JTIwU2VydmljZXMsQ049U2VydmljZXMsQ049Q29uZmlndXJhdGlvbixE
# Qz1DSVNBRE9NLERDPVBSST9jQUNlcnRpZmljYXRlP2Jhc2U/b2JqZWN0Q2xhc3M9
# Y2VydGlmaWNhdGlvbkF1dGhvcml0eTBTBggrBgEFBQcwAoZHaHR0cDovL2NzcHd2
# Y2F3ZWIwMS5jaXNhZG9tLnByaS9wa2kvQ1NQV1ZDQUkwMy5DSVNBRE9NLlBSSV9D
# U0lTU1UwMy5jcnQwMAYIKwYBBQUHMAGGJGh0dHA6Ly9jc3B3dmNhd2ViMDEuY2lz
# YWRvbS5wcmkvb2NzcDATBgNVHSUEDDAKBggrBgEFBQcDAzAbBgkrBgEEAYI3FQoE
# DjAMMAoGCCsGAQUFBwMDMDEGA1UdEQQqMCigJgYKKwYBBAGCNxQCA6AYDBZhZG1f
# aWJhY2NvQGNpc2Fkb20ucHJpMFAGCSsGAQQBgjcZAgRDMEGgPwYKKwYBBAGCNxkC
# AaAxBC9TLTEtNS0yMS0xNzA5OTM0NzIzLTE4NzA1OTI4MzgtMjczOTQ5NTY0NS0y
# NTk5NDANBgkqhkiG9w0BAQsFAAOCAQEAHLputSVPpeQ3LwTCD+qMYjiCMIxOGwgU
# w2gCRui/1yFLjK8EO1iIr3RAeTxQs6lI1hIvfyRP//JaB30nMmvSUU4BdxHgOFQx
# vm4mpl4vJAGW3jgGmaaguud9lvlUX3wW1YINVxOxerMET+MCBcUzFt+hxElpi+bv
# olcQvexrDFkBv+Prldw2VR+m3DJGB3roa4au31bMUoDYc6ORarrur6lSclXTaMAL
# doOU+QtErQtnquhvy8obtWbtRd/gPxJAsyvOg0rJcGmP8MO7RtZGyS9sc72P2fV6
# POEqjXR6k6fY/EzIuR91XBqV3TEGUfLHN3+luGKQ+ZhEPK/06ek0XzGCAfkwggH1
# AgEBMFgwQTETMBEGCgmSJomT8ixkARkWA1BSSTEXMBUGCgmSJomT8ixkARkWB0NJ
# U0FET00xETAPBgNVBAMTCENTSVNTVTAzAhNTAACSzn2sLveE9TTRAAAAAJLOMAkG
# BSsOAwIaBQCgeDAYBgorBgEEAYI3AgEMMQowCKACgAChAoAAMBkGCSqGSIb3DQEJ
# AzEMBgorBgEEAYI3AgEEMBwGCisGAQQBgjcCAQsxDjAMBgorBgEEAYI3AgEVMCMG
# CSqGSIb3DQEJBDEWBBRpBVUFcjzj15i7KUGlxSRfVRw9yTANBgkqhkiG9w0BAQEF
# AASCAQBZY/l3CHVe5Y0W5bK2BvobfMxAIhKYhyLhwX642WOfZi89jNaPRXveJpa7
# +AQ4PAs7DSgM49t6T0Pqom17tU66N0hqpGLoYwY8fJKJVS5XuLeCb0JOwRak0tj0
# 3gDASexwjwpw9SilU3BrIdQMi4IOYrmrvv571vOYVWQfGQspZVm5jjeW5sookHy6
# kWN6Vvur9UfYHN1GySplXCavvRyh7O0bIyX7rxryafZ1J4/+uLVte5SDLq7sVVsg
# Uk7aGAANpovPPwDFCvSSYII+QcvGJL9n5gMqxwCbl1vlIk2B86IXzeeQOhBsdr8G
# qUZnVmr4UKBr4lX/cty5bTJ/sQTd
# SIG # End signature block
