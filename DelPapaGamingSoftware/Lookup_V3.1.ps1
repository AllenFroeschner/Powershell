Import-Module ActiveDirectory
$App=$true
$text = @"
·▄▄▄▄  ▄▄▄ .▄▄▌   ▄▄▄· ▄▄▄·  ▄▄▄· ▄▄▄·     ·▄▄▄▄  ▪  .▄▄ · ▄▄▄▄▄   
██▪ ██ ▀▄.▀·██•  ▐█ ▄█▐█ ▀█ ▐█ ▄█▐█ ▀█     ██▪ ██ ██ ▐█ ▀. •██     
▐█· ▐█▌▐▀▀▪▄██▪   ██▀·▄█▀▀█  ██▀·▄█▀▀█     ▐█· ▐█▌▐█·▄▀▀▀█▄ ▐█.▪   
██. ██ ▐█▄▄▌▐█▌▐▌▐█▪·•▐█ ▪▐▌▐█▪·•▐█ ▪▐▌    ██. ██ ▐█▌▐█▄▪▐█ ▐█▌·   
▀▀▀▀▀•  ▀▀▀ .▀▀▀ .▀    ▀  ▀ .▀    ▀  ▀     ▀▀▀▀▀• ▀▀▀ ▀▀▀▀  ▀▀▀  ▀ 
"@,@"
________         .__ __________                       ________  .__          __     
\______ \   ____ |  |\______   \_____  ___________    \______ \ |__| _______/  |_   
 |    |  \_/ __ \|  | |     ___/\__  \ \____ \__  \    |    |  \|  |/  ___/\   __\  
 |    `   \  ___/|  |_|    |     / __ \|  |_> > __ \_  |    `   \  |\___ \  |  |    
/_______  /\___  >____/____|    (____  /   __(____  / /_______  /__/____  > |__| /\ 
        \/     \/                    \/|__|       \/          \/        \/       \/ 
"@,@"
 ____  ____  __    ____   __   ____   __     ____  __  ____  ____     
(    \(  __)(  )  (  _ \ / _\ (  _ \ / _\   (    \(  )/ ___)(_  _)    
 ) D ( ) _) / (_/\ ) __//    \ ) __//    \   ) D ( )( \___ \  )(  _   
(____/(____)\____/(__)  \_/\_/(__)  \_/\_/  (____/(__)(____/ (__)(_)  
"@,@"
    ___     _   ___                       ___ _     _   
   /   \___| | / _ \__ _ _ __   __ _     /   (_)___| |_ 
  / /\ / _ \ |/ /_)/ _` | '_ \ / _` |   / /\ / / __| __|
 / /_//  __/ / ___/ (_| | |_) | (_| |  / /_//| \__ \ |_ 
/___,' \___|_\/    \__,_| .__/ \__,_| /___,' |_|___/\__|
                        |_|                             
"@
$curDir = Get-Location
Write-Host "Current Working Directory: $curDir"
FUNCTION findFile{
    # Full path of the file
    $file = 'C:\DelPapaGamingSoftware\files\Log\OFFLINE.txt','C:\DelPapaGamingSoftware\files\List\computers.txt','C:\DelPapaGamingSoftware\files\List\names.txt','C:\DelPapaGamingSoftware\files\List\names1.txt'

    for($i=0; $i -lt $file.length;$i++){

    if (!(Test-Path $file[$i]))
    {
        New-Item -path $file[$i]-type "file" -value "my new text"
        Write-Host "Created new file and text content added"
    }
    else
    {
        Add-Content -path $file[$i] -value 0
        Write-Host "File already exists: ",$file[$i]
        $content = Get-Content $file[$i]
        $content[0..($content.length-2)]|Out-File $file[$i] -Force
    }
    }
}
FUNCTION createLogFile{
    param($CompName,$incr)
    $file = 'C:\DelPapaGamingSoftware\files\Log\computerInfo'+[string]$CompName
    $temp=$file+$incr+'.txt'
    "tempfile:", $temp
    
    if (!(Test-Path $temp))
    {
        New-Item -path $temp -type "file" -value 0
        Write-Host "Created new file and text content added"
        $content = Get-Content $temp
        $content[0..($content.length-2)]|Out-File $temp -Force
        return $temp
    }
    elseif($incr -lt 10)
    {
        $incr++
        createLogFile($CompName,$incr)
    }
    else{
        Write-Host "maximum ammount of log files reached. please clear log files for this computer then try again."
    }
}
findFile
$indx= Get-Random -Minimum 0 -Maximum 3
Write-Host $text[$indx]
[string[]]$array = Get-Content -Path 'C:\DelPapaGamingSoftware\files\List\names1.txt'; Write-Host "Computers in List: ", $array.length
[string[]]$arrayOffline = Get-Content -Path 'C:\DelPapaGamingSoftware\files\Log\OFFLINE.txt'; Write-Host "Number of Logged offline: ", $arrayOffline.length
[string[]]$arrayDev = Get-Content -Path 'C:\DelPapaGamingSoftware\files\List\computers.txt'; Write-Host "Number of company devices TD: ", $arrayDev.length

FUNCTION NotFound{
    param($CompName)
    Invoke-Command -ComputerName $compName -ScriptBlock {Get-PSDrive | Where {$_.Free -gt 0}}
    }

FUNCTION ResetAdPass{
    $userName = Read-Host -Prompt 'Enter username to change password for'
    $newPassword = Read-Host -Prompt 'Enter new password' -AsSecureString

    $userDN = (Get-ADUser -Identity $userName).DistinguishedName

    do {
        $passwordKnown = Read-Host -Prompt 'Do you know the existing account password? (Y/N)'
    }  until ($passwordKnown -match '[yY|nN]')

    if ($passwordKnown -eq 'Y') {
        $oldPassword = Read-Host -Prompt 'Enter old password:' -AsSecureString
        Set-ADAccountPassword -Identity $userDN -OldPassword $oldPassword -NewPassword $newPassword
    }
    else {
        $adminCred = Get-Credential -Message  'Without the existing password, admin access is required. Please enter admin account credentials.'
        Set-ADAccountPassword -Identity $userDN -NewPassword $newPassword -Reset -Credential $adminCred
    }
}
FUNCTION Full_List {
param($array)
"EMPTY `rEMPTY" | Out-File -FilePath C:\DelPapaGamingSoftware\Log\OFFLINE.txt
[int]$start= Read-Host "Enter start point of look up.(2-" $array.length  ")"
for($i=$start; $i -lt $array.length; $i++){
        $RUser = $array[$i]
        $i
        $RUser
        if (Test-Connection -BufferSize 32 -Count 1 -ComputerName $RUser -Quiet) {
            Write-Host "The remote machine is Online"
            Invoke-Command -ComputerName $RUser -ScriptBlock {Get-WmiObject -ClassName win32_product | Where{$_.name -like 'Microsoft Office*'}| FT Name,Version}
        } 
        else {
            "The remote machine is Down"
             $RUser | Out-File -FilePath C:\DelPapaGamingSoftware\Log\OFFLINE.txt -Append
        }
        
      
 
    }
}
FUNCTION compInfo {
param($array)
[int]$start= Read-Host "Enter start point of look up.(2-" $array.length  ")"
for($i=$start; $i -lt $array.length; $i++){
        $CompName = $array[$i]
        $i
        $CompName
        if (Test-Connection -BufferSize 32 -Count 1 -ComputerName $CompName -Quiet) {
            Write-Host "The remote machine is Online"
            $computerSystem = Get-CimInstance -ClassName Win32_ComputerSystem -Computer $CompName
    		$computerBIOS = Get-CimInstance -ClassName Win32_BIOS -Computer $CompName
    		$computerOS = Get-CimInstance -ClassName Win32_OperatingSystem -Computer $CompName
    		$computerCPU = Get-CimInstance -ClassName Win32_Processor -Computer $CompName
    		$computerHDD = Get-CimInstance -ClassName Win32_LogicalDisk -ComputerName $CompName -Filter drivetype=3
        	write-host "System Information for: " $computerSystem.Name -BackgroundColor DarkCyan
        	"-------------------------------------------------------"
        	"Manufacturer: " + $computerSystem.Manufacturer
        	"Model: " + $computerSystem.Model
        	"Serial Number: " + $computerBIOS.SerialNumber
        	"CPU: " + $computerCPU.Name
                Try{
        		    "HDD Capacity: "  + "{0:N2}" -f ($computerHDD.Size/1GB) + "GB" 
                }
                Catch {
                    Write-Output "oops"
                    NotFound($CompName)
                }
        	"HDD Space: " + "{0:P2}" -f ($computerHDD.FreeSpace/$computerHDD.Size) + " Free (" + "{0:N2}" -f ($computerHDD.FreeSpace/1GB) + "GB)"
        	"RAM: " + "{0:N2}" -f ($computerSystem.TotalPhysicalMemory/1GB) + "GB"
        	"Operating System: " + $computerOS.caption + ", Service Pack: " + $computerOS.ServicePackMajorVersion
        	"User logged In: " + $computerSystem.UserName
        	"Last Reboot: " + $computerOS.ConvertToDateTime($computerOS.LastBootUpTime)
        	""
        	"-------------------------------------------------------"
        } 
        else {
            "The remote machine is Down"
        }
        
      
 
    }
}
FUNCTION compInfoToFile {
param($array)
$incr=0
[int]$start= Read-Host "Enter start point of look up.(2-" $array.length  ")"
for($i=$start; $i -lt $array.length; $i++){
        $CompName = $array[$i]
        $i
        $CompName
        if (Test-Connection -BufferSize 32 -Count 1 -ComputerName $CompName -Quiet) {
            Write-Host "The remote machine is Online"
            $filePath=createLogFile($CompName,$incr)
            $computerSystem = Get-CimInstance -ClassName Win32_ComputerSystem -Computer $CompName
    		$computerBIOS = Get-CimInstance -ClassName Win32_BIOS -Computer $CompName
    		$computerOS = Get-CimInstance -ClassName Win32_OperatingSystem -Computer $CompName
    		$computerCPU = Get-CimInstance -ClassName Win32_Processor -Computer $CompName
    		$computerHDD = Get-CimInstance -ClassName Win32_LogicalDisk -ComputerName $CompName -Filter drivetype=3
        	write-host "System Information for: " $computerSystem.Name -BackgroundColor DarkCyan
        	"-------------------------------------------------------" | Out-File -FilePath $filePath -Append
        	"Manufacturer: " + $computerSystem.Manufacturer | Out-File -FilePath $filePath -Append
        	"Model: " + $computerSystem.Model | Out-File -FilePath $filePath -Append
        	"Serial Number: " + $computerBIOS.SerialNumber | Out-File -FilePath $filePath -Append
        	"CPU: " + $computerCPU.Name | Out-File -FilePath $filePath -Append
                Try{
        		    "HDD Capacity: "  + "{0:N2}" -f ($computerHDD.Size/1GB) + "GB"  | Out-File -FilePath $filePath -Append
                }
                Catch {
                    Write-Output "HDD Capacity not found trying second method." | Out-File -FilePath $filePath -Append
                    NotFound($CompName)
                }
        	"HDD Space: " + "{0:P2}" -f ($computerHDD.FreeSpace/$computerHDD.Size) + " Free (" + "{0:N2}" -f ($computerHDD.FreeSpace/1GB) + "GB)" | Out-File -FilePath $filePath -Append
        	"RAM: " + "{0:N2}" -f ($computerSystem.TotalPhysicalMemory/1GB) + "GB" | Out-File -FilePath $filePath -Append
        	"Operating System: " + $computerOS.caption + ", Service Pack: " + $computerOS.ServicePackMajorVersion | Out-File -FilePath $filePath -Append
        	"User logged In: " + $computerSystem.UserName | Out-File -FilePath $filePath -Append
        	"Last Reboot: " + $computerOS.ConvertToDateTime($computerOS.LastBootUpTime) | Out-File -FilePath $filePath -Append
        	""
        	"-------------------------------------------------------" | Out-File -FilePath $filePath -Append
        } 
        else {
            $filePath=createLogFile($CompName,$incr)
            "The remote machine is Down" | Out-File -FilePath $filePath -Append
             
        }
        
      
 
    }
}
FUNCTION Disconnect(){
    param($connected,$compName)
    $connected=$false;
    "connection to "+$compName+"= "+$connected;
    $compName="";
    Run;               
}
FUNCTION Run(){

while($App=$true){
    $enter1=Read-Host "Enter Command"


    switch($enter1){
        1 {
            do {
            $OfflineOrOnline = Read-Host -Prompt 'Do you want to only search offline computers? (Y/N)'
            }  until ($OfflineOrOnline -match '[yY|nN]')
                    if($OfflineOrOnline -eq 'y'){
                    Full_List($arrayOffline)
                    }
                    else{
                    Full_List($array)
                    }
                }
                2 { 
                    $enter=Read-Host "Enter Computer Name to look up."
                    if (Test-Connection -BufferSize 32 -Count 1 -ComputerName $enter -Quiet) {
                        "The remote machine is Online"
                    Invoke-Command -ComputerName $enter -ScriptBlock {Get-WmiObject -ClassName win32_product | Where{$_.name -like 'Microsoft Office*'}| FT Name,Version}
                    } 
                    else {
                    "The remote machine is Down"
                    $enter | Out-File -FilePath C:\DelPapaGamingSoftware\files\Log\OFFLINE.txt -Append
                    }
                }
        
		"CompInfoAll"{
		    do {
            $OfflineOrOnline = Read-Host -Prompt 'Do you want to search all computers? (Y/N)'
            }  until ($OfflineOrOnline -match '[yY|nN]')
                    if($OfflineOrOnline -eq 'y'){
                    compInfo($arrayDev)
                    }
               
		}
        "CompInfoAllToFile"{
		    do {
            $OfflineOrOnline = Read-Host -Prompt 'Do you want to search all computers? (Y/N)'
            }  until ($OfflineOrOnline -match '[yY|nN]')
                    if($OfflineOrOnline -eq 'y'){
                    compInfoToFile($arrayDev)
                    }
               
		}
        "select computer" {
            $connected=$false;
            $compName=Read-Host "Enter Computer Name to look up."
            if (Test-Connection -BufferSize 32 -Count 1 -ComputerName $compName -Quiet) {
                        "The remote machine is Online"
                    $connected=$true;
                    } 
                    else {
                    "The remote machine is Down"
                    Disconnect($connected,$compName);
                    }
            while($connected=$true){
            $enter=Read-Host "Enter remote Command"
            switch($enter){
                

		"CompInfo"{
		#$CompName=Read-Host "Enter Computer Name to look up."
		$computerSystem = get-wmiobject Win32_ComputerSystem -Computer $CompName
    		$computerBIOS = get-wmiobject Win32_BIOS -Computer $CompName
    		$computerOS = get-wmiobject Win32_OperatingSystem -Computer $CompName
    		$computerCPU = get-wmiobject Win32_Processor -Computer $CompName
    		$computerHDD = Get-WmiObject Win32_LogicalDisk -ComputerName $CompName -Filter drivetype=3 -ErrorVariable NotFound
        		write-host "System Information for: " $computerSystem.Name -BackgroundColor DarkCyan
        		"-------------------------------------------------------"
        		"Manufacturer: " + $computerSystem.Manufacturer
        		"Model: " + $computerSystem.Model
        		"Serial Number: " + $computerBIOS.SerialNumber
        		"CPU: " + $computerCPU.Name
        		"HDD Capacity: "  + "{0:N2}" -f ($computerHDD.Size/1GB) + "GB"
        		"HDD Space: " + "{0:P2}" -f ($computerHDD.FreeSpace/$computerHDD.Size) + " Free (" + "{0:N2}" -f ($computerHDD.FreeSpace/1GB) + "GB)"
        		"RAM: " + "{0:N2}" -f ($computerSystem.TotalPhysicalMemory/1GB) + "GB"
        		"Operating System: " + $computerOS.caption + ", Service Pack: " + $computerOS.ServicePackMajorVersion
        		"User logged In: " + $computerSystem.UserName
        		"Last Reboot: " + $computerOS.ConvertToDateTime($computerOS.LastBootUpTime)
        		""
        		"-------------------------------------------------------"
		}
                "PSSession" {
                    do {
                        $enter = Read-Host -Prompt 'this will exit the script and create a remote powershell connection with the selected computer. Are you sure you wish to continue? (Y/N)'
                    }  until ($enter -match '[yY|nN]')
                    if($enter -eq 'y'){
                    
                        Enter-PSSession -ComputerName $compName
                        "press ctrl+c"
                    }
                    else{
                        Run;
                    }
                    
                    
                }
                "free space" {
                    
                    Invoke-Command -ComputerName $compName -ScriptBlock {Get-PSDrive | Where {$_.Free -gt 0}}
		            Invoke-Command -ComputerName $CompName {Get-WmiObject win32_computersystem} | select @{name="RAM (GB)";e={[math]::Round($_.totalphysicalmemory/1GB,0)}}
		            $WmiObject = Get-WmiObject -computername $CompName Win32_ComputerSystem | select Manufacturer, Model, Name 
                }
                "disconnect"{
                   Disconnect($connected,$compName);
                }

                "remote session" {
                    $enter=Read-Host "Enter Computer Name to look up."
            
                    $enter1=Read-Host "Enter powershell Command"
                    $b = {$enter1}
                    & $b
                    Invoke-Command -ComputerName $enter -ScriptBlock $b
                }

                "reset Password"{
                    ResetAdPass;
                }

                "h" { "","PSSession= remote powershell session;","free space= check free space in C drive;","remote session= not working just use PSSession","Disconnect= disconnects from remote pc","";
                }
                
       
            }
            }
        }
        "reset Password"{ResetAdPass;}

        "clearLog"{"EMPTY `rEMPTY" | Out-File -FilePath C:\DelPapaGamingSoftware\files\Log\OFFLINE.txt}

        "h" { "","1= check office install version of full list; ","2= check office install version single computer; ","clearLog= clear log of offline computers;","select computer= connect to computer and perform commands;","reset password= reset AD user password","";
                }
    }
}
}
Run