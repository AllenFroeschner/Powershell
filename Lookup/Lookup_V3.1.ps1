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
$indx= Get-Random -Minimum 0 -Maximum 3
Write-Host $text[$indx]
[string[]]$array = Get-Content -Path 'C:\Lookup\List\names1.txt'; Write-Host "Computers in List: ", $array.length
[string[]]$arrayOffline = Get-Content -Path 'C:\Lookup\Log\OFFLINE.txt'; Write-Host "Number of Logged offline: ", $array.length


FUNCTION Full_List {
param($array)
"EMPTY `rEMPTY" | Out-File -FilePath C:\Lookup\Log\OFFLINE.txt
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
             $RUser | Out-File -FilePath C:\Lookup\Log\OFFLINE.txt -Append
        }
        
      
 
    }
}
while($App=$true){
    $enter=Read-Host "Enter Command"

    switch($enter){
        1 {
            [Char]$OfflineOrOnline= Read-Host "Check only offline computers?(y or n)"
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
              $enter | Out-File -FilePath C:\Lookup\Log\OFFLINE.txt -Append
            }
        }
        3 {
            $enter=Read-Host "Enter Computer Name to look up."
            Enter-PSSession -ComputerName $enter
        }
        4 {
            $enter=Read-Host "Enter Computer Name to look up."
            Invoke-Command -ComputerName $enter -ScriptBlock {Get-PSDrive C} 
        }
        "remote session" {
            $enter=Read-Host "Enter Computer Name to look up."
            
            $enter1=Read-Host "Enter powershell Command"
            $b = {$enter1}
            & $b
            Invoke-Command -ComputerName $enter -ScriptBlock $b
        }
        "h" { "1=try full list; 2= try single computer; clearLog= clear log of offline computers"
        }
        "clearLog"{"EMPTY `rEMPTY" | Out-File -FilePath C:\Lookup\Log\OFFLINE.txt}
       
    }
}