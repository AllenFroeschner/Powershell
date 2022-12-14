$App=$true
[System.Collections.ArrayList]$array = Get-Content -Path 'C:\Lookup\List\names.txt'; $array.Size()
[System.Collections.ArrayList]$arrayOffline = Get-Content -Path 'C:\Lookup\Log\OFFLINE.txt'; $array.length

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
FUNCTION Full_ListOffline {
param($array)
#[int]$start= Read-Host "Enter start point of look up.(2-" $array.length  ")"
for($i=0; $i -lt $array.length; $i++){
        $RUser = $array[$i]
        $i
        $RUser
        if (Test-Connection -BufferSize 32 -Count 1 -ComputerName $RUser -Quiet) {
            Write-Host "The remote machine is Online"
            $newLog= "C:\Lookup\Log\OnlineUsers" + $RUser + ".txt"
            New-Item $newLog
            Invoke-Command -ComputerName $RUser -ScriptBlock {Get-WmiObject -ClassName win32_product | Where{$_.name -like 'Microsoft Office*'}| FT Name,Version}
            Invoke-Command -ComputerName $RUser -ScriptBlock {Get-WmiObject -ClassName win32_product | Where{$_.name -like 'Microsoft Office*'}| FT Name,Version} | Out-File -FilePath $newLog 
       
            for($j=0; $j -lt $array.length; $j++){
            
            }

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
        "h" { "1=try full list; 2= try single computer; clearLog= clear log of offline computers"
        }
        "clearLog"{"EMPTY `rEMPTY" | Out-File -FilePath C:\Lookup\Log\OFFLINE.txt}
        "loopOffline"{
        Full_List($arrayOffline)
        }
    }
}