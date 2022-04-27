Write-Host "            
--------------------------------------------------------------------------------           

/(/(/(/      //(/(/(/    (/(/(/      //(/(/(/(/(/(/(            /(/(/(//        
///////     ///////      //////    /////////  ////////         /////////.       
/(/(/(/   /(/(/(/        (/(/(/    (/(/(/       *(/(/(/       /(/(/(/(/(/       
/////// ///////          //////    ////////                  ////// //////      
/(/(/(/(/(/(/(           (/(/(/      //(/(/(/(/(//          ((/(/(   //(/(/     
///////////////          //////          ////////////       /////(    //////    
/(/(/(/(/ /(/(/(/        (/(/(/                (/(/(/(,    (/(/(/(/(/(/(/(/(/   
///////    ///////       //////   //////(        //////   ////////////////////  
/(/(/(/     /(/(/(//     (/(/(/    (/(/(/(//  /(/(/(/(.  (/(/(/         /(/(/(/ 
///////       ///////    //////      ////////////////   ///////          ///////

---------------------------------------------------------------------------------
"
start-sleep 2

$Kunde = Read-Host "Fuer welchen Kunden wird das ODT bereit gestellt?"
$MAK = Read-Host "Welcher MAK-Schluessel soll verwendet werden?"
$CounterMAK = Read-Host "Wie viele Lizenzen wurden vom Kunden bestellt?"

set-Location $PSScriptroot
#TXT erstellung von Lizenzliste
if (!(Test-Path .\Lizenzliste.txt)) {
    New-Item -path .\Lizenzliste.txt -itemType File
    Add-Content -Path .\Lizenzliste.txt -Value "Kunde; Lizenzanzahl"
}
$output = "$Kunde -> $CounterMAK"
Add-Content -Path .\Lizenzliste.txt -Value $output

#Kunden ODT Verzeichnis im selben Dir anlegen
$KundePath = ".\" + $Kunde
New-Item -path $KundePath -itemType Directory
Copy-Item -Path ".\ODT" -Recurse -Destination $KundePath -Container

start-sleep 2
#Erstellen des MAK Counters
$MAKCounterPath = $KundePath + "\ODT\Exec\CounterMAK.txt"

New-Item -path $MAKCounterPath -itemType File
Add-Content -Path $MAKCounterPath -Value $CounterMAK

#Erstellen eines Encryption Keys
if (!(Test-Path $PSScriptroot\Setup-Data\AES.key)) {
    ".\Setup-Data\KeyGen.ps1"
}


$ExecPath = $KundePath + "\ODT\Exec"
Copy-Item -Path .\Setup-Data\AES.key -Destination $ExecPath
$InstallPath = $KundePath + "\ODT"
Copy-Item -Path .\Setup-Data\Install.ps1 -Destination $InstallPath

.\Setup-Data\MAKEncryptGen.ps1 $MAK $ExecPath


Write-Host "

ODT Paket fuer Kunden wurde erfolgreich erstellt!

_                
| |               
| |__  _   _  ___ 
| '_ \| | | |/ _ \
| |_) | |_| |  __/
|_.__/ \__, |\___|
        __/ |     
       |___/      
"

start-sleep 2



