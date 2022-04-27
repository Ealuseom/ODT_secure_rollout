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

Write-Host "Beginne Installation von Office"

start-sleep 2

Set-Location $PSScriptRoot

$EncryptionKeyData = Get-Content ".\AES.key"
$PasswordSecureString = Get-Content ".\MAK.txt" | ConvertTo-SecureString -Key $EncryptionKeyData
$PlainTextPassword = [System.Runtime.InteropServices.Marshal]::PtrToStringBSTR([System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($PasswordSecureString))

$xmlObjectsettings = New-Object System.Xml.XmlWriterSettings
$xmlObjectsettings.Indent = $true
$xmlObjectsettings.IndentChars = "  "
$ParentDir = (get-item $PSScriptRoot).parent
$XmlFilePath = $ParentDir + "\Office2021Std.xml"
$XmlObjectWriter = [System.XML.XmlWriter]::Create($XmlFilePath, $xmlObjectsettings)
$XmlObjectWriter.WriteStartDocument()
$XmlObjectWriter.WriteStartElement('Configuration')
$XmlObjectWriter.WriteAttributeString('ID', '5105a8f7-d7b7-418b-940f-0f7555d2a5f1')
    $XmlObjectWriter.WriteStartElement('Add') 
    $XmlObjectWriter.WriteAttributeString('OfficeClientEdition', '64')
    $XmlObjectWriter.WriteAttributeString('Channel', 'PerpetualVL2021')
        $XmlObjectWriter.WriteStartElement('Product')
        $XmlObjectWriter.WriteAttributeString('ID', 'Standard2021Volume')
        $XmlObjectWriter.WriteAttributeString('PIDKEY', $PlainTextPassword)
            $XmlObjectWriter.WriteStartElement('Language')
            $XmlObjectWriter.WriteAttributeString('ID', 'de-de')
            $XmlObjectWriter.WriteEndElement()
            $XmlObjectWriter.WriteStartElement('ExcludeApp')
            $XmlObjectWriter.WriteAttributeString('ID', 'OneDrive')
            $XmlObjectWriter.WriteEndElement()
            $XmlObjectWriter.WriteStartElement('ExcludeApp')
            $XmlObjectWriter.WriteAttributeString('ID', 'Teams')
            $XmlObjectWriter.WriteEndElement()
       $XmlObjectWriter.WriteEndElement()
    $XmlObjectWriter.WriteEndElement()
        $XmlObjectWriter.WriteStartElement('Property')
        $XmlObjectWriter.WriteAttributeString('Name', 'SharedComputerLicensing')
        $XmlObjectWriter.WriteAttributeString('Value', '0')
        $XmlObjectWriter.WriteEndElement()
        $XmlObjectWriter.WriteStartElement('Property')
        $XmlObjectWriter.WriteAttributeString('Name', 'FORCEAPPSHUTDOWN')
        $XmlObjectWriter.WriteAttributeString('Value', 'FALSE')
        $XmlObjectWriter.WriteEndElement()
        $XmlObjectWriter.WriteStartElement('Property')
        $XmlObjectWriter.WriteAttributeString('Name', 'DeviceBasedLicensing')
        $XmlObjectWriter.WriteAttributeString('Value', '0')
        $XmlObjectWriter.WriteEndElement()
        $XmlObjectWriter.WriteStartElement('Property')
        $XmlObjectWriter.WriteAttributeString('Name', 'SCLCacheOverride')
        $XmlObjectWriter.WriteAttributeString('Value', '0')
        $XmlObjectWriter.WriteEndElement()
        $XmlObjectWriter.WriteStartElement('Property')
        $XmlObjectWriter.WriteAttributeString('Name', 'AUTOACTIVATE')
        $XmlObjectWriter.WriteAttributeString('Value', '1')
        $XmlObjectWriter.WriteEndElement()
        $XmlObjectWriter.WriteStartElement('Updates')
        $XmlObjectWriter.WriteAttributeString('Enabled', 'TRUE')
        $XmlObjectWriter.WriteEndElement()
        $XmlObjectWriter.WriteStartElement('RemoveMSI')
        $XmlObjectWriter.WriteEndElement()
        $XmlObjectWriter.WriteStartElement('AppSettings')
            $XmlObjectWriter.WriteStartElement('User')
            $XmlObjectWriter.WriteAttributeString('Key', 'software\microsoft\office\16.0\excel\options')
            $XmlObjectWriter.WriteAttributeString('Name', 'defaultformat')
            $XmlObjectWriter.WriteAttributeString('Value', '51')
            $XmlObjectWriter.WriteAttributeString('Type', 'REG_DWORD')
            $XmlObjectWriter.WriteAttributeString('App', 'excel16')
            $XmlObjectWriter.WriteAttributeString('Id', 'L_SaveExcelfilesas')
            $XmlObjectWriter.WriteEndElement()
            $XmlObjectWriter.WriteStartElement('User')
            $XmlObjectWriter.WriteAttributeString('Key', 'software\microsoft\office\16.0\powerpoint\options')
            $XmlObjectWriter.WriteAttributeString('Name', 'defaultformat')
            $XmlObjectWriter.WriteAttributeString('Value', '27')
            $XmlObjectWriter.WriteAttributeString('Type', 'REG_DWORD')
            $XmlObjectWriter.WriteAttributeString('App', 'ppt16')
            $XmlObjectWriter.WriteAttributeString('Id', 'L_SavePowerPointfilesas')
            $XmlObjectWriter.WriteEndElement()
            $XmlObjectWriter.WriteStartElement('User')
            $XmlObjectWriter.WriteAttributeString('Key', 'software\microsoft\office\16.0\word\options')
            $XmlObjectWriter.WriteAttributeString('Name', 'defaultformat')
            $XmlObjectWriter.WriteAttributeString('Value', '')
            $XmlObjectWriter.WriteAttributeString('Type', 'REG_SZ')
            $XmlObjectWriter.WriteAttributeString('App', 'word16')
            $XmlObjectWriter.WriteAttributeString('Id', 'L_SaveWordfilesas')
            $XmlObjectWriter.WriteEndElement()
        $XmlObjectWriter.WriteEndElement()
        $XmlObjectWriter.WriteStartElement('Display')
        $XmlObjectWriter.WriteAttributeString('Level', 'None')
        $XmlObjectWriter.WriteAttributeString('AcceptEULA', 'TRUE')
$XmlObjectWriter.WriteEndDocument()
$XmlObjectWriter.Flush()
$XmlObjectWriter.Close()

$Counter = get-content ".\CounterMAK.txt"
while($Counter -ge 1)
{

    cmd.exe /c "cd $ParentDir | setup.exe /configure Office2021Std.xml"

    $Counter--
    Set-Content ".\CounterMAK.txt" $Counter

    Remove-Item -Path $XmlFilePath -Force

    Remove-Variable * -ErrorAction SilentlyContinue; Remove-Module *; $error.Clear(); Clear-Host

}
else
{

    Remove-Item -Path ".\CounterMAK.txt" -Force
    Remove-Item -Path ".\MAK.txt" -Force
    Remove-Item -Path ".\AES.key" -Force

    Remove-Variable * -ErrorAction SilentlyContinue; Remove-Module *; $error.Clear(); Clear-Host

    Write-Host "Ihre Lizenzen sind aufgebraucht, bitte wenden Sie sich an schulen@kisa.it oder Telefonisch unter 0351 86682670 "
    Start-Sleep 10

}