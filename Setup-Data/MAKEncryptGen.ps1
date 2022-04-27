param ($MAK, $ExecPath)

$KeyFile = ".\Setup-Data\AES.key"
$Key = Get-Content $KeyFile
$Password = $MAK | ConvertTo-SecureString -AsPlainText -Force
$Outpath = $ExecPath + "\MAK.txt"
$Password | ConvertFrom-SecureString -key $Key | Out-File $Outpath