param($moveableFile)
$MyDir = [System.IO.Path]::GetDirectoryName($myInvocation.MyCommand.Definition)
#Ha a szerveren más mappába szeretnénk helyezni a fájlokat, akkor az URL változót kell megváltoztatni
$URL = "https://ipsolzrt.sharepoint.com/sites/Projekt"
$passFile = $MyDir + "\creds\tavmeres-cred.txt"
$Pass = Get-Content $passFile | ConvertTo-SecureString
Add-PnPStoredCredential -Name $URL -Username tavmeres@ipsol.hu -Password $Pass
import-Module SharePointPnPPowerShellOnline
Connect-PnPOnline $URL


#$moveableFile = "C:\KS_fajlok\node-test\output\Bizonylatok tételes lekérdezése.xlsx"
#$moveableFile = ".\output\output-2020-04-23.json"
#$moveableFile = "C:\Auditlogs\nfx_activity_tracker\output\tevekenyseg-naplo-2020-04-16.pdf"
Write-Output $moveableFile

$fullPath = "Shared Documents/General/KS_projekt_statusz_vizsgalo/input/" + $finalFolder

Add-PnPFile -Folder $fullPath -Path $moveableFile
