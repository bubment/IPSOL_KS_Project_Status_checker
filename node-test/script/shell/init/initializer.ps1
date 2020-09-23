#Ha első futtatáskor az a hiba lép fel, hogy The file <fájl neve> is not digitally signed, akkor az alábbi kódot kell egy üres scriptben lefuttatni
#-------------------------------------------
#Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
#-------------------------------------------
#Windows 7-en történő inicializáláshoz kellenek ezeka kódok
#-------------------------------------------
#[System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]'Tls11,Tls12'
#Install-PackageProvider -Name "Nuget" -RequiredVersion "2.8.5.201" -Force
#-------------------------------------------
#HA a NuGet provider hiányára panaszkodik a program akkor is az egyel fentebbi kódot kell futtatni
Set-ExecutionPolicy RemoteSigned
Install-Module MSOnline
Install-Module AzureAD
Install-Module PowerShellGet -Force
Install-Module -Name ExchangeOnlineManagement
Install-Module -Name Microsoft.Online.SharePoint.PowerShell
Install-Module SharePointPnPPowerShellOnline
Install-Module -Name MicrosoftTeams

<#Test#>
<#Connect-MsolService
Get-MsolUser#>

