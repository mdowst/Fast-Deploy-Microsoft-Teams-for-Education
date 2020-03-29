param(
$DllBasePath = $env:dllBasePath)

If ($DllBasePath -eq $Null -or $DllBasePath -eq '' )
{
    $DllBasePath = $PSScriptRoot
}
$DllBasePath += "\"

Import-Module "$($DllBasePath)Newtonsoft.Json.dll"
Import-Module "$($DllBasePath)Microsoft.IdentityModel.Clients.ActiveDirectory.dll"
Import-Module "$($DllBasePath)Microsoft.IdentityModel.Clients.ActiveDirectory.Platform.dll"
Import-Module "$($DllBasePath)Microsoft.Open.Teams.CommonLibrary.dll"
Import-Module "$($DllBasePath)Microsoft.TeamsCmdlets.PowerShell.Custom.dll"