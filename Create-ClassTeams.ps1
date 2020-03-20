[CmdletBinding()]
[OutputType([string])]
Param(
    [parameter(Mandatory=$true)]
    [String]$csvPath
)

if(-not (Get-Module MicrosoftTeams)){
    if(Get-Module MicrosoftTeams -ListAvailable){
        Import-Module MicrosoftTeams
    } else {
        throw "The MicrosoftTeams module is required for this script. You can install the module by running the command 'Install-Module MicrosoftTeams -Scope CurrentUser'"
    }
}

Connect-MicrosoftTeams | Out-Null

$CreateTeams = Import-Csv -Path $csvPath
$Wp = 0
foreach($newTeam in $CreateTeams){
    Write-Progress -Activity "Creating class $($newTeam.ClassName)" -Status "$wp of $($CreateTeams.count)" -PercentComplete $(($wp/$($CreateTeams.count))*100) -id 1
    $team = Get-Team -DisplayName $newTeam.ClassName
    if(-not $team){
        $MailNickname = [regex]::Replace($newTeam.ClassName,"[^0-9a-zA-Z]","")
        $team = New-Team -MailNickname $MailNickname -displayname $newTeam.ClassName -Visibility Private
    }

    Add-TeamUser -GroupId $team.GroupId -User $newTeam.Teacher -Role Owner
    $team
    $wp++
}




