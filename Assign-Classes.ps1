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

$Assignments = Import-Csv -Path $csvPath

# Group on the classes
$Classes = $Assignments | Group-Object Class

$Wp = 0
foreach($class in $Classes){
    Write-Progress -Activity "Assigning user to class $($class.Name)" -Status "$wp of $($Classes.count)" -PercentComplete $(($wp/$($Classes.count))*100) -id 1
    # confirm team is found for the class
    $team = Get-Team -DisplayName $Class.Name
    if(-not $team){
        Write-Error "$($Class.Name) - Not Found"
        continue
    }

    $st = 0
    # Add each student to the class
    foreach($Student in $class.Group.Student){
        Write-Progress -Activity "Adding students" -Status "$st of $($class.count)" -PercentComplete $(($st/$($class.count))*100) -id 2
        Add-TeamUser -GroupId $team.GroupId -User $Student -Role Member
        $st++
    }
    Write-Progress -Activity "Done" -Id 2 -Completed
    $wp++
}
Write-Progress -Activity "Done" -Id 1 -Completed
