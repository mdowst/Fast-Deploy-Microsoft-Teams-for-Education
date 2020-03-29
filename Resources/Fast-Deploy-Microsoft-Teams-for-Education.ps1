param(
    [parameter(Mandatory=$true)]
    $WorkbookPath
)

$yes = new-Object System.Management.Automation.Host.ChoiceDescription "&Yes","help"
$no = new-Object System.Management.Automation.Host.ChoiceDescription "&No","help"
$choices = [System.Management.Automation.Host.ChoiceDescription[]]($yes,$no)
$message = "This script is designed to be run after you have entered all your information in the " + 
"'Fast Deploy Microsoft Teams for Education.xlsx' workbook.`n'Yes' to continue with the upload process, or 'No' to cancel"
$answer = $host.ui.PromptForChoice("Confirmation",$message,$choices,0)

if($answer -ne 0){exit}

Import-Module (Join-Path $PSScriptRoot "ImportExcel\7.1.0\ImportExcel.psd1") -ErrorAction Stop
Import-Module (Join-Path $PSScriptRoot "AzureAD\2.0.2.76\AzureAD.psd1") -ErrorAction Stop
Import-Module (Join-Path $PSScriptRoot "MicrosoftTeams\0.9.6\MicrosoftTeams.psd1") -ErrorAction Stop

$Credential = Get-Credential -Message 'Enter Office 365 username and password'

# Connect to AzureAD
#Connect to Exchange Online
Write-Host "Connect to Exchange Online..." -NoNewline
$NewPSSession = @{
    ConfigurationName = 'Microsoft.Exchange'
    ConnectionUri     = "https://ps.outlook.com/powershell" 
    Credential        = $Credential 
    Authentication    = 'Basic'
    AllowRedirection  = $true
    WarningAction     = 'SilentlyContinue'
}
$PS_Session = New-PSSession @NewPSSession  -ErrorAction stop
$ImportResults = Import-PSSession -Session $PS_Session  -DisableNameChecking -AllowClobber -ErrorAction stop
Write-Host "connected" -ForegroundColor Green

Write-Host "Connect to AzureAD..." -NoNewline
try{ Get-AzureADCurrentSessionInfo -ErrorAction stop | Out-Null }
catch{ Connect-AzureAD -Credential $Credential -ErrorAction stop | Out-Null }
Write-Host "connected" -ForegroundColor Green

# Connect to Teams
Write-Host "Connect to Teams..." -NoNewline
Connect-MicrosoftTeams -Credential $Credential  -ErrorAction stop | Out-Null
Write-Host "connected" -ForegroundColor Green

# set the domain for the username
Write-Host "Determine email domain..." -NoNewline
$FoundDomains = @(Get-AzureADTenantDetail | Select-Object -ExpandProperty VerifiedDomains | Select-Object -Property Name, Capabilities)

# prompt if more than 1 domain is found
If($FoundDomains.Count -gt 1){
    $script:TenantDomain = $FoundDomains | Out-GridView -Title 'Select email domain' -PassThru | Select-Object -ExpandProperty Name
} else {
    $script:TenantDomain = $FoundDomains | Select-Object -ExpandProperty Name
}
Write-Host "$script:TenantDomain" -ForegroundColor Green

# Get the naming pattern to use
$script:Pattern = Import-Excel -Path $WorkbookPath -WorksheetName 'School Details' -NoHeader -StartRow 5 -EndRow 5 -EndColumn 4 -StartColumn 4 | 
    Select-Object  @{l='Pattern';e={$_.P1.Split()[0]}} | Select-Object -ExpandProperty Pattern

# if pattern is not found prompt for which one to use
if([string]::IsNullOrEmpty($script:Pattern)){
    $NamePatterns = ('[{"Pattern":"First.Last","Example":"Diana.Price"},{"Pattern":"FLast","Example":"DPrice"}]' | ConvertFrom-Json)
    $script:Pattern = $NamePatterns | Out-GridView -Title 'Select naming pattern' -PassThru | Select-Object -ExpandProperty Pattern
}

# Get the naming pattern to use
$UsageLocation = Import-Excel -Path $WorkbookPath -WorksheetName 'School Details' -NoHeader -StartRow 7 -EndRow 7 -EndColumn 5 -StartColumn 5 | 
    Select-Object -ExpandProperty P1

# if pattern is not found just set to US
if([string]::IsNullOrEmpty($UsageLocation)){
    $UsageLocation = 'US'
}

Function Import-StudentsAndTeachers {
<#
.SYNOPSIS
Will create user accounts in Office 365 for students and teachers

.PARAMETER UserImport
The object that contains the users to import

.PARAMETER Role
Teacher ot Student

.PARAMETER UsageLocation
A two letter country code (ISO standard 3166). Required for users that will be assigned licenses due 
to legal requirement to check for availability of services in countries. 
Examples include: "US", "JP", and "GB".
For a complete list go to - https://www.iso.org/obp/ui/#search


#>
    [CmdletBinding()]
    [OutputType([string])]
    Param(
        [parameter(Mandatory = $true)]
        [object]$UserImport,

        [parameter(Mandatory = $true)]
        [ValidateSet('Teacher', 'Student')]
        [string]$Role,
    
        [parameter(Mandatory = $false)]
        [String]$UsageLocation = 'US'
    )

    $Sku = @{
        "M365EDU_A5_FACULTY" = "Office 365 A5 for faculty"
        "M365EDU_A5_STUDENT" = "Office 365 A5 for students"
        "M365EDU_A3_FACULTY" = "Office 365 A3 for faculty"
        "M365EDU_A3_STUDENT" = "Office 365 A3 for students"
        "M365EDU_A1_FACULTY" = "Office 365 A1 for faculty"
        "M365EDU_A1_STUDENT" = "Office 365 A1 for students"
    }

    # Create the objects we'll need to add licenses
    $AllLicenses = Get-AzureADSubscribedSku | Select-Object @{l = 'License'; e = { $sku.Item($_.SkuPartNumber) } }, SkuPartNumber, 
        @{l = 'Available'; e = { $_.PrepaidUnits.Enabled - $_.ConsumedUnits } }, @{l = 'Purchased'; e = { $_.PrepaidUnits.Enabled } }, SkuID 
    
    # attempt to find the license to use
    if($Role -eq 'Teacher'){
        $LicenseSku = @($AllLicenses | Where-Object{ $_.SkuPartNumber -like '*_FACULTY' })
    } elseif($Role -eq 'Student'){
        $LicenseSku = @($AllLicenses | Where-Object{ $_.SkuPartNumber -like '*_STUDENT' })
    }

    # if license not found or more than one type is found prompt the user
    if(-not $LicenseSku){
        $LicenseSku = $AllLicenses | Out-GridView -Title "Select license type for $Role" -PassThru
    } elseif($LicenseSku.count -gt 1){
        $LicenseSku = $LicenseSku | Out-GridView -Title "Select license type for $Role" -PassThru
    }

    # create the user accounts
    [System.Collections.Generic.List[PSObject]]$CreatedUsers = @()
    foreach ($user in $UserImport) {
        Write-Progress -Activity "Creating users" -Status "$($CreatedUsers.count) of $($UserImport.count)" -PercentComplete $(($($CreatedUsers.count)/$($UserImport.count))*100) -id 1

        # determine user login name
        $userCheck = @{
            FirstName    = $user.'First Name'
            LastName     = $user.'Last Name'
            Pattern      = $Pattern
            TenantDomain = $TenantDomain
            UniqueId     = $user."$($Role) #"
        }
        $UPNCheck = Get-Office356UserName @userCheck

        if ($UPNCheck.Exists -eq $true) {
            Write-Host "`tConfirmed $($Role): $($user.'First Name') $($user.'Last Name')" -ForegroundColor Cyan
            $userObject = Get-AzureADUser -ObjectId $UPNCheck.UPN
            $AssignedSku = @(Get-AzureADUserLicenseDetail -ObjectId $userObject.ObjectId)[0].SkuPartNumber
            if ($sku.Item($AssignedSku)) {
                $AssignedSku = $sku.Item($AssignedSku)
            }
        }
        else {
            Write-Host "`tCreating $($Role): $($user.'First Name') $($user.'Last Name')" -ForegroundColor Green
            $PasswordProfile = New-Object -TypeName Microsoft.Open.AzureAD.Model.PasswordProfile
            $PasswordProfile.Password = 'Welcome' + $($user."$($Role) #" -replace '\s', '') + '!'
            $PasswordProfile.ForceChangePasswordNextLogin = $true
        
            # create the user
            $userParams = @{
                DisplayName              = "$($user.'First Name') $($user.'Last Name')" 
                GivenName                = $user.'First Name'
                SurName                  = $user.'Last Name'
                UserPrincipalName        = $UPNCheck.UPN
                JobTitle                 = $Role
                UsageLocation            = $UsageLocation
                MailNickName             = $user.'First Name'
                PasswordProfile          = $PasswordProfile 
                AccountEnabled           = $true
                FacsimileTelephoneNumber = $user."$($Role) #"
            }
            $userObject = New-AzureADUser @userParams
        
            # Set the license.
            $license = New-Object -TypeName Microsoft.Open.AzureAD.Model.AssignedLicense
            $AssignedLicenses = New-Object -TypeName Microsoft.Open.AzureAD.Model.AssignedLicenses
            $license.SkuId = $LicenseSku.SkuID
            $AssignedLicenses.AddLicenses = $license
            try {
                Set-AzureADUserLicense -ObjectId $userObject.ObjectId -AssignedLicenses $AssignedLicenses -ErrorAction Stop | Out-Null
                $AssignedSku = $LicenseSku.License
            }
            catch {
                $AssignedSku = $null
            }
        }

        $newUser = $userObject | Select-Object @{l = 'FirstName'; e = { $_.GivenName } }, @{l = 'LastName'; e = { $_.Surname } }, @{l = 'Role'; e = { $_.JobTitle } },
            @{l = 'AssignedLicenses'; e = { $AssignedSku } }, @{l = 'Id'; e = { $_.FacsimileTelephoneNumber } }, @{l = 'Lookup'; e = { $user.Lookup } },
            @{l = 'Password'; e = { 'Welcome' + $($user."$($Role) #" -replace '\s', '') + '!' } }, UserPrincipalName
            
        $CreatedUsers.Add($newUser)
        
    }
    Write-Progress -Activity "Done" -Id 1 -Completed

    $CreatedUsers
}

Function Get-Office356UserName {
<#
.SYNOPSIS
Use to determine a user's Office 365 username

.DESCRIPTION
This function will determine the user’s Office 365 username based on the person’s name, and the pattern specified. 

.PARAMETER FirstName
The user’s first name.

.PARAMETER LastName
The user’s last name.

.PARAMETER Pattern
The pattern to use when creating the username. FLast or First.Last

.PARAMETER TenantDomain
The TenantDomain that will be appended to the end of the user's account

.PARAMETER digits
Use to specify if you want a number added to the end, how many digits it should be. 

.PARAMETER number
If you want a number added to the end, what value it should start at

.PARAMETER IncludeNumber
Pass if you want to include a number at the end by default. Otherwise a number will only be
include if there is a duplicate.

#>
    [CmdletBinding()]
    [OutputType([string])]
    Param(
        [parameter(Mandatory = $true)]
        [String]$FirstName,
            
        [parameter(Mandatory = $true)]
        [String]$LastName,
    
        [parameter(Mandatory = $false)]
        [String]$UniqueId,
            
        [parameter(Mandatory = $true)]
        [ValidateSet('First.Last', 'FLast')]
        [String]$Pattern,
    
        [parameter(Mandatory = $true)]
        [String]$TenantDomain,
    
        [parameter(Mandatory = $false)]
        [ValidateRange(1, 20)]
        [int]$digits = 2,
    
        [parameter(Mandatory = $false)]
        [int]$number = 0,
    
        [parameter(Mandatory = $false)]
        [switch]$IncludeNumber
    )
    $username = [string]::Empty
    
    if ($Pattern -eq 'First.Last') {
        $username = $FirstName + '.' + $LastName
    }
    elseif ($Pattern -eq 'FLast') {
        $username = $FirstName.Substring(0,1) + $LastName
    }
        
    
    if ($IncludeNumber) {
        $username += ($number.ToString().PadLeft($digits)).Replace(" ", "0")
    }
    
    $User = Get-AzureADUser -Filter "userPrincipalName eq '$($username)@$($TenantDomain)'"
    if (-not [string]::IsNullOrEmpty($UniqueId) -and $User.FacsimileTelephoneNumber -eq $UniqueId) {
        [pscustomobject]@{Exists = $true; UPN = "$($username)@$($TenantDomain)" }
    }
    elseif ($User) {
        $number++
        if (-not $PSBoundParameters.ContainsKey('number')) {
            $PSBoundParameters.Add('number', $number)
        }
        else {
            $PSBoundParameters['number'] = $number
        }
    
        if (-not $PSBoundParameters.ContainsKey('IncludeNumber')) {
            $PSBoundParameters.Add('IncludeNumber', $true)
        } 
        Get-Office356UserName @PSBoundParameters
    }
    else {
        [pscustomobject]@{Exists = $false; UPN = "$($username)@$($TenantDomain)" }
    }
}

Function Import-ClassTeams {
<#
.SYNOPSIS
Creates a Team for each class and set the Education class template

.PARAMETER Classrooms
The object that contains the classes to create Team sites for

#>
    [CmdletBinding()]
    [OutputType([string])]
    Param(
        [parameter(Mandatory = $true)]
        [object]$Classrooms
    )
    $Wp = 0
    foreach($newTeam in $Classrooms){
        Write-Progress -Activity "Creating class $($newTeam)" -Status "$wp of $($Classrooms.count)" -PercentComplete $(($wp/$($Classrooms.count))*100) -id 1
        
        $team = Get-Team | Where-Object{ $_.DisplayName -eq $newTeam }
        if(-not $team){
            $MailNickname = [regex]::Replace($newTeam,"[^0-9a-zA-Z]","")
            $team = New-Team -Alias $MailNickname -DisplayName $newTeam -Template EDU_Class 
            Write-Host "`tCreated class $($newTeam)" -ForegroundColor Green
        } else {
            Write-Host "`tConfirmed class $($newTeam)" -ForegroundColor Cyan
        }
        $team | Select-Object GroupId, @{l='DisplayName';e={$newTeam}}
        $wp++
    }
}

Write-Host "Gathering data from spreadsheet..." -NoNewline
$Teachers = Import-Excel -Path $WorkbookPath -WorksheetName 'Teachers' -StartRow 3 | Where-Object { -not [string]::IsNullOrEmpty($_.Lookup) }
$Students = Import-Excel -Path $WorkbookPath -WorksheetName 'Students' -StartRow 3 | Where-Object { -not [string]::IsNullOrEmpty($_.Lookup) }
$Classes  = Import-Excel -Path $WorkbookPath -WorksheetName 'Classes'  -StartRow 3 | Where-Object { -not [string]::IsNullOrEmpty($_.Lookup) }
Write-Host "done" -ForegroundColor Green

Write-Host "Importing Teachers...."
$ImportedTeachers = Import-StudentsAndTeachers -UserImport $Teachers -Role Teacher -UsageLocation $UsageLocation
Write-Host "Importing Students...."
$ImportedStudents = Import-StudentsAndTeachers -UserImport $Students -Role Student -UsageLocation $UsageLocation

Write-Host "Creating Team sites for each class...."
$Classrooms = $Classes | Where-Object { -not [string]::IsNullOrEmpty($_.'Class Name') } | Select-Object -ExpandProperty 'Class Name'
$ImportedClasses = Import-ClassTeams -Classrooms $Classrooms

$Wp = 0
foreach($class in $ImportedClasses){
    Write-Progress -Activity "Assigning user to class $($class.DisplayName)" -Status "$wp of $($ImportedClasses.count)" -PercentComplete $(($wp/$($ImportedClasses.count))*100) -id 1
    Write-Host "Assigning teachers to: $($class.DisplayName)"
    
    $ClassTeachers = $Classes | Where-Object{ $_.Lookup -eq $class.DisplayName -and -not [string]::IsNullOrEmpty($_.Teachers) } | Select-Object -ExpandProperty Teachers
    foreach($Teacher in $ClassTeachers){
        $UserPrincipalName = $ImportedTeachers | Where-Object{ $_.Lookup -eq $Teacher } | Select-Object -ExpandProperty UserPrincipalName
        if([string]::IsNullOrEmpty($UserPrincipalName)){
            Write-Host "No user account found for $Teacher" -ForegroundColor Red
            continue
        }
        
        try{
            Add-TeamUser -GroupId $class.GroupId -User $UserPrincipalName -Role Member -ErrorAction Stop
            Write-Host "`tAdded $($Teacher.Split('-')[0].Trim())" -ForegroundColor Green
        } catch {
            if($_.Exception.Message -like '*One or more added object references already exist for the following modified properties:*'){
                Write-Host "`tConfirmed $($Teacher.Split('-')[0].Trim())" -ForegroundColor Cyan
            } else {
                Write-Host "Error adding $($Teacher.Split('-')[0].Trim())" -ForegroundColor Red
            }
        }
    }
    
    Write-Host "Assigning students to: $($class.DisplayName)"
    $ClassStudents = $Classes | Where-Object{ $_.Lookup -eq $class.DisplayName -and -not [string]::IsNullOrEmpty($_.Students) } | Select-Object -ExpandProperty Students
    foreach($student in $ClassStudents){
        $UserPrincipalName = $ImportedStudents | Where-Object{ $_.Lookup -eq $student } | Select-Object -ExpandProperty UserPrincipalName
        if([string]::IsNullOrEmpty($UserPrincipalName)){
            Write-Host "No user account found for $student" -ForegroundColor Red
            continue
        }
        
        try{
            Add-TeamUser -GroupId $class.GroupId -User $UserPrincipalName -Role Member -ErrorAction Stop
            Write-Host "`tAdded $($student.Split('-')[0].Trim())" -ForegroundColor Green
        } catch {
            if($_.Exception.Message -like '*One or more added object references already exist for the following modified properties:*'){
                Write-Host "`tConfirmed $($student.Split('-')[0].Trim())" -ForegroundColor Cyan
            } else {
                Write-Host "Error adding $($student.Split('-')[0].Trim())" -ForegroundColor Red
            }
        }
    }

    
}
Write-Progress -Activity "Done" -Id 1 -Completed

# Enable schedules to show in Teams
$timer =  [system.diagnostics.stopwatch]::StartNew()

Write-Host "Apply calendar fix for classes"
foreach($class in $ImportedClasses){
    $timer.Restart()
    do{
        $GroupCheck = Get-UnifiedGroup -Identity $class.GroupId -ErrorAction SilentlyContinue
        
        if($timer.Elapsed.TotalSeconds -gt 30){break}
        elseif(-not $GroupCheck){
            Start-Sleep -Seconds 5
            Write-Host "The exchange group for $($class.DisplayName) could not be found. Will recheck in 5 seconds." -ForegroundColor Yellow
        }
    }while(-not $GroupCheck)
    
    if($GroupCheck){
        Set-UnifiedGroup -Identity $class.GroupId -AlwaysSubscribeMembersToCalendarEvents:$True -WarningAction SilentlyContinue
        Write-Host "`tApplied calendar fix for $($class.DisplayName)" -ForegroundColor Green
    } else {
        Write-Host "The exchange group for $($class.DisplayName) could not be found after 30 seconds. Try rerunning this script" -ForegroundColor Red
    }
}
Remove-PSSession -Session $PS_Session

