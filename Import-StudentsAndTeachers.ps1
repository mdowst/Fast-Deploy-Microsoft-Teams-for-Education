<#
.SYNOPSIS
Use this script to import a CSV for students or teachers to a Office 365 for education subscription

.PARAMETER csvPath
The full path to the CSV import file

.PARAMETER UsageLocation
A two letter country code (ISO standard 3166). Required for users that will be assigned licenses due 
to legal requirement to check for availability of services in countries. 
Examples include: "US", "JP", and "GB".
For a complete list go to - https://www.iso.org/obp/ui/#search

.EXAMPLE
.\Import-StudentsAndTeachers.ps1 -csvPath .\Import-Teachers.csv

Imports the Import-Teacher.csv file where the file is in the same folder as the script

.EXAMPLE
.\Import-StudentsAndTeachers.ps1 -csvPath C:\School-Hydration\Import-Students.csv

Imports the Import-Students.csv file where the file is in a different folder then the script

.EXAMPLE
.\Import-StudentsAndTeachers.ps1 -csvPath .\Import-Teachers.csv -UsageLocation 'GB'

Imports the CSV file where the file and sets the usage location to the United Kingdom

#>
[CmdletBinding()]
[OutputType([string])]
Param(
    [parameter(Mandatory=$true)]
    [String]$csvPath,
    
    [parameter(Mandatory=$false)]
    [String]$UsageLocation = 'US'
)

# Check for the AzureAD module
if(-not (Get-Module AzureAD)){
    if(Get-Module AzureAD -ListAvailable){
        Import-Module AzureAD
    } else {
        throw "The AzureAD module is required for this script. You can install the module by running the command 'Install-Module AzureAD -Scope CurrentUser'"
    }
}

# Connect to AzureAD
try{ Get-AzureADCurrentSessionInfo -ErrorAction stop | Out-Null }
catch{ Connect-AzureAD | Out-Null }

$UserImport = Import-Csv -Path $csvPath

$CheckRoles = $UserImport | Where-Object{ 'Teacher','Student','Faculty' -notcontains $_.Role }
if($CheckRoles){
    throw "Bad user roles found for the users. Roles should only be 'Teacher','Student','Faculty': $($CheckRoles | Out-String)"
}


# set the domain for the username
$script:TenantDomain = Get-AzureADTenantDetail | Select-Object -ExpandProperty VerifiedDomains | Select-Object -Property Name, Capabilities  | 
    Out-GridView -Title 'Select email domain' -PassThru | Select-Object -ExpandProperty Name


$Sku = @{
    "M365EDU_A5_FACULTY" = "Office 365 A5 for faculty"
    "M365EDU_A5_STUDENT" = "Office 365 A5 for students"
    "M365EDU_A3_FACULTY" = "Office 365 A3 for faculty"
    "M365EDU_A3_STUDENT" = "Office 365 A3 for students"
    "M365EDU_A1_FACULTY" = "Office 365 A1 for faculty"
    "M365EDU_A1_STUDENT" = "Office 365 A1 for students"
}

# Create the objects we'll need to add licenses
$LicenseSku = Get-AzureADSubscribedSku | Select-Object @{l='License';e={$sku.Item($_.SkuPartNumber)}}, SkuPartNumber, @{l='Available';e={$_.PrepaidUnits.Enabled - $_.ConsumedUnits}},
    @{l='Purchased';e={$_.PrepaidUnits.Enabled}}, SkuID | Out-GridView -Title 'Select default license' -PassThru

# Get the naming pattern to use
$NamePatterns = ('[{"Pattern":"F.L","Example":"Diana.Price"},
    {"Pattern":"fL","Example":"DPrice"},{"Pattern":"F.m.L",
    "Example":"Diana.L.Price"}]' | ConvertFrom-Json)
$Pattern = $NamePatterns | Out-GridView -Title 'Select naming pattern' -PassThru | Select-Object -ExpandProperty Pattern

Function Check-Office356UserName{
<#
.SYNOPSIS
Use to determine a user's Office 365 username

.DESCRIPTION
This function will determine the user’s Office 365 username based on the person’s name, and the pattern specified. 

.PARAMETER First
The user’s first name.

.PARAMETER Middle
The user’s middle name.

.PARAMETER Last
The user’s last name.

.PARAMETER Pattern
The pattern to use when creating the samaccountname. It uses the letters (F, M, L) to set the pattern. 
By specifying the uppercase letter it will use the full word. By specifying a lowercase it will use a 
single char from the word starting with the first letter. Lowercase letters can be specified multiple 
times. Each instance will use the next letter. Any other character specified will be passed as part of 
the username. See the examples below:
	F.m.L = First.M.Last
	fL = FLast
	flllm = FLasM

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
        [parameter(Mandatory=$true)]
        [String]$FirstName,
        
        [parameter(Mandatory=$false)]
        [String]$MiddleName,
        
        [parameter(Mandatory=$true)]
        [String]$LastName,

        [parameter(Mandatory=$false)]
        [String]$UniqueId,
        
        [parameter(Mandatory=$true)]
        [String]$Pattern,

        [parameter(Mandatory=$true)]
        [String]$TenantDomain,

        [parameter(Mandatory=$false)]
        [ValidateRange(1,20)]
        [int]$digits=2,

        [parameter(Mandatory=$false)]
        [int]$number=0,

        [parameter(Mandatory=$false)]
        [switch]$IncludeNumber
    )
    $username = [string]::Empty
    $FirstInt = 0; $MiddleInt = 0; $LastInt = 0

    foreach ($item in $Pattern.ToCharArray()){
        if($item.Equals([char]'F')){
            $username += $FirstName
        }
        elseif($item.Equals([char]'f')){
            $username += $($FirstName.ToCharArray())[$FirstInt]
            $FirstInt++
        }
        elseif($item.Equals([char]'M') -and $MiddleName){
            $username += $MiddleName
        }
        elseif($item.Equals([char]'m') -and $MiddleName){
            $username += $($MiddleName.ToCharArray())[$MiddleInt]
            $MiddleInt++
        }
        elseif($item.Equals([char]'L')){
            $username += $LastName
        }
        elseif($item.Equals([char]'l')){
            $username += $($LastName.ToCharArray())[$LastInt]
            $LastInt++
        }
        elseif($item -notmatch "[1-9a-zA-Z]"){
            $username += $item
        }

    }

    if($IncludeNumber){
        $username += ($number.ToString().PadLeft($digits)).Replace(" ","0")
    }

    $User = Get-AzureADUser -Filter "userPrincipalName eq '$($username)@$($TenantDomain)'"
    if(-not [string]::IsNullOrEmpty($UniqueId) -and $User.FacsimileTelephoneNumber -eq $UniqueId){
        [pscustomobject]@{Exists=$true;UPN="$($username)@$($TenantDomain)"}
    } elseif($User){
        $number++
        if(-not $PSBoundParameters.ContainsKey('number')){
            $PSBoundParameters.Add('number',$number)
        } else {
            $PSBoundParameters['number'] = $number
        }

        if(-not $PSBoundParameters.ContainsKey('IncludeNumber')){
            $PSBoundParameters.Add('IncludeNumber',$true)
        } 
        Check-Office356UserName @PSBoundParameters
    }
    else{
        [pscustomobject]@{Exists=$false;UPN="$($username)@$($TenantDomain)"}
    }
}


# create the user accounts
[System.Collections.Generic.List[PSObject]]$CreatedUsers = @()
foreach($user in $UserImport){
    Write-Progress -Activity "Creating users" -Status "$($CreatedUsers.count) of $($UserImport.count)" -PercentComplete $(($($CreatedUsers.count)/$($UserImport.count))*100) -id 1

    # determine user login name
    $userCheck = @{
        FirstName = $user.FirstName
        LastName = $user.LastName
        MiddleName = $UPNCheck.UPN
        Pattern = $Pattern
        TenantDomain = $TenantDomain
        UniqueId = $user.id
    }
    $UPNCheck = Check-Office356UserName @userCheck

    if($UPNCheck.Exists -eq $true){
        Write-Host "User account for '$($user.FirstName) $($user.LastName)' was found." -ForegroundColor Cyan
        $userObject = Get-AzureADUser -ObjectId $UPNCheck.UPN
        $AssignedSku = @(Get-AzureADUserLicenseDetail -ObjectId $userObject.ObjectId)[0].SkuPartNumber
        if($sku.Item($AssignedSku)){
            $AssignedSku = $sku.Item($AssignedSku)
        }
    } else {
        Write-Host "Creating user account for '$($user.FirstName) $($user.LastName)'"
        $PasswordProfile=New-Object -TypeName Microsoft.Open.AzureAD.Model.PasswordProfile
        $PasswordProfile.Password = 'Welcome' + $user.id + '!'
        $PasswordProfile.ForceChangePasswordNextLogin = $true
    
        # create the user
        $userParams = @{
            DisplayName = "$($user.FirstName) $($user.LastName)" 
            GivenName = $user.FirstName
            SurName = $user.LastName
            UserPrincipalName = $UPNCheck.UPN
            JobTitle = $user.Role
            UsageLocation = $UsageLocation
            MailNickName = "$($user.FirstName)$($user.LastName)"
            PasswordProfile = $PasswordProfile 
            AccountEnabled = $true
            FacsimileTelephoneNumber = $user.id
        }
        $userObject = New-AzureADUser @userParams
    
        # Set the license.
        $license = New-Object -TypeName Microsoft.Open.AzureAD.Model.AssignedLicense
        $AssignedLicenses = New-Object -TypeName Microsoft.Open.AzureAD.Model.AssignedLicenses
        $license.SkuId = $LicenseSku.SkuID
        $AssignedLicenses.AddLicenses = $license
        try{
            Set-AzureADUserLicense -ObjectId $userObject.ObjectId -AssignedLicenses $AssignedLicenses -ErrorAction Stop | Out-Null
            $AssignedSku = $LicenseSku.License
        } catch {
            $AssignedSku = $null
        }
    }

    $newUser = $userObject | Select-Object @{l='FirstName';e={$_.GivenName}}, @{l='LastName';e={$_.Surname}}, @{l='Role';e={$_.JobTitle}},
        @{l='AssignedLicenses';e={$AssignedSku}}, @{l='Id';e={$_.FacsimileTelephoneNumber}}, @{l='Password';e={'Welcome' + $user.id + '!'}}, UserPrincipalName
        
    $CreatedUsers.Add($newUser)
    
}
Write-Progress -Activity "Done" -Id 1 -Completed

# Export the results to a new CSV
$fileName = [System.IO.Path]::GetFileNameWithoutExtension($csvPath)
$exportCsv = Join-Path (Split-Path $csvPath) "$($fileName)-imported.csv"
$CreatedUsers | Export-Csv -Path $exportCsv -NoTypeInformation

Write-Output "A list of the created accounts has been exported to '$exportCsv'.`nPlease use the UserPrincipalNames to map students and teachers to classes"

