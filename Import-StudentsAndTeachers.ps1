[CmdletBinding()]
[OutputType([string])]
Param(
    [parameter(Mandatory=$true)]
    [String]$csvPath,
    
    [parameter(Mandatory=$true)]
    [String]$defaultPassword
)


$UserImport = Import-Csv -Path $csvPath

$CheckRoles = $UserImport | Where-Object{ 'Teacher','Student','Faculty' -notcontains $_.Role }
if($CheckRoles){
    throw "Bad user roles found for the users. Roles should only be 'Teacher','Student','Faculty': $($CheckRoles | Out-String)"
}

# Connect to AzureAD
Connect-AzureAD

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
    @{l='Purchased';e={$_.PrepaidUnits.Enabled}}, SkuID | Out-GridView -Title 'Select default license' -PassThru | Select-Object -ExpandProperty SkuID



Function New-Office365Person{
<#
.SYNOPSIS
Creates the office 365 user account

.PARAMETER FirstName
The user’s first name.

.PARAMETER MiddleName
The user’s middle name.

.PARAMETER LastName
The user’s last name.

.PARAMETER Password
The user's password to set

#>
    [CmdletBinding()]
    [OutputType([object])]
    Param(
        [parameter(Mandatory=$true)]
        [String]$FirstName,
        
        [parameter(Mandatory=$false)]
        [String]$MiddleName,
        
        [parameter(Mandatory=$true)]
        [String]$LastName,
        
        [parameter(Mandatory=$true)]
        [String]$Password,
        
        [parameter(Mandatory=$true)]
        [String]$Role,
        
        [parameter(Mandatory=$true)]
        [String]$LicenseSku
    )

    # determine user login name
    $UserPrincipalName = Set-Office356UserName -First $FirstName -Middle $MiddleName -Last $LastName -Pattern $Pattern -TenantDomain $TenantDomain

    $PasswordProfile=New-Object -TypeName Microsoft.Open.AzureAD.Model.PasswordProfile
    $PasswordProfile.Password = $Password

    # create the user
    $userParams = @{
        DisplayName = "$FirstName $LastName" 
        GivenName = $FirstName
        SurName = $LastName
        UserPrincipalName = $UserPrincipalName
        JobTitle = $Role
        UsageLocation = 'US'
        MailNickName = $FirstName
        PasswordProfile = $PasswordProfile 
        AccountEnabled = $true
    }
    $userObject = New-AzureADUser @userParams

    # Set the license.
    $license = New-Object -TypeName Microsoft.Open.AzureAD.Model.AssignedLicense
    $AssignedLicenses = New-Object -TypeName Microsoft.Open.AzureAD.Model.AssignedLicenses
    $license.SkuId = $LicenseSku
    $AssignedLicenses.AddLicenses = $license
    Set-AzureADUserLicense -ObjectId $userObject.ObjectId -AssignedLicenses $AssignedLicenses | Out-Null

    $userObject | Select-Object @{l='FirstName';e={$_.GivenName}}, @{l='LastName';e={$_.Surname}}, @{l='Role';e={$_.JobTitle}}, UserPrincipalName
}

Function Set-Office356UserName{
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
        [String]$First,
        
        [parameter(Mandatory=$false)]
        [String]$Middle,
        
        [parameter(Mandatory=$true)]
        [String]$Last,
        
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
            $username += $First
        }
        elseif($item.Equals([char]'f')){
            $username += $($First.ToCharArray())[$FirstInt]
            $FirstInt++
        }
        elseif($item.Equals([char]'M') -and $Middle){
            $username += $Middle
        }
        elseif($item.Equals([char]'m') -and $Middle){
            $username += $($Middle.ToCharArray())[$MiddleInt]
            $MiddleInt++
        }
        elseif($item.Equals([char]'L')){
            $username += $Last
        }
        elseif($item.Equals([char]'l')){
            $username += $($Last.ToCharArray())[$LastInt]
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
    if($User){
        $number++
        Set-Office356UserName -First $First -Middle $Middle -Last $Last -Pattern $Pattern -digits $digits -number $number -TenantDomain $TenantDomain -IncludeNumber
    }
    else{
        "$($username)@$($TenantDomain)"
    }
}

$script:Pattern = 'F.L'
# create the user accounts
[System.Collections.Generic.List[PSObject]]$CreatedUsers = @()
foreach($user in $UserImport){
    Write-Progress -Activity "Creating users" -Status "$($CreatedUsers.count) of $($UserImport.count)" -PercentComplete $(($($CreatedUsers.count)/$($UserImport.count))*100) -id 1
    $Office365Person = @{
        FirstName = $user.FirstName
        MiddleName = $user.MiddleName
        LastName = $user.LastName
        Password = $defaultPassword
        Role = $user.Role
        LicenseSku = $LicenseSku
    }
    
    $newUser = New-Office365Person @Office365Person
    $CreatedUsers.Add($newUser)
}
Write-Progress -Activity "Done" -Id 1 -Completed

$fileName = [System.IO.Path]::GetFileNameWithoutExtension($csvPath)
$exportCsv = Join-Path (Split-Path $csvPath) "$($fileName)-imported.csv"
$CreatedUsers | Export-Csv -Path $exportCsv -NoTypeInformation

Write-Output "A list of the created accounts has been exported to '$exportCsv'.`nPlease use the UserPrincipalNames to map students and teachers to classes"

