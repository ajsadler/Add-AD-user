# Clears current variables
Remove-Variable * -ErrorAction SilentlyContinue
Get-Variable -Exclude PWD,*Preference | Remove-Variable -EA 0

# Defines the log file. Adds to existing file instead of over-writing
$Logfile = "C:\temp\Add AD user log.log"
function WriteLog {
Param ([string]$LogString)
$Stamp = (Get-Date).toString("yyyy/MM/dd HH:mm:ss")
$LogMessage = "$Stamp $LogString"
Add-content $LogFile -value $LogMessage }

# Picks out the current logged in user (ie. username-adm)
$logonName = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
$logonNameSplit = $logonName.Split("\")
$admUsername = $logonNameSplit[1]

WriteLog "`n"
WriteLog "`n"
WriteLog "The script started by user $admUsername."

# Defines the text culture necessary for 'ToTitleCase'
$textInfo = (Get-Culture).TextInfo

# Validates the name to be precisely two names (hyphenated names are accepted, eg. Mary-Jane)
do
{   $newName = Read-Host "New user's full name"
    $newFullName = $newName.split(" ")
    if ($newFullName.Length -eq 2) {continue}
    "Invalid name entry. (Firstname Surname)"
    WriteLog "Invalid name entered: '$newName'"
}   while ($null -eq $newFullName -or $newFullName.Length -ne 2)

# Converts the name data to all necessary formats
$newFirstName = $textInfo.ToTitleCase($newFullName[0].ToLower())
$newSurname = $textInfo.ToTitleCase($newFullName[1].ToLower())
$newFirstInitial = $newFirstName.Substring(0,1).ToLower()
$newSurnameInitial = $newSurname.Substring(0,1).ToLower()
$newFirstNameLower = $newFirstName.ToLower()
$newSurnameLower = $newSurname.ToLower()
$newSurnameTruncate = $newSurnameLower.Replace("-", "")
$newSurnameTruncateLength = $newSurnameTruncate.Length
if ($newSurnameTruncateLength -gt 9) {$newSurnameTruncate = $newSurnameTruncate.Substring(0,9)}
$newUsername = $newFirstInitial + $newSurnameTruncate
$newEmail = $newFirstNameLower + "." + $newSurnameLower + "@igdoors.co.uk"
$newDisplayName = $newSurname + ", " + $newFirstName

# Error checking for a duplicate User Principal Name (UPN) / username
# Tests for an existing AD user with the created username
# If one is found, then it loops with sequential numbers until a gap is found
$error.Clear()
$i = 1
do
{   try {Get-ADUser $newUsername}
    catch {$i--
        "Duplicate user(s) found: $i"}
    if ($error) {continue}
    else {
        $i++
        if ($newSurnameTruncateLength -gt 8) {$newSurnameTruncate = $newSurnameTruncate.Substring(0,8)}
        $newUsername = $newFirstInitial + $newSurnameTruncate + $i
        $newEmail = $newFirstNameLower + "." + $newSurnameLower + $i + "@igdoors.co.uk"
        $newDisplayName = $newSurname + ", " + $newFirstName + $i
        }
}   while (!$error)

WriteLog "New user: $newFirstName $newSurname"
WriteLog "Display name: $newDisplayName"
WriteLog "Login name: $newUsername"
WriteLog "Email address: $newEmail"

# Validates date format with any delimiter
# Must contain dd, mm, yyyy
# Although practically the yyyy value is not necessarily needed
$delimiters = "/",".","-"," "
do
{   $newStartDate = Read-Host "New user's start date (as dd/mm/yyyy)"
    $newStartDateSplit = $newStartDate -Split {$delimiters -contains $_}
    if ($newStartDateSplit[0].Length -eq 2 -and $newStartDateSplit[1].Length -eq 2 -and $newStartDateSplit[2].Length -eq 4) {continue}
    "Invalid date format."
    WriteLog "Invalid date format entered: '$newStartDate'"
}   while ($newStartDateSplit[0].Length -ne 2 -or $newStartDateSplit[1].Length -ne 2 -or $newStartDateSplit[2].Length -ne 4)
$newStartDateDD = $newStartDate.Substring(0,2)
$newStartDateMM = $newStartDate.Substring(3,2)
$newStartDateYYYY = $newStartDate.Substring(6,4)

$newPasswordString = $newFirstInitial + $newSurnameInitial + $newStartDateDD + $newStartDateMM
$newPassword = "IGD$newPasswordString#"

WriteLog "Start date: $newStartDateDD/$newStartDateMM/$newStartDateYYYY"

"`nDepartment list:"
"3Tec                   IT                      Production"
"Accounts               Logistics               Project Engineering"
"After Sales            Logistics (Bridgetime)  Purchasing"
"CNC Team               Materials & Warehouse   Quality & HSE"
"Customer Care          New Build               Research & Development"
"Engineering            Operations              Sales"
"Finance                Order Processing        Social Housing"
"Health & Safety        Payroll                 Technical"
"Human Resources        Planning                Trade`n"

# Loop to add more departments to an array, exits when a blank entry is made
# The first department added is designated the 'Primary', eg. used for the email signature
# Any further departments added are for zPermissions
$newDepartment = @()
$addDepartment = "y"
do
{   if ($addDepartment -eq "n") {continue}
    $newDepartment += Read-Host "New user's department"
    if ($newDepartment.Contains(""))
    {   $addDepartment = "n";
        $newDepartment = $newDepartment -ne ""
        continue
    }
}   while ($addDepartment -eq "y")
$newDepartmentList = $newDepartment -join ', '
$newDepartmentList = $textInfo.ToTitleCase($newDepartmentList.ToLower())

$newDepartmentFull = @()
$newDepartmentAD = @()
foreach ($department in $newDepartment)
{
	switch -wildcard ($department)
    {
        '3t*'       {$newDepartmentFull += "3Tec";                                  $newDepartmentAD += "3tec"}
        'acc*'      {$newDepartmentFull += "Accounts Department";                   } # $newDepartmentAD += ""}
        'aft*'      {$newDepartmentFull += "After Sales Department";                $newDepartmentAD += "aftersales-team"}
        'cnc*'      {$newDepartmentFull += "CNC Team";                              $newDepartmentAD += "CNC-Team"}
        'cus*'      {$newDepartmentFull += "Customer Care Department";              $newDepartmentAD += "customercare-team"}
        'eng*'      {$newDepartmentFull += "Engineering Department";                $newDepartmentAD += "engineering-team"}
        'fin*'      {$newDepartmentFull += "Finance Department";                    $newDepartmentAD += "finance-team"}
        'hea*'      {$newDepartmentFull += "Health & Safety Department";            $newDepartmentAD += "HealthandSafety"}
        'hum*'      {$newDepartmentFull += "Human Resources Department";            $newDepartmentAD += "human-resources"}
        'hr'        {$newDepartmentFull += "Human Resources Department";            $newDepartmentAD += "human-resources"}
        'it*'       {$newDepartmentFull += "I.T. Systems";                          $newDepartmentAD += "it-team"}
        'log*'      {$newDepartmentFull += "Logistics";                             $newDepartmentAD += "Logistics"}
        '*brid*'    {$newDepartmentFull += "Logistics (Bridgetime)";                $newDepartmentAD += "bridgetime"}
        'mat*'      {$newDepartmentFull += "Materials & Warehouse Management";      $newDepartmentAD += "MaterialsandWarehouse-Management"}
        'new*'      {$newDepartmentFull += "New Build Division";                    $newDepartmentAD += "newbuild"}
        'ope*'      {$newDepartmentFull += "Operations Department";                 } # $newDepartmentAD += ""}
        'ord*'      {$newDepartmentFull += "Order Processing Department";           $newDepartmentAD += "orderprocessing-team"}
        'pla*'      {$newDepartmentFull += "Planning Department";                   $newDepartmentAD += "planning"}
        'prod*'     {$newDepartmentFull += "Production Department";                 $newDepartmentAD += "production"}
        'proj*'     {$newDepartmentFull += "Project Engineering Department";        $newDepartmentAD += "project-engineering"}
        'pur*'      {$newDepartmentFull += "Purchasing Department";                 } # $newDepartmentAD += ""}
        'qua*'      {$newDepartmentFull += "Quality & HSE Department";              $newDepartmentAD += "quality-team"}
        'res*'      {$newDepartmentFull += "Research & Development";                $newDepartmentAD += "randd-team"}
        'rma*'      {$newDepartmentFull += "RMA";                                   $newDepartmentAD += "RMA"}
        'sag*'      {$newDepartmentFull += "Payroll Department";                    $newDepartmentAD += "sage-payroll"}
        'pay*'      {$newDepartmentFull += "Payroll Department";                    $newDepartmentAD += "sage-payroll"}
        'sal*'      {$newDepartmentFull += "Sales Department";                      $newDepartmentAD += "sales-team"}
        'soc*'      {$newDepartmentFull += "Social Housing Department";             $newDepartmentAD += "socialhousing-team"}
        'tec*'      {$newDepartmentFull += "Technical Department";                  $newDepartmentAD += "technical-team"}
        'tra*'      {$newDepartmentFull += "Trade Division";                        $newDepartmentAD += "trade-team"}
        Default     {"A department could not be determined: '$department'";         WriteLog "A department could not be determined: '$department'"}
    }
}
$newDepartmentPrimary = $newDepartmentFull[0]
$newDepartmentADList = $newDepartmentAD -join ', '

WriteLog "Department: $newDepartmentPrimary"
WriteLog "Active Directory groups: $newDepartmentADList"

# Lists the currently provided details and asks to confirm
"`nNew user:                  $newFirstName $newSurname"
"Display name:              $newDisplayName"
"Login name:                $newUsername"
"Email address:             $newEmail"
"Start date:                $newStartDateDD/$newStartDateMM/$newStartDateYYYY"
"Department:                $newDepartmentPrimary"
"Active Directory groups:   $newDepartmentADList"
"`nPlease check and confirm the above details."

do
{   $confirmDetails = (Read-Host "Are the details correct: Y/N?").Substring(0,1).ToLower()
    if ($confirmDetails -eq "y" -or $confirmDetails -eq "n") {continue}
    "`nInvalid Y/N entry."
}   while ($confirmDetails -ne "y" -and $confirmDetails -ne "n")

if ($confirmDetails -eq "n")
{
    "`nPlease exit and restart.`a"
    WriteLog "Are the details correct: selected 'NO'"
    if ($Host.Name -eq "ConsoleHost")
    {
        Write-Host "Press any key to continue..."
        $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyUp") > $null
    }
    exit
}

WriteLog "Are the details correct: selected 'YES'"

# Creates a new user in Active Directory under /gb/blw/user/internal users, with the default new password
New-ADUser -Name "$newUsername" -path "OU=internal Users,OU=user,OU=blw,OU=gb,OU=hgroup-production,DC=hgroup,DC=intra" -AccountPassword $(ConvertTo-SecureString -AsPlainText $newPassword -Force) -Enabled $true

# Sets the new user properties
Set-ADUser -Identity "$newUsername" -Replace @{
'Company' = "IG Doors Ltd.";
'StreetAddress' = "Lon Gellideg, Oakdale Business Park";
'l' = "Blackwood";
'st' = "South Wales";
'PostalCode' = "NP12 4AE";
'c' = "GB";
'co' = "United Kingdom";
'countryCode' = "826";
'GivenName' = "$newFirstName";
'sn' = "$newSurname";
'DisplayName' = "$newDisplayName";
'mailNickname' = "$newUsername";
'mail' = "$newEmail";
'UserPrincipalName' = "$newEmail";
'targetAddress' = "smtp:$newFirstNameLower.$newSurnameLower.igdoors.co.uk@hgrouponline.mail.onmicrosoft.com";
'Department' = "$newDepartmentPrimary";
'extensionAttribute1' = "Included";
'extensionAttribute15' = "AADSync";
}

# Renames the user
Get-ADUser "$newUsername" | Rename-ADObject -NewName "$newDisplayName"

# Sets the attributes for the proxyAddresses
$User = Get-ADUser $newUsername -Properties proxyAddresses
$User.proxyAddresses.Add("SMTP:$newEmail")
$User.proxyAddresses.Add("smtp:$newFirstNameLower.$newSurnameLower.igdoors.co.uk@hgrouponline.mail.onmicrosoft.com")
Set-ADUser -instance $User

# Sets 'Allow Trusted Locations'
Add-ADGroupMember -Identity zzx-gbblw-m365-ConditionalAccess-allowTrustedLocations -Members $newUsername

# Added to zPermissions for each department added
foreach ($department in $newDepartmentAD)
{Add-ADGroupMember -Identity zzd-gbblw-igdoors-$department-rw -Members $newUsername}

# Displays the new details before closing
Get-ADUser -Identity $newUsername -Properties *

# Writes to the log a script to set up mailbox settings in the Exchange Management Shell
WriteLog "`n
Enable-Remotemailbox -Identity '$newUsername' -PrimarySmtpAddress '$newEmail' -Remoteroutingaddress '$newFirstNameLower.$newSurnameLower.igdoors.co.uk@hgrouponline.mail.onmicrosoft.com'
Enable-RemoteMailbox $newEmail -Archive
Connect-ExchangeOnline -UserPrincipalName '$admUsername@igdoors.co.uk'
Set-Mailbox -Identity $newEmail -RetentionPolicy 'HGROUP Default RetentionPolicy'"

# Add user to the zPermissions excel file
# $xl = Import-Excel -path \\hgroup.intra\data\GB-IGDoors\zPermissions\copyDepartment-Data_Permissions_IGDoors.xlsx 
# $ws = $xl.Workbook.Worksheets["Permissions"]
# Set-ExcelRange
# $xl.InsertRow(1,1)
# Close-ExcelPackage -ExcelPackage $xl -Show

WriteLog "The task has run successfully."
"The task has run successfully."
"Mailbox settings script has been output to C:\temp\Add AD user log."

# If running in the console, wait for input before closing.
if ($Host.Name -eq "ConsoleHost")
{
    Write-Host "Press any key to continue..."
    $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyUp") > $null
}