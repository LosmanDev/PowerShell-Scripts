# INPUT: Either SamAccountName or EmployeeID
$userName = ""  # e.g. "john.doe" or "123456"

#Get-ADUser -Identity $userName -Server "" -Properties *


$commonParams = @{
    Server     = ""
    Properties = @(
        'Enabled', 'Created', 'whenChanged', 'CanonicalName', 'accountExpires', 'AccountExpirationDate',
        'pwdLastSet', 'City', 'Department', 'directReports', 'EmployeeID', 'HomeDrive',
        'homePostalAddress', 'LastBadPasswordAttempt', 'Manager',
        'msDS-cloudExtensionAttribute7', 'msDS-cloudExtensionAttribute6', 'msDS-cloudExtensionAttribute14',
        'Office', 'OfficePhone', 'otherMobile', 'PostalCode', 'proxyAddresses',
        'SamAccountName', 'State', 'StreetAddress', 'Title', 'UserPrincipalName'
    )
}


if ($userName -match '^\d+$') {
   
    $user = Get-ADUser -Filter "EmployeeID -eq '$userName'" @commonParams |
    Select-Object -First 1  
}
else {
   
    $user = Get-ADUser -Identity $userName @commonParams
}

if (-not $user) {
    Throw "No user found matching '$userName'."
}

$pwdLastSetDt = [DateTime]::FromFileTime($user.pwdLastSet)
$pwdExpiryDt = $pwdLastSetDt.AddDays(90)
$daysLeft = ($pwdExpiryDt - (Get-Date)).Days


if (-not $user.Enabled) {
    
    [PSCustomObject]@{
        SamAccountName = $user.SamAccountName
        Status         = 'Disabled'
        LastModifiedOn = $user.whenChanged
        OUPath         = $user.CanonicalName
        CloudExtAttr7  = $user.'msDS-cloudExtensionAttribute7'
        Manager        = $user.Manager
    } | Format-List
}
else {
    # Enabled account: all the fields
    $report = [ordered]@{
        'User Principal Name'  = $user.UserPrincipalName
        'SamAccountName'       = $user.SamAccountName
        'Employee ID'          = $user.EmployeeID
        'Title'                = $user.Title       
        'Department'           = $user.Department
        'Office'               = $user.Office
        'Office Phone'         = $user.OfficePhone
        'Mobile'               = $user.otherMobile
        'Street Address'       = $user.StreetAddress
        'City, State, ZIP'     = "$($user.City), $($user.State) $($user.PostalCode)"
        'Home Postal Address'  = $user.homePostalAddress
        'Employment'           = $user.'msDS-cloudExtensionAttribute6'
        'Home Drive'           = $user.HomeDrive
        'Proxy Addresses'      = ($user.proxyAddresses -join ", ")
        'Direct Reports'       = ($user.directReports -join ", ")
        'Last Bad Password At' = $user.LastBadPasswordAttempt
        'Manager'              = $user.Manager
        'Cloud Ext Attr 14'    = $user.'msDS-cloudExtensionAttribute14'
        'OU Path'              = $user.CanonicalName
        'Password Last Set'    = $pwdLastSetDt
        'Password Expires'     = "$pwdExpiryDt ($daysLeft days left)"
        'Account Expires'      = $user.accountExpires
        'Enabled?'             = $user.Enabled
        'Created'              = $user.Created
    }
    
    # Display as a neat vertical list—you can swap to Format-Table if you prefer columns
    [PSCustomObject]$report | Format-List
        
}
 
#Set-ADUser -Identity $user.SamAccountName -Replace @{pwdLastSet=0}
 
#Get-ADUser -Identity $user.SamAccountName -Properties pwdLastSet | Select-Object SamAccountName, pwdLastSet

#Set-ADUser -Identity $user.SamAccountName -Replace @{pwdLastSet=-1}

#Get-ADUser -Identity $user.SamAccountName -Properties pwdLastSet | Select-Object SamAccountName, pwdLastSet
