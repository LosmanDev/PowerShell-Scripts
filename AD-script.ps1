$userName = "liban.osman"

# Fetch user and all needed properties in one go
$user = Get-ADUser -Identity $userName -Server ".net" `
    -Properties Enabled, whenChanged, CanonicalName, accountExpires, AccountExpirationDate, `
    pwdLastSet, City, Department, directReports, EmployeeID, HomeDrive, homePostalAddress, `
    LastBadPasswordAttempt, Manager, msDS-cloudExtensionAttribute7, msDS-cloudExtensionAttribute14, `
    Office, OfficePhone, otherMobile, PostalCode, proxyAddresses, SamAccountName, State, `
    StreetAddress, Title, UserPrincipalName

# Compute password dates
$pwdLastSetDt = [DateTime]::FromFileTime($user.pwdLastSet)
$pwdExpiryDt = $pwdLastSetDt.AddDays(90)
$daysLeft = ($pwdExpiryDt - (Get-Date)).Days

# Compute account expiration display
$acctExp = if ($user.AccountExpirationDate) { $user.AccountExpirationDate } else { 'Never' }

if (-not $user.Enabled) {
    # Disabled account: show minimal info
    $disabledReport = [ordered]@{
        'SamAccountName'   = $user.SamAccountName 
        'Status'           = 'Disabled'
        'Last Modified On' = $user.whenChanged
        'OU Path'          = $user.CanonicalName
        'Cloud Ext Attr 7' = $user.'msDS-cloudExtensionAttribute7'
    }

    [PSCustomObject]$disabledReport | Format-List

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
        'Home Drive'           = $user.HomeDrive
        'Proxy Addresses'      = ($user.proxyAddresses -join ", ")
        'Direct Reports'       = ($user.directReports -join ", ")
        'Last Bad Password At' = $user.LastBadPasswordAttempt
        'Manager'              = $user.Manager
        'Cloud Ext Attr 14'    = $user.'msDS-cloudExtensionAttribute14'
        'OU Path'              = $user.CanonicalName
        'Password Last Set'    = $pwdLastSetDt
        'Password Expires'     = "$pwdExpiryDt ($daysLeft days left)"
        'Account Expires'      = $acctExp
        'Enabled?'             = $user.Enabled
    }

    # Display as a neat vertical list—you can swap to Format-Table if you prefer columns
    [PSCustomObject]$report | Format-List
    
}
