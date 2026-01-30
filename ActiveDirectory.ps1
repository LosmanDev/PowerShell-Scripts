# INPUT: Either SamAccountName or EmployeeID

$userName = ""  # e.g. "john.doe" or "123456"

#Get-ADUser -Identity $userName -Server $dc -Properties *

$dc = (Get-ADDomain).PDCEmulator
$dc.MaxPasswordAge

$commonParams = @{
    Server     = $dc
    Properties = @(
        'Enabled', 'Created', 'whenChanged', 'CanonicalName', 'DisplayName', 'accountExpires', 'AccountExpirationDate',
        'pwdLastSet', 'City', 'Department', 'directReports', 'EmployeeID', 'HomeDrive',
        'homePostalAddress', 'LastBadPasswordAttempt', 'LastLogonDate', 'Manager',
        'msDS-cloudExtensionAttribute7', 'msDS-cloudExtensionAttribute6', 'msDS-cloudExtensionAttribute14',
        'Office', 'OfficePhone', 'otherMobile', 'PostalCode', 'proxyAddresses',
        'SamAccountName', 'State', 'StreetAddress', 'Title', 'UserPrincipalName', 'msDS-cloudExtensionAttribute3', 'msDS-cloudExtensionAttribute15'
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
        'SamAccountName' = $user.SamAccountName
        'Status'         = 'Disabled'
        'LastModifiedOn' = $user.whenChanged
        'OUPath'         = $user.CanonicalName
        'Legal Hold'     = $user.'msDS-cloudExtensionAttribute7'
        'Manager'        = $user.Manager
        'Title'          = $user.Title 
        'Department'     = $user.Department
        'Home Email'     = $user.'msDS-cloudExtensionAttribute3'
    } | Format-List
}
else {
    # Enabled account: all the fields
    $report = [ordered]@{
        'Created'              = $user.Created
        'DisplayName'          = $user.DisplayName
        'User Principal Name'  = $user.UserPrincipalName
        'SamAccountName'       = $user.SamAccountName
        'Employee ID'          = $user.EmployeeID
        'Title'                = $user.Title       
        'Department'           = $user.Department
        'OU Path'              = $user.CanonicalName
        'Home Email'           = $user.'msDS-cloudExtensionAttribute3'
        'Office'               = $user.Office
        'Office Phone'         = $user.OfficePhone
        'Mobile'               = $user.otherMobile
        'Street Address'       = $user.StreetAddress
        'City, State, ZIP'     = "$($user.City), $($user.State) $($user.PostalCode)"
        'Home Postal Address'  = $user.homePostalAddress
        'Employment'           = $user.'msDS-cloudExtensionAttribute6'
        'Home Drive'           = $user.HomeDrive
        'Proxy Addresses'      = ($user.proxyAddresses -join ", ")
        #'Direct Reports'       = ($user.directReports -join ", ")
        'Manager'              = $user.Manager
        'Cloud Ext Attr 14'    = $user.'msDS-cloudExtensionAttribute14'
        'Last Bad Password At' = $user.LastBadPasswordAttemptS    
        'Password Last Set'    = $pwdLastSetDt
        'Password Expires'     = "$pwdExpiryDt ($daysLeft days left)"
        'Enabled?'             = $user.Enabled
        'LastModifiedOn'       = $user.whenChanged
        'Legal Hold'           = $user.'msDS-cloudExtensionAttribute7'
        'Account Expires'      = $user.AccountExpirationDate
    }
    
    # Display as a neat vertical list—you can swap to Format-Table if you prefer columns
    [PSCustomObject]$report | Format-List
        
}

#Set-ADAccountPassword -Identity $username -NewPassword (ConvertTo-SecureString '' -AsPlainText -Force) -Reset -COnfirm:$true


#Get-ADDefaultDomainPasswordPolicy | Select-Object MaxPasswordAge, MinPasswordAge, PasswordHistoryCount, MinPasswordLength, LockoutThreshold, LockoutDuration, LockoutObservationWindow
