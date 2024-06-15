# Define variables
$smarterMailAdmin = "admin"
$smarterMailAdminPassword = "NewAdminPassword123!"
$smarterMailServer = "localhost"
$domainToAdd = "b.com"
$emailAccount = "user@b.com"
$emailPassword = "UserPassword123!"

# Function to reset SmarterMail admin password
function Reset-AdminPassword {
    $adminUrl = "http://$smarterMailServer/Services/svcUserAdmin.asmx"
    $body = @{
        username = $smarterMailAdmin
        password = $smarterMailAdminPassword
    }

    Invoke-WebRequest -Uri "$adminUrl/UpdateUserPassword" -Method Post -Body $body | Out-Null
    Write-Host "Admin password reset successfully."
}

# Function to add a new domain
function Add-Domain {
    $domainAdminUrl = "http://$smarterMailServer/Services/svcDomainAdmin.asmx"
    $body = @{
        domain = $domainToAdd
        password = $smarterMailAdminPassword
    }

    Invoke-WebRequest -Uri "$domainAdminUrl/AddDomain" -Method Post -Body $body | Out-Null
    Write-Host "Domain added successfully."
}

# Function to create a new email account
function Create-EmailAccount {
    $userAdminUrl = "http://$smarterMailServer/Services/svcUserAdmin.asmx"
    $body = @{
        username = $emailAccount
        password = $emailPassword
    }

    Invoke-WebRequest -Uri "$userAdminUrl/AddUser" -Method Post -Body $body | Out-Null
    Write-Host "Email account created successfully."
}

# Function to test email reception (external testing is manual)
function Test-EmailReception {
    Write-Host "Send a test email to $emailAccount from an external email provider to verify the setup."
}

# Execute functions
Reset-AdminPassword
Add-Domain
Create-EmailAccount
Test-EmailReception