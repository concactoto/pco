# Define variables
$hMailServerInstallerUrl = "https://www.hmailserver.com/files/hMailServer-5.6.7-B2425.exe"
$hMailServerPath = "C:\hMailServer"
$downloadPath = "$env:TEMP\hMailServerSetup.exe"
$domainName = "secure.vermontlaw.edu"
$emailAccount = "users"
$emailPassword = "Kuteno12@"  # Change this to a secure password
$adminPassword = "Kuteno12@"  # Set this to the hMailServer admin password you set during installation

# Download hMailServer installer
Write-Output "Downloading hMailServer installer..."
try {
    Invoke-WebRequest -Uri $hMailServerInstallerUrl -OutFile $downloadPath -UseBasicParsing
} catch {
    Write-Output "Error downloading hMailServer installer: $_"
    exit
}

# Install hMailServer
Write-Output "Installing hMailServer..."
try {
    Start-Process -FilePath $downloadPath -ArgumentList "/silent" -Wait
    Start-Sleep -Seconds 30  # Wait for installation to complete
} catch {
    Write-Output "Error installing hMailServer: $_"
    exit
}

# Import hMailServer COM library
Write-Output "Importing hMailServer COM library..."
try {
    Add-Type -Path "$hMailServerPath\Bin\hMailServerNetAPILib.dll"
} catch {
    Write-Output "Error importing hMailServer COM library: $_"
    exit
}

# Create hMailServer application object
Write-Output "Creating hMailServer application object..."
try {
    $hMailServerApp = New-Object -ComObject hMailServer.Application
} catch {
    Write-Output "Error creating hMailServer application object: $_"
    exit
}

# Authenticate
Write-Output "Authenticating with hMailServer..."
try {
    $hMailServerApp.Authenticate("Administrator", $adminPassword)
} catch {
    Write-Output "Error authenticating with hMailServer: $_"
    exit
}

# Create domain
Write-Output "Creating domain $domainName..."
try {
    $domain = $hMailServerApp.Domains.Add()
    $domain.Name = $domainName
    $domain.Save()
} catch {
    Write-Output "Error creating domain: $_"
    exit
}

# Create account
Write-Output "Creating email account $emailAccount@$domainName..."
try {
    $account = $domain.Accounts.Add()
    $account.Address = "$emailAccount@$domainName"
    $account.Password = $emailPassword
    $account.Active = $true
    $account.Save()
} catch {
    Write-Output "Error creating email account: $_"
    exit
}

# Configure SMTP, POP3, and IMAP settings
Write-Output "Configuring protocols..."
try {
    $smtp = $hMailServerApp.Settings.Protocols.SMTP
    $smtp.SMTPRelayer = ""
    $smtp.RelayerRequiresAuthentication = $false
    $smtp.Save()

    $pop3 = $hMailServerApp.Settings.Protocols.POP3
    $pop3.MaxPOP3Connections = 50
    $pop3.Save()

    $imap = $hMailServerApp.Settings.Protocols.IMAP
    $imap.MaxIMAPConnections = 50
    $imap.Save()
} catch {
    Write-Output "Error configuring protocols: $_"
    exit
}

# Open firewall ports
Write-Output "Opening firewall ports..."
try {
    New-NetFirewallRule -DisplayName "Allow SMTP" -Direction Inbound -Protocol TCP -LocalPort 25 -Action Allow
    New-NetFirewallRule -DisplayName "Allow POP3" -Direction Inbound -Protocol TCP -LocalPort 110 -Action Allow
    New-NetFirewallRule -DisplayName "Allow IMAP" -Direction Inbound -Protocol TCP -LocalPort 143 -Action Allow
    New-NetFirewallRule -DisplayName "Allow SMTP Submission" -Direction Inbound -Protocol TCP -LocalPort 587 -Action Allow
} catch {
    Write-Output "Error opening firewall ports: $_"
    exit
}

# Confirm configuration
Write-Output "hMailServer installation and configuration complete."
Write-Output "Domain: $domainName"
Write-Output "Email Account: $emailAccount@$domainName"
