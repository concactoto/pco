# Define variables
$hMailServerInstallerUrl = "https://www.hmailserver.com/download_latest.aspx"
$hMailServerPath = "C:\Program Files (x86)\hMailServer"
$downloadPath = "$env:TEMP\hMailServerSetup.exe"
$domainName = "secure.vermontlaw.edu"
$emailAccount = "supports"
$emailPassword = "Kuteno12@"  # Change this to a secure password
$adminPassword = "Kuteno12@"  # Set this to the hMailServer admin password you set during installation

# Download hMailServer installer
Write-Output "Downloading hMailServer installer..."
Invoke-WebRequest -Uri $hMailServerInstallerUrl -OutFile $downloadPath

# Install hMailServer
Write-Output "Installing hMailServer..."
Start-Process -FilePath $downloadPath -ArgumentList "/silent" -Wait

# Wait for installation to complete
Start-Sleep -Seconds 10

# Import hMailServer COM library
Write-Output "Importing hMailServer COM library..."
Add-Type -Path "$hMailServerPath\Bin\hMailServerNetAPILib.dll"

# Create hMailServer application object
Write-Output "Creating hMailServer application object..."
$hMailServerApp = New-Object -ComObject hMailServer.Application

# Authenticate
Write-Output "Authenticating with hMailServer..."
$hMailServerApp.Authenticate("Administrator", $adminPassword)

# Create domain
Write-Output "Creating domain $domainName..."
$domain = $hMailServerApp.Domains.Add()
$domain.Name = $domainName
$domain.Save()

# Create account
Write-Output "Creating email account $emailAccount@$domainName..."
$account = $domain.Accounts.Add()
$account.Address = "$emailAccount@$domainName"
$account.Password = $emailPassword
$account.Active = $true
$account.Save()

# Configure SMTP, POP3, and IMAP settings
Write-Output "Configuring protocols..."
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

# Open firewall ports
Write-Output "Opening firewall ports..."
New-NetFirewallRule -DisplayName "Allow SMTP" -Direction Inbound -Protocol TCP -LocalPort 25 -Action Allow
New-NetFirewallRule -DisplayName "Allow POP3" -Direction Inbound -Protocol TCP -LocalPort 110 -Action Allow
New-NetFirewallRule -DisplayName "Allow IMAP" -Direction Inbound -Protocol TCP -LocalPort 143 -Action Allow
New-NetFirewallRule -DisplayName "Allow SMTP Submission" -Direction Inbound -Protocol TCP -LocalPort 587 -Action Allow

# Confirm configuration
Write-Output "hMailServer installation and configuration complete."
Write-Output "Domain: $domainName"
Write-Output "Email Account: $emailAccount@$domainName"
