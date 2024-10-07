<#
.SYNOPSIS
Retrieves mailbox legacy exchange DN data for a list of users from Exchange Online and exports the information into a CSV file.
.DESCRIPTION
This script connects to Exchange Online, reads email addresses from a predefined CSV file located in the same directory as the script, retrieves mailbox details for each user (such as PrimarySmtpAddress, LegacyExchangeDN, Forwarding settings, etc.), and exports the results into a CSV file. The script manages the installation and importing of required modules, handles errors during mailbox data retrieval, and ensures proper disconnection from Exchange Online after execution.
.EXAMPLE
Simply run the script from its folder:
.\Export-MailboxLegacyDN.ps1
The script automatically reads email addresses from 'users.csv' in the same directory and exports mailbox data to 'Users_LEDN_Report.csv'.
.NOTES
Written by: SubjectData
#>


# Module names
$ExchangeOnlineModule = "ExchangeOnlineManagement"

# Check if the module is already installed
if (-not(Get-Module -Name $ExchangeOnlineModule)) {
    # Install the module
    Install-Module -Name $ExchangeOnlineModule -Force
}

Import-Module $ExchangeOnlineModule -Force


$myDir = Split-Path -Parent $MyInvocation.MyCommand.Path

$XLloc = "$myDir\"
#$ReportsPath = "$myDir\Outputs"
$CSVPath = $XLloc + "users.csv"

# Authenticate to Exchange Online
try {
    Connect-ExchangeOnline -ErrorAction Stop
    Write-Host "Connected to Exchange Online successfully." -ForegroundColor Green
}
catch {
    Write-Host "Failed to connect to Exchange Online: $_" -ForegroundColor Red
    exit 1
}

# Read email IDs from the input CSV file
try {
    $emailList = Import-Csv -Path $CSVPath | Select-Object -ExpandProperty Email
    Write-Host "Email IDs loaded successfully from $InputCsv." -ForegroundColor Green
}
catch {
    Write-Host "Error reading the input CSV file: $_" -ForegroundColor Red
    exit 1
}

# Initialize an array to store mailbox data
$mailboxData = @()

# Retrieve mailbox data for each email ID and add it to the array
foreach ($email in $emailList) {
    try {
        $mailbox = Get-Mailbox -Identity $email -ErrorAction Stop | Select-Object `
            PrimarySmtpAddress, 
            LegacyExchangeDN, 
            UserPrincipalName, 
            ForwardingAddress, 
            ForwardingSMTPAddress, 
            DeliverToMailboxAndForward,
            ArchiveStatus
                    
        $mailboxData += $mailbox

        Write-Host "Retrieved mailbox data for $email." -ForegroundColor Green
    }
    catch {
        Write-Host "Error retrieving mailbox data for $($email): $_" -ForegroundColor Red
        continue
    }
}

# Export all mailbox data to CSV
try {
    $mailboxData | Export-Csv -Path "$($XLloc)\Users_LEDN_Report.csv" -NoTypeInformation
    Write-Host "Mailbox data exported successfully" -ForegroundColor Green
}
catch {
    Write-Host "Error exporting mailbox data to CSV: $_" -ForegroundColor Red
}

# Disconnect from Exchange Online
try {
    Disconnect-ExchangeOnline -Confirm:$false
    Write-Host "Disconnected from Exchange Online." -ForegroundColor Green
}
catch {
    Write-Host "Error disconnecting from Exchange Online: $_" -ForegroundColor Red
}

