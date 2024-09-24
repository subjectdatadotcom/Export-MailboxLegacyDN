# Export Mailbox Legacy Exchange DN Script

This PowerShell script retrieves mailbox details, including the Legacy Exchange DN, for a list of users from Exchange Online and exports the information into a CSV file.

## Prerequisites

- **PowerShell**: Ensure you have PowerShell installed.
- **Exchange Online Management Shell**: The script uses the ExchangeOnlineManagement module, which will be automatically installed if it's not already present.

## Instructions

### Edit the Script:

1. Open the script file.
2. Verify that the email list CSV (`users.csv`) is in the same directory as the script. This file should contain a list of user email addresses you wish to process.

### Prepare the CSV File: (sample file attached)

- Ensure you have a `users.csv` file in the same directory as the script.
- The CSV file should have the following structure:

```plaintext
Email
user1@example.com
user2@example.com
```
3. **Run the Script**:
   - Open PowerShell as an administrator.
   - Navigate to the directory containing the script.
   - Run the script:
     ```powershell
     .\Export-MailboxLegacyDN.ps1
     ```
   - Authenticate using Admin account
   - The script will connect to your Exchange Online tenant, retrieve the mailbox details for the users listed in the CSV, and export the data to Users_LEDN_Report.csv in the same directory

4. **Check the Output**:
   - The results will be saved in Users_LEDN_Report.csv in the script directory.

## Troubleshooting

- **No CSV file to read**: Ensure the `Users.csv` file is present and correctly formatted.
- **Permission Issues**: Ensure that you have the necessary permissions to connect to Exchange Online.
- **Module Installation**: If the script fails to install the module, try manually installing it:
  ```powershell
  Install-Module -Name ExchangeOnlineManagement -Force
  ```

## Additional Notes

- This script is designed to be run in an environment with access to Exchange Online.
- Ensure proper permissions and administrative access to retrieve mailbox data for the listed users.


