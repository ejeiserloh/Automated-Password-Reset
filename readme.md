Powershell script I created in my role as IT Helpdesk Technician to help with password reset requests.

# How it works
*  Start by inputting employee number
*  Pull user account by the 'Office' field under the account in the connected Active-Directory Domain
*  Asks for confirmation to reset account using temporary specified password by company (for security reasons, encrypted to securestring to prevent plain text password in script)
*  Also unlocks account if locked, and forces password change at next logon
*  Local outlook will open with a template including user documentation and guide to change password
*  Replaces fields inside template named {{LogonName}} to account username in Active-Directory
*  Pulls email from clipboard which is copied from company email repository including employee personal emails

The reason I used the clipboard function is because personal emails are not saved inside domain accounts
