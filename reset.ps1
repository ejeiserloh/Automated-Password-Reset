# Load Active Directory module
Import-Module ActiveDirectory

# Title Text
Write-Host "$([char]27)[4mAuto Password Reset Script made by Evan$([char]27)[24m" -ForegroundColor Blue 

do {
# Ask for Employee Number
$employeeNumber = Read-Host "`nEnter Employee Number"

# Search Active Directory by Office attribute
$user = Get-ADUser -Filter "Office -eq '$employeeNumber'" -Properties sAMAccountName, displayName

if (-not $user) {
    Write-Host "No user found with Employee number $employeeNumber" -ForegroundColor Red
    Pause
    exit
}

$logonName = $user.sAMAccountName
$fullName = $user.displayName
$clipboard = Get-Clipboard    

Write-Host "User found: $($fullName)" -ForegroundColor Magenta
Write-Host "Logon Name: $logonName" -ForegroundColor Green

# ===== CONFIRM USER BEFORE RESET =====
do {
    $confirmation = Read-Host "Confirm reset for $fullName ($logonName)? (Y/N)"
} until ($confirmation -match "^[YyNn]$")

if ($confirmation -match "^[Nn]$") {
    Write-Host "Operation cancelled. No password reset performed." -ForegroundColor Yellow
    Pause
    exit
}

# ===== PASSWORD RESET SECTION =====
# Path to encrypted password file (created earlier with ConvertFrom-SecureString)
$SecureFilePath = "C:\Users\User\defaultpass.txt"   # <-- CHANGE if needed

if (-not (Test-Path $SecureFilePath)) {
    Write-Host "Secure password file not found at $SecureFilePath" -ForegroundColor Red
    Pause
    exit
}

try {
    # Read encrypted password and convert back to SecureString
    $DefaultPassword = Get-Content $SecureFilePath | ConvertTo-SecureString

    # Reset password
    Set-ADAccountPassword -Identity $logonName -NewPassword $DefaultPassword -Reset

    # Force password change at next logon
    Set-ADUser -Identity $logonName -ChangePasswordAtLogon $true

    # Unlock account if locked
    Unlock-ADAccount -Identity $logonName -ErrorAction SilentlyContinue

    Write-Host "Password reset successfully for $logonName" -ForegroundColor Green
}
catch {
    Write-Host "Password reset failed: $_" -ForegroundColor Red
    Pause
    exit
}


# Start Outlook
$outlook = New-Object -ComObject Outlook.Application

$mail = $outlook.CreateItem(0)  # 0 = MailItem

# Path to your Outlook template (.oft file)
$templatePath = "C:\Users\User\template.oft"   # <-- CHANGE THIS PATH

if (-not (Test-Path $templatePath)) {
    Write-Host "Template not found at $templatePath" -ForegroundColor Red
    Pause
    exit
}

# Open template
$mail = $outlook.CreateItemFromTemplate($templatePath)

# Replace placeholder in subject field
$mail.Subject = $mail.Subject -replace "{{LogonName}}", $fullName

# Replace placeholder in template
$mail.HTMLBody = $mail.HTMLBody -replace "{{LogonName}}", $logonName

# Replace recipient using clipboard content (assumes email address is copied to clipboard)
$mail.To = $clipboard

# Show email (does NOT auto-send to verify information is correct first)
$mail.Display()

Write-Host "Email ready to send." -ForegroundColor Cyan

    } while ($true)   # repeat indefinitely until user types 'exit' or closes window