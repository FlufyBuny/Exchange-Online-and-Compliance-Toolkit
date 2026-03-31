# Check if the ExchangeOnlineManagement module is installed, and install it if not
if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
    Write-Host "ExchangeOnlineManagement module not found. Installing now..."
    Install-Module -Name ExchangeOnlineManagement -Force -AllowClobber
    Import-Module ExchangeOnlineManagement
} else {
    Write-Host "ExchangeOnlineManagement module is already installed."
}

function Show-Menu {
    param (
        [string]$Title = 'Exchange Online Administration Menu'
    )
    Clear-Host
    Write-Host "================ $Title ================"

    Write-Host "1: List all mailboxes"
    Write-Host "2: Get mailbox details"
    Write-Host "3: Create new mailbox"
    Write-Host "4: Delete a mailbox"
    Write-Host "5: Turn on AutoExpandingArchiving for a mailbox"
    Write-Host "6: Force Archiving for a mailbox NOW"
    Write-Host "7: Check Rules for SaaS Alerts"
    Write-Host "8: Change Global Admin"
    Write-Host "Q: Quit"
}

function Connect-ToExchangeOnline {
    Connect-ExchangeOnline -ShowBanner:$false
}

function List-Mailboxes {
    Get-Mailbox -ResultSize Unlimited | Select DisplayName, UserPrincipalName, PrimarySmtpAddress
}

function Get-MailboxDetails {
    $mailbox = Read-Host "Enter the email address of the mailbox"
    Get-Mailbox $mailbox | Format-List
}

function Create-Mailbox {
    $userPrincipalName = Read-Host "Enter user principal name"
    $password = Read-Host "Enter password" -AsSecureString
    $displayName = Read-Host "Enter display name"
    New-Mailbox -UserPrincipalName $userPrincipalName -Password $password -DisplayName $displayName -ResetPasswordOnNextLogon $false
}

function Delete-Mailbox {
    $mailbox = Read-Host "Enter the email address of the mailbox to delete"
    Remove-Mailbox -Identity $mailbox -Confirm:$false
}

function Enable-Archiving {
    $mailbox = Read-Host "Enter the email address of the mailbox to enable AutoExpandingArchiving"
    Enable-Mailbox -Identity $mailbox -AutoExpandArchive
    Write-Host "AutoExpandingArchiving has been enabled for $mailbox"
}

function Configure-AutoArchiving {
    $mailbox = Read-Host "Enter the email address of the mailbox"
    Start-ManagedFolderAssistant -Identity $mailbox
    Write-Host "Auto-archiving Started NOW for $mailbox"
}

function Rules {
    $mailbox = Read-Host "Enter the email address of the mailbox"
    Get-InboxRule -Mailbox $mailbox | Select-Object -ExpandProperty description | Format-List
}

function Change-GA {
    Write-Host "Reconnecting to Exchange Online..."
    Connect-ExchangeOnline -ShowBanner:$false
    Write-Host "Reconnected to Exchange Online."
}

# Entry point of the script
Connect-ToExchangeOnline

do {
    Show-Menu
    $input = Read-Host "Please select an option"
    switch ($input) {
        '1' { List-Mailboxes }
        '2' { Get-MailboxDetails }
        '3' { Create-Mailbox }
        '4' { Delete-Mailbox }
        '5' { Enable-Archiving }
        '6' { Configure-AutoArchiving }
        '7' { Rules }
        '8' { Change-GA }
        'Q' { Write-Host "Exiting..."; break }
        default { Write-Host "Invalid option, please try again." }
    }
    Read-Host -Prompt "Press Enter to continue"
} while ($input -ne 'Q')
