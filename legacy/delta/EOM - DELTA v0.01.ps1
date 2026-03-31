# Check if the ExchangeOnlineManagement module is installed, and install it if not
if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
    Write-Host "ExchangeOnlineManagement module not found. Installing now..."
    Install-Module -Name ExchangeOnlineManagement -Force -AllowClobber
    Import-Module ExchangeOnlineManagement
} else {
    Write-Host "ExchangeOnlineManagement module is already installed."
}

function Show-MainMenu {
    param ([string]$Title = 'Exchange Online Administration Menu DELTA v0.1')

    Clear-Host
    Write-Host "================ $Title ================"
    Write-Host "1: Mailbox Management"
    Write-Host "2: Archiving Options"
    Write-Host "3: Rules & Attributes"
    Write-Host "4: Admin & Configuration"
    Write-Host "Q: Quit"
}

function Show-MailboxMenu {
    param ([string]$Title = 'Mailbox Management')

    Clear-Host
    Write-Host "================ $Title ================"
    Write-Host "1: List all mailboxes"
    Write-Host "2: Get mailbox details"
    Write-Host "3: Create a new mailbox"
    Write-Host "4: Delete a mailbox"
    Write-Host "5: Delegate user to a mailbox"
    Write-Host "6: Remove delegate user from a mailbox"
    Write-Host "B: Back to Main Menu"
}

function Show-ArchivingMenu {
    param ([string]$Title = 'Archiving Options')

    Clear-Host
    Write-Host "================ $Title ================"
    Write-Host "1: Enable Archiving"
    Write-Host "2: Enable Auto-Expanding Archiving"
    Write-Host "3: Start Archiving Now"
    Write-Host "B: Back to Main Menu"
}

function Show-RulesAttributesMenu {
    param ([string]$Title = 'Rules & Attributes')

    Clear-Host
    Write-Host "================ $Title ================"
    Write-Host "1: View Mailbox Rules"
    Write-Host "2: View Mailbox Attributes"
    Write-Host "B: Back to Main Menu"
}

function Show-AdminMenu {
    param ([string]$Title = 'Admin & Configuration')

    Clear-Host
    Write-Host "================ $Title ================"
    Write-Host "1: Change Global Admin"
    Write-Host "2: Add Mail User Contact"
    Write-Host "B: Back to Main Menu"
}

# Function for Exchange Online Connection
function Connect-ToExchangeOnline {
    Clear-Host     
    Connect-ExchangeOnline -ShowBanner:$false
}

# Mailbox Management Functions
function List-Mailboxes { Get-Mailbox -ResultSize Unlimited | Select DisplayName, UserPrincipalName, PrimarySmtpAddress }
function Get-MailboxDetails { $mailbox = Read-Host "Enter email"; Get-Mailbox $mailbox | Format-List }
function Create-Mailbox { Write-Host "Function to create mailbox" }
function Delete-Mailbox { $mailbox = Read-Host "Enter email to delete"; Remove-Mailbox -Identity $mailbox -Confirm:$false }
function Delegate-Access { Write-Host "Function to delegate access" }
function Remove-Access { Write-Host "Function to remove access" }

# Archiving Functions
function Enable-Archiving { $mailbox = Read-Host "Enter email"; Enable-Mailbox -Identity $mailbox -Archive }
function Enable-ExpandingArchiving { $mailbox = Read-Host "Enter email"; Enable-Mailbox -Identity $mailbox -AutoExpandArchive }
function Configure-AutoArchiving { $mailbox = Read-Host "Enter email"; Start-ManagedFolderAssistant -Identity $mailbox }

# Rules & Attributes Functions
function View-Rules { $mailbox = Read-Host "Enter email"; Get-InboxRule -Mailbox $mailbox | Select Name, Identity, Description | Format-List }
function View-Attributes { $mailbox = Read-Host "Enter email"; Get-Recipient $mailbox | Format-List Name,RecipientType,RecipientTypeDetails,CustomAttribute1 }

# Admin Functions
function Change-GA { Connect-ExchangeOnline -ShowBanner:$false }
function Add-MailUserContact { Write-Host "Function to add mail user contact" }

# Handling Menus
function Handle-MailboxMenu {
    do {
        Show-MailboxMenu
        $choice = Read-Host "Select an option"
        switch ($choice) {
            '1' { List-Mailboxes }
            '2' { Get-MailboxDetails }
            '3' { Create-Mailbox }
            '4' { Delete-Mailbox }
            '5' { Delegate-Access }
            '6' { Remove-Access }
            'B' { return }
            default { Write-Host "Invalid option, try again." }
        }
        Read-Host "Press Enter to continue"
    } while ($choice -ne 'B')
}

function Handle-ArchivingMenu {
    do {
        Show-ArchivingMenu
        $choice = Read-Host "Select an option"
        switch ($choice) {
            '1' { Enable-Archiving }
            '2' { Enable-ExpandingArchiving }
            '3' { Configure-AutoArchiving }
            'B' { return }
            default { Write-Host "Invalid option, try again." }
        }
        Read-Host "Press Enter to continue"
    } while ($choice -ne 'B')
}

function Handle-RulesAttributesMenu {
    do {
        Show-RulesAttributesMenu
        $choice = Read-Host "Select an option"
        switch ($choice) {
            '1' { View-Rules }
            '2' { View-Attributes }
            'B' { return }
            default { Write-Host "Invalid option, try again." }
        }
        Read-Host "Press Enter to continue"
    } while ($choice -ne 'B')
}

function Handle-AdminMenu {
    do {
        Show-AdminMenu
        $choice = Read-Host "Select an option"
        switch ($choice) {
            '1' { Change-GA }
            '2' { Add-MailUserContact }
            'B' { return }
            default { Write-Host "Invalid option, try again." }
        }
        Read-Host "Press Enter to continue"
    } while ($choice -ne 'B')
}

# Main Menu Handler
function Handle-MainMenu {
    do {
        Show-MainMenu
        $choice = Read-Host "Select an option"
        switch ($choice) {
            '1' { Handle-MailboxMenu }
            '2' { Handle-ArchivingMenu }
            '3' { Handle-RulesAttributesMenu }
            '4' { Handle-AdminMenu }
            'Q' { Write-Host "Exiting..."; break }
            default { Write-Host "Invalid option, try again." }
        }
        Read-Host "Press Enter to continue"
    } while ($choice -ne 'Q')
}

# Start the script
Connect-ToExchangeOnline
Handle-MainMenu
