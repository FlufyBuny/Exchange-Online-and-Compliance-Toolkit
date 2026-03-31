# Check if the ExchangeOnlineManagement module is installed, and install it if not
if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
    Write-Host "ExchangeOnlineManagement module not found. Installing now..."
    Install-Module -Name ExchangeOnlineManagement -Force -AllowClobber
    Import-Module ExchangeOnlineManagement
} else {
    Write-Host "ExchangeOnlineManagement module is already installed."
}

function Show-MainMenu {
    param (
        [string]$Title = 'Exchange Online Administration Menu (DELTA v0.2)'
    )
    Clear-Host
    Write-Host "================ $Title ================"
    Write-Host "1: Mailbox Management"
    Write-Host "2: Archiving Options"
    Write-Host "3: Mail User Contacts"
    Write-Host "4: Global Administration"
    Write-Host "Q: Quit"
}

function Show-MailboxMenu {
    Clear-Host
    Write-Host "===== Mailbox Management ====="
    Write-Host "1: List all mailboxes"
    Write-Host "2: Get mailbox details"
    Write-Host "3: Create new mailbox"
    Write-Host "4: Delete a mailbox"
    Write-Host "5: Delegate user to a mailbox"
    Write-Host "6: Remove delegate user from a mailbox"
    Write-Host "7: View Mailbox Rules"
    Write-Host "B: Back to Main Menu"
}

function Show-ArchivingMenu {
    Clear-Host
    Write-Host "===== Archiving Options ====="
    Write-Host "1: Enable Archiving"
    Write-Host "2: Enable AutoExpanding Archiving"
    Write-Host "3: Start Archiving Immediately"
    Write-Host "B: Back to Main Menu"
}

function Show-ContactsMenu {
    Clear-Host
    Write-Host "===== Mail User Contacts ====="
    Write-Host "1: Add a Mail User Contact"
    Write-Host "B: Back to Main Menu"
}

function Show-AdminMenu {
    Clear-Host
    Write-Host "===== Global Administration ====="
    Write-Host "1: Change Global Admin"
    Write-Host "B: Back to Main Menu"
}

function Connect-ToExchangeOnline {
    Clear-Host     
    Connect-ExchangeOnline -ShowBanner:$false
}

function List-Mailboxes {
    Get-Mailbox -ResultSize Unlimited | Select DisplayName, UserPrincipalName, PrimarySmtpAddress
    Read-Host -Prompt "Press Enter to continue"
}

function Get-MailboxDetails {
    $mailbox = Read-Host "Enter the email address of the mailbox"
    Get-Mailbox $mailbox | Format-List
    Read-Host -Prompt "Press Enter to continue"
}

function Create-Mailbox {
    $name = Read-Host "Enter the full name"
    $password = Read-Host "Enter password" -AsSecureString
    $alias = Read-Host "Enter alias"
    $firstName = Read-Host "Enter first name"
    $lastName = Read-Host "Enter last name"
    $primarySmtpAddress = Read-Host "Enter primary SMTP address"
    New-Mailbox -Alias $alias -Name $name -FirstName $firstName -LastName $lastName -DisplayName $name -MicrosoftOnlineServicesID $primarySmtpAddress -Password $password -ResetPasswordOnNextLogon $true
    Write-Host "$name Created!"
    Read-Host -Prompt "Press Enter to continue"
}

function Enable-Archiving {
    $mailbox = Read-Host "Enter the email address of the mailbox (Archive)"
    Enable-Mailbox -Identity $mailbox -Archive
    Write-Host "Archiving has been enabled for $mailbox"
    Read-Host -Prompt "Press Enter to continue"
}

function Add-MailUserContact {
    $mailboxusername = Read-Host "Enter Name of the user, first last"
    $extuser = Read-Host "Enter external username WITHOUT @domain"
    $extdomain = Read-Host "Enter external domain"
    $onmsdomain = Read-Host "Enter the .onmicrosoft domain name"
    $extmailbox = "$extuser@$extdomain"
    New-MailUser -Name "$mailboxusername" -ExternalEmailAddress $extmailbox -MicrosoftOnlineServicesID "$extuser@$onmsdomain"
    Read-Host -Prompt "Press Enter to continue"
}

function Rules {
    Clear-Host   
    $mailbox = Read-Host "Enter the email address of the mailbox"
    Get-InboxRule -Mailbox $mailbox | Select Name, Identity, Description | fl
    Read-Host -Prompt "Press Enter to continue"
}

# Entry point of the script
Connect-ToExchangeOnline

do {
    Show-MainMenu
    $input = Read-Host "Please select an option"
    switch ($input) {
        '1' {
            do {
                Show-MailboxMenu
                $subInput = Read-Host "Select an option"
                if ($subInput -eq 'B') { break }
                switch ($subInput) {
                    '1' { List-Mailboxes }
                    '2' { Get-MailboxDetails }
                    '3' { Create-Mailbox }
                    '7' { Rules }
                }
            } while ($true)
        }
        '2' {
            do {
                Show-ArchivingMenu
                $subInput = Read-Host "Select an option"
                if ($subInput -eq 'B') { break }
                switch ($subInput) {
                    '1' { Enable-Archiving }
                }
            } while ($true)
        }
        '3' {
            do {
                Show-ContactsMenu
                $subInput = Read-Host "Select an option"
                if ($subInput -eq 'B') { break }
                switch ($subInput) {
                    '1' { Add-MailUserContact }
                }
            } while ($true)
        }
        '4' {
            do {
                Show-AdminMenu
                $subInput = Read-Host "Select an option"
                if ($subInput -eq 'B') { break }
                switch ($subInput) {
                    '1' { Connect-ToExchangeOnline }
                }
            } while ($true)
        }
        'Q' {
            Clear-Host 
            Write-Host "Exiting..."; break 
        }
        default { Write-Host "Invalid option, please try again." }
    }
} while ($input -ne 'Q')
