# Check if the ExchangeOnlineManagement module is installed, and install it if not
if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
    Write-Host "ExchangeOnlineManagement module not found. Installing now..."
    Install-Module -Name ExchangeOnlineManagement -Force -AllowClobber
    Import-Module ExchangeOnlineManagement
} else {
    Write-Host "ExchangeOnlineManagement module is already installed."
}

# Global flag to indicate if user wants to quit
$global:DeltaQuit = $false

function Show-StartupMenu {
    Clear-Host
    Write-Host "========== DELTA v0.6 =========="
    Write-Host "1: Exchange Online Tasks"
    Write-Host "2: Compliance Center (Purview) Tasks"
    Write-Host "Q: Quit"
}

function Show-MainMenu {
    param (
        [string]$Title = 'Exchange Online Administration Menu (DELTA v0.6)'
    )
    Clear-Host
    Write-Host "================ $Title ================"
    Write-Host "1: Mailbox Management"
    Write-Host "2: Archiving Options"
    Write-Host "3: Mail User Contacts"
    Write-Host "4: Global Administration"
    Write-Host "Q: Quit to Mode Selection"
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
    Write-Host "Q: Quit Program"
    Write-Host "B: Back"
}

function Show-ArchivingMenu {
    Clear-Host
    Write-Host "===== Archiving Options ====="
    Write-Host "1: Enable Archiving"
    Write-Host "2: Enable AutoExpanding Archiving"
    Write-Host "3: Configure Auto-Archiving (Full Process)"
    Write-Host "Q: Quit Program"
    Write-Host "B: Back"
}

function Show-ContactsMenu {
    Clear-Host
    Write-Host "===== Mail User Contacts ====="
    Write-Host "1: Add a Mail User Contact"
    Write-Host "Q: Quit Program"
    Write-Host "B: Back"
}

function Show-AdminMenu {
    Clear-Host
    Write-Host "===== Global Administration ====="
    Write-Host "1: Change Global Admin"
    Write-Host "Q: Quit Program"
    Write-Host "B: Back"
}

function Show-ComplianceMenu {
    Clear-Host
    Write-Host "===== Compliance / Purview Tasks ====="
    Write-Host "1: Purge Emails by Subject"
    Write-Host "Q: Quit Program"
    Write-Host "B: Back"
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

function Delete-Mailbox {
    $mailbox = Read-Host "Enter the email address of the mailbox to delete"
    Remove-Mailbox -Identity $mailbox -Confirm:$false
    Write-Host "$mailbox DELETED!!!"
    Read-Host -Prompt "Press Enter to continue"
}

function Delegate-Access {
    $mailbox = Read-Host "Enter the email address of the mailbox"
    $delegate = Read-Host "Enter the email address of the delegate"
    Add-MailboxPermission -Identity $mailbox -User $delegate -AccessRights FullAccess -InheritanceType All
    Write-Host "Full Access has been given to $delegate on mailbox $mailbox"
    Read-Host -Prompt "Press Enter to continue"
}

function Remove-Access {
    $mailbox = Read-Host "Enter the email address of the mailbox"
    $delegate = Read-Host "Enter the email address of the delegate"
    Remove-MailboxPermission -Identity $mailbox -User $delegate -AccessRights FullAccess -InheritanceType All
    Write-Host "Full Access has been removed from $delegate on mailbox $mailbox"
    Read-Host -Prompt "Press Enter to continue"
}

function Enable-Archiving {
    $mailbox = Read-Host "Enter the email address of the mailbox"
    Enable-Mailbox -Identity $mailbox -Archive
    Write-Host "Archiving has been enabled for $mailbox"
    Read-Host -Prompt "Press Enter to continue"
}

function Enable-ExpandingArchiving {
    $mailbox = Read-Host "Enter the email address of the mailbox"
    Enable-Mailbox -Identity $mailbox -AutoExpandingArchive
    Write-Host "Auto-Expanding Archiving has been enabled for $mailbox"
    Read-Host -Prompt "Press Enter to continue"
}

function Configure-AutoArchiving {
    Clear-Host
    $mailbox = Read-Host "Enter the email address of the mailbox to configure archiving"

    $mbx = Get-Mailbox -Identity $mailbox -ErrorAction Stop
    if ($mbx.ArchiveStatus -eq 'Active') {
        Write-Host "Archive mailbox is already enabled for $mailbox"
    } else {
        Enable-Mailbox -Identity $mailbox -Archive
        Write-Host "Archive mailbox has been enabled for $mailbox"
    }

    Enable-Mailbox -Identity $mailbox -AutoExpandingArchive
    Write-Host "Auto-expanding archive has been enabled for $mailbox"

    Start-ManagedFolderAssistant -Identity $mailbox
    Write-Host "Managed Folder Assistant has been triggered for $mailbox."

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
    Get-InboxRule -Mailbox $mailbox | Select Name, Identity, Description | Format-List
    Read-Host -Prompt "Press Enter to continue"
}

function Change-GA {
    Clear-Host
    Write-Host "Reconnecting to Exchange Online..."
    Connect-ExchangeOnline -ShowBanner:$false
    Write-Host "Reconnected to Exchange Online."
    Read-Host -Prompt "Press Enter to continue"
}

function Purge-EmailsBySubject {
    Clear-Host
    Write-Host "Connecting to Microsoft Purview Compliance Center..."
    Connect-IPPSSession

    $subjectInput = Read-Host "Enter the exact subject line to search and purge"
    $escapedSubject = '"' + $subjectInput + '"'

    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $searchName = "Purge_Subject_$timestamp"

    Write-Host "Creating compliance search: $searchName"
    New-ComplianceSearch -Name $searchName -ExchangeLocation All -ContentMatchQuery "subject:$escapedSubject"

    Start-ComplianceSearch -Identity $searchName
    Write-Host "Search started. Waiting for completion..."

    do {
        Start-Sleep -Seconds 10
        $search = Get-ComplianceSearch -Identity $searchName
        $status = $search.Status
        Write-Host "Search status: $status"
    } while ($status -ne "Completed")

    Write-Host "`nSearch complete."
    Write-Host "Items found: $($search.ItemsFound)"
    Write-Host "Partially indexed items: $($search.ItemsPartiallyIndexed)`n"

    if ($search.ItemsFound -gt 0) {
        Write-Host "Starting purge action..."
        $purge = New-ComplianceSearchAction -SearchName $searchName -Purge -PurgeType SoftDelete
        $purgeId = $purge.Identity

        do {
            Start-Sleep -Seconds 10
            $purgeStatus = Get-ComplianceSearchAction -Identity $purgeId
            Write-Host "Purge status: $($purgeStatus.Status)"
        } while ($purgeStatus.Status -ne "Completed")

        Write-Host "`nPurge complete.`n"
        $purgeStatus.Results | Out-String | Write-Host
    }
    else {
        Write-Host "No items found matching subject. Nothing purged."
    }

    Read-Host -Prompt "Press Enter to continue"
}

# =========================
# MAIN SCRIPT ENTRY POINT
# =========================

do {
    Show-StartupMenu
    $mode = Read-Host "Choose a mode"

    switch ($mode) {
        '1' {
            Connect-ToExchangeOnline

            do {
                Show-MainMenu
                $input = Read-Host "Select an option"
                if ($input -eq 'Q') {
                    $global:DeltaQuit = $true
                    break
                }
                switch ($input) {
                    '1' {
                        do {
                            Show-MailboxMenu
                            $subInput = Read-Host "Select an option"
                            if ($subInput -eq 'Q') { $global:DeltaQuit = $true; break }
                            if ($subInput -eq 'B') { break }
                            switch ($subInput) {
                                '1' { List-Mailboxes }
                                '2' { Get-MailboxDetails }
                                '3' { Create-Mailbox }
                                '4' { Delete-Mailbox }
                                '5' { Delegate-Access }
                                '6' { Remove-Access }
                                '7' { Rules }
                            }
                            if ($global:DeltaQuit) { break }
                        } while ($true)
                    }
                    '2' {
                        do {
                            Show-ArchivingMenu
                            $subInput = Read-Host "Select an option"
                            if ($subInput -eq 'Q') { $global:DeltaQuit = $true; break }
                            if ($subInput -eq 'B') { break }
                            switch ($subInput) {
                                '1' { Enable-Archiving }
                                '2' { Enable-ExpandingArchiving }
                                '3' { Configure-AutoArchiving }
                            }
                            if ($global:DeltaQuit) { break }
                        } while ($true)
                    }
                    '3' {
                        do {
                            Show-ContactsMenu
                            $subInput = Read-Host "Select an option"
                            if ($subInput -eq 'Q') { $global:DeltaQuit = $true; break }
                            if ($subInput -eq 'B') { break }
                            switch ($subInput) {
                                '1' { Add-MailUserContact }
                            }
                            if ($global:DeltaQuit) { break }
                        } while ($true)
                    }
                    '4' {
                        do {
                            Show-AdminMenu
                            $subInput = Read-Host "Select an option"
                            if ($subInput -eq 'Q') { $global:DeltaQuit = $true; break }
                            if ($subInput -eq 'B') { break }
                            switch ($subInput) {
                                '1' { Change-GA }
                            }
                            if ($global:DeltaQuit) { break }
                        } while ($true)
                    }
                    default {
                        Write-Host "Invalid option."
                    }
                }
                if ($global:DeltaQuit) { break }
            } while ($true)
        }
        '2' {
            Connect-IPPSSession
            do {
                Show-ComplianceMenu
                $subChoice = Read-Host "Select an option"
                if ($subChoice -eq 'Q') { $global:DeltaQuit = $true; break }
                if ($subChoice -eq 'B') { break }
                switch ($subChoice) {
                    '1' { Purge-EmailsBySubject }
                    default { Write-Host "Invalid option." }
                }
                if ($global:DeltaQuit) { break }
            } while ($true)
        }
        'Q' {
            $global:DeltaQuit = $true
            break
        }
        default {
            Write-Host "Invalid selection."
        }
    }

    if ($global:DeltaQuit) {
        Clear-Host
        Write-Host "Goodbye!"
        break
    }

} while ($true)
