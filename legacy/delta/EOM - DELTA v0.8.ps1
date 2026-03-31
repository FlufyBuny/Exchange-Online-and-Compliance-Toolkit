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
    Write-Host "========== DELTA v0.8 =========="
    Write-Host "1: Exchange Online Tasks"
    Write-Host "2: Compliance Center (Purview) Tasks"
    Write-Host "Q: Quit"
}

function Show-MainMenu {
    param (
        [string]$Title = 'Exchange Online Administration Menu (DELTA v0.8)'
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
    Clear-Host
    Get-Mailbox -ResultSize Unlimited | Select DisplayName, UserPrincipalName, PrimarySmtpAddress
    Read-Host -Prompt "Press Enter to continue"
}

function Get-MailboxDetails {
    Clear-Host
    $mailbox = Read-Host "Enter the email address of the mailbox"
    Get-Mailbox $mailbox | Format-List
    Read-Host -Prompt "Press Enter to continue"
}

function Create-Mailbox {
    Clear-Host
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
    Clear-Host
    $mailbox = Read-Host "Enter the email address of the mailbox to delete"
    Remove-Mailbox -Identity $mailbox -Confirm:$false
    Write-Host "$mailbox DELETED!!!"
    Read-Host -Prompt "Press Enter to continue"
}

function Delegate-Access {
    Clear-Host
    $mailbox = Read-Host "Enter the email address of the mailbox"
    $delegate = Read-Host "Enter the email address of the delegate"
    Add-MailboxPermission -Identity $mailbox -User $delegate -AccessRights FullAccess -InheritanceType All
    Write-Host "Full Access has been given to $delegate on mailbox $mailbox"
    Read-Host -Prompt "Press Enter to continue"
}

function Remove-Access {
    Clear-Host
    $mailbox = Read-Host "Enter the email address of the mailbox"
    $delegate = Read-Host "Enter the email address of the delegate"
    Remove-MailboxPermission -Identity $mailbox -User $delegate -AccessRights FullAccess -InheritanceType All
    Write-Host "Full Access has been removed from $delegate to mailbox $mailbox"
    Read-Host -Prompt "Press Enter to continue"
}

function Enable-Archiving {
    Clear-Host
    $mailbox = Read-Host "Enter the email address of the mailbox"
    Enable-Mailbox -Identity $mailbox -Archive
    Write-Host "Archiving has been enabled for $mailbox"
    Read-Host -Prompt "Press Enter to continue"
}

function Enable-ExpandingArchiving {
    Clear-Host
    $mailbox = Read-Host "Enter the email address of the mailbox"
    Enable-Mailbox -Identity $mailbox -AutoExpandingArchive
    Write-Host "Auto-Expanding Archiving has been enabled for $mailbox"
    Read-Host -Prompt "Press Enter to continue"
}

function Configure-AutoArchiving {
    Clear-Host
    $mailbox = Read-Host "Enter the email address of the mailbox to configure archiving"

    # Validate mailbox exists
    try {
        $mbx = Get-Mailbox -Identity $mailbox -ErrorAction Stop
    } catch {
        Write-Host "Mailbox not found or error: $($_.Exception.Message)"
        Read-Host -Prompt "Press Enter to continue"
        return
    }

    # Step 1: Enable archive if needed
    if ($mbx.ArchiveStatus -ne 'Active') {
        Write-Host "Archive not active. Issuing Enable-Mailbox..."
        try {
            Enable-Mailbox -Identity $mailbox -Archive -ErrorAction Stop
            Write-Host "Enable-Mailbox command issued. Waiting for provisioning..."
        } catch {
            Write-Host "Enable-Mailbox failed: $($_.Exception.Message)"
            Read-Host -Prompt "Press Enter to continue"
            return
        }
    } else {
        Write-Host "Archive already active for $mailbox."
    }

    # Step 2: Enable Auto-Expanding Archive (best-effort)
    try {
        Enable-Mailbox -Identity $mailbox -AutoExpandingArchive -ErrorAction Stop
        Write-Host "Auto-expanding archive request issued for $mailbox."
    } catch {
        Write-Host "Auto-expanding archive request failed or not available: $($_.Exception.Message)"
    }

    # Step 3: Poll for archive status (wait up to configurable timeout)
    $pollInterval = 5
    $maxWait = 300   # seconds (5 minutes)
    $elapsed = 0
    do {
        Start-Sleep -Seconds $pollInterval
        $elapsed += $pollInterval
        try {
            $mbx = Get-Mailbox -Identity $mailbox -ErrorAction Stop
        } catch {
            Write-Host "Error checking mailbox status: $($_.Exception.Message)"
            break
        }
        if ($mbx.ArchiveStatus -eq 'Active') { break }
        Write-Host "Waiting for archive to appear... ($elapsed sec elapsed)"
    } while ($elapsed -lt $maxWait)

    if ($mbx.ArchiveStatus -ne 'Active') {
        Write-Host "Archive provisioning did not finish within $maxWait seconds. You can retry later."
    } else {
        Write-Host "Archive is active for $mailbox."
    }

    # Step 4: Attempt to run Managed Folder Assistant with retries & exponential backoff
    $maxAttempts = 6
    $attempt = 1
    $success = $false
    $delaySeconds = 10

    while (-not $success -and $attempt -le $maxAttempts) {
        try {
            Write-Host ("Attempt {0}: Starting Managed Folder Assistant for {1}" -f $attempt, $mailbox)
            Start-ManagedFolderAssistant -Identity $mailbox -ErrorAction Stop
            Write-Host "Managed Folder Assistant successfully triggered for $mailbox."
            $success = $true
        } catch {
            Write-Host ("Attempt {0} failed: {1}" -f $attempt, $_.Exception.Message) -ForegroundColor Yellow
            if ($attempt -lt $maxAttempts) {
                Write-Host ("Waiting {0} seconds before retry..." -f $delaySeconds)
                Start-Sleep -Seconds $delaySeconds
                # exponential backoff, cap at 120s
                $delaySeconds = [math]::Min($delaySeconds * 2, 120)
            }
            $attempt++
        }
    }

    if (-not $success) {
        Write-Host ""
        Write-Host "Managed Folder Assistant could not be started after $maxAttempts attempts."
        Write-Host "Common causes: backend provisioning in progress, mailbox move, or temporary service throttling."
        Write-Host "Recommendation: wait 10-30 minutes and try again, or let Microsoft run MFA (it runs automatically periodically)."
    }

    Read-Host -Prompt "Press Enter to continue"
}

function Add-MailUserContact {
    Clear-Host
    $mailboxusername = Read-Host "Enter Name of the user, first last"
    $extuser = Read-Host "Enter external username WITHOUT @domain"
    $extdomain = Read-Host "Enter external domain"
    $onmsdomain = Read-Host "Enter the .onmicrosoft domain name ex: cs365.onmicrosoft.com"
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

        Write-Host "`nPurge complete."
        # Ensure the results are fully printed before continuing
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
