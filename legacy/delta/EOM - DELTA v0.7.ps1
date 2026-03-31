ExchangeOnlineManagement - DELTA v0.7

# Check if the ExchangeOnlineManagement module is installed, and install it if not
if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
    Write-Host "ExchangeOnlineManagement module not found. Installing now..."
    Install-Module -Name ExchangeOnlineManagement -Force -AllowClobber
    Import-Module ExchangeOnlineManagement
} else {
    Write-Host "ExchangeOnlineManagement module is already installed."
}

# Global flags to indicate session state and quit intent
$global:DeltaQuit = $false
$global:ExchangeConnected = $false
$global:ComplianceConnected = $false

# ----------------------------
# UI: Menus
# ----------------------------
function Show-StartupMenu {
    Clear-Host
    Write-Host "========== DELTA v0.7 =========="
    Write-Host "1: Exchange Online Tasks"
    Write-Host "2: Compliance Center (Purview) Tasks"
    Write-Host "Q: Quit"
}

function Show-MainMenu {
    param (
        [string]$Title = 'Exchange Online Administration Menu (DELTA v0.7)'
    )
    Clear-Host
    Write-Host "================ $Title ================"
    Write-Host "1: Mailbox Management"
    Write-Host "2: Archiving"
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
    Write-Host "===== Archiving ====="
    Write-Host "1: Enable Archive"
    Write-Host "2: Enable Archive + Auto-Expanding"
    Write-Host "3: Configure Auto-Archiving (Full Process: enable -> autoexpand -> run MFA with retries)"
    Write-Host "4: View Archive Status"
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
    Write-Host "1: Change Global Admin (reconnect)"
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

# ----------------------------
# Connection helpers
# ----------------------------
function Test-ExchangeConnection {
    try {
        # Try a lightweight command to see if session works
        Get-Mailbox -ResultSize 1 -ErrorAction Stop | Out-Null
        return $true
    } catch {
        return $false
    }
}

function Connect-ToExchangeOnline {
    if ($global:ExchangeConnected -and (Test-ExchangeConnection)) {
        Write-Host "Already connected to Exchange Online."
        return
    }

    Clear-Host
    Write-Host "Connecting to Exchange Online..."
    try {
        Connect-ExchangeOnline -ShowBanner:$false
        # quick test
        if (Test-ExchangeConnection) {
            $global:ExchangeConnected = $true
            Write-Host "Connected to Exchange Online."
        } else {
            $global:ExchangeConnected = $false
            Write-Host "Connected but verification failed. Some commands may error."
        }
    } catch {
        $global:ExchangeConnected = $false
        Write-Host "Failed to connect to Exchange Online:`n$($_.Exception.Message)"
    }
    Read-Host -Prompt "Press Enter to continue"
}

function Test-ComplianceConnection {
    try {
        # lightweight compliance command
        Get-ComplianceSearch -ErrorAction Stop | Out-Null
        return $true
    } catch {
        return $false
    }
}

function Connect-ToCompliance {
    if ($global:ComplianceConnected -and (Test-ComplianceConnection)) {
        Write-Host "Already connected to Compliance (IPPSSession)."
        return
    }

    Clear-Host
    Write-Host "Connecting to Microsoft Purview (Compliance) via Connect-IPPSSession ..."
    try {
        Connect-IPPSSession
        if (Test-ComplianceConnection) {
            $global:ComplianceConnected = $true
            Write-Host "Connected to Compliance (Purview)."
        } else {
            $global:ComplianceConnected = $false
            Write-Host "Connected but verification failed. Some compliance commands may error."
        }
    } catch {
        $global:ComplianceConnected = $false
        Write-Host "Failed to connect to Compliance:`n$($_.Exception.Message)"
    }
    Read-Host -Prompt "Press Enter to continue"
}

# ----------------------------
# Exchange: Mailbox functions
# ----------------------------
function List-Mailboxes {
    Get-Mailbox -ResultSize Unlimited | Select DisplayName, UserPrincipalName, PrimarySmtpAddress
    Read-Host -Prompt "Press Enter to continue"
}

function Get-MailboxDetails {
    $mailbox = Read-Host "Enter the email address of the mailbox"
    Get-Mailbox -Identity $mailbox | Format-List
    Read-Host -Prompt "Press Enter to continue"
}

function Create-Mailbox {
    $name = Read-Host "Enter the full name"
    $password = Read-Host "Enter password" -AsSecureString
    $alias = Read-Host "Enter alias"
    $firstName = Read-Host "Enter first name"
    $lastName = Read-Host "Enter last name"
    $primarySmtpAddress = Read-Host "Enter primary SMTP address"
    try {
        New-Mailbox -Alias $alias -Name $name -FirstName $firstName -LastName $lastName -DisplayName $name -MicrosoftOnlineServicesID $primarySmtpAddress -Password $password -ResetPasswordOnNextLogon $true -ErrorAction Stop
        Write-Host "$name Created!"
    } catch {
        Write-Host "Error creating mailbox: $($_.Exception.Message)"
    }
    Read-Host -Prompt "Press Enter to continue"
}

function Delete-Mailbox {
    $mailbox = Read-Host "Enter the email address of the mailbox to delete"
    try {
        Remove-Mailbox -Identity $mailbox -Confirm:$false -ErrorAction Stop
        Write-Host "$mailbox DELETED!!!"
    } catch {
        Write-Host "Error deleting mailbox: $($_.Exception.Message)"
    }
    Read-Host -Prompt "Press Enter to continue"
}

function Delegate-Access {
    $mailbox = Read-Host "Enter the email address of the mailbox"
    $delegate = Read-Host "Enter the email address of the delegate"
    try {
        Add-MailboxPermission -Identity $mailbox -User $delegate -AccessRights FullAccess -InheritanceType All -ErrorAction Stop
        Write-Host "Full Access has been given to $delegate on mailbox $mailbox"
    } catch {
        Write-Host "Error adding mailbox permission: $($_.Exception.Message)"
    }
    Read-Host -Prompt "Press Enter to continue"
}

function Remove-Access {
    $mailbox = Read-Host "Enter the email address of the mailbox"
    $delegate = Read-Host "Enter the email address of the delegate"
    try {
        Remove-MailboxPermission -Identity $mailbox -User $delegate -AccessRights FullAccess -InheritanceType All -ErrorAction Stop
        Write-Host "Full Access has been removed from $delegate on mailbox $mailbox"
    } catch {
        Write-Host "Error removing mailbox permission: $($_.Exception.Message)"
    }
    Read-Host -Prompt "Press Enter to continue"
}

function Rules {
    Clear-Host
    $mailbox = Read-Host "Enter the email address of the mailbox"
    try {
        Get-InboxRule -Mailbox $mailbox | Select Name, Identity, Description | Format-List
    } catch {
        Write-Host "Error getting inbox rules: $($_.Exception.Message)"
    }
    Read-Host -Prompt "Press Enter to continue"
}

# ----------------------------
# Exchange: Archiving functions & submenu
# ----------------------------
function Enable-Archiving {
    $mailbox = Read-Host "Enter the email address of the mailbox"
    try {
        Enable-Mailbox -Identity $mailbox -Archive -ErrorAction Stop
        Write-Host "Archiving has been enabled for $mailbox"
    } catch {
        Write-Host "Error enabling archive: $($_.Exception.Message)"
    }
    Read-Host -Prompt "Press Enter to continue"
}

function Enable-Archiving-AutoExpand {
    $mailbox = Read-Host "Enter the email address of the mailbox"
    try {
        Enable-Mailbox -Identity $mailbox -Archive -ErrorAction Stop
    } catch { 
        # ignore if already enabled, we'll check below
    }
    try {
        Enable-Mailbox -Identity $mailbox -AutoExpandingArchive -ErrorAction Stop
        Write-Host "Auto-Expanding archive enabled for $mailbox"
    } catch {
        Write-Host "Error enabling Auto-Expanding archive: $($_.Exception.Message)"
    }
    Read-Host -Prompt "Press Enter to continue"
}

function Get-ArchiveStatus {
    $mailbox = Read-Host "Enter the email address of the mailbox to view archive status"
    try {
        $mbx = Get-Mailbox -Identity $mailbox -ErrorAction Stop
        Write-Host "ArchiveStatus: $($mbx.ArchiveStatus)"
        if ($mbx.ArchiveStatus -eq 'Active') {
            Write-Host "Archive GUID: $($mbx.ArchiveGuid)"
        }
    } catch {
        Write-Host "Error retrieving mailbox: $($_.Exception.Message)"
    }
    Read-Host -Prompt "Press Enter to continue"
}

function Configure-AutoArchiving-Full {
    Clear-Host
    $mailbox = Read-Host "Enter the email address of the mailbox to configure archiving (full process)"

    # Step 1: Enable archive if needed
    try {
        $mbx = Get-Mailbox -Identity $mailbox -ErrorAction Stop
    } catch {
        Write-Host "Mailbox not found or error: $($_.Exception.Message)"
        Read-Host -Prompt "Press Enter to continue"
        return
    }

    if ($mbx.ArchiveStatus -ne 'Active') {
        Write-Host "Archive not active. Enabling archive for $mailbox..."
        try {
            Enable-Mailbox -Identity $mailbox -Archive -ErrorAction Stop
            Write-Host "Enable-Mailbox issued. Waiting for archive provisioning..."
        } catch {
            Write-Host "Enable-Mailbox returned an error: $($_.Exception.Message)"
            Read-Host -Prompt "Press Enter to continue"
            return
        }
    } else {
        Write-Host "Archive already enabled for $mailbox."
    }

    # Step 2: Wait for archive to be reported active (poll)
    $waitSeconds = 0
    $maxWait = 300  # total seconds to wait (5 minutes)
    $pollInterval = 5
    do {
        Start-Sleep -Seconds $pollInterval
        $waitSeconds += $pollInterval
        try {
            $mbx = Get-Mailbox -Identity $mailbox -ErrorAction Stop
        } catch {
            Write-Host "Error checking mailbox status: $($_.Exception.Message)"
            break
        }
        if ($mbx.ArchiveStatus -eq 'Active') { break }
        Write-Host "Waiting for archive provisioning... ($waitSeconds seconds elapsed)"
    } while ($waitSeconds -lt $maxWait)

    if ($mbx.ArchiveStatus -ne 'Active') {
        Write-Host "Archive provisioning did not complete within $maxWait seconds. You can try again later."
    } else {
        Write-Host "Archive is active for $mailbox."
    }

    # Step 3: Enable AutoExpandingArchive (best-effort)
    try {
        Enable-Mailbox -Identity $mailbox -AutoExpandingArchive -ErrorAction Stop
        Write-Host "Auto-expanding archive enabled (or request issued) for $mailbox."
    } catch {
        Write-Host "Auto-expanding archive request failed or not available: $($_.Exception.Message)"
    }

    # Step 4: Attempt to run Managed Folder Assistant with retries
    $maxRetries = 5
    $attempt = 0
    $successMFA = $false
    $delaySeconds = 5

    while ($attempt -lt $maxRetries) {
        $attempt++
        try {
            Start-Sleep -Seconds 1  # small pause before call
            Start-ManagedFolderAssistant -Identity $mailbox -ErrorAction Stop
            Write-Host "Start-ManagedFolderAssistant succeeded on attempt $attempt."
            $successMFA = $true
            break
        } catch {
            $errMsg = $_.Exception.Message
            Write-Host "Attempt $attempt : Start-ManagedFolderAssistant failed: $errMsg"
            if ($attempt -lt $maxRetries) {
                Write-Host "Waiting $delaySeconds seconds before next attempt..."
                Start-Sleep -Seconds $delaySeconds
                $delaySeconds = [math]::Min($delaySeconds * 2, 60)  # exponential backoff, cap at 60s
            }
        }
    }

    if (-not $successMFA) {
        Write-Host ""
        Write-Host "Managed Folder Assistant could not be triggered after $maxRetries attempts."
        Write-Host "Common reasons: backend busy, mailbox just provisioned, server-side throttling, mailbox move in progress."
        Write-Host "Recommendation: wait 10-30 minutes and try 'Configure Auto-Archiving (Full Process)' again, or let Microsoft run MFA automatically (they run it routinely)."
    }

    Read-Host -Prompt "Press Enter to continue"
}

# ----------------------------
# Exchange: Contacts & Admin
# ----------------------------
function Add-MailUserContact {
    $mailboxusername = Read-Host "Enter Name of the user, first last"
    $extuser = Read-Host "Enter external username WITHOUT @domain"
    $extdomain = Read-Host "Enter external domain"
    $onmsdomain = Read-Host "Enter the .onmicrosoft domain name (e.g., contoso.onmicrosoft.com)"
    $extmailbox = "$extuser@$extdomain"
    try {
        New-MailUser -Name "$mailboxusername" -ExternalEmailAddress $extmailbox -MicrosoftOnlineServicesID "$extuser@$onmsdomain" -ErrorAction Stop
        Write-Host "Mail user contact created: $mailboxusername -> $extmailbox"
    } catch {
        Write-Host "Error creating mail user contact: $($_.Exception.Message)"
    }
    Read-Host -Prompt "Press Enter to continue"
}

function Change-GA {
    Clear-Host
    Write-Host "Reconnecting to Exchange Online..."
    try {
        Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop
        $global:ExchangeConnected = $true
        Write-Host "Reconnected to Exchange Online."
    } catch {
        Write-Host "Error reconnecting: $($_.Exception.Message)"
    }
    Read-Host -Prompt "Press Enter to continue"
}

# ----------------------------
# Compliance: Purge by subject
# ----------------------------
function Purge-EmailsBySubject {
    Clear-Host
    Write-Host "Connecting to Microsoft Purview Compliance Center..."
    if (-not $global:ComplianceConnected) {
        try {
            Connect-IPPSSession
            $global:ComplianceConnected = $true
        } catch {
            Write-Host "Failed to connect to Compliance: $($_.Exception.Message)"
            Read-Host -Prompt "Press Enter to continue"
            return
        }
    }

    $subjectInput = Read-Host "Enter the exact subject line to search and purge"
    $escapedSubject = '"' + $subjectInput + '"'

    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $searchName = "Purge_Subject_$timestamp"

    Write-Host "Creating compliance search: $searchName"
    try {
        New-ComplianceSearch -Name $searchName -ExchangeLocation All -ContentMatchQuery "subject:$escapedSubject" -ErrorAction Stop
        Start-ComplianceSearch -Identity $searchName -ErrorAction Stop
    } catch {
        Write-Host "Error creating/starting compliance search: $($_.Exception.Message)"
        Read-Host -Prompt "Press Enter to continue"
        return
    }

    Write-Host "Search started. Waiting for completion..."
    do {
        Start-Sleep -Seconds 10
        $search = Get-ComplianceSearch -Identity $searchName -ErrorAction SilentlyContinue
        if (-not $search) { Write-Host "Waiting for search to be visible..." ; continue }
        $status = $search.Status
        Write-Host "Search status: $status"
    } while ($status -ne "Completed")

    Write-Host "`nSearch complete."
    Write-Host "Items found: $($search.ItemsFound)"
    Write-Host "Partially indexed items: $($search.ItemsPartiallyIndexed)`n"

    if ($search.ItemsFound -gt 0) {
        Write-Host "Starting purge action..."
        try {
            $purge = New-ComplianceSearchAction -SearchName $searchName -Purge -PurgeType SoftDelete -ErrorAction Stop
        } catch {
            Write-Host "Error starting purge action: $($_.Exception.Message)"
            Read-Host -Prompt "Press Enter to continue"
            return
        }

        $purgeId = $purge.Identity
        do {
            Start-Sleep -Seconds 10
            $purgeStatus = Get-ComplianceSearchAction -Identity $purgeId -ErrorAction SilentlyContinue
            if (-not $purgeStatus) { Write-Host "Waiting for purge action to be visible..." ; continue }
            Write-Host "Purge status: $($purgeStatus.Status)"
        } while ($purgeStatus.Status -ne "Completed")

        Write-Host "`nPurge complete.`n"
        $purgeStatus.Results | Out-String | Write-Host
    } else {
        Write-Host "No items found matching subject. Nothing purged."
    }

    Read-Host -Prompt "Press Enter to continue"
}

# ----------------------------
# MAIN: mode selection and loops
# ----------------------------
do {
    Show-StartupMenu
    $mode = Read-Host "Choose a mode"

    switch ($mode) {
        '1' {
            # Exchange mode
            if (-not $global:ExchangeConnected) {
                Connect-ToExchangeOnline
            } else {
                Write-Host "Using existing Exchange session."
            }

            do {
                Show-MainMenu
                $input = Read-Host "Select an option"
                if ($input -eq 'Q') {
                    $global:DeltaQuit = $true
                    break
                }
                switch ($input) {
                    '1' {
                        # Mailbox management
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
                                default { Write-Host "Invalid option." }
                            }
                            if ($global:DeltaQuit) { break }
                        } while ($true)
                    }
                    '2' {
                        # Archiving submenu
                        do {
                            Show-ArchivingMenu
                            $subInput = Read-Host "Select an option"
                            if ($subInput -eq 'Q') { $global:DeltaQuit = $true; break }
                            if ($subInput -eq 'B') { break }
                            switch ($subInput) {
                                '1' { Enable-Archiving }
                                '2' { Enable-Archiving-AutoExpand }
                                '3' { Configure-AutoArchiving-Full }
                                '4' { Get-ArchiveStatus }
                                default { Write-Host "Invalid option." }
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
                                default { Write-Host "Invalid option." }
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
                                default { Write-Host "Invalid option." }
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
            # Compliance mode
            if (-not $global:ComplianceConnected) {
                Connect-ToCompliance
            } else {
                Write-Host "Using existing Compliance session."
                Read-Host -Prompt "Press Enter to continue"
            }

            do {
                Show-ComplianceMenu
                $subChoice = Read-Host "Select an option"
                if ($subChoice -eq 'Q') { $global:DeltaQuit = $true; break }
                if ($subChoice -eq 'B') { break }
                switch ($subChoice) {
                    '1' { Purge-EmailsBySubject }
                    default { Write-Host "Invalid option." ; Read-Host -Prompt "Press Enter to continue" }
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
