<img width="550" height="453" alt="epsilon_dark" src="https://github.com/user-attachments/assets/3066fbff-b612-43fd-8d53-392bf3210ca0" />

## Changelog
See [CHANGELOG.md](./CHANGELOG.md) for version history.

---

## What's New in v1.2

- Added **Calendar Permissions module**
  - Grant, update, and remove calendar access
  - View existing calendar permissions
  - Built-in access level selection (Reviewer, Editor, Owner, etc.)
  - Automatic detection and updating of existing permissions  

- Improved **delegation workflows**
  - Streamlined assistant / executive calendar management scenarios  

- Enhanced **permission handling**
  - Prevents duplicate permission errors by updating existing entries  

- Continued improvements to **script stability and usability**

---

## Project

**EPSILON – Exchange Online & Compliance Toolkit**

EPSILON is a PowerShell-based administrative toolkit designed to simplify and streamline common tasks in Microsoft 365, specifically across Exchange Online and the Compliance (Purview) Center.

Built with MSP workflows in mind, EPSILON provides a menu-driven interface for fast, repeatable operations without needing to remember complex PowerShell commands.

---

## Project Structure

This repository includes three deployment formats:

### Windows PowerShell Script
- `EOM - v1.2.ps1`  
- Full interactive menu for Exchange Online & Compliance tasks  

### macOS PowerShell Script
- `EOM - v1.2 - MAC.ps1`  
- Compatible with PowerShell Core (pwsh) on macOS  

### Executable Version
- `EOM - v1.2.exe`  
- Packaged version of the PowerShell script for simplified execution  
- **Note:** This is a wrapped `.ps1`, not a compiled binary  

---

## Features

### Exchange Online
- List mailboxes  
- View mailbox details  
- Create mailboxes  
- Remove mailboxes  
- Enable archive mailbox  
- Enable auto-expanding archive  
- Trigger Managed Folder Assistant  
- View inbox rules  
- Remove inbox rules  

### Calendar & Delegation (New in v1.2)
- Grant calendar access (Reviewer, Editor, Owner, etc.)  
- Modify existing calendar permissions  
- Remove calendar access  
- View calendar permissions  
- Built-in access level selection  
- Automatic update of existing permissions  

### Compliance (Purview)
- Create Compliance Searches (by subject)  
- Start and monitor searches  
- Purge emails from tenant  
- Status tracking and completion loops  

---

## Requirements

- PowerShell 5.1+ (Windows) or PowerShell Core (macOS/Linux)

### Microsoft Modules
- ExchangeOnlineManagement  
- Microsoft.Graph *(optional depending on usage)*  

If modules are missing, EPSILON can prompt or install them automatically (depending on version).

---

## Permissions Required

- Exchange Admin or Global Admin  
- Compliance Admin *(for Purview features)*  

---

## Usage

### Windows (PS1)
```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File ".\EOM - v1.2.ps1"
```

### macOS (PowerShell Core)
```bash
pwsh ./EOM - v1.2 - MAC.ps1
```

### Executable
```
EOM - v1.2.exe
```

---

## Execution Policy / Security

If you encounter script blocking:

```powershell
Set-ExecutionPolicy Bypass -Scope Process -Force
```

This allows the script to run without permanently lowering system security.

---

## Notes

- The EXE version is a wrapped PowerShell script and may still trigger security warnings  
- Some features (like Compliance purge) can take time — EPSILON includes built-in progress loops  
- Managed Folder Assistant errors (RPC issues) are typically server-side and not caused by the script  

---

## Known Issues

- `Start-ManagedFolderAssistant` may return:

  ```
  RPC Error -2147220992
  ```

  This is a Microsoft service-side issue and usually resolves on retry.

- Inbox rule creation timestamps are not directly exposed via standard cmdlets  

- Calendar permissions do not automatically configure full Outlook delegate behavior (meeting handling requires additional configuration)

---

## Roadmap / Future Enhancements

- CSV reporting / export options  
- GUI version  
- Enhanced rule auditing (creation metadata if available)  
- Better error handling and retry logic  
- Integration with Microsoft Graph for deeper insights  

---

## Contributing

Contributions, improvements, and ideas are welcome.

If you’re an MSP or admin using this in production, feel free to submit:

- Feature requests  
- Bug reports  
- Enhancements  

---

## License

This project is provided **as-is** for administrative use.

Customize freely for internal or client environments.

Developed for real-world MSP operations to reduce friction and increase efficiency in Microsoft 365 administration.

---

![Version](https://img.shields.io/badge/version-v1.2-blue)
![Platform](https://img.shields.io/badge/platform-Windows%20%7C%20macOS-lightgrey)
![PowerShell](https://img.shields.io/badge/powershell-5.1%2B%20%7C%20Core-blue)
