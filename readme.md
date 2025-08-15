# Add-GroupToChat PowerShell Scripts

## Overview
This repository contains PowerShell scripts for managing Microsoft Teams group chats and Microsoft 365 (O365) group memberships using Microsoft Graph API. The primary script, `Add-Group2Chat.ps1`, allows you to add all members of an M365 group to a Teams group chat with app-only authentication, without requiring external modules.

## Features
- **Add all members of an M365 group to a Teams group chat**
- **App-only authentication** (client credentials flow)
- **No external dependencies** (uses only built-in `Invoke-RestMethod`)
- **Retry logic** for user, chat, and group lookups
- **Duplicate avoidance** (skips users already in the chat)
- **Configurable chat history sharing**
- **Handles group name or email input**
- **Rate limiting and error handling**

## Prerequisites
- PowerShell 5.1 or later
- Azure AD App Registration with the following Application permissions:
  - `ChatMember.ReadWrite.All`
  - `Chat.ReadBasic.All`
  - `Group.Read.All`
  - `GroupMember.Read.All`
- Admin consent granted for the above permissions
- A JSON file with your app credentials (see below)

## Setup
1. **Clone this repository**
2. **Create a secrets file** at `.secrets/graph-app.json` in the following format:
   ```json
   {
     "tenantId": "<your-tenant-guid>",
     "clientId": "<your-app-client-guid>",
     "clientSecret": "<your-client-secret>"
   }
   ```
3. **Run the script**
   ```powershell
   .\Add-GroupToChat\Add-Group2Chat.ps1
   ```
   The script will prompt for:
   - The UPN of a user in the target chat
   - The Teams group chat name (topic)
   - The Microsoft 365 group name or email address

## File Descriptions
- `Add-GroupToChat/Add-Group2Chat.ps1` — Main script to add group members to a Teams chat
- `Add-GroupToChat/Add-GroupToChat.ps1` — (Legacy/alternate version)
- `Add-GroupToChat/readme.md` — This file
- Other scripts in the root folder provide additional Teams, Graph, and utility functions

## Security
- **Do not commit secrets!** The `.secrets/graph-app.json` file should be gitignored.
- All authentication is handled securely via a local secrets file.

## Troubleshooting
- Ensure your app registration has the correct permissions and admin consent.
- If you encounter errors, check the error message for missing permissions or invalid input.
- The script will prompt up to 3 times for user, chat, or group if not found.

## License
MIT License

## Author
Dylan 

---

For questions or contributions, please open an issue or pull request.
