<#
.SYNOPSIS
Adds all members of an O365 (M365) group to a Microsoft Teams group chat using Microsoft Graph (app-only authentication).

.DESCRIPTION
This script prompts for:
  - A **User Principal Name (UPN)** of a user who is in the target chat.
  - The **Teams group chat name** (chat topic) to identify the chat.
  - The **Microsoft 365 group name** (or email address) whose members you want to add to the chat.

The script then:
  1. Logs into Microsoft Graph with **app-only** permissions (client credentials flow). 
     - The client secret is read from a local file (to avoid storing secrets in the script).
  2. Finds the chat **ID** based on the provided chat name and user’s UPN.
  3. Finds the M365 group’s **ID** based on the provided group name (or email).
  4. Retrieves the members of that M365 group.
  5. Adds each member to the chat via the Graph API.

FEATURES:
  - **Secure secret handling:** Expects app credentials in a separate JSON file that can be .gitignored.
  - **No external modules:** Uses only built-in `Invoke-RestMethod`.
  - **Retry logic:** If the provided UPN, chat name, or group name is not found, the script will prompt up to 3 times to re-enter the correct value before aborting.
  - **Email input handling for group:** If you provide a group’s email address instead of its display name, the script will handle it by using the part before the "@" for lookup (and also attempts direct email match).
  - **Duplicate avoidance:** Skips adding users who are already in the chat.
  - **Visibility of history:** By default, new chat members **do not** get any chat history. This can be changed via a configuration variable.
  - **Rate limiting:** Implements retries with backoff if Graph API returns 429 (too many requests) or similar errors.

.NOTES
Permissions needed (as **Application** permissions in Azure AD App Registration):
  - ChatMember.ReadWrite.All  (to add chat members)
  - Chat.ReadBasic.All        (to read chat info like chat names for a user)
  - Group.Read.All            (to read group info)
  - GroupMember.Read.All      (to read group members)
You might also use broader permissions like Chat.ReadWrite.All or Directory.Read.All depending on your scenario.

Make sure to grant admin consent for these application permissions.
#>

# --------------------------- CONFIGURATION ---------------------------
# Path to the JSON file storing app credentials (tenant ID, client ID, client secret).
# Prepare this file manually with the format:
# {
#   "tenantId": "<your-tenant-guid>",
#   "clientId": "<your-app-client-guid>",
#   "clientSecret": "<your-client-secret>"
# }
$Script:SecretFilePath = Join-Path $PSScriptRoot ".secrets/graph-app.json"   # Adjust path as needed.

# Microsoft Graph endpoints and settings
$Script:GraphBase      = "https://graph.microsoft.com/v1.0"
$Script:TokenEndpoint  = "https://login.microsoftonline.com/{0}/oauth2/v2.0/token"  # {0} will be tenant ID
$Script:UserAgent      = "AddGroupMembersToChatScript/1.0"

# Behavior toggles
$Script:IncludeNestedGroupMembers = $false    # Set to $true to include nested group members (uses /transitiveMembers if true).
$Script:ShareHistoryMode          = "None"    # Options: "None" (no history), "All" (all history), or "Since:<ISO8601 timestamp>".
$Script:MaxRetries                = 6         # Max retry attempts for transient errors (HTTP 429/503).
$Script:InitialDelaySeconds       = 2         # Initial delay for backoff (doubles with each retry, up to 60s max).

# --------------------------- FUNCTION DEFINITIONS ---------------------------

# (1) Read-GraphAppSecret: Reads client credentials from the JSON file.
function Read-GraphAppSecret {
    [CmdletBinding()]
    param()

    if (-not (Test-Path -LiteralPath $Script:SecretFilePath)) {
        throw "App secret file not found at '$Script:SecretFilePath'. Please create the file with tenantId, clientId, clientSecret."
    }
    try {
        # Read and parse the JSON
        $secret = Get-Content -LiteralPath $Script:SecretFilePath -Raw | ConvertFrom-Json
    } catch {
        throw "Failed to read or parse the app secret JSON file. $_"
    }

    # Validate required fields
    foreach ($field in @("tenantId", "clientId", "clientSecret")) {
        if ([string]::IsNullOrEmpty($secret.$field)) {
            throw "The field '$field' is missing or empty in the app secret file."
        }
    }
    return $secret
}

# (2) Get-GraphToken: Obtains an OAuth 2.0 token using client credentials (app-only auth).
function Get-GraphToken {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [string] $TenantId,
        [Parameter(Mandatory)] [string] $ClientId,
        [Parameter(Mandatory)] [string] $ClientSecret
    )

    $tokenUrl = [string]::Format($Script:TokenEndpoint, $TenantId)
    $body     = @{
        client_id     = $ClientId
        client_secret = $ClientSecret
        grant_type    = "client_credentials"
        scope         = "https://graph.microsoft.com/.default"
    }
    try {
        $response = Invoke-RestMethod -Method POST -Uri $tokenUrl -Body $body -ContentType "application/x-www-form-urlencoded"
    } catch {
        throw "Failed to acquire token. $_"
    }
    if (-not $response.access_token) {
        throw "Token response did not contain an access_token."
    }
    return $response.access_token
}

# (3) Invoke-GraphRequest: Invokes a REST call to Microsoft Graph with the current token, handling retries for rate limiting.
function Invoke-GraphRequest {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [ValidateSet('GET','POST','PATCH','DELETE','PUT')] [string] $Method,
        [Parameter(Mandatory)] [string] $Uri,
        [Parameter()] $Body,
        [Parameter()] [Hashtable] $Headers,
        [Parameter()] [switch] $ExpectNoContent
    )

    # Prepare headers (include Authorization and any custom headers)
    $mergedHeaders = @{
        "Authorization" = "Bearer $($Script:AccessToken)"
        "Accept"        = "application/json"
        "User-Agent"    = $Script:UserAgent
    }
    if ($Headers) {
        foreach ($key in $Headers.Keys) {
            $mergedHeaders[$key] = $Headers[$key]
        }
    }

    # Execute with retry logic for transient errors (429, 503, 504)
    $attempt = 0
    $delay   = $Script:InitialDelaySeconds
    while ($true) {
        try {
            if ($PSBoundParameters.ContainsKey('Body')) {
                # If Body is provided, convert it to JSON (unless already a string)
                $jsonBody = if ($Body -is [string]) { $Body } else { $Body | ConvertTo-Json -Depth 10 }
                $result = Invoke-RestMethod -Method $Method -Uri $Uri -Headers $mergedHeaders -ContentType "application/json" -Body $jsonBody
            } else {
                $result = Invoke-RestMethod -Method $Method -Uri $Uri -Headers $mergedHeaders
            }
            # If a response is expected to have no content (e.g., 204 NoContent), treat a null result as success.
            if ($ExpectNoContent -and $null -eq $result) {
                return $true
            }
            return $result
        }
        catch {
            # Capture the HTTP status (if available) to check for transient errors
            $statusCode = $null
            $retryAfter = $null
            if ($_.Exception.Response) {
                try { $statusCode = [int]$_.Exception.Response.StatusCode } catch {}
                $retryAfter = $_.Exception.Response.Headers["Retry-After"]
            }

            if ($statusCode -in 429,503,504) {
                if ($attempt -ge $Script:MaxRetries) {
                    throw "Graph API request failed after $Script:MaxRetries retries (last status: $statusCode)."
                }
                $attempt++
                # Honor "Retry-After" header if present, otherwise use exponential backoff
                if ($retryAfter) {
                    Write-Host "Received retryable error $statusCode. Waiting $retryAfter seconds before retry #$attempt..." -ForegroundColor Yellow
                    Start-Sleep -Seconds ([int]$retryAfter)
                } else {
                    Write-Host "Received retryable error $statusCode. Waiting $delay seconds before retry #$attempt..." -ForegroundColor Yellow
                    Start-Sleep -Seconds $delay
                    # Exponential backoff for next attempt (capped at 60 seconds)
                    $delay = [Math]::Min([int][Math]::Ceiling($delay * 1.8), 60)
                }
                continue  # Retry the loop
            }
            # If not a transient error or max retries exceeded, propagate the error
            throw $_
        }
    }
}

# (4) Resolve-UserId: Resolves a user's Microsoft Entra (Azure AD) object ID from their UPN (email address).
function Resolve-UserId {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string] $Upn)
    # Use Graph to get the user by UPN; select only the id (object ID) to minimize payload.
    $uri = "$($Script:GraphBase)/users/$([uri]::EscapeDataString($Upn))?`$select=id"
    try {
        $user = Invoke-GraphRequest -Method GET -Uri $uri
    } catch {
        throw "User '$Upn' not found or not accessible."
    }
    return $user.id
}

# (5) Get-UserChatByTopic: Finds a group chat for a given user by the chat's topic (name).
function Get-UserChatByTopic {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string] $UserId,
        [Parameter(Mandatory)][string] $Topic
    )
    # List all chats for the user (Chat.ReadBasic.All required). We filter client-side for a matching topic.
    $uri = "$($Script:GraphBase)/users/$UserId/chats?`$select=id,topic,chatType"
    $chatsResponse = Invoke-GraphRequest -Method GET -Uri $uri
    $chats = $chatsResponse.value

    if (-not $chats) { return $null }

    # Filter to group chats with matching topic (case-insensitive contains for flexibility).
    $matches = $chats | Where-Object { $_.chatType -eq "group" -and $null -ne $_.topic -and $_.topic.ToLower().Contains($Topic.ToLower()) }
    if ($matches.Count -gt 1) {
        Write-Warning "Multiple chats found containing '$Topic'. Using the first match."
    }
    return $matches | Select-Object -First 1
}

# (6) Resolve-GroupId: Resolves a group's Azure AD object ID from its name or email.
function Resolve-GroupId {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string] $GroupNameOrEmail)
    # If input looks like an email (contains '@'), strip the domain to use the alias for displayName matching.
    $nameInput = $GroupNameOrEmail
    if ($GroupNameOrEmail -like "*@*") {
        # Try full email match first
        $filterEmail = $GroupNameOrEmail.Replace("'", "''")  # escape quotes for OData
        $uriEmail = "$($Script:GraphBase)/groups?`$select=id,displayName,mail&`$filter=mail eq '$filterEmail'"
        $respEmail = Invoke-GraphRequest -Method GET -Uri $uriEmail
        if ($respEmail.value.Count -gt 0) {
            return $respEmail.value[0].id  # return first match on email
        }
        # No direct email match – fall back to using the part before '@' as the name key
        $nameInput = $GroupNameOrEmail.Split("@")[0]
        Write-Host "No group found by email. Trying name '$nameInput'..." -ForegroundColor Yellow
    }
    # Escape single quotes in the name for OData
    $escapedName = $nameInput.Replace("'", "''")

    # Try exact match on displayName
    $uri = "$($Script:GraphBase)/groups?`$select=id,displayName,mail&`$filter=displayName eq '$escapedName'"
    $resp = Invoke-GraphRequest -Method GET -Uri $uri
    if ($resp.value.Count -eq 1) {
        return $resp.value[0].id
    }
    if ($resp.value.Count -gt 1) {
        Write-Warning "Multiple groups named '$nameInput'. Using the first match (ID: $($resp.value[0].id))."
        return $resp.value[0].id
    }

    # Fallback: partial match (startsWith) on displayName
    $uri2 = "$($Script:GraphBase)/groups?`$select=id,displayName,mail&`$filter=startsWith(displayName,'$escapedName')"
    $resp2 = Invoke-GraphRequest -Method GET -Uri $uri2
    if ($resp2.value.Count -ge 1) {
        Write-Warning "No exact match. Using first partial match: '$($resp2.value[0].displayName)' (ID: $($resp2.value[0].id))."
        return $resp2.value[0].id
    }

    throw "Group '$GroupNameOrEmail' not found."
}

# (7) Get-GroupUsers: Retrieves user members of the specified group (optionally including nested members).
function Get-GroupUsers {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string] $GroupId,
        [Parameter()][switch] $IncludeTransitive
    )

    # PowerShell 5.1-friendly way to set the endpoint
    $memberEndpoint = if ($IncludeTransitive.IsPresent) { 'transitiveMembers' } else { 'members' }
    if ([string]::IsNullOrWhiteSpace($memberEndpoint)) { $memberEndpoint = 'members' }

    # Use OData cast to return only users, so we can safely select userType, etc.
    # Docs: OData cast supported for transitiveMembers; members is the standard relationship.
    # - https://learn.microsoft.com/graph/api/group-list-transitivemembers (OData cast enabled)
    # - https://learn.microsoft.com/graph/api/group-list-members
    $baseUri = "$($Script:GraphBase)/groups/$GroupId/$memberEndpoint/microsoft.graph.user"
    $uri     = $baseUri + "?`$select=id,displayName,userPrincipalName,userType"

    Write-Verbose "GET $uri"

    # ConsistencyLevel=eventual is generally required for $count/$search/advanced filters.
    # It's harmless here; keep it if you later add filtering.
    $headers = @{ "ConsistencyLevel" = "eventual" }

    $results = @()
    do {
        $page = Invoke-GraphRequest -Method GET -Uri $uri -Headers $headers
        if ($page.value) { $results += $page.value }
        $uri = $page.'@odata.nextLink'
    } while ($uri)

    return $results
}

# (8) Get-ExistingChatMemberIds: Gets a set of user IDs already present in the chat (to avoid duplicate adds).
function Get-ExistingChatMemberIds {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string] $ChatId,
        [Parameter(Mandatory)][string] $UserId
    )

    $uri = "$($Script:GraphBase)/$UserId/chats/$ChatId/members"

    Write-Verbose "URI: $uri"

    $allMembers = @()
    do {
        $resp = Invoke-GraphRequest -Method GET -Uri $uri
        if ($resp.value) {
            $allMembers += $resp.value
        }
        $uri = $resp.'@odata.nextLink'
    } while ($uri)

    $memberIdSet = New-Object 'System.Collections.Generic.HashSet[string]'
    foreach ($m in $allMembers) {
        if ($m.userId) {
            $memberIdSet.Add($m.userId) | Out-Null
        }
    }
    return $memberIdSet
}

# (9) Compute-VisibleHistoryStart: Determines the timestamp for chat history sharing based on $Script:ShareHistoryMode.
function Compute-VisibleHistoryStart {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string] $Mode)
    switch -Regex ($Mode) {
        "None"  { return $null }  # No history will be shared (omit the property in API call).
        "All"   { return "0001-01-01T00:00:00Z" }  # Graph uses year 0001 to indicate all history.
        "^Since:(.+)" {
            # Validate the provided date/time format (ISO 8601 expected)
            $ts = $Matches[1]
            try { [DateTimeOffset]::Parse($ts) | Out-Null } catch { throw "Invalid date format for ShareHistory 'Since:' timestamp." }
            return $ts
        }
        default { throw "Invalid ShareHistoryMode '$Mode'. Use 'None', 'All', or 'Since:YYYY-MM-DDThh:mm:ssZ'." }
    }
}

# (10) Add-User-To-Chat: Adds a single user to the specified chat.
function Add-User-To-Chat {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string] $ChatId,
        [Parameter(Mandatory)][string] $UserId,
        [Parameter()][string] $UserType    # e.g., 'Member' or 'Guest' (from Azure AD userType property)
    )
    # Determine roles: if the user is a guest in the tenant, assign the "guest" role; otherwise no roles (defaults to member).
    $roles = @()
    if ($UserType -eq "Guest") {
        $roles = @("guest")
    } else {
        $roles = @("owner")
    }
    # Prepare the JSON body for adding a chat member.
    $body = @{
        "@odata.type"            = "#microsoft.graph.aadUserConversationMember"
        "user@odata.bind"        = "$Script:GraphBase/users/$UserId"  # bind user by ID
        "roles"                  = $roles
    }
    # Optionally include history start (if not sharing history, we leave it out to default to "no history").
    $historyTimestamp = Compute-VisibleHistoryStart -Mode $Script:ShareHistoryMode
    if ($historyTimestamp) {
        $body.visibleHistoryStartDateTime = $historyTimestamp
    }

    # Call Graph API to add the member.
    $uri = "$($Script:GraphBase)/chats/$ChatId/members"

    Write-Verbose "URI: $Uri"
    Write-Verbose "Body: $( $body | ConvertTo-Json -Depth 5 )"

    try {
        Invoke-GraphRequest -Method POST -Uri $uri -Body $body -ExpectNoContent
        return $true
    }
    catch {
        # Common errors: 409 Conflict if user already in chat, 403 Forbidden if permissions missing, 400 for bad request.
        Write-Warning "Failed to add user $UserId to chat $ChatId : $($_.Exception.Message)"
        return $false
    }
}

# (11) Add-GroupMembers-To-Chat: Orchestrates the process by using all above functions and handles user input & retries.
function Add-GroupMembers-To-Chat {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string] $Upn,
        [Parameter(Mandatory)][string] $ChatTopic,
        [Parameter(Mandatory)][string] $GroupName
    )

    # --- Step 1: Resolve User UPN to UserId, with retry logic ---
    $localUpn = $Upn
    for ($attempt = 1; $attempt -le 3; $attempt++) {
        try {
            Write-Host "Resolving user '$localUpn'..." -ForegroundColor Cyan
            $userId = Resolve-UserId -Upn $localUpn
            break  # success, exit loop
        }
        catch {
            Write-Warning $_.Exception.Message
            if ($attempt -eq 3) {
                throw "Unable to find user '$localUpn' after 3 attempts."
            }
            # Prompt for correction and retry
            $localUpn = Read-Host "Please enter a valid user UPN"
        }
    }

    # --- Step 2: Find the chat by topic (name) for that user, with retry logic ---
    $localChatTopic = $ChatTopic
    for ($attempt = 1; $attempt -le 3; $attempt++) {
        Write-Host "Finding chat named '$localChatTopic' for user $localUpn..." -ForegroundColor Cyan
        $chat = Get-UserChatByTopic -UserId $userId -Topic $localChatTopic
        if ($chat) {
            $chatId = $chat.id
            Write-Host "Chat found: '$($chat.topic)' (ID: $chatId)" -ForegroundColor Green
            break
        } else {
            Write-Warning "No group chat named '$localChatTopic' was found for user $localUpn."
            if ($attempt -eq 3) {
                throw "Unable to find a chat named '$localChatTopic' after 3 attempts."
            }
            $localChatTopic = Read-Host "Please enter the correct chat name (topic)"
        }
    }

    # --- Step 3: Resolve the M365 group name to groupId, with retry logic ---
    $localGroupName = $GroupName
    for ($attempt = 1; $attempt -le 3; $attempt++) {
        try {
            Write-Host "Resolving group '$localGroupName'..." -ForegroundColor Cyan
            $groupId = Resolve-GroupId -GroupNameOrEmail $localGroupName
            Write-Host "Group resolved. ID: $groupId" -ForegroundColor Green
            break
        }
        catch {
            Write-Warning $_.Exception.Message
            if ($attempt -eq 3) {
                throw "Unable to find group '$localGroupName' after 3 attempts."
            }
            $localGroupName = Read-Host "Please enter a valid group name or email"
        }
    }

    # --- Step 4: Get group members ---
    Write-Host "Retrieving members of group (ID: $groupId)..." -ForegroundColor Cyan
    $groupMembers = Get-GroupUsers -GroupId $groupId -IncludeTransitive:$Script:IncludeNestedGroupMembers
    if (-not $groupMembers -or $groupMembers.Count -eq 0) {
        throw "The group '$localGroupName' has no members or members are not accessible."
    }
    Write-Host "Found $($groupMembers.Count) user(s) in the group." -ForegroundColor Green

    # --- Step 5: Load existing chat members to avoid duplicates ---
    Write-Host "Checking current chat members to avoid duplicate adds..." -ForegroundColor Cyan
    $existingMemberIds = Get-ExistingChatMemberIds -ChatId $chatId -UserId $userId

    # --- Step 6: Add each member to the chat ---
    Write-Host "Adding group members to the chat..." -ForegroundColor Cyan
    $added   = @()
    $skipped = @()
    $failed  = @()
    foreach ($member in $groupMembers) {
        $uid = $member.id
        if ($existingMemberIds.Contains($uid)) {
            $skipped += $member
            continue
        }
        $success = Add-User-To-Chat -ChatId $chatId -UserId $uid -UserType $member.userType
        if ($success) {
            $added += $member
        } else {
            $failed += $member
        }
        Start-Sleep -Milliseconds 200  # small delay to be polite (avoid spamming requests)
    }

    # --- Step 7: Summary output ---
    Write-Host "`n========== SUMMARY ==========" -ForegroundColor Yellow
    Write-Host "Total members in group   : $($groupMembers.Count)"
    Write-Host "Added to chat            : $($added.Count)"
    Write-Host "Already in chat (skipped): $($skipped.Count)"
    Write-Host "Failed to add            : $($failed.Count)" 
    if ($failed.Count -gt 0) {
        Write-Host "`nFailed to add the following users:" -ForegroundColor Red
        $failed | Select-Object displayName, userPrincipalName, id, userType | Format-Table -AutoSize
    }
}
# --------------------------- SCRIPT ENTRY POINT ---------------------------
try {
    # Read app credentials and get an access token
    $app = Read-GraphAppSecret
    $Script:AccessToken = Get-GraphToken -TenantId $app.tenantId -ClientId $app.clientId -ClientSecret $app.clientSecret

    # Prompt for inputs
    $upnInput   = Read-Host "Enter the UPN of a user in the target chat"
    $chatInput  = Read-Host "Enter the Teams group chat name (topic)"
    $groupInput = Read-Host "Enter the Microsoft 365 group name (or email address)"

    # Execute the add-members process
    Add-GroupMembers-To-Chat -Upn $upnInput -ChatTopic $chatInput -GroupName $groupInput
}
catch {
    Write-Error "ERROR: $($_.Exception.Message)"
    exit 1
}