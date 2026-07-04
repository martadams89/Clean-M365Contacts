<#
==============================================================================
 Clean-M365Contacts.ps1
------------------------------------------------------------------------------
 Menu-driven de-duplication of PERSONAL contacts in an Exchange Online mailbox
 (new Outlook / cloud mailbox), via the Microsoft Graph REST API.

 WHY THIS DESIGN:
   * Uses delegated device-code sign-in, so it only ever touches the mailbox
     of the account you sign in with. Run it ON the user's machine, sign in
     AS that user, and you physically cannot hit the wrong mailbox.
   * Pure Invoke-RestMethod against Graph - NO modules to install.
     Runs on Windows PowerShell 5.1 (incl. the ISE) AND PowerShell 7.
   * Deletions move contacts to the mailbox's Deleted Items (recoverable),
     they are NOT hard-deleted.

 HOW TO RUN:
   PS7 host:   pwsh -File .\Clean-M365Contacts.ps1
   ISE host:   open the file in PowerShell ISE and press F5
               (or:  powershell -ExecutionPolicy Bypass -File .\Clean-M365Contacts.ps1)

 Repeat on each user's machine, signing in as that user.
==============================================================================
#>

# --- Force TLS 1.2 (Windows PowerShell 5.1 / ISE often defaults too low) -----
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

# --- Configuration -----------------------------------------------------------
# Public client used by the Microsoft Graph PowerShell SDK. Supports device code.
$ClientId = '14d82eec-204b-4c2f-b7e8-296a70dab67e'
# 'organizations' = any work/school account (blocks personal MSA). You can
# hard-code the tenant domain here instead, e.g. 'dennisfreight.co.uk'.
$TenantId = 'organizations'
$Scopes   = 'https://graph.microsoft.com/Contacts.ReadWrite https://graph.microsoft.com/User.Read offline_access openid profile'

$GraphBase = 'https://graph.microsoft.com/v1.0'
$Select    = 'id,displayName,givenName,surname,emailAddresses,mobilePhone,businessPhones,companyName,jobTitle'

# --- Script-scoped state -----------------------------------------------------
$script:Token        = $null      # full token response (access + refresh)
$script:TokenExpiry  = [datetime]::MinValue
$script:Identity     = $null      # /me result: who we're signed in as
$script:Contacts     = $null      # last fetched contact list
$script:MatchMode    = 'NameEmail'  # NameEmail | Name | Email

# Output/cache locations. Guarded so the script can also be dot-sourced on
# non-Windows CI runners (for the test suite) without throwing on empty paths.
$desktop = [Environment]::GetFolderPath('Desktop')
if ([string]::IsNullOrEmpty($desktop)) { $desktop = [Environment]::GetFolderPath('UserProfile') }
if ([string]::IsNullOrEmpty($desktop)) { $desktop = (Get-Location).Path }
$script:LogDir = Join-Path $desktop 'ContactCleanup'

# Cached refresh token (DPAPI-encrypted, current-user + this-machine only) so
# re-runs sign in silently instead of prompting the authenticator every time.
$localApp = $env:LOCALAPPDATA
if ([string]::IsNullOrEmpty($localApp)) { $localApp = [IO.Path]::GetTempPath() }
$script:CacheFile = Join-Path (Join-Path $localApp 'ContactCleanup') 'session.dat'

# =============================================================================
#  Helpers
# =============================================================================

function Get-GraphErrorBody {
    param($ErrorRecord)
    # PS7 (and often 5.1): body is in ErrorDetails.Message
    if ($ErrorRecord.ErrorDetails -and $ErrorRecord.ErrorDetails.Message) {
        try { return ($ErrorRecord.ErrorDetails.Message | ConvertFrom-Json) } catch { Write-Verbose $_.Exception.Message }
    }
    # Windows PowerShell 5.1 fallback: read the WebException response stream
    $resp = $ErrorRecord.Exception.Response
    if ($resp -and ($resp -is [System.Net.HttpWebResponse])) {
        try {
            $reader = New-Object System.IO.StreamReader($resp.GetResponseStream())
            $text   = $reader.ReadToEnd(); $reader.Close()
            return ($text | ConvertFrom-Json)
        } catch { Write-Verbose $_.Exception.Message }
    }
    return $null
}

function Get-JwtClaims {
    # Decode the payload of a JWT (e.g. the id_token) without any Graph call.
    param([string]$Jwt)
    if (-not $Jwt) { return $null }
    $parts = $Jwt.Split('.')
    if ($parts.Count -lt 2) { return $null }
    $p = $parts[1].Replace('-', '+').Replace('_', '/')
    switch ($p.Length % 4) { 2 { $p += '==' } 3 { $p += '=' } }
    try {
        $bytes = [Convert]::FromBase64String($p)
        return ([Text.Encoding]::UTF8.GetString($bytes) | ConvertFrom-Json)
    } catch { return $null }
}

function Resolve-Identity {
    # Prefer the id_token (no network call); fall back to Graph /me if needed.
    param($TokenResponse)
    $claims = Get-JwtClaims $TokenResponse.id_token
    if ($claims -and $claims.preferred_username) {
        $script:Identity = [pscustomobject]@{
            displayName       = $claims.name
            userPrincipalName = $claims.preferred_username
        }
        return
    }
    try {
        $me = Invoke-RestMethod -Method Get -Uri "$GraphBase/me" `
            -Headers @{ Authorization = "Bearer $($TokenResponse.access_token)" }
        $script:Identity = [pscustomobject]@{
            displayName       = $me.displayName
            userPrincipalName = $me.userPrincipalName
        }
    } catch {
        $script:Identity = [pscustomobject]@{ displayName = '(unknown)'; userPrincipalName = '(unknown)' }
    }
}

function Save-TokenCache {
    # Store the refresh token, DPAPI-encrypted to THIS user on THIS machine.
    if (-not $script:Token -or -not $script:Token.refresh_token) { return }
    try {
        $dir = Split-Path $script:CacheFile -Parent
        if (-not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
        $payload = @{ refresh_token = $script:Token.refresh_token } | ConvertTo-Json
        $sec = ConvertTo-SecureString $payload -AsPlainText -Force
        ConvertFrom-SecureString $sec | Set-Content -Path $script:CacheFile -Encoding ASCII
    } catch { Write-Verbose $_.Exception.Message }   # caching is a convenience; never let it break the run
}

function Clear-TokenCache {
    if (Test-Path $script:CacheFile) { Remove-Item $script:CacheFile -Force -ErrorAction SilentlyContinue }
}

function Get-TokenFromRefresh {
    param([string]$RefreshToken)
    return Invoke-RestMethod -Method Post -ContentType 'application/x-www-form-urlencoded' `
        -Uri "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token" -Body @{
            grant_type    = 'refresh_token'
            client_id     = $ClientId
            refresh_token = $RefreshToken
            scope         = $Scopes
        }
}

function Restore-Session {
    # Called at startup: if a cached refresh token exists, sign in silently.
    if (-not (Test-Path $script:CacheFile)) { return $false }
    try {
        $enc = Get-Content -Path $script:CacheFile -Raw
        $sec = ConvertTo-SecureString $enc   # DPAPI decrypt (fails if wrong user/machine)
        $plain = (New-Object System.Management.Automation.PSCredential('x', $sec)).GetNetworkCredential().Password
        $refresh = ($plain | ConvertFrom-Json).refresh_token
        if (-not $refresh) { return $false }

        $tok = Get-TokenFromRefresh $refresh
        $script:Token       = $tok
        $script:TokenExpiry = (Get-Date).AddSeconds([int]$tok.expires_in)
        Resolve-Identity $tok
        Save-TokenCache   # refresh tokens rotate; persist the new one
        Write-Host "Resumed saved session for: $($script:Identity.userPrincipalName)" -ForegroundColor Green
        return $true
    } catch {
        Clear-TokenCache   # stale/invalid cache - wipe it and fall back to interactive
        return $false
    }
}

function Connect-Graph {
    Write-Host ''
    $body = @{ client_id = $ClientId; scope = $Scopes }
    try {
        $dc = Invoke-RestMethod -Method Post -ContentType 'application/x-www-form-urlencoded' `
            -Uri "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/devicecode" -Body $body
    } catch {
        Write-Host "Could not start sign-in: $($_.Exception.Message)" -ForegroundColor Red
        return
    }

    Write-Host '------------------------------------------------------------' -ForegroundColor Cyan
    Write-Host $dc.message -ForegroundColor Yellow
    Write-Host '------------------------------------------------------------' -ForegroundColor Cyan

    # Convenience: copy the code and open the sign-in page automatically
    try { Set-Clipboard -Value $dc.user_code; Write-Host '(code copied to clipboard)' -ForegroundColor DarkGray } catch { Write-Verbose $_.Exception.Message }
    try { Start-Process $dc.verification_uri } catch { Write-Verbose $_.Exception.Message }

    $interval = [int]$dc.interval
    $deadline = (Get-Date).AddSeconds([int]$dc.expires_in)
    Write-Host 'Waiting for you to complete sign-in in the browser...' -ForegroundColor DarkGray

    while ((Get-Date) -lt $deadline) {
        Start-Sleep -Seconds $interval
        try {
            $tok = Invoke-RestMethod -Method Post -ContentType 'application/x-www-form-urlencoded' `
                -Uri "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token" -Body @{
                    grant_type  = 'urn:ietf:params:oauth:grant-type:device_code'
                    client_id   = $ClientId
                    device_code = $dc.device_code
                }
            $script:Token       = $tok
            $script:TokenExpiry = (Get-Date).AddSeconds([int]$tok.expires_in)
            break
        } catch {
            $err = Get-GraphErrorBody $_
            if     ($err.error -eq 'authorization_pending') { continue }
            elseif ($err.error -eq 'slow_down')             { $interval += 5; continue }
            else {
                Write-Host "Sign-in failed: $($err.error) - $($err.error_description)" -ForegroundColor Red
                return
            }
        }
    }

    if (-not $script:Token) { Write-Host 'Sign-in timed out.' -ForegroundColor Red; return }

    # Confirm identity so you can SEE which mailbox you're about to touch
    Resolve-Identity $script:Token
    Save-TokenCache            # remember this session for silent re-runs
    Write-Host ''
    Write-Host "Signed in as: $($script:Identity.displayName) <$($script:Identity.userPrincipalName)>" -ForegroundColor Green
    $script:Contacts = $null   # force a fresh scan for the new account
}

function Test-Auth {
    if (-not $script:Token) { Write-Host 'Not signed in. Choose option 1 first.' -ForegroundColor Yellow; return $false }
    # Refresh silently if the token is near expiry (long review sessions)
    if ((Get-Date) -ge $script:TokenExpiry.AddMinutes(-5) -and $script:Token.refresh_token) {
        try {
            $tok = Get-TokenFromRefresh $script:Token.refresh_token
            $script:Token       = $tok
            $script:TokenExpiry = (Get-Date).AddSeconds([int]$tok.expires_in)
            Save-TokenCache   # rotated refresh token - keep the cache current
        } catch { Write-Host 'Token refresh failed - please sign in again (option 1).' -ForegroundColor Yellow; return $false }
    }
    return $true
}

function Get-AuthHeader { @{ Authorization = "Bearer $($script:Token.access_token)" } }

function Get-AllContacts {
    $uri = "$GraphBase/me/contacts?`$top=100&`$select=$Select"
    $all = New-Object System.Collections.Generic.List[object]
    while ($uri) {
        $resp = Invoke-RestMethod -Method Get -Uri $uri -Headers (Get-AuthHeader)
        foreach ($c in $resp.value) { $all.Add($c) }
        $uri = $resp.'@odata.nextLink'
    }
    return $all
}

function Get-FirstEmail {
    param($c)
    if ($c.emailAddresses -and $c.emailAddresses.Count -gt 0) {
        return ('' + $c.emailAddresses[0].address).Trim().ToLower()
    }
    return ''
}

function Get-ContactKey {
    param($c, $Mode)
    $name  = ('' + $c.displayName).Trim().ToLower()
    $email = Get-FirstEmail $c
    switch ($Mode) {
        'NameEmail' { return "$name|$email" }
        'Name'      { return $name }
        'Email'     { return $email }
    }
}

function Get-Score {
    # "Completeness" - so we KEEP the richest record and delete the stubs
    param($c)
    $s = 0
    if ($c.displayName)  { $s++ }
    if ($c.givenName)    { $s++ }
    if ($c.surname)      { $s++ }
    if ($c.emailAddresses)  { $s += $c.emailAddresses.Count }
    if ($c.mobilePhone)  { $s++ }
    if ($c.businessPhones)  { $s += $c.businessPhones.Count }
    if ($c.companyName)  { $s++ }
    if ($c.jobTitle)     { $s++ }
    return $s
}

function Find-Duplicates {
    param($Contacts, $Mode)
    $result = @()
    $groups = $Contacts | Group-Object { Get-ContactKey $_ $Mode }
    foreach ($g in $groups) {
        if ($g.Count -lt 2) { continue }
        # Never auto-delete on an empty match value (e.g. blank name / no email)
        $emptyKey = ($Mode -eq 'NameEmail' -and $g.Name -eq '|') -or `
                    ($Mode -ne 'NameEmail' -and [string]::IsNullOrWhiteSpace($g.Name))
        if ($emptyKey) { continue }

        $sorted = $g.Group | Sort-Object { Get-Score $_ } -Descending
        $result += [pscustomobject]@{
            Key    = $g.Name
            Keep   = $sorted[0]
            Delete = @($sorted | Select-Object -Skip 1)
        }
    }
    return $result
}

function Show-ContactRow {
    param($c, $Action)
    [pscustomobject]@{
        Action  = $Action
        Name    = $c.displayName
        Email   = (Get-FirstEmail $c)
        Mobile  = $c.mobilePhone
        Company = $c.companyName
        Id      = $c.id
    }
}

function Initialize-LogDir { if (-not (Test-Path $script:LogDir)) { New-Item -ItemType Directory -Path $script:LogDir | Out-Null } }

# =============================================================================
#  Menu actions
# =============================================================================

function Invoke-Scan {
    if (-not (Test-Auth)) { return }
    Write-Host 'Fetching contacts...' -ForegroundColor DarkGray
    $script:Contacts = Get-AllContacts
    $dupes = Find-Duplicates $script:Contacts $script:MatchMode
    $delCount = ($dupes | ForEach-Object { $_.Delete.Count } | Measure-Object -Sum).Sum
    if (-not $delCount) { $delCount = 0 }
    Write-Host ''
    Write-Host "Total contacts in mailbox : $($script:Contacts.Count)" -ForegroundColor Cyan
    Write-Host "Duplicate groups found    : $($dupes.Count)   (match mode: $($script:MatchMode))" -ForegroundColor Cyan
    Write-Host "Contacts flagged to delete: $delCount" -ForegroundColor Cyan
}

function Invoke-Preview {
    if (-not (Test-Auth)) { return }
    if (-not $script:Contacts) { Invoke-Scan }
    $dupes = Find-Duplicates $script:Contacts $script:MatchMode
    if (-not $dupes.Count) { Write-Host 'No duplicates found with the current match mode.' -ForegroundColor Green; return }

    $rows = foreach ($grp in $dupes) {
        Show-ContactRow $grp.Keep 'KEEP'
        foreach ($d in $grp.Delete) { Show-ContactRow $d 'delete' }
    }
    $rows | Format-Table -AutoSize | Out-Host
    $delCount = ($dupes | ForEach-Object { $_.Delete.Count } | Measure-Object -Sum).Sum
    Write-Host ""
    Write-Host "Would KEEP $($dupes.Count) master record(s) and DELETE $delCount duplicate(s)." -ForegroundColor Yellow
    Write-Host "Nothing has been changed. Use option 5 to actually delete." -ForegroundColor Yellow
}

function Invoke-Export {
    if (-not (Test-Auth)) { return }
    if (-not $script:Contacts) { Invoke-Scan }
    $dupes = Find-Duplicates $script:Contacts $script:MatchMode
    if (-not $dupes.Count) { Write-Host 'No duplicates to export.' -ForegroundColor Green; return }
    Initialize-LogDir
    $stamp = Get-Date -Format 'yyyyMMdd-HHmmss'
    $who   = if ($script:Identity) { $script:Identity.userPrincipalName } else { 'unknown' }
    $file  = Join-Path $script:LogDir "preview_${who}_$stamp.csv"
    $rows = foreach ($grp in $dupes) {
        Show-ContactRow $grp.Keep 'KEEP'
        foreach ($d in $grp.Delete) { Show-ContactRow $d 'delete' }
    }
    $rows | Export-Csv -Path $file -NoTypeInformation -Encoding UTF8
    Write-Host "Preview written to: $file" -ForegroundColor Green
}

function Invoke-Deletion {
    if (-not (Test-Auth)) { return }
    if (-not $script:Contacts) { Invoke-Scan }
    $dupes = Find-Duplicates $script:Contacts $script:MatchMode
    $toDelete = @($dupes | ForEach-Object { $_.Delete } )
    if (-not $toDelete.Count) { Write-Host 'Nothing to delete.' -ForegroundColor Green; return }

    Write-Host ''
    Write-Host "About to delete $($toDelete.Count) duplicate contact(s) from:" -ForegroundColor Red
    Write-Host "  $($script:Identity.userPrincipalName)" -ForegroundColor Red
    Write-Host "They will go to Deleted Items and can be recovered." -ForegroundColor DarkGray
    $confirm = Read-Host "Type  DELETE  to proceed (anything else cancels)"
    if ($confirm -ne 'DELETE') { Write-Host 'Cancelled.' -ForegroundColor Yellow; return }

    Initialize-LogDir
    $stamp = Get-Date -Format 'yyyyMMdd-HHmmss'
    $who   = if ($script:Identity) { $script:Identity.userPrincipalName } else { 'unknown' }
    $log   = Join-Path $script:LogDir "deleted_${who}_$stamp.csv"
    $logRows = @()

    $ok = 0; $fail = 0
    foreach ($c in $toDelete) {
        try {
            Invoke-RestMethod -Method Delete -Uri "$GraphBase/me/contacts/$($c.id)" -Headers (Get-AuthHeader) | Out-Null
            $ok++
            Write-Host "  deleted: $($c.displayName)  <$(Get-FirstEmail $c)>" -ForegroundColor DarkGray
            $logRows += Show-ContactRow $c 'deleted'
        } catch {
            $fail++
            Write-Host "  FAILED : $($c.displayName) - $($_.Exception.Message)" -ForegroundColor Red
            $logRows += Show-ContactRow $c 'FAILED'
        }
    }
    if ($logRows.Count) { $logRows | Export-Csv -Path $log -NoTypeInformation -Encoding UTF8 }
    Write-Host ''
    Write-Host "Done. Deleted: $ok   Failed: $fail" -ForegroundColor Green
    Write-Host "Audit log: $log" -ForegroundColor Green
    $script:Contacts = $null   # force a re-scan next time
}

function Select-MatchMode {
    Write-Host ''
    Write-Host 'Match mode decides what counts as a duplicate:' -ForegroundColor Cyan
    Write-Host '  1. Name + Email   (SAFEST - same display name AND same first email)'
    Write-Host '  2. Name only      (looser - same display name, ignores email)'
    Write-Host '  3. Email only     (same first email, ignores name)'
    $m = Read-Host 'Choose 1-3'
    switch ($m) {
        '1' { $script:MatchMode = 'NameEmail' }
        '2' { $script:MatchMode = 'Name' }
        '3' { $script:MatchMode = 'Email' }
        default { Write-Host 'Unchanged.' -ForegroundColor Yellow; return }
    }
    Write-Host "Match mode is now: $($script:MatchMode)" -ForegroundColor Green
    $script:Contacts = $null
}

function Disconnect-Session {
    $script:Token = $null; $script:Identity = $null; $script:Contacts = $null
    $script:TokenExpiry = [datetime]::MinValue
    Clear-TokenCache
    Write-Host 'Signed out and cleared the saved session on this machine.' -ForegroundColor Green
}

# =============================================================================
#  Main menu loop
# =============================================================================

function Show-Menu {
    Write-Host ''
    Write-Host '==================================================' -ForegroundColor White
    Write-Host '   M365 Contact De-duplication' -ForegroundColor White
    Write-Host '==================================================' -ForegroundColor White
    if ($script:Identity) {
        Write-Host " Account : $($script:Identity.userPrincipalName)" -ForegroundColor Green
    } else {
        Write-Host ' Account : (not signed in)' -ForegroundColor Yellow
    }
    Write-Host " Match   : $($script:MatchMode)" -ForegroundColor Green
    Write-Host '--------------------------------------------------'
    Write-Host ' 1. Sign in / switch account (only needed if not already resumed)'
    Write-Host ' 2. Scan (show counts)'
    Write-Host ' 3. Preview duplicates (dry run)'
    Write-Host ' 4. Export preview to CSV'
    Write-Host ' 5. Delete duplicates'
    Write-Host ' 6. Change match mode'
    Write-Host ' 7. Sign out (clears saved session on this machine)'
    Write-Host ' 0. Exit'
    Write-Host '--------------------------------------------------'
}

function Invoke-ContactCleanup {
    # Try to resume a previously saved session (no prompt if the cache is valid)
    Write-Host 'Checking for a saved session...' -ForegroundColor DarkGray
    [void](Restore-Session)

    do {
        Show-Menu
        $choice = Read-Host 'Select'
        switch ($choice) {
            '1' { Connect-Graph }
            '2' { Invoke-Scan }
            '3' { Invoke-Preview }
            '4' { Invoke-Export }
            '5' { Invoke-Deletion }
            '6' { Select-MatchMode }
            '7' { Disconnect-Session }
            '0' { Write-Host 'Bye.' -ForegroundColor Cyan }
            default { Write-Host 'Invalid choice.' -ForegroundColor Yellow }
        }
        if ($choice -ne '0') { Write-Host ''; [void](Read-Host 'Press Enter to return to the menu') }
    } while ($choice -ne '0')
}

# Auto-run the interactive menu when executed directly. When the test suite
# dot-sources this file it sets CLEANM365_NOAUTORUN=1 to load functions only.
if (-not $env:CLEANM365_NOAUTORUN) {
    Invoke-ContactCleanup
}
