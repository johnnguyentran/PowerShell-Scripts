#Requires -Modules Az.Accounts, Az.KeyVault

function Send-GraphMailMessage {
<#
.SYNOPSIS
    Sends an HTML email via the Microsoft Graph sendMail endpoint.

.DESCRIPTION
    Reusable mail-send helper used by SERVICEACCOUNT automation scripts. Acquires a
    client-credentials access token from Entra, assembles an HTML message
    body (optional intro text, optional styled table, optional outro text),
    optionally attaches one or more files, and posts the message to the
    Graph /users/{sender}/sendMail endpoint.

    The Graph application client secret is retrieved from Azure Key Vault
    instead of a local DPAPI-encrypted file. This allows the function to run
    under different execution identities, including Azure Functions using
    Managed Identity, as long as the running identity has access to the
    configured Key Vault secret.

    Because this function is itself the alerting pipeline for the scripts
    that call it, it does not call out to any higher-level Send-Alert
    function. Instead it logs every run (success or failure) to its own
    persistent log file and, when -LogFile is supplied, mirrors those
    entries into the caller's log so Graph-mail activity appears inline
    with the rest of the caller's run.

.PARAMETER sender
    The mailbox the message is sent from. Defaults to ServiceAccount@contoso.com.

.PARAMETER subject
    Email subject line. Required.

.PARAMETER body1
    Optional text shown above the table, if any. Line breaks become <br>.

.PARAMETER table
    Optional object or array of objects rendered as an HTML table.

.PARAMETER body2
    Optional text shown below the table, if any. Line breaks become <br>.

.PARAMETER attachments
    Optional array of file paths to attach. Missing/unreadable paths are
    logged and skipped. The send still proceeds with the remaining attachments.

.PARAMETER recipients
    Array of recipient email addresses. Required; must contain at least one entry.

.PARAMETER LogFile
    Optional path to the caller's log file. When supplied, this function
    mirrors its own log entries into that file in addition to writing its
    own dedicated log.

.PARAMETER KeyVaultName
    Name of the Azure Key Vault that stores the Graph app client secret.
    Defaults to <KEYVAULT-NAME>.

.PARAMETER SecretName
    Name of the Key Vault secret containing the Graph app client secret.
    Defaults to Send-GraphMailSecret.

.OUTPUTS
    On success:
        - Returns the object returned by Invoke-RestMethod if one is provided.
        - If Invoke-RestMethod succeeds but Graph returns no response body,
          returns a small success object with Success=$true and LogFile.

    On failure:
        - Returns a PSCustomObject with Success=$false, ErrorMessage, and
          LogFile properties.

.NOTES
    Execution context:
        Called inline by automion scripts.

    Dependencies:
        Az.Accounts
        Az.KeyVault
        Outbound HTTPS to:
            login.microsoftonline.com
            graph.microsoft.com
            configured Key Vault endpoint

    Key Vault access:
        The running identity must have permission to read the configured secret.
        Recommended RBAC role: Key Vault Secrets User.

    Auth behavior:
        Uses an existing Az context if one is available.
        If no context exists, attempts Connect-AzAccount -Identity, which is
        appropriate for Azure Functions, Azure VMs, or Arc-enabled servers
        using Managed Identity.
#>

    [CmdletBinding()]
    Param(
        [string]$sender = "ServiceAccount@contoso.com",

        [Parameter(Mandatory)]
        [string]$subject,

        [string]$body1,

        $table,

        [string]$body2,

        [string[]]$attachments,

        [Parameter(Mandatory)]
        [string[]]$recipients,

        [string]$LogFile,

        [string]$KeyVaultName = "<KEYVAULT-NAME>",

        [string]$SecretName = "<SECRET-NAME>"
    )

    #region ========================= CONFIGURATION =========================

    $Config = @{
        AppId          = "<APP-ID>"
        TenantId       = "<TENANT-ID>"
        TokenUriFormat = "https://login.microsoftonline.com/{0}/oauth2/v2.0/token"
        SendMailUriFmt = "https://graph.microsoft.com/v1.0/users/{0}/sendMail"
        GraphScope     = "https://graph.microsoft.com/.default"
        LogDir         = "<LOG DIRECTORY>GraphMailMessage"

        KeyVaultName   = $KeyVaultName
        SecretName     = $SecretName
    }

    # 24-hour timestamp prevents collisions between AM/PM runs.
    $Timestamp  = Get-Date -Format "yyyyMMdd-HHmmss"
    $OwnLogFile = Join-Path $Config.LogDir "GraphMail_$Timestamp.log"

    #endregion

    #region ========================= LOGGING =========================

    try {
        if (-not (Test-Path $Config.LogDir)) {
            New-Item -Path $Config.LogDir -ItemType Directory -Force | Out-Null
        }
    }
    catch {
        $OwnLogFile = Join-Path $env:TEMP "GraphMail_$Timestamp.log"
    }

    function Write-GraphMailLog {
        param(
            [Parameter(Mandatory)]
            [string]$Message,

            [ValidateSet("INFO", "WARN", "ERROR")]
            [string]$Level = "INFO"
        )

        $Entry = "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] [$Level] $Message"

        try {
            Add-Content -Path $OwnLogFile -Value $Entry -Encoding UTF8 -ErrorAction Stop
        }
        catch {
            # If own logging fails, do not throw from the logger and mask the original issue.
        }

        if ($LogFile) {
            try {
                Add-Content -Path $LogFile -Value "$Entry [Send-GraphMailMessage]" -Encoding UTF8 -ErrorAction Stop
            }
            catch {
                # Caller log mirroring is optional. Failure to mirror should not stop mail send.
            }
        }

        switch ($Level) {
            "ERROR" { Write-Host $Entry -ForegroundColor Red }
            "WARN"  { Write-Host $Entry -ForegroundColor Yellow }
            default { Write-Host $Entry }
        }
    }

    function New-FailureResult {
        param(
            [Parameter(Mandatory)]
            [string]$ErrorMessage
        )

        [PSCustomObject]@{
            Success      = $false
            ErrorMessage = $ErrorMessage
            LogFile      = $OwnLogFile
        }
    }

    function New-SuccessResult {
        param(
            [string]$Message = "Invoke-RestMethod completed successfully, but no response body was returned."
        )

        [PSCustomObject]@{
            Success = $true
            Message = $Message
            LogFile = $OwnLogFile
        }
    }

    #endregion

    #region ========================= HELPER FUNCTIONS =========================

    function Ensure-AzContext {
        <#
        .SYNOPSIS
            Ensures an Azure PowerShell context is available.

        .DESCRIPTION
            Uses an existing Az context if one already exists. If not, attempts
            Managed Identity authentication via Connect-AzAccount -Identity.

            This supports:
                - Interactive/admin runs where Connect-AzAccount was already used
                - Scheduled-task/service-account runs with an existing context
                - Azure Functions using Managed Identity
                - Azure VMs / Arc-enabled servers using Managed Identity
        #>

        try {
            $ExistingContext = Get-AzContext -ErrorAction SilentlyContinue

            if ($ExistingContext -and $ExistingContext.Account) {
                Write-GraphMailLog "Connected to Az using Managed Identity: $($ManagedIdentityContext.Account.Id)"
                return $true
            }
        }
        catch {
            Write-GraphMailLog "Unable to check existing Az context: $($_.Exception.Message)" -Level WARN
        }

        try {
            Connect-AzAccount -Identity -ErrorAction Stop | Out-Null
            $ManagedIdentityContext = Get-AzContext -ErrorAction SilentlyContinue

            if ($ManagedIdentityContext -and $ManagedIdentityContext.Account) {
                Write-GraphMailLog "Connected to Az using Managed Identity: $($ManagedIdentityContext.Account.Id)"
            }
            else {
                Write-GraphMailLog "Connected to Az using Managed Identity."
            }

            return $true
        }
        catch {
            Write-GraphMailLog "No existing Az context and Connect-AzAccount -Identity failed: $($_.Exception.Message)" -Level ERROR
            return $false
        }
    }

    function Get-GraphErrorDetails {
        param(
            [Parameter(Mandatory)]
            $ErrorRecord
        )

        $RawResponse       = "No HTTP response available."
        $StatusCode        = "N/A"
        $StatusDescription = "N/A"

        $ErrorDetails = [PSCustomObject]@{
            Code            = "N/A"
            Message         = $ErrorRecord.Exception.Message
            Date            = (Get-Date -Format 'yyyyMMdd-HHmmss').ToString()
            RequestId       = "N/A"
            ClientRequestId = "N/A"
        }

        # PowerShell 7 / newer Invoke-RestMethod commonly exposes response body here.
        if ($ErrorRecord.ErrorDetails -and $ErrorRecord.ErrorDetails.Message) {
            $RawResponse = $ErrorRecord.ErrorDetails.Message
        }

        # Windows PowerShell WebException pattern.
        if ($null -ne $ErrorRecord.Exception.Response) {
            try {
                if ($ErrorRecord.Exception.Response.StatusCode) {
                    try {
                        $StatusCode = $ErrorRecord.Exception.Response.StatusCode.value__
                    }
                    catch {
                        $StatusCode = [string]$ErrorRecord.Exception.Response.StatusCode
                    }
                }

                if ($ErrorRecord.Exception.Response.StatusDescription) {
                    $StatusDescription = $ErrorRecord.Exception.Response.StatusDescription
                }

                if ($ErrorRecord.Exception.Response.GetResponseStream) {
                    $ResponseStream = $ErrorRecord.Exception.Response.GetResponseStream()

                    if ($ResponseStream) {
                        $StreamReader = New-Object System.IO.StreamReader($ResponseStream)
                        $StreamBody   = $StreamReader.ReadToEnd()
                        $StreamReader.Close()

                        if (-not [string]::IsNullOrWhiteSpace($StreamBody)) {
                            $RawResponse = $StreamBody
                        }
                    }
                }
            }
            catch {
                Write-GraphMailLog "Failed to read HTTP error response stream: $($_.Exception.Message)" -Level WARN
            }
        }

        if (-not [string]::IsNullOrWhiteSpace($RawResponse) -and $RawResponse -ne "No HTTP response available.") {
            try {
                $ErrorInfo = ConvertFrom-Json $RawResponse -ErrorAction Stop

                if ($ErrorInfo.error) {
                    $ErrorDetails = [PSCustomObject]@{
                        Code            = $ErrorInfo.error.code
                        Message         = $ErrorInfo.error.message
                        Date            = $ErrorInfo.error.innerError.date
                        RequestId       = $ErrorInfo.error.innerError.'request-id'
                        ClientRequestId = $ErrorInfo.error.innerError.'client-request-id'
                    }
                }
            }
            catch {
                $ErrorDetails.Message = "Unable to parse JSON error response. Original exception: $($ErrorRecord.Exception.Message)"
            }
        }

        [PSCustomObject]@{
            RawResponse       = $RawResponse
            StatusCode        = $StatusCode
            StatusDescription = $StatusDescription
            ErrorDetails      = $ErrorDetails
        }
    }

    #endregion

    Write-GraphMailLog "========== Send-GraphMailMessage Started =========="
    Write-GraphMailLog "Subject:    $subject"
    Write-GraphMailLog "Sender:     $sender"
    Write-GraphMailLog "Recipients: $($recipients -join ', ')"
    Write-GraphMailLog "Key Vault:  $($Config.KeyVaultName)"
    Write-GraphMailLog "Secret:     $($Config.SecretName)"

    #region ========================= INITIALIZATION =========================

    if (-not $recipients -or $recipients.Count -eq 0) {
        Write-GraphMailLog "No recipients supplied, cannot send." -Level ERROR
        return (New-FailureResult -ErrorMessage "No recipients supplied.")
    }

    if ([string]::IsNullOrWhiteSpace($Config.KeyVaultName)) {
        Write-GraphMailLog "Key Vault name is missing from configuration." -Level ERROR
        return (New-FailureResult -ErrorMessage "Key Vault name is missing from configuration.")
    }

    if ([string]::IsNullOrWhiteSpace($Config.SecretName)) {
        Write-GraphMailLog "Secret name is missing from configuration." -Level ERROR
        return (New-FailureResult -ErrorMessage "Secret name is missing from configuration.")
    }

    #endregion

    #region ========================= SECRET RETRIEVAL =========================

    <#
        The Graph app client secret is retrieved from Key Vault and cached
        at script scope for the lifetime of the current PowerShell session
        or Azure Functions runspace.

        This reduces repeated Key Vault calls when multiple messages are
        sent within the same script run or warm Functions runspace.

        This cache is not persistent. A new process/runspace/invocation may
        need to retrieve the secret again.
    #>

    
    $KvNameForKey  = $Config.KeyVaultName
    $SecNameForKey = $Config.SecretName
    $SecretCacheKey = "$($Config.KeyVaultName)|$($Config.SecretName)"

    if (-not $script:GraphMailSecretCache) {
        $script:GraphMailSecretCache = @{}
    }

    if ($script:GraphMailSecretCache.ContainsKey($SecretCacheKey)) {
        $AppSec = $script:GraphMailSecretCache[$SecretCacheKey]
        Write-GraphMailLog "Using cached Graph client secret from current session."
    }
    else {
        if (-not (Ensure-AzContext)) {
            return (New-FailureResult -ErrorMessage "No Az context available; cannot retrieve Graph client secret from Key Vault.")
        }

        try {
            $AppSec = Get-AzKeyVaultSecret `
                -VaultName $Config.KeyVaultName `
                -Name $Config.SecretName `
                -AsPlainText `
                -ErrorAction Stop

            if ([string]::IsNullOrWhiteSpace($AppSec)) {
                Write-GraphMailLog "Key Vault returned an empty secret value." -Level ERROR
                return (New-FailureResult -ErrorMessage "Key Vault returned an empty secret value.")
            }

            $script:GraphMailSecretCache[$SecretCacheKey] = $AppSec
            Write-GraphMailLog "Retrieved Graph client secret from Key Vault '$($Config.KeyVaultName)'."
        }
        catch {
            Write-GraphMailLog "Failed to retrieve secret from Key Vault: $($_.Exception.Message)" -Level ERROR
            return (New-FailureResult -ErrorMessage "Key Vault secret retrieval failed: $($_.Exception.Message)")
        }
    }

    #endregion

    #region ========================= TOKEN ACQUISITION =========================

    <#
        Token caching:
            Cached at script scope only.
            Reused only if the cached token has more than 10 minutes of
            remaining trusted lifetime.

        In SERVICEACCOUNT scheduled-task context:
            Helpful when one script sends multiple messages in a single run.

        In Azure Functions:
            Helpful within a single invocation and possibly across warm
            runspace reuse, but no cross-invocation persistence is assumed.
    #>

    $TokenCacheKey = "$($Config.TenantId)|$($Config.AppId)|$($Config.GraphScope)"

    if (-not $script:GraphMailTokenCache) {
        $script:GraphMailTokenCache = @{}
    }

    $UseCachedToken = $false

    if ($script:GraphMailTokenCache.ContainsKey($TokenCacheKey)) {
        $CachedTokenEntry = $script:GraphMailTokenCache[$TokenCacheKey]

        if ($CachedTokenEntry.Token -and $CachedTokenEntry.ExpiresAfter -gt (Get-Date).AddMinutes(10)) {
            $Token = $CachedTokenEntry.Token
            $UseCachedToken = $true
            Write-GraphMailLog "Using cached Graph access token."
        }
        else {
            Write-GraphMailLog "Cached Graph access token is missing or too close to expiry; requesting a new token."
        }
    }

    if (-not $UseCachedToken) {
        $TokenUri = [string]::Format($Config.TokenUriFormat, $Config.TenantId)

        $AuthBody = @{
            client_id     = $Config.AppId
            scope         = $Config.GraphScope
            client_secret = $AppSec
            grant_type    = "client_credentials"
        }

        try {
            $TokenRequest = Invoke-WebRequest `
                -Method Post `
                -Uri $TokenUri `
                -ContentType "application/x-www-form-urlencoded" `
                -Body $AuthBody `
                -UseBasicParsing `
                -ErrorAction Stop

            $Token = ($TokenRequest.Content | ConvertFrom-Json).access_token
        }
        catch {
            Write-GraphMailLog "Token request failed: $($_.Exception.Message)" -Level ERROR
            return (New-FailureResult -ErrorMessage "Token request failed: $($_.Exception.Message)")
        }

        if (-not $Token) {
            Write-GraphMailLog "Access token was not obtained; token response contained no access_token value." -Level ERROR
            return (New-FailureResult -ErrorMessage "Access token was not obtained.")
        }

        $script:GraphMailTokenCache[$TokenCacheKey] = [PSCustomObject]@{
            Token        = $Token
            ExpiresAfter = (Get-Date).AddMinutes(50)
        }

        Write-GraphMailLog "Access token acquired and cached for current session."
    }

    $Headers = @{
        'Content-Type'  = 'application/json'
        'Authorization' = "Bearer $Token"
    }

    #endregion

    #region ========================= HTML ASSEMBLY =========================

    $HtmlBodySegment1    = ""
    $HtmlBodySegment2    = ""
    $HtmlTableWithBorder = ""

    try {
        if (-not [string]::IsNullOrWhiteSpace($body1)) {
            $HtmlBodySegment1 = "<p>" + ($body1 -replace "`r?`n", "<br>") + "</p>"
        }

        if (-not [string]::IsNullOrWhiteSpace($body2)) {
            $HtmlBodySegment2 = "<p>" + ($body2 -replace "`r?`n", "<br>") + "</p>"
        }

        if ($null -ne $table) {
            $HtmlTable = $table | ConvertTo-Html -Fragment -As Table

            $HtmlTableWithBorder = (
                $HtmlTable `
                    -replace "<table>", "<table style='border-collapse: collapse; border: 1px solid black;'>" `
                    -replace "<td>",    "<td style='border: 1px solid black; padding: 5px;'>" `
                    -replace "<th>",    "<th style='border: 1px solid black; padding: 5px;'>"
            ) -join "`r`n"
        }
    }
    catch {
        Write-GraphMailLog "HTML assembly failed: $($_.Exception.Message)" -Level ERROR
        return (New-FailureResult -ErrorMessage "HTML assembly failed: $($_.Exception.Message)")
    }

    #endregion

    #region ========================= ATTACHMENT ASSEMBLY =========================

    $AttachmentList = [System.Collections.Generic.List[hashtable]]::new()

    foreach ($AttachmentPath in $attachments) {
        if ([string]::IsNullOrWhiteSpace($AttachmentPath)) {
            continue
        }

        if (-not (Test-Path $AttachmentPath)) {
            Write-GraphMailLog "Attachment not found, skipping: $AttachmentPath" -Level WARN
            continue
        }

        try {
            $AttachmentName    = [System.IO.Path]::GetFileName($AttachmentPath)
            $AttachmentContent = [System.IO.File]::ReadAllBytes($AttachmentPath)

            $AttachmentList.Add(@{
                "@odata.type"  = "#microsoft.graph.fileAttachment"
                "name"         = $AttachmentName
                "contentBytes" = [Convert]::ToBase64String($AttachmentContent)
            })

            Write-GraphMailLog "Attached: $AttachmentName"
        }
        catch {
            Write-GraphMailLog "Failed to read attachment '$AttachmentPath': $($_.Exception.Message)" -Level WARN
        }
    }

    #endregion

    #region ========================= MESSAGE BUILD =========================

    $EmailMessage = @{
        message = @{
            subject = $subject
            body    = @{
                contentType = 'HTML'
                content     = "$HtmlBodySegment1$HtmlTableWithBorder$HtmlBodySegment2"
            }
            toRecipients = @(
                $recipients | ForEach-Object {
                    @{
                        emailAddress = @{
                            address = $_
                        }
                    }
                }
            )
            attachments = @($AttachmentList)
        }
    }

    $MessageParams = @{
        URI         = [string]::Format($Config.SendMailUriFmt, $sender)
        Headers     = $Headers
        Method      = "POST"
        ContentType = 'application/json'
        Body        = $EmailMessage | ConvertTo-Json -Depth 6
    }

    #endregion

    #region ========================= SEND =========================

    try {
        $Result = Invoke-RestMethod @MessageParams -ErrorAction Stop

        Write-GraphMailLog "Message sent successfully to: $($recipients -join ', ')"
        Write-GraphMailLog "========== Send-GraphMailMessage Completed Successfully =========="

        if ($null -ne $Result) {
            return $Result
        }
        else {
            return (New-SuccessResult)
        }
    }
    catch {
        $ParsedError = Get-GraphErrorDetails -ErrorRecord $_

        $CallingApp = "Interactive / Unknown"

        try {
            $Stack      = Get-PSCallStack
            $OuterFrame = $Stack[$Stack.Length - 1]

            if ($OuterFrame.ScriptName) {
                $CallingApp = $OuterFrame.ScriptName
            }
        }
        catch {
            # Defensive only. If call-stack inspection fails, keep default.
        }

        $FormattedErrorDetails = ($ParsedError.ErrorDetails | Out-String).Trim()

        $FailureMessage = @"
API send failed.
StatusCode        : $($ParsedError.StatusCode)
StatusDescription : $($ParsedError.StatusDescription)
DuringScript      : $CallingApp
ErrorDetails      :
$FormattedErrorDetails
RawResponse       : $($ParsedError.RawResponse)
MessageParameters :
Subject    : $subject
Sender     : $sender
Recipients : $($recipients -join ', ')
"@

        Write-GraphMailLog $FailureMessage -Level ERROR
        Write-GraphMailLog "========== Send-GraphMailMessage Completed With Errors =========="

        return (New-FailureResult -ErrorMessage "Graph sendMail failed: $($ParsedError.ErrorDetails.Message)")
    }

    #endregion
}
