# =============================================================================
# Automated Inbox Sorting-TEMPLATE.ps1
# Automatically sorts emails in a shared support inbox into customer folders.
# Uses Outlook COM automation - no Microsoft Graph API or admin approval needed.
# Requires: Outlook desktop app installed and running on this machine.
#
# SETUP INSTRUCTIONS:
# 1. Fill in your mailbox address in the MAILBOX SETTINGS section
# 2. Fill in your internal domains in the INTERNAL DOMAINS section
# 3. Add your customers and their domains to $CustomerMap
# 4. Add subject keywords for each customer to $SubjectKeywords
# 5. Add any vendor/third party domains to $VendorDomains
# 6. Save as a .ps1 file locally (do not download - save via Notepad to avoid
#    Windows blocking the script)
# 7. Set your age threshold in the AGE THRESHOLD section (default: 24 hours)
# 8. Run via the provided .bat file
# =============================================================================

# -----------------------------------------------------------------------------
# CUSTOMER DOMAIN MAP
# Format: "senderdomain.com" = "Exact Folder Name"
# The folder name must match exactly how it appears in Outlook.
# You can map multiple domains to the same folder (e.g. rebranded companies).
# -----------------------------------------------------------------------------
$CustomerMap = @{
    # "customer1domain.com"     = "Customer 1 Folder Name"
    # "customer2domain.com.au"  = "Customer 2 Folder Name"
    # "customer2old.com.au"     = "Customer 2 Folder Name"   # old domain, same folder
    # Add your customers below:

}

# -----------------------------------------------------------------------------
# MAILBOX SETTINGS
# Set $TargetMailbox to the email address of the shared inbox you want to sort.
# If running directly as the inbox account, set it to $null.
# -----------------------------------------------------------------------------
$TargetMailbox = "support@yourcompany.com"

# -----------------------------------------------------------------------------
# INTERNAL DOMAINS
# Emails from these domains are treated as internal.
# Internal emails trigger a body scan to find the original external customer.
# Add all domains used by your organisation (e.g. parent company, sister brands).
# -----------------------------------------------------------------------------
$InternalDomains = @(
    "yourcompany.com"
    # "relatedcompany.com"
)

# -----------------------------------------------------------------------------
# VENDOR DOMAINS
# Third party vendors that appear in email threads but are not customers.
# Emails from these domains will be skipped during body scan so they don't
# accidentally trigger a match ahead of the real customer domain.
# -----------------------------------------------------------------------------
$VendorDomains = @(
    # "vendordomain.com"
)

# -----------------------------------------------------------------------------
# SUBJECT KEYWORD MAP
# Used as a last resort for fully internal threads where no external customer
# domain exists anywhere in the email chain.
# Format: "keyword" = "Exact Folder Name"
# Keywords are case insensitive. Add common abbreviations as separate entries.
# -----------------------------------------------------------------------------
$SubjectKeywords = @{
    # "Customer 1"      = "Customer 1 Folder Name"
    # "Cust1"           = "Customer 1 Folder Name"   # common abbreviation
    # "Customer 2"      = "Customer 2 Folder Name"
    # Add your keywords below:

}


# -----------------------------------------------------------------------------
# AGE THRESHOLD (OPTIONAL)
# Emails newer than this many hours will be skipped and left in the inbox.
# This gives you a buffer to read fresh emails before they are sorted.
# Set to 0 to disable and sort all emails regardless of age.
# Default is 24 hours.
# -----------------------------------------------------------------------------
$AgeThresholdHours = 24

# -----------------------------------------------------------------------------
# SCRIPT — No changes needed below this line
# -----------------------------------------------------------------------------

# Start logging - creates a Logs folder next to the script, one file per run
$LogFolder = Join-Path $PSScriptRoot "Logs"
if (-not (Test-Path $LogFolder)) { New-Item -ItemType Directory -Path $LogFolder | Out-Null }
$LogFile = Join-Path $LogFolder ("Inbox Sorter-" + (Get-Date -Format "yyyy-MM-dd_HH-mm-ss") + ".txt")
Start-Transcript -Path $LogFile

Write-Host ""
Write-Host "=============================================" -ForegroundColor Cyan
Write-Host "  Automated Inbox Sorting" -ForegroundColor Cyan
Write-Host "=============================================" -ForegroundColor Cyan
Write-Host ""

# Connect to Outlook
Write-Host "Connecting to Outlook..." -ForegroundColor Yellow
try {
    $Outlook   = New-Object -ComObject Outlook.Application -ErrorAction Stop
    $Namespace = $Outlook.GetNamespace("MAPI")
    Write-Host "Connected to Outlook successfully." -ForegroundColor Green
} catch {
    Write-Host "ERROR: Could not connect to Outlook. Is Outlook installed and running?" -ForegroundColor Red
    Write-Host $_.Exception.Message
    Stop-Transcript
    exit
}

Write-Host ""

# Get the inbox
Write-Host "Locating inbox..." -ForegroundColor Yellow
try {
    if ($TargetMailbox) {
        $Recipient = $Namespace.CreateRecipient($TargetMailbox)
        $Recipient.Resolve() | Out-Null
        $Inbox = $Namespace.GetSharedDefaultFolder($Recipient, 6)
    } else {
        $Inbox = $Namespace.GetDefaultFolder(6)
    }
    Write-Host "Inbox located: $($Inbox.FolderPath)" -ForegroundColor Green
} catch {
    Write-Host "ERROR: Could not locate inbox. Check the mailbox address in configuration." -ForegroundColor Red
    Write-Host $_.Exception.Message
    Stop-Transcript
    exit
}

# Build a lookup of all subfolders inside the Inbox
$FolderLookup = @{}
foreach ($Folder in $Inbox.Folders) {
    $FolderLookup[$Folder.Name.ToLower()] = $Folder
}

# Verify all configured folders exist
Write-Host "Mapping customer folders..." -ForegroundColor Yellow
$UniqueFolders = $CustomerMap.Values | Sort-Object -Unique
foreach ($FolderName in $UniqueFolders) {
    if (-not $FolderLookup.ContainsKey($FolderName.ToLower())) {
        Write-Host "  WARNING: Folder '$FolderName' not found in mailbox." -ForegroundColor Yellow
    }
}
Write-Host "  All customer folders verified." -ForegroundColor Green

Write-Host ""

# Get emails from inbox
$Emails      = $Inbox.Items

Write-Host "Found $($Emails.Count) email(s) in inbox. Processing all..." -ForegroundColor White
Write-Host ""
Write-Host "Processing..." -ForegroundColor Yellow
Write-Host ""

$MovedCount   = 0
$SkippedCount = 0
$ErrorCount   = 0
$DeletedCount = 0
$MovedSummary = @()
$Processed    = 0

for ($i = $Emails.Count; $i -ge 1; $i--) {
    $Email = $Emails.Item($i)

    if ($Email.Class -ne 43) { continue }

    # Age threshold check - skip emails newer than $AgeThresholdHours
    if ($AgeThresholdHours -gt 0) {
        $EmailAge = (Get-Date) - $Email.ReceivedTime
        if ($EmailAge.TotalHours -lt $AgeThresholdHours) {
            Write-Host "  SKIPPED [$($Email.SenderEmailAddress)] Too recent ($([math]::Round($EmailAge.TotalHours, 1))h old, threshold: ${AgeThresholdHours}h)" -ForegroundColor DarkYellow
            Write-Host "          Subject: $($Email.Subject)" -ForegroundColor DarkGray
            $SkippedCount++
            continue
        }
    }

    $SenderAddress    = $Email.SenderEmailAddress
    $Subject          = $Email.Subject
    $TargetFolderName = $null
    $MatchMethod      = $null
    $Processed++

    # Exchange DN format means the sender is internal
    if ($SenderAddress -match "^/O=") {
        $SenderDomain = $InternalDomains[0]
    } else {
        $SenderDomain = ($SenderAddress -split "@")[-1].ToLower()
    }

    # Pre-filter: Move automatic replies to Deleted Items
    if ($Subject -match "(?i)automatic reply") {
        $Email.Delete()
        Write-Host "  DELETED [$SenderAddress] Automatic reply removed" -ForegroundColor DarkGray
        Write-Host "          Subject: $Subject" -ForegroundColor DarkGray
        $DeletedCount++
        continue
    }

    # Step 1: Direct domain match on sender
    if ($CustomerMap.ContainsKey($SenderDomain)) {
        $TargetFolderName = $CustomerMap[$SenderDomain]
        $MatchMethod      = "direct domain match"
    }

    # Step 1b: Direct domain match on recipients (To/CC fields)
    if (-not $TargetFolderName) {
        $Recipients = $Email.Recipients
        foreach ($Recipient in $Recipients) {
            $RecipientAddress = $Recipient.Address
            $RecipientDomain  = ($RecipientAddress -split "@")[-1].ToLower()
            if ($InternalDomains -contains $RecipientDomain) { continue }
            if ($CustomerMap.ContainsKey($RecipientDomain)) {
                $TargetFolderName = $CustomerMap[$RecipientDomain]
                $MatchMethod      = "recipient domain match"
                break
            }
        }
    }

    # Step 2: Internal sender - scan body From: lines for original customer sender
    if (-not $TargetFolderName) {
        $IsInternal = $InternalDomains -contains $SenderDomain

        if ($IsInternal) {
            $BodyText = $Email.Body
            if ([string]::IsNullOrWhiteSpace($BodyText)) {
                $BodyText = $Email.HTMLBody -replace "<[^>]+>", " "
            }

            $FromLines = [regex]::Matches($BodyText, '(?i)From\s*:\s*.+?([a-zA-Z0-9._%+-]+@([a-zA-Z0-9.-]+\.[a-zA-Z]{2,}))')

            foreach ($Match in $FromLines) {
                $FoundDomain = $Match.Groups[2].Value.ToLower()
                if ($InternalDomains -contains $FoundDomain) { continue }
                if ($VendorDomains -contains $FoundDomain) { continue }

                if ($CustomerMap.ContainsKey($FoundDomain)) {
                    $TargetFolderName = $CustomerMap[$FoundDomain]
                    $MatchMethod      = "via body scan"
                    break
                }
            }
        }
    }

    # Step 3: Subject keyword match - last resort for fully internal threads
    if (-not $TargetFolderName) {
        foreach ($Keyword in $SubjectKeywords.Keys) {
            if ($Subject -imatch [regex]::Escape($Keyword)) {
                $TargetFolderName = $SubjectKeywords[$Keyword]
                $MatchMethod      = "subject keyword match"
                break
            }
        }
    }

    # Step 4: Move or skip
    if ($TargetFolderName) {
        $TargetFolder = $FolderLookup[$TargetFolderName.ToLower()]

        if (-not $TargetFolder) {
            Write-Host "  SKIP  [$SenderAddress] -> '$TargetFolderName' (folder not found)" -ForegroundColor Yellow
            Write-Host "        Subject: $Subject" -ForegroundColor DarkGray
            $SkippedCount++
            continue
        }

        try {
            $Email.Move($TargetFolder) | Out-Null
            Write-Host "  MOVED [$SenderAddress] -> $TargetFolderName ($MatchMethod)" -ForegroundColor Green
            Write-Host "        Subject: $Subject" -ForegroundColor DarkGray
            $MovedSummary += [PSCustomObject]@{
                From    = $SenderAddress
                Subject = $Subject
                Folder  = $TargetFolderName
                Method  = $MatchMethod
            }
            $MovedCount++
        } catch {
            Write-Host "  ERROR [$SenderAddress] Could not move email." -ForegroundColor Red
            Write-Host "        $($_.Exception.Message)" -ForegroundColor DarkRed
            $ErrorCount++
        }
    } else {
        Write-Host "  LEFT  [$SenderAddress] No matching customer (left in inbox)" -ForegroundColor DarkGray
        Write-Host "        Subject: $Subject" -ForegroundColor DarkGray
        $SkippedCount++
    }
}

# -----------------------------------------------------------------------------
# SUMMARY
# -----------------------------------------------------------------------------
Write-Host ""
Write-Host "=============================================" -ForegroundColor Cyan
Write-Host "  Run Complete" -ForegroundColor Cyan
Write-Host "=============================================" -ForegroundColor Cyan
Write-Host "  Total emails processed     : $Processed" -ForegroundColor White
Write-Host "  Moved to folders           : $MovedCount" -ForegroundColor Green
Write-Host "  Left in inbox              : $SkippedCount" -ForegroundColor DarkGray
    Write-Host "    (includes emails skipped as too recent)" -ForegroundColor DarkGray
Write-Host "  Automatic replies deleted  : $DeletedCount" -ForegroundColor DarkGray
if ($ErrorCount -gt 0) {
    Write-Host "  Errors                     : $ErrorCount" -ForegroundColor Red
}
Write-Host ""

if ($MovedCount -gt 0) {
    Write-Host "Moved emails breakdown:" -ForegroundColor Cyan
    $MovedSummary | Group-Object Folder | ForEach-Object {
        Write-Host "  $($_.Name): $($_.Count) email(s)" -ForegroundColor White
    }
    $BodyScanMoves = $MovedSummary | Where-Object { $_.Method -eq "via body scan" }
    if ($BodyScanMoves.Count -gt 0) {
        Write-Host ""
        Write-Host "  ($($BodyScanMoves.Count) matched via body scan)" -ForegroundColor DarkGray
    }
}

Write-Host ""
Stop-Transcript
