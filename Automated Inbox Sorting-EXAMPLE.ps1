# =============================================================================
# Automated Inbox Sorting-EXAMPLE.ps1
# DRY RUN / MOCK VERSION - Tests matching logic only, no Outlook connection.
# No emails are moved. Run this to verify logic before going live.
# =============================================================================

$CustomerMap = @{
    "apple.com"             = "Apple"
    "nike.com"              = "Nike"
    "tesla.com"             = "Tesla"
    "spotify.com"           = "Spotify"
    "netflix.com"           = "Netflix"
    "samsung.com"           = "Samsung"
    "google.com"            = "Google"
    "microsoft.com"         = "Microsoft"
    "amazon.com"            = "Amazon"
    "adidas.com"            = "Adidas"
}

$InternalDomains = @("yourcompany.com", "relatedcompany.com")

$VendorDomains = @(
    "somevendor.com"
)

$SubjectKeywords = @{
    "Apple"         = "Apple"
    "Nike"          = "Nike"
    "Tesla"         = "Tesla"
    "Spotify"       = "Spotify"
    "Netflix"       = "Netflix"
    "Samsung"       = "Samsung"
    "Google"        = "Google"
    "Microsoft"     = "Microsoft"
    "Amazon"        = "Amazon"
    "Adidas"        = "Adidas"
}

# Age threshold - emails newer than this many hours are skipped (0 = disabled)
$AgeThresholdHours = 24

# -----------------------------------------------------------------------------
# MOCK EMAILS
# -----------------------------------------------------------------------------
$MockEmails = @(

    # Should match: direct sender domain (48h old - will be sorted)
    [PSCustomObject]@{
        Subject    = "Order issue"
        From       = "taylorswift@apple.com"
        Recipients = @("support@yourcompany.com")
        Body       = "Hi, I am having an issue with my order."
        AgeHours   = 48
    },

    # Should be skipped: too recent (2h old, under 24h threshold)
    [PSCustomObject]@{
        Subject    = "RMA Request"
        From       = "elonmusk@tesla.com"
        Recipients = @("support@yourcompany.com")
        Body       = "Please process the following RMA."
        AgeHours   = 2
    },

    # Should match: recipient domain
    [PSCustomObject]@{
        Subject    = "Following up"
        From       = "staff@yourcompany.com"
        Recipients = @("beyonce@spotify.com")
        Body       = "Hi, just following up on your case."
    
        AgeHours   = 48
    },

    # Should match: internal sender, external domain in body (RE: thread)
    [PSCustomObject]@{
        Subject    = "RE: Streaming issue case 1234"
        From       = "staff@yourcompany.com"
        Recipients = @("support@yourcompany.com")
        Body       = "Thanks for your email. From: Rihanna <rihanna@netflix.com> Subject: Re: Streaming issue case 1234"
    
        AgeHours   = 48
    },

    # Should match: internal sender, external domain in body (FW: thread)
    [PSCustomObject]@{
        Subject    = "FW: Faulty device - RMA98765"
        From       = "staff@yourcompany.com"
        Recipients = @("support@yourcompany.com")
        Body       = "FYI. From: Billie Eilish <billie@samsung.com> To: support@yourcompany.com Subject: Faulty device"
    
        AgeHours   = 48
    },

    # Should match: Exchange DN sender treated as internal, body scan finds customer
    [PSCustomObject]@{
        Subject    = "RE: Faulty device - RMA98765"
        From       = "/O=YOURCOMPANY/OU=EXCHANGE ADMINISTRATIVE GROUP/CN=RECIPIENTS/CN=STAFF"
        Recipients = @("support@yourcompany.com")
        Body       = "From: Billie Eilish <billie@samsung.com> Subject: Re: Faulty device"
    
        AgeHours   = 48
    },

    # Should match: related company sender treated as internal, body scan finds customer
    [PSCustomObject]@{
        Subject    = "RE: Delivery query"
        From       = "drake@relatedcompany.com"
        Recipients = @("support@yourcompany.com")
        Body       = "From: support@yourcompany.com Subject: RE: Delivery query From: Kendrick Lamar <kendrick@amazon.com>"
    
        AgeHours   = 48
    },

    # Should match: subject keyword (lowercase - tests case insensitivity)
    [PSCustomObject]@{
        Subject    = "RE: microsoft licensing discussion"
        From       = "staff@yourcompany.com"
        Recipients = @("other.staff@yourcompany.com")
        Body       = "All internal, no external addresses here."
    
        AgeHours   = 48
    },

    # Should match: fully internal thread with brand in subject
    [PSCustomObject]@{
        Subject    = "RE: Google account access for Adele"
        From       = "staff@yourcompany.com"
        Recipients = @("manager@yourcompany.com", "support@yourcompany.com")
        Body       = "All internal discussion about the Google account."
    
        AgeHours   = 48
    },

    # Should match: vendor domain skipped, real customer found after
    [PSCustomObject]@{
        Subject    = "RE: Shoe sizing complaint"
        From       = "staff@yourcompany.com"
        Recipients = @("support@yourcompany.com")
        Body       = "From: Vendor Contact <contact@somevendor.com> From: Kanye West <kanye@adidas.com>"
    
        AgeHours   = 48
    },

    # Should NOT match: internal email with no customer domain
    [PSCustomObject]@{
        Subject    = "RE: Team lunch Friday"
        From       = "staff@yourcompany.com"
        Recipients = @("everyone@yourcompany.com")
        Body       = "Sounds good, see you there."
    
        AgeHours   = 48
    },

    # Should NOT match: unknown external domain
    [PSCustomObject]@{
        Subject    = "Partnership inquiry"
        From       = "hello@unknowncompany.com"
        Recipients = @("support@yourcompany.com")
        Body       = "We would like to discuss a partnership."
    
        AgeHours   = 48
    },

    # Should be deleted: automatic reply
    [PSCustomObject]@{
        Subject    = "Automatic Reply: Out of office"
        From       = "arianagrande@apple.com"
        Recipients = @("support@yourcompany.com")
        Body       = "I am out of the office until Monday."
    
        AgeHours   = 48
    }
)

# -----------------------------------------------------------------------------
# EXAMPLE RUNNER
# -----------------------------------------------------------------------------

# Start logging
$LogFolder = Join-Path $PSScriptRoot "Logs"
if (-not (Test-Path $LogFolder)) { New-Item -ItemType Directory -Path $LogFolder | Out-Null }
$LogFile = Join-Path $LogFolder ("Inbox Sorter-" + (Get-Date -Format "yyyy-MM-dd_HH-mm-ss") + ".txt")
Start-Transcript -Path $LogFile

Write-Host ""
Write-Host "=============================================" -ForegroundColor Cyan
Write-Host "  Automated Inbox Sorting - DRY RUN / EXAMPLE" -ForegroundColor Cyan
Write-Host "=============================================" -ForegroundColor Cyan
Write-Host "  No emails will be moved. Logic test only." -ForegroundColor Yellow
Write-Host "=============================================" -ForegroundColor Cyan
Write-Host ""

$MatchedCount = 0
$SkippedCount = 0
$DeletedCount = 0

foreach ($Email in $MockEmails) {
    $SenderAddress    = $Email.From
    $Subject          = $Email.Subject
    $BodyText         = $Email.Body
    $TargetFolderName = $null
    $MatchMethod      = $null

    # Exchange DN format means sender is internal
    if ($SenderAddress -match "^/O=") {
        $SenderDomain = $InternalDomains[0]
    } else {
        $SenderDomain = ($SenderAddress -split "@")[-1].ToLower()
    }

    # Age threshold check - skip emails newer than $AgeThresholdHours
    if ($AgeThresholdHours -gt 0) {
        $SimulatedAge = $Email.AgeHours
        if ($SimulatedAge -lt $AgeThresholdHours) {
            Write-Host "  WOULD SKIP [$SenderAddress] Too recent (simulated ${SimulatedAge}h old, threshold: ${AgeThresholdHours}h)" -ForegroundColor DarkYellow
            Write-Host "             Subject: $Subject" -ForegroundColor DarkGray
            $SkippedCount++
            continue
        }
    }

    # Pre-filter: automatic replies
    if ($Subject -match "(?i)automatic reply") {
        Write-Host "  WOULD DELETE [$SenderAddress] Automatic reply" -ForegroundColor DarkGray
        Write-Host "               Subject: $Subject" -ForegroundColor DarkGray
        $DeletedCount++
        continue
    }

    # Step 1: Direct sender domain match
    if ($CustomerMap.ContainsKey($SenderDomain)) {
        $TargetFolderName = $CustomerMap[$SenderDomain]
        $MatchMethod      = "direct domain match"
    }

    # Step 1b: Recipient domain match
    if (-not $TargetFolderName) {
        foreach ($RecipientAddress in $Email.Recipients) {
            $RecipientDomain = ($RecipientAddress -split "@")[-1].ToLower()
            if ($InternalDomains -contains $RecipientDomain) { continue }
            if ($CustomerMap.ContainsKey($RecipientDomain)) {
                $TargetFolderName = $CustomerMap[$RecipientDomain]
                $MatchMethod      = "recipient domain match"
                break
            }
        }
    }

    # Step 2: Internal sender - scan body From: lines
    if (-not $TargetFolderName) {
        $IsInternal = $InternalDomains -contains $SenderDomain

        if ($IsInternal) {
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

    # Step 3: Subject keyword match
    if (-not $TargetFolderName) {
        foreach ($Keyword in $SubjectKeywords.Keys) {
            if ($Subject -imatch [regex]::Escape($Keyword)) {
                $TargetFolderName = $SubjectKeywords[$Keyword]
                $MatchMethod      = "subject keyword match"
                break
            }
        }
    }

    # Output result
    if ($TargetFolderName) {
        $MatchedCount++
        Write-Host "  WOULD MOVE  [$SenderAddress] -> $TargetFolderName ($MatchMethod)" -ForegroundColor Green
        Write-Host "              Subject: $Subject" -ForegroundColor DarkGray
    } else {
        $SkippedCount++
        Write-Host "  WOULD LEAVE [$SenderAddress] - no match (stays in inbox)" -ForegroundColor DarkGray
        Write-Host "              Subject: $Subject" -ForegroundColor DarkGray
    }
}

Write-Host ""
Write-Host "=============================================" -ForegroundColor Cyan
Write-Host "  Test Complete" -ForegroundColor Cyan
Write-Host "=============================================" -ForegroundColor Cyan
Write-Host "  Emails tested          : $($MockEmails.Count)" -ForegroundColor White
Write-Host "  Would be moved         : $MatchedCount" -ForegroundColor Green
Write-Host "  Would stay             : $SkippedCount" -ForegroundColor DarkGray
Write-Host "  Would be deleted       : $DeletedCount" -ForegroundColor DarkGray
Write-Host ""

Stop-Transcript

