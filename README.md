# Automated Inbox Sorting

A PowerShell script that automatically sorts emails from a shared Outlook inbox into customer folders — no Microsoft Graph API permissions, no admin approval, no third-party tools required.

It uses Outlook COM automation, meaning it runs directly against your locally installed Outlook app. If Outlook is open and you have access to the mailbox, the script will work.

---

## The Problem

Outlook's built-in Rules work well for simple cases, but they fall apart quickly in a shared support or business inbox:

- Rules can't reliably sort **forwarded or replied threads** where the original customer email is buried in the chain
- Rules are limited to **256 total** across the mailbox
- **Internal emails** about a customer — where every address in the To/From fields is your own domain — have no external signal for a rule to match on
- Managing dozens of rules for dozens of customers becomes a maintenance burden

---

## How It Works

Each email is evaluated against four matching steps in order. As soon as a match is found, the email is moved and the script moves on.

### Step 1 — Sender Domain Match
The simplest case. If the sender's email domain is in your customer map, the email goes to that customer's folder.

```
From: contact@customerdomain.com  →  matches "customerdomain.com"  →  moved to Customer Folder
```

### Step 2 — Recipient Domain Match
Catches outbound emails and replies sent *to* a customer. Checks the To and CC fields for any external customer domain.

```
To: contact@customerdomain.com  →  matches "customerdomain.com"  →  moved to Customer Folder
```

### Step 3 — Body Scan
The most powerful step. When the sender is internal (your own domain), the script strips the email body and scans every `From:` line in the quoted chain for an external customer email address. This catches:

- Forwarded customer emails (`FW:`)
- Internal replies to customer threads (`RE:`)
- Escalations passed between staff members

Vendor domains can be excluded so they don't trigger false matches ahead of the real customer.

```
From: staff@yourcompany.com
Body: "...From: Customer Contact <contact@customerdomain.com>..."
→  body scan finds customerdomain.com  →  moved to Customer Folder
```

### Step 4 — Subject Keyword Match
Last resort for fully internal threads where no external address exists anywhere in the chain. If the subject contains a configured keyword, the email is matched on that basis.

```
Subject: "RE: CustomeName account query"  →  keyword match  →  moved to Customer Folder
```

### Automatic Reply Deletion
Any email with "Automatic Reply" in the subject is moved to Deleted Items before matching begins.

---

## Files

| File | Purpose |
|------|---------|
| `Automated Inbox Sorting-TEMPLATE.ps1` | Blank script — fill in your own details |
| `Automated Inbox Sorting-TEMPLATE.bat` | Launcher for the template script |
| `Automated Inbox Sorting-EXAMPLE.ps1` | Pre-filled example using fictional data |
| `Automated Inbox Sorting-EXAMPLE.bat` | Launcher for the example script |

> Your production script (filled-in template) should be kept locally and **never committed to a public repository** as it will contain your mailbox address, domain names, and folder structure.

---

## Requirements

- Windows PC with Microsoft Outlook installed and running
- Access to the inbox you want to sort (delegated access or logged in as that account)
- PowerShell (built into Windows — no installation needed)
- Customer folders must already exist inside the inbox in Outlook before running

---

## Getting Started

### 1. Download the template files
Download `Automated Inbox Sorting-TEMPLATE.ps1` and `Automated Inbox Sorting-TEMPLATE.bat` and place them in the same folder.

### 2. Rename both files
Rename them to something meaningful for your context, keeping the same naming pattern for both:

```
My Company Inbox Sorting-COM.ps1
My Company Inbox Sorting-COM.bat
```

> The .bat file references the .ps1 by name, so both must be renamed consistently and kept in the same folder.

### 3. Open the .ps1 in Notepad
Right-click the .ps1 file and choose **Open with > Notepad**. Do not double-click — this will try to run it.

### 4. Fill in the configuration

#### Mailbox address
Set this to the email address of the inbox you want to sort:
```powershell
$TargetMailbox = "support@yourcompany.com"
```
If you are logged in as that account directly, set it to `$null`.

#### Internal domains
Add every domain used by your organisation. Emails from these are treated as internal and trigger the body scan:
```powershell
$InternalDomains = @(
    "yourcompany.com"
    "relatedcompany.com"
)
```

#### Customer domain map
Map each customer's email domain to their folder name. The folder name must match exactly what exists in Outlook:
```powershell
$CustomerMap = @{
    "customerdomain.com"     = "Customer Folder Name"
    "anotherdomain.com.au"   = "Another Customer"
    "olddomain.com"          = "Another Customer"   # multiple domains, same folder
}
```

#### Subject keywords
Used when no external address exists anywhere in the thread:
```powershell
$SubjectKeywords = @{
    "Customer Name"    = "Customer Folder Name"
    "CustAbbrev"       = "Customer Folder Name"
}
```

#### Vendor domains (optional)
Third-party vendors that appear in threads but are not customers. Their domains are skipped during body scanning:
```powershell
$VendorDomains = @(
    "vendordomain.com"
)
```

### 5. Save and run
Save the file, then double-click the .bat file to run. Make sure Outlook is open first.

---

## Output

The script prints results to the screen as it runs:

```
  MOVED  [contact@customerdomain.com] -> Customer Folder (direct domain match)
         Subject: Your subject here

  LEFT   [hello@unknowncompany.com] No matching customer (left in inbox)
         Subject: Your subject here

  DELETED [someone@domain.com] Automatic reply removed
          Subject: Automatic Reply: Out of office
```

A summary is shown at the end of each run:

```
  Total emails processed     : 24
  Moved to folders           : 21
  Left in inbox              : 2
  Automatic replies deleted  : 1
```

---

## Logs

Every run creates a timestamped log file in a `Logs` folder next to the script:

```
Inbox Sorter-2026-03-10_11-30-00.txt
```

The log captures everything printed to the screen. If an email ever ends up in the wrong folder, the log will show exactly what the script saw and what decision it made.

---

## Ongoing Maintenance

**Adding a new customer**
1. Create the folder in Outlook first
2. Add their domain to `$CustomerMap`
3. Add their name to `$SubjectKeywords`

**Adding an extra domain for an existing customer**
Just add another line in `$CustomerMap` pointing to the same folder name.

**Adding a vendor to exclude**
Add their domain to `$VendorDomains`.

---

## Using an AI Assistant to Configure the Script

If you are not comfortable editing the script manually, you can paste the contents of the template into any AI chat tool (Claude, ChatGPT, Gemini, etc.) and ask it to fill in the configuration for you. For example:

> *"Here is a PowerShell script. Please fill in the configuration section with the following: my mailbox is support@mycompany.com, my internal domain is mycompany.com, and my customers are: [list of customers and domains]"*

The AI will return a fully populated script ready to paste into Notepad and save.

---

## Troubleshooting

**"Could not connect to Outlook"**
Outlook is not open. Open Outlook and sign in before running the script.

**"Could not locate inbox"**
The mailbox address in `$TargetMailbox` is wrong, or you do not have delegated access. Check the address and your Outlook permissions.

**WARNING: Folder not found**
The folder name in `$CustomerMap` or `$SubjectKeywords` does not match what exists in Outlook. Check for typos, extra spaces, or capitalisation differences.

**Script is blocked / will not run**
If downloaded from the internet, Windows may block the file. Instead, open the .ps1 in Notepad, copy all the content, paste it into a new Notepad window, and save it locally as a new .ps1 file. Files created locally are not subject to this restriction.

**Emails not matching**
Check the LEFT entries in the output. If the sender domain looks correct, verify it is spelled exactly right in `$CustomerMap`. For internal threads, make sure the sender's domain is listed in `$InternalDomains`.

---

## License

MIT — free to use, modify, and share.
