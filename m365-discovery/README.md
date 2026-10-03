# Microsoft 365 tenant discovery

Reads a Microsoft 365 tenant and produces a report describing its size and shape, so a migration
estimate rests on measured figures rather than assumptions.

**[Full documentation (PDF)](docs/What-the-discovery-script-does.pdf)** — written for whoever has
to approve running it, not just for whoever runs it.

---

## Try it before you trust it

```powershell
.\EnTech-M365-Discovery.ps1 -SelfTest
```

Builds a complete report from obviously fictional sample data. **Contacts no tenant, requires no
sign-in, asks for no permission.** You see exactly what the output looks like first.

---

## What it does not do, and how to check

Every claim below can be checked by searching this file. We would rather you check than take our
word for it.

| Claim | How to verify |
|---|---|
| **Makes no change to your tenant** | Search for `Invoke-MgGraphRequest` — there are four, all `-Method GET`. Every Exchange cmdlet is a `Get-`. There is no `Set-`, `Update-` or `Disable-` cmdlet in the file |
| **Reads no message or document content** | No scope granting content access is requested. It reads counts, sizes, names and configuration only |
| **Sends nothing anywhere** | Search for `Invoke-WebRequest` and `Invoke-RestMethod` — there are none. The only outbound calls are to `graph.microsoft.com` and Exchange Online |
| **Stores no credential** | Sign-in is interactive through Microsoft's prompt; the session is disconnected at the end |
| **Only writes your own files** | `New-Item` creates the output folder, `Remove-Item` replaces a previous zip of the same name. Both local |

---

## What it collects

| Area | Why it matters for a migration |
|---|---|
| Tenant, verified domains, directory sync | A domain lives in one tenant at a time — this is what makes cutover a fixed window |
| Licensing, purchased vs assigned | Which workloads are actually entitled |
| Accounts: total, enabled, disabled, guests | Disabled and unlicensed accounts often need not move at all |
| Mailbox count, sizes, largest, archives, holds | A few large mailboxes drive the schedule more than the average |
| **Mail data split by mailbox type** | Where the data really sits — user vs shared vs rooms vs system accounts |
| OneDrive storage and file counts | File *count* matters as much as volume |
| SharePoint sites and storage | |
| **SharePoint composition** | Document libraries are copied; custom lists, pages and forms are rebuilt by hand |
| Teams, private channels, groups | Private channels each have a hidden site and migrate badly |
| Transport rules, connectors, public folders | None of these migrate — each is rebuilt |
| Applications with SSO, conditional access | The most commonly under-estimated item |

Anything it could not collect is listed in the report under **Sections that could not be
collected**, with the reason. A silently missing section would mean an estimate built on a gap
nobody knew about.

---

## Requirements

- Windows, PowerShell 5.1 (built in) or PowerShell 7
- An account with **Global Reader** (or Global Administrator)
- Two Microsoft modules, installed for the current user only — the script asks first:
  `Microsoft.Graph.Authentication` and `ExchangeOnlineManagement`

All requested Graph scopes are read scopes. They are listed with a reason for each in the PDF.

---

## Running it

```powershell
Unblock-File -Path .\EnTech-M365-Discovery.ps1       # downloaded files are blocked by default
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
.\EnTech-M365-Discovery.ps1
```

### Options

```powershell
.\EnTech-M365-Discovery.ps1 -SelfTest                # sample data, no tenant contacted
.\EnTech-M365-Discovery.ps1 -SkipExchange            # Graph sections only
.\EnTech-M365-Discovery.ps1 -ReportPeriod D90        # 90-day usage window (default D30)
.\EnTech-M365-Discovery.ps1 -OutputPath C:\Temp\Disc # somewhere other than the Desktop
```

---

## Output

- `M365-Discovery-Report.pdf` — the readable summary
- `M365-Discovery-Report.html` — the same report, if you prefer to print it yourself
- `data\*.csv` — the extracts behind every figure, so nothing in the report is unsourced
- a `.zip` of all of it

A few minutes for a small organisation. Most of that is the Teams and SharePoint sections, which
ask about each team and site in turn.

---

## Known behaviours worth expecting

**Hashes instead of names.** Microsoft 365 can be set to conceal user, group and site names in
reports. The script detects this and says so. Totals stay accurate; the detail becomes
unattributable until an administrator turns it off under Settings → Org settings → Reports.

**No PDF produced.** Conversion uses Microsoft Edge in headless mode. If neither Edge nor Chrome
is found, the HTML is left in place — open it and print to PDF.

**Conditional access blocks sign-in.** Some policies block PowerShell. Run from a compliant
device or have an administrator add a temporary exclusion. The script does not work around this,
and should not be able to.
