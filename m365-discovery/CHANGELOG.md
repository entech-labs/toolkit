# Changelog

## 1.0.0 — 3 October 2026

First public release.

- Tenant, verified domains and directory-sync status
- Licensing: purchased versus assigned by SKU
- Accounts: total, enabled, disabled, guests, licensed, directory-synced
- Mailbox sizing from the Microsoft 365 usage reports, with the largest listed
- Mail data split by mailbox type — user, shared, room, equipment
- OneDrive and SharePoint storage and file counts, largest listed
- SharePoint composition: document libraries versus custom lists, per site
- Teams and groups, including private channel counts
- Exchange structure: shared and resource mailboxes, archives, litigation holds,
  transport rules, connectors, public folders
- Applications using single sign-on, and conditional access policy counts
- Branded HTML report, PDF conversion via headless Edge or Chrome, CSV extracts, zip
- `-SelfTest` builds a full sample report without contacting any tenant
- Collection failures are reported in the output rather than silently skipped
- Detects whether tenant reports are anonymised, and says so
