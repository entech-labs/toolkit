# EnTech Labs Toolkit

Small, read-only tools we use when assessing an environment — published openly so the people
being assessed can read them before running them.

EnTech Labs is the technology practice of [EnTech Engineering, P.C.](https://entech.nyc), an
engineering and construction management firm working across the New York metropolitan region.

---

## Why these are public

When we ask a client to run something inside their own environment, they are entitled to know
exactly what it does first. A script that arrives as an email attachment cannot be checked. One
that sits in a public repository can be — by their IT staff, by their security reviewer, or by
anyone they choose to ask.

So everything here follows the same rules:

| | |
|---|---|
| **Read-only** | No tool in this repository writes to, changes, or deletes anything in a client system |
| **No transmission** | Output is written locally. The client reads it, and decides whether to send it |
| **No stored credentials** | Authentication is interactive, through the vendor's own sign-in |
| **Verifiable claims** | Every safety claim in the documentation says how to check it in the source |

If you find something in here that contradicts any of the above, that is a bug and we want to
know. See [SECURITY.md](SECURITY.md).

---

## Tools

### [`m365-discovery/`](m365-discovery/) — Microsoft 365 tenant discovery

Reads a Microsoft 365 tenant and produces a report describing its size and shape: mailboxes and
their sizes, OneDrive, SharePoint, Teams, groups, mail flow, and connected applications. Built
for sizing a tenant-to-tenant migration without anybody counting mailboxes by hand.

Has a `-SelfTest` mode that produces a complete sample report without contacting any tenant or
requiring a sign-in — so you can see the output before granting any permission.

**⬇️ [Download the script](https://github.com/entech-labs/toolkit/releases/latest/download/EnTech-M365-Discovery.ps1)**
&middot; **[Step-by-step instructions, with pictures](m365-discovery/)**
&middot; **[What it does, in full (PDF)](m365-discovery/docs/What-the-discovery-script-does.pdf)**

<sub>New to this? The [step-by-step guide](m365-discovery/) walks through it from downloading the
file to sending the report, assuming no PowerShell experience.</sub>

---

## Using these

Download from [Releases](../../releases) rather than cloning, and check the SHA256 published
with each release against the file you downloaded:

```powershell
Get-FileHash .\EnTech-M365-Discovery.ps1 -Algorithm SHA256
```

Windows marks files downloaded from the internet as untrusted and will refuse to run them.
Unblock before running:

```powershell
Unblock-File -Path .\EnTech-M365-Discovery.ps1
```

Each tool's own README covers what it needs and how to run it.

---

## Licence

[MIT](LICENSE). Use them, read them, adapt them. They are provided as-is — read the source and
satisfy yourself before running anything against an environment you care about, which is the
entire reason this repository exists.

---

<sub>EnTech Engineering, P.C. &middot; 17 State St, 36th Fl, New York, NY 10004 &middot;
[entech.nyc](https://entech.nyc) &middot; [entechnology.io](https://entechnology.io)</sub>
