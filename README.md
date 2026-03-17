# mailsniper-go

A cross-platform Go port of [MailSniper](https://github.com/dafthack/MailSniper) — a penetration testing tool for searching email in Microsoft Exchange / Office 365 environments.

Single self-contained binary. No PowerShell. No dependencies to install. Runs on Linux, macOS, and Windows.

---

## Release Notes

### v1.1.0

- `open-inbox` bug fix: correctly parses EWS XML `ResponseCode` (namespace-aware) instead of relying on HTTP status — no longer falsely reports all mailboxes as accessible
- `spray-owa` and `spray-ews` now accept `--username` (single) or `--user-list` (file) for users
- `spray-owa` and `spray-ews` now accept `--password` (single) or `--password-list` (file) for passwords
- `spray-owa` and `spray-ews` now accept `--domain` to automatically prepend `DOMAIN\` to plain usernames
- `spray-owa` and `spray-ews` now accept `--verbose` to print per-attempt progress (silent by default)

### v1.0.0

- Full port of all MailSniper PowerShell modules to Go
- NTLM / Negotiate authentication (transparent via `go-ntlmssp`) — works against on-prem Exchange without Basic auth
- OAuth2 Bearer token support for O365
- Concurrent goroutine-based spraying and searching (`--threads`)
- Output to **CSV**, **JSON**, and **plain text**
- EWS autodiscover support
- Attachment content searching with optional download
- All-folders traversal (`--folder all`)

**Modules ported:**

| Go command | Original PowerShell |
|---|---|
| `self-search` | `Invoke-SelfSearch` |
| `global-search` | `Invoke-GlobalMailSearch` |
| `get-gal` | `Get-GlobalAddressList` |
| `spray-owa` | `Invoke-PasswordSprayOWA` |
| `spray-ews` | `Invoke-PasswordSprayEWS` |
| `harvest-users` | `Invoke-UsernameHarvestOWA` |
| `harvest-domain` | `Invoke-DomainHarvestOWA` |
| `open-inbox` | `Invoke-OpenInboxFinder` |
| `list-folders` | `Get-MailboxFolders` |
| `get-aduser` | `Get-ADUsernameFromEWS` |
| `send-email` | `Send-EWSEmail` |

---

## Build

```bash
git clone <repo>
cd mailsniper-go
go build -o mailsniper .
```

Requires Go 1.21+.

Cross-compile examples:

```bash
# Windows 64-bit
GOOS=windows GOARCH=amd64 go build -o mailsniper.exe .

# Linux 64-bit
GOOS=linux GOARCH=amd64 go build -o mailsniper .

# macOS ARM
GOOS=darwin GOARCH=arm64 go build -o mailsniper .
```

---

## Usage

```
mailsniper [command] [flags]
```

Global help:

```
./mailsniper --help
./mailsniper [command] --help
```

---

## Commands

### `self-search` — Search your own mailbox

Search the current user's mailbox via EWS for specific terms.

```bash
./mailsniper self-search \
  --hostname mail.domain.com \
  --mailbox user@domain.com \
  --username "DOMAIN\user" \
  --password Password1 \
  --folder all \
  --mails-per-user 200 \
  --output results.csv
```

**Flags:**

| Flag | Description | Default |
|---|---|---|
| `--mailbox` | Email address of the mailbox to search | |
| `--hostname` | Exchange server hostname | |
| `--ews-url` | Full EWS URL (overrides `--hostname`) | |
| `--username` | Username (`DOMAIN\user` or UPN) | |
| `--password` | Password | |
| `--access-token` | OAuth2 Bearer token | |
| `--exchange-version` | Exchange version string | `Exchange2010` |
| `--folder` | Folder to search; use `all` for all folders | `inbox` |
| `--mails-per-user` | Max emails to retrieve | `100` |
| `--terms` | Search terms (repeatable) | `*password*,*creds*,*credentials*` |
| `--regex` | Regex pattern (overrides `--terms`) | |
| `--check-attachments` | Search attachment content | `false` |
| `--download-dir` | Save matched attachments here | |
| `--output` | Output file path | |
| `--output-format` | `csv`, `json`, or `txt` | `csv` |
| `--skip-tls` | Skip TLS certificate verification | `false` |
| `--other-mailbox` | Read a different user's mailbox | |

---

### `global-search` — Search all mailboxes (admin / impersonation)

Requires the `ApplicationImpersonation` role assigned to the authenticating account.

```bash
./mailsniper global-search \
  --hostname mail.domain.com \
  --username "DOMAIN\admin" \
  --password Password1 \
  --impersonation-account "DOMAIN\admin" \
  --email-list mailboxes.txt \
  --folder all \
  --threads 10 \
  --output global-results.csv
```

**Key flags:**

| Flag | Description |
|---|---|
| `--impersonation-account` | Account with ApplicationImpersonation role |
| `--email-list` | File of email addresses to search (one per line) |
| `--autodiscover-email` | Email for EWS autodiscovery |
| `--threads` | Concurrent mailbox searches (default `5`) |

All `self-search` flags also apply.

---

### `get-gal` — Enumerate the Global Address List

```bash
./mailsniper get-gal \
  --hostname mail.domain.com \
  --username "DOMAIN\user" \
  --password Password1 \
  --exchange-version Exchange2016 \
  --output gal.txt \
  --output-format txt
```

| Flag | Description | Default |
|---|---|---|
| `--exchange-version` | Use `Exchange2013` or higher for FindPeople | `Exchange2013` |
| `--max` | Max entries to return (0 = all) | `0` |
| `--owa` | Try OWA FindPeople API first | `false` |

---

### `spray-owa` — Password spray against OWA

```bash
# Single user, single password
./mailsniper spray-owa \
  --hostname mail.domain.com \
  --username "DOMAIN\user" \
  --password Summer2025 \
  --skip-tls

# User list, password list, with domain auto-prefix
./mailsniper spray-owa \
  --hostname mail.domain.com \
  --user-list users.txt \
  --password-list wordlist.txt \
  --domain DOMAIN \
  --threads 5 \
  --delay 500 \
  --output hits.txt \
  --verbose
```

| Flag | Description | Default |
|---|---|---|
| `--hostname` | OWA server hostname | required |
| `--username` | Single username to spray | |
| `--user-list` | File of usernames, one per line | |
| `--password` | Single password to spray | |
| `--password-list` | File of passwords, one per line | |
| `--domain` | Prepend `DOMAIN\` to usernames that lack a domain prefix | |
| `--threads` | Concurrent threads | `5` |
| `--delay` | Delay between requests per thread (ms) | `0` |
| `--output` | Output file path | |
| `--output-format` | `csv`, `json`, or `txt` | `txt` |
| `--skip-tls` | Skip TLS certificate verification | `false` |
| `--verbose` | Print each password attempt and errors | `false` |

> Either `--username` or `--user-list` is required. Either `--password` or `--password-list` is required.

---

### `spray-ews` — Password spray against EWS

```bash
# Single user, single password
./mailsniper spray-ews \
  --hostname mail.domain.com \
  --username "DOMAIN\user" \
  --password Summer2025 \
  --skip-tls

# User list, password list, with domain auto-prefix
./mailsniper spray-ews \
  --hostname mail.domain.com \
  --user-list users.txt \
  --password-list wordlist.txt \
  --domain DOMAIN \
  --threads 5 \
  --output hits.txt \
  --verbose
```

| Flag | Description | Default |
|---|---|---|
| `--hostname` | Exchange server hostname | |
| `--ews-url` | Full EWS URL (overrides `--hostname`) | |
| `--username` | Single username to spray | |
| `--user-list` | File of usernames, one per line | |
| `--password` | Single password to spray | |
| `--password-list` | File of passwords, one per line | |
| `--domain` | Prepend `DOMAIN\` to usernames that lack a domain prefix | |
| `--exchange-version` | Exchange version string | `Exchange2010` |
| `--threads` | Concurrent threads | `5` |
| `--delay` | Delay between requests per thread (ms) | `0` |
| `--output` | Output file path | |
| `--output-format` | `csv`, `json`, or `txt` | `txt` |
| `--skip-tls` | Skip TLS certificate verification | `false` |
| `--verbose` | Print each password attempt | `false` |

> Either `--username` or `--user-list` is required. Either `--password` or `--password-list` is required.

---

### `harvest-users` — Enumerate valid OWA usernames via timing

```bash
./mailsniper harvest-users \
  --hostname mail.domain.com \
  --user-list users.txt \
  --threads 1 \
  --output valid-users.txt
```

> Keep `--threads 1` for accurate timing baseline comparisons.

---

### `harvest-domain` — Discover AD domain from OWA headers

```bash
./mailsniper harvest-domain \
  --hostname mail.domain.com \
  --skip-tls
```

Inspects the `WWW-Authenticate` response header for `realm=` to identify the Active Directory domain name.

---

### `open-inbox` — Find accessible mailboxes

```bash
./mailsniper open-inbox \
  --hostname mail.domain.com \
  --username "DOMAIN\user" \
  --password Password1 \
  --email-list gal.txt \
  --threads 10 \
  --skip-tls
```

Tests whether the authenticated user can read the Inbox of each address in the list.

---

### `list-folders` — List all folders in a mailbox

```bash
./mailsniper list-folders \
  --hostname mail.domain.com \
  --mailbox user@domain.com \
  --username "DOMAIN\user" \
  --password Password1 \
  --output folders.txt \
  --skip-tls
```

---

### `get-aduser` — Resolve AD usernames from email addresses

```bash
./mailsniper get-aduser \
  --email-list gal.txt \
  --hostname mail.domain.com \
  --username "DOMAIN\user" \
  --password Password1 \
  --output adusers.txt \
  --skip-tls
```

Uses EWS `ResolveNames` to map SMTP addresses to Active Directory display names.

---

### `send-email` — Send email via EWS

```bash
./mailsniper send-email \
  --hostname mail.domain.com \
  --username "DOMAIN\user" \
  --password Password1 \
  --recipient target@domain.com \
  --subject "Hello" \
  --body "<b>Test</b>" \
  --skip-tls
```

---

## Authentication

| Method | Flags |
|---|---|
| **NTLM** (default, on-prem Exchange) | `--username DOMAIN\user --password ...` |
| **Basic** | Same flags — falls back if NTLM is rejected |
| **OAuth2 Bearer** | `--access-token <token>` |

`--skip-tls` disables certificate verification (useful for self-signed Exchange certs).

---

## Output Formats

| Format | Flag | Notes |
|---|---|---|
| CSV | `--output-format csv` | Default for search commands |
| JSON | `--output-format json` | Machine-readable |
| Plain text | `--output-format txt` | Default for spray/harvest |

---

## License

MIT
