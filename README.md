# Export-ZammadKB

A PowerShell script that exports all articles from a Zammad Knowledge Base as Markdown files, mirroring the KB category structure as folders on disk.

Intended for importing into Open WebUI's Knowledge Base (RAG) or similar tools.

> **Note:** This is an unofficial community project, not affiliated with or endorsed by the Zammad Foundation.

---

## What it does

- Fetches all KB articles via the Zammad Search API (paginated)
- Converts article HTML to Markdown with YAML front matter
- Downloads file attachments alongside each article
- Mirrors the KB category hierarchy as folders
- Without a token: exports all public articles
- With a token: exports everything, including internal and draft articles

Output structure:

```
OutputPath/
    Category/
        Subcategory/
            Article Title.md
            Article Title_attachments/
                file.pdf
```

---

## Requirements

- PowerShell 7+
- A running Zammad instance

---

## Usage

**Run with defaults (public articles, with attachments):**

```powershell
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
.\Export-ZammadKB.ps1
```

**Custom output path, no attachments:**

```powershell
.\Export-ZammadKB.ps1 -OutputPath "C:\kb-export" -DownloadAttachments $false
```

**Export internal and draft articles (requires an API token):**

```powershell
.\Export-ZammadKB.ps1 -ApiToken "your-token-here"
```

When `-ApiToken` is provided, the script automatically switches to agent flavor. This includes internal and draft articles that are not publicly accessible. Attachment downloads also use the token in this mode.

You can also just edit the .ps1 file directly if you prefer.

---

## Parameters

| Parameter | Default | Description |
|---|---|---|
| `-ZammadUrl` | `https://helpdesk.example.de` | Base URL of your Zammad instance |
| `-OutputPath` | `C:\temp\zammad-kb-export` | Where to save the exported files |
| `-Language` | `de-de` | KB locale (must match your Zammad config) |
| `-KBId` | `1` | Knowledge Base ID |
| `-Flavor` | `public` | `public` or `agent` (auto-set when token is provided) |
| `-PageSize` | `100` | Articles per API page |
| `-ApiToken` | *(empty)* | Zammad API token - auto-switches to agent flavor when set |
| `-DownloadAttachments` | `$true` | Set to `$false` to skip attachments |

---

## Creating an API token

In Zammad, go to your profile and open **Token Access**. Create a token with Knowledge Base read permissions and paste it into `-ApiToken`.

Don't commit your token. Either pass it as a parameter or set it as an environment variable.

---

## Known limitations

- Article titles containing `[` or `]` will cause a write error on Windows (reserved wildcard characters in paths)
- Only single-locale knowledge bases are fully supported

---

## License

MIT
