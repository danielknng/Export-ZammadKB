#Requires -Version 7.0
<#
.SYNOPSIS
    Exports all articles from a Zammad Knowledge Base as Markdown files,
    mirroring the KB category structure as folders on disk.

.DESCRIPTION
    Fetches all KB articles via the Zammad Search API (with pagination),
    retrieves the full article content, converts HTML to Markdown, and
    downloads any file attachments alongside the article.

    Without an API token, only public articles are exported using the public
    HTML endpoint (no authentication required).

    When -ApiToken is provided, the flavor is automatically switched to "agent",
    which includes internal and draft articles. Article content is then fetched
    via the authenticated REST API instead of the public HTML endpoint.
    Attachment downloads also use the token in this mode.

    Output folder structure mirrors the KB hierarchy:
        OutputPath/
            Category/
                Subcategory/
                    Article Title.md
                    Article Title_attachments/
                        file.pdf

    Intended for importing into Open WebUI's Knowledge Base (RAG).

.PARAMETER ZammadUrl
    Base URL of your Zammad instance. No trailing slash.

.PARAMETER OutputPath
    Local folder where Markdown files will be saved.

.PARAMETER Language
    KB locale to export. Must match what's configured in Zammad (e.g. "de-de").

.PARAMETER KBId
    Knowledge Base ID. Almost always 1 unless you deleted and recreated your KB.

.PARAMETER Flavor
    "public" for public KB articles (no auth required).
    "agent"  for internal articles (requires -ApiToken).

.PARAMETER PageSize
    Number of articles to fetch per API request. Max is usually 100.

.PARAMETER ApiToken
    Optional Zammad API token. When set, flavor is automatically switched to
    "agent" to include internal and draft articles. Create one in Zammad under
    Profile > Token Access.

.PARAMETER DownloadAttachments
    Set to $false to skip attachment downloads entirely.

.EXAMPLE
    .\Export-ZammadKB.ps1

.EXAMPLE
    .\Export-ZammadKB.ps1 -OutputPath "C:\kb-export" -DownloadAttachments $false

.EXAMPLE
    .\Export-ZammadKB.ps1 -Flavor "agent" -ApiToken "your-token-here"

.NOTES
    Author: danielknng
    Requires: PowerShell 7+
#>

# ── Configuration ─────────────────────────────────────────────────────────────

param(
    [string] $ZammadUrl           = "https://helpdesk.example.de",  # Your Zammad instance
    [string] $OutputPath          = "C:\temp\zammad-kb-export",       # Where to save the files
    [string] $Language            = "de-de",                          # KB language/locale
    [int]    $KBId                = 1,                                 # Knowledge Base ID
    [string] $Flavor              = "public",                          # "public" or "agent"
    [int]    $PageSize            = 100,                               # Articles per API page
    [string] $ApiToken            = "",                                # Only needed for agent flavor
    [bool]   $DownloadAttachments = $true                              # Download attached files
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

# If a token is provided, switch to agent flavor automatically to include
# internal and draft articles. The user does not need to set this manually.
if ($ApiToken -and $Flavor -eq "public") {
    $Flavor = "agent"
    Write-Host "  (ApiToken set – switching to agent flavor)" -ForegroundColor DarkGray
}


# ── Helpers ───────────────────────────────────────────────────────────────────

function Get-Headers {
    $headers = @{ "Content-Type" = "application/json" }
    if ($ApiToken) { $headers["Authorization"] = "Token token=$ApiToken" }
    return $headers
}

function Get-SafeName {
    param([string] $Name, [int] $MaxLength = 60)
    $invalid = [System.IO.Path]::GetInvalidFileNameChars() -join ''
    $result  = ($Name -replace "[{0}]" -f [RegEx]::Escape($invalid), '_').Trim('.')
    if ([string]::IsNullOrWhiteSpace($result)) { $result = "_" }
    if ($result.Length -gt $MaxLength) { $result = $result.Substring(0, $MaxLength).TrimEnd() }
    return $result
}

# Splits "Category > Subcategory > .." into a clean path array.
# Filters out ".." placeholders Zammad uses for missing hierarchy levels.
function Get-CategoryPath {
    param([string] $Subtitle)
    if ([string]::IsNullOrEmpty($Subtitle)) { return @("General") }
    $s = $Subtitle -replace '&amp;','&' -replace '&gt;','>' -replace '&lt;','<' -replace '<[^>]+>',''
    $parts = @(($s -split '\s*>\s*') | Where-Object { $_ -and $_.Trim() -notin @('..', '.', '') })
    if ($parts.Count -eq 0) { return @("General") }
    return @($parts | ForEach-Object { $_.Trim() })
}

function Get-ShortUrl {
    param([string] $RelativeUrl)
    if ($RelativeUrl -match '/help/[^/]+/(\d+)[^/]*/(\d+)') {
        return "$ZammadUrl/help/$Language/$($Matches[1])/$($Matches[2])"
    }
    return "$ZammadUrl$RelativeUrl"
}


# ── HTML to Markdown ──────────────────────────────────────────────────────────

function ConvertTo-Markdown {
    param([string] $Html)
    if ([string]::IsNullOrEmpty($Html)) { return "" }

    $t = $Html

    $t = $t -replace '(?s)<script[^>]*>.*?</script>', ''
    $t = $t -replace '(?s)<style[^>]*>.*?</style>',  ''

    $t = $t -replace '(?s)<h1[^>]*>(.*?)</h1>', "`n# `$1`n"
    $t = $t -replace '(?s)<h2[^>]*>(.*?)</h2>', "`n## `$1`n"
    $t = $t -replace '(?s)<h3[^>]*>(.*?)</h3>', "`n### `$1`n"
    $t = $t -replace '(?s)<h4[^>]*>(.*?)</h4>', "`n#### `$1`n"
    $t = $t -replace '(?s)<h5[^>]*>(.*?)</h5>', "`n##### `$1`n"
    $t = $t -replace '(?s)<h6[^>]*>(.*?)</h6>', "`n###### `$1`n"

    $t = $t -replace '(?s)<strong[^>]*>(.*?)</strong>', '**$1**'
    $t = $t -replace '(?s)<b[^>]*>(.*?)</b>',           '**$1**'
    $t = $t -replace '(?s)<em[^>]*>(.*?)</em>',         '*$1*'
    $t = $t -replace '(?s)<i[^>]*>(.*?)</i>',           '*$1*'

    $t = $t -replace '(?s)<pre[^>]*><code[^>]*>(.*?)</code></pre>', "``````\`n`$1\`n``````"
    $t = $t -replace '(?s)<code[^>]*>(.*?)</code>', '`$1`'

    $t = $t -replace '(?s)<a[^>]*href="([^"]*)"[^>]*>(.*?)</a>', '[$2]($1)'
    $t = $t -replace '(?s)<img[^>]*alt="([^"]*)"[^>]*src="([^"]*)"[^>]*/?>', '![$1]($2)'
    $t = $t -replace '(?s)<img[^>]*src="([^"]*)"[^>]*alt="([^"]*)"[^>]*/?>',  '![$2]($1)'
    $t = $t -replace '(?s)<img[^>]*src="([^"]*)"[^>]*/?>',                     '![]($1)'

    $t = [regex]::Replace($t, '(?s)<tr[^>]*>(.*?)</tr>', {
        param($m)
        $cells = [regex]::Matches($m.Groups[1].Value, '(?s)<t[hd][^>]*>(.*?)</t[hd]>')
        "| " + (($cells | ForEach-Object { ($_.Groups[1].Value -replace '<[^>]+>', '').Trim() }) -join " | ") + " |`n"
    })

    $t = [regex]::Replace($t, '(?s)<ul[^>]*>(.*?)</ul>', {
        param($m)
        $items = [regex]::Matches($m.Groups[1].Value, '(?s)<li[^>]*>(.*?)</li>')
        ($items | ForEach-Object { "- " + ($_.Groups[1].Value -replace '<[^>]+>', '').Trim() }) -join "`n"
    })

    $t = [regex]::Replace($t, '(?s)<ol[^>]*>(.*?)</ol>', {
        param($m)
        $items = [regex]::Matches($m.Groups[1].Value, '(?s)<li[^>]*>(.*?)</li>')
        $n = 1
        ($items | ForEach-Object { "$n. " + ($_.Groups[1].Value -replace '<[^>]+>', '').Trim(); $n++ }) -join "`n"
    })

    $t = $t -replace '<br\s*/?>', "`n"
    $t = $t -replace '(?s)<p[^>]*>(.*?)</p>',    "`n`$1`n"
    $t = $t -replace '(?s)<div[^>]*>(.*?)</div>', "`n`$1`n"
    $t = $t -replace '<hr\s*/?>', "`n---`n"

    $t = [regex]::Replace($t, '(?s)<blockquote[^>]*>(.*?)</blockquote>', {
        param($m)
        ($m.Groups[1].Value -split "`n" | ForEach-Object { "> " + $_.Trim() }) -join "`n"
    })

    $t = $t -replace '<[^>]+>', ''

    $t = $t -replace '&amp;',     '&'
    $t = $t -replace '&lt;',      '<'
    $t = $t -replace '&gt;',      '>'
    $t = $t -replace '&nbsp;',    ' '
    $t = $t -replace '&quot;',    '"'
    $t = $t -replace '&#39;',     "'"
    $t = $t -replace '&#[0-9]+;', ''
    $t = $t -replace '&[a-z]+;',  ''

    $t = $t -replace '(\r?\n\s*){3,}', "`n`n"

    return $t.Trim()
}


# ── Article content extraction ────────────────────────────────────────────────

function Get-ArticleHtml {
    param([string] $PageHtml)
    if ($PageHtml -match '(?s)<article[^>]*>(.*?)</article>')                  { return $Matches[1] }
    if ($PageHtml -match '(?s)class="[^"]*answer[^"]*"[^>]*>(.*?)</section>') { return $Matches[1] }
    if ($PageHtml -match '(?s)<main[^>]*>(.*?)</main>')                        { return $Matches[1] }
    return $PageHtml
}

# Finds all attachment links in the raw HTML.
function Get-AttachmentLinks {
    param([string] $Html)

    $matches = [regex]::Matches($Html, '(/api/v1/attachments/(\d+))')

    $results = @()

    foreach ($m in $matches) {
        $id  = $m.Groups[2].Value
        $url = $m.Groups[1].Value

        $results += [PSCustomObject]@{
            RelativeUrl = $url
            Id          = $id
            FileName    = "attachment_$id"
        }
    }

    # Remove duplicates (same attachment may appear multiple times)
    $results = $results | Sort-Object Id -Unique

    return $results
}


# ── Pagination ────────────────────────────────────────────────────────────────

function Get-AllArticles {
    $all  = [System.Collections.Generic.List[object]]::new()
    $page = 1

    do {
        $body = @{
            knowledge_base_id = $KBId
            locale            = $Language
            query             = "*"
            flavor            = $Flavor
            page              = $page
            per_page          = $PageSize
        } | ConvertTo-Json

        $result      = Invoke-RestMethod -Uri "$ZammadUrl/api/v1/knowledge_bases/search" -Method POST -Headers (Get-Headers) -Body $body
        $pageDetails = @($result.details)

        if ($pageDetails.Count -eq 0) { break }

        $all.AddRange($pageDetails)
        Write-Host "     Page ${page}: $($pageDetails.Count) articles" -ForegroundColor DarkGray

        if ($pageDetails.Count -lt $PageSize) { break }
        $page++
    } while ($true)

    return $all
}


# ── Main ──────────────────────────────────────────────────────────────────────

Write-Host ""
Write-Host "  Zammad KB -> Markdown Export" -ForegroundColor Cyan
Write-Host "  $ZammadUrl  ->  $OutputPath" -ForegroundColor DarkGray
Write-Host ""

New-Item -ItemType Directory -Force -Path $OutputPath | Out-Null


# Step 1: Article list

Write-Host "[1/3] Fetching article list..." -ForegroundColor Yellow
$details = Get-AllArticles
Write-Host "     $($details.Count) articles found in total" -ForegroundColor Green


# Step 2: Load content and find attachments

Write-Host "[2/3] Loading article content..." -ForegroundColor Yellow

$articles   = [System.Collections.Generic.List[object]]::new()
$loadErrors = 0
$counter    = 0

foreach ($detail in $details) {
    $counter++

    $title = $detail.title -replace '</?em>','' -replace '<[^>]+>',''
    $title = $title -replace '&amp;','&' -replace '&quot;','"' -replace '&#39;',"'" -replace '&gt;','>' -replace '&lt;','<'
    $title = $title.Trim()

    $categoryPath = Get-CategoryPath ([string]($detail.PSObject.Properties['subtitle']?.Value))
    $url          = Get-ShortUrl $detail.url

    Write-Host "   [$counter/$($details.Count)] $title" -ForegroundColor DarkGray -NoNewline

    $markdown    = ""
    $attachments = @()

    try {
        if ($Flavor -eq "agent") {
            # Extract the article ID from the URL pattern /help/{locale}/{categoryId}/{articleId}
            # Category URLs like /help/de-de/43-rechnungsbearbeitung have no article ID – skip them.
            $articleId = $null
            if ($detail.url -match '/help/[^/]+/\d+[^/]*/(\d+)') {
                $articleId = $Matches[1]
            }
            if (-not $articleId) {
                Write-Host "  SKIP (category)" -ForegroundColor DarkGray
                continue
            }

            # First call to get the answer metadata and translation IDs
            $apiAnswer     = Invoke-RestMethod -Uri "$ZammadUrl/api/v1/knowledge_bases/$KBId/answers/$articleId" -Headers (Get-Headers) -TimeoutSec 15
            $answerAssets  = $apiAnswer.assets
            $answerObj     = $answerAssets.KnowledgeBaseAnswer."$articleId"

            # Use the first available translation (single-locale setups have exactly one)
            $translationIds = @($answerObj.translation_ids)
            if ($translationIds.Count -eq 0) {
                throw "No translation IDs found for article $articleId"
            }
            $translationId = $translationIds[0]

            # Second call with include_contents to get the HTML body
            $apiContent    = Invoke-RestMethod -Uri "$ZammadUrl/api/v1/knowledge_bases/$KBId/answers/$articleId`?include_contents=$translationId" -Headers (Get-Headers) -TimeoutSec 15
            $contentAssets = $apiContent.assets
            $contentObj    = $contentAssets.KnowledgeBaseAnswerTranslation."$translationId"
            $contentId     = $contentObj.content_id
            $contentHtml   = $contentAssets.KnowledgeBaseAnswerTranslationContent."$contentId".body
            $markdown      = ConvertTo-Markdown $contentHtml

            # Attachments come directly from the answer object, no HTML scraping needed
            if ($DownloadAttachments) {
                $rawAtts = @($answerObj.attachments)
                $attachments = $rawAtts | Where-Object { $_ } | ForEach-Object {
                    [PSCustomObject]@{
                        RelativeUrl = $_.url
                        Id          = $_.id
                        FileName    = $_.filename
                    }
                }
                $attachments = @($attachments)
            }
        }
        else {
            $response    = Invoke-WebRequest -Uri $url -UseBasicParsing -TimeoutSec 15
            $contentHtml = Get-ArticleHtml $response.Content
            $markdown    = ConvertTo-Markdown $contentHtml

            if ($DownloadAttachments) {
                $attachments = @(Get-AttachmentLinks $contentHtml)
            }

            if ($markdown.Length -lt 30) {
                $rawBody  = [string]($detail.PSObject.Properties['body']?.Value)
                $markdown = ConvertTo-Markdown ($rawBody -replace '</?em>','')
                $markdown += "`n`n> Preview only – [full article]($url)"
            }
        }

        Write-Host "  OK" -ForegroundColor Green
    }
    catch {
        Write-Host "  FAILED" -ForegroundColor Red
        Write-Warning "  $url -> $_"
        $rawBody  = [string]($detail.PSObject.Properties['body']?.Value)
        $markdown = ConvertTo-Markdown ($rawBody -replace '</?em>','')
        if ($url) { $markdown += "`n`n> Preview only – [full article]($url)" }
        $loadErrors++
    }

    $articles.Add([PSCustomObject]@{
        Title        = $title
        Content      = $markdown
        CategoryPath = $categoryPath
        Url          = $url
        Attachments  = $attachments
    })
}


# Step 3: Write Markdown files and download attachments
# No auth header is sent for attachments – they are publicly accessible
# on public Zammad KB instances just like the article pages themselves.

Write-Host "[3/3] Writing files..." -ForegroundColor Yellow

$written         = 0
$writeErrors     = 0
$attachmentCount = 0

foreach ($article in $articles) {
    try {
		$categoryPath = @($article.CategoryPath)
        # Build folder path from category hierarchy
        $folder = $OutputPath
        foreach ($segment in $categoryPath) {
            $folder = Join-Path $folder (Get-SafeName $segment)
        }
        New-Item -ItemType Directory -Force -Path $folder | Out-Null

        $filename = Get-SafeName $article.Title
        $filepath = Join-Path $folder "$filename.md"

        $suffix = 1
        while (Test-Path $filepath) {
            $filepath = Join-Path $folder "${filename}_${suffix}.md"
            $suffix++
        }

        # Download attachments – no token needed for public KB
        $attachmentMd = ""
		$attachments = @($article.Attachments)
        if ($attachments.Count -gt 0) {
            $attachFolder = Join-Path $folder "${filename}_attachments"
            New-Item -ItemType Directory -Force -Path $attachFolder | Out-Null

            $attachmentLines = @()
foreach ($att in $attachments) {
    try {
        $uri = "$ZammadUrl$($att.RelativeUrl)"

        $attParams = @{
            Uri               = $uri
            TimeoutSec        = 60
            MaximumRedirection = 5
        }
        if ($Flavor -eq "agent") { $attParams["Headers"] = Get-Headers }
        $response = Invoke-WebRequest @attParams

        # In agent mode the filename comes directly from the API.
        # For public mode, read it from the Content-Disposition response header.
        $fileName = $null
        if ($att.PSObject.Properties['FileName'] -and $att.FileName -and $att.FileName -match '\.') {
            $fileName = $att.FileName
        }
        else {
            $cd = $response.Headers["Content-Disposition"]
            if ($cd) {
                $m = [regex]::Match($cd, 'filename\*?=(?:UTF-8'''')?(?:"([^"]+)"|([^\s;]+))')
                if ($m.Success) {
                    $fileName = if ($m.Groups[1].Value) { $m.Groups[1].Value } else { $m.Groups[2].Value }
                    $fileName = [System.Uri]::UnescapeDataString($fileName.Trim('"'))
                }
            }
        }
        if (-not $fileName -or $fileName -notmatch '\.') {
            $fileName = "attachment_$($att.Id)"
        }

        $attFile = Join-Path $attachFolder (Get-SafeName $fileName)

        [System.IO.File]::WriteAllBytes($attFile, $response.Content)

        $attachmentLines += "- [$fileName](./${filename}_attachments/$(Get-SafeName $fileName))"
        $attachmentCount++

        Write-Host "     Attachment: $fileName" -ForegroundColor DarkGray
    }
    catch {
        Write-Warning "  Could not download attachment '$($att.Id)': $_"
        $attachmentLines += "- attachment_$($att.Id) *(download failed)*"
    }
}

            $attachmentMd = "`n`n## Attachments`n`n" + ($attachmentLines -join "`n")
        }

        @"
---
title: "$($article.Title)"
category: "$($categoryPath[0])"
subcategory: "$(if ($categoryPath.Count -gt 1) { $categoryPath[1] } else { '' })"
source: "$($article.Url)"
exported: "$(Get-Date -Format "yyyy-MM-dd")"
---

# $($article.Title)

$($article.Content)$attachmentMd
"@ | Out-File -FilePath $filepath -Encoding UTF8 -Force

        $written++
    }
    catch {
        Write-Warning "Failed to write '$($article.Title)': $_"
        $writeErrors++
    }
}


# ── Summary ───────────────────────────────────────────────────────────────────

Write-Host ""
Write-Host "  Done" -ForegroundColor Cyan
Write-Host "  $written articles exported to: $OutputPath"
if ($attachmentCount -gt 0) { Write-Host "  $attachmentCount attachments downloaded"      -ForegroundColor Green  }
if ($loadErrors      -gt 0) { Write-Host "  $loadErrors articles used snippet fallback"    -ForegroundColor Yellow }
if ($writeErrors     -gt 0) { Write-Host "  $writeErrors write errors"                     -ForegroundColor Red    }

Write-Host ""
Write-Host "  Folder overview:" -ForegroundColor DarkGray
@(Get-ChildItem -Path $OutputPath -Directory) | ForEach-Object {
    $count = @(Get-ChildItem $_.FullName -Recurse -Filter "*.md").Count
    Write-Host "    $($_.Name)/ ($count articles)" -ForegroundColor DarkGray
}

Write-Host ""