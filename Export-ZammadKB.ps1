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
    HTML pages (no authentication required). The public search API returns no
    category information, so the category hierarchy is read from the
    breadcrumbs of each article page.

    When -ApiToken is provided, the flavor is automatically switched to "agent",
    which includes internal and draft articles. Article content is then fetched
    via the authenticated REST API instead of the public HTML pages.
    Attachment downloads also use the token in this mode.

    The output contains no run-specific values such as export dates, so an
    unchanged article produces an identical file on every run. Content-hash
    based sync tools (e.g. oikb for Open WebUI) therefore only pick up real
    changes.

    The export is written to "<OutputPath>.new" first and swapped in when it
    is complete. Articles deleted in Zammad disappear from the export, and a
    failed run leaves the previous export untouched.

    Output folder structure mirrors the KB hierarchy:
        OutputPath/
            Category/
                Subcategory/
                    Article Title.md
                    Article Title_attachments/
                        file.pdf

    Articles with the same title in the same folder get their answer ID
    appended, e.g. "Article Title (123).md".

    Intended for importing into Open WebUI's Knowledge Base (RAG).

.PARAMETER ZammadUrl
    Base URL of your Zammad instance.

.PARAMETER OutputPath
    Folder where the Markdown files will be saved. Its previous content is
    replaced on every successful run.

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
    .\Export-ZammadKB.ps1 -ZammadUrl "https://helpdesk.example.de"

.EXAMPLE
    .\Export-ZammadKB.ps1 -OutputPath "C:\kb-export" -DownloadAttachments $false

.EXAMPLE
    .\Export-ZammadKB.ps1 -ApiToken "your-token-here"

.NOTES
    Author: danielknng
    Requires: PowerShell 7+
#>

param(
    [string] $ZammadUrl           = "https://helpdesk.example.de",
    [string] $OutputPath          = "C:\temp\zammad-kb-export",
    [string] $Language            = "de-de",
    [int]    $KBId                = 1,
    [ValidateSet("public", "agent")]
    [string] $Flavor              = "public",
    [int]    $PageSize            = 100,
    [string] $ApiToken            = "",
    [bool]   $DownloadAttachments = $true
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$ZammadUrl  = $ZammadUrl.TrimEnd("/")
$OutputPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($OutputPath).TrimEnd("\", "/")
$Utf8NoBom  = [System.Text.UTF8Encoding]::new($false)

if ($ApiToken -and $Flavor -eq "public") {
    $Flavor = "agent"
    Write-Host "  (ApiToken set, switching to agent flavor)" -ForegroundColor DarkGray
}


function Get-Headers {
    $headers = @{ "Content-Type" = "application/json" }
    if ($ApiToken) { $headers["Authorization"] = "Token token=$ApiToken" }
    return $headers
}

function ConvertFrom-HtmlText {
    param([string] $Html)
    $text = [System.Net.WebUtility]::HtmlDecode(($Html -replace '<[^>]+>', ''))
    return ($text -replace '\s+', ' ').Trim()
}

# Restricted to characters that are valid on Windows and SMB shares, because exports
# created on Linux are often copied there. -KeepExtension stops truncation from cutting
# off the file type of long attachment names.
function Get-SafeName {
    param([string] $Name, [int] $MaxLength = 80, [switch] $KeepExtension)
    $clean = ($Name -replace '[\x00-\x1f<>:"/\\|?*]', '_').Trim(' ', '.')
    if (-not $clean) { return "_" }
    if ($clean.Length -le $MaxLength) { return $clean }

    $extension = if ($KeepExtension) { [System.IO.Path]::GetExtension($clean) } else { "" }
    return $clean.Substring(0, $MaxLength - $extension.Length).TrimEnd(' ', '.') + $extension
}

# Category hits appear in the same search results but have no answer ID in their URL.
function Get-AnswerId {
    param([string] $Url)
    $match = [regex]::Match($Url, '/help/[^/]+/\d+[^/]*/(\d+)')
    if ($match.Success) { return [int] $match.Groups[1].Value }
    return $null
}

# Agent search results render the hierarchy as "Category > Subcategory > ..",
# with ".." standing for elided levels.
function Get-CategoryPathFromSubtitle {
    param([string] $Subtitle)
    $parts = @((ConvertFrom-HtmlText $Subtitle) -split '\s*>\s*' | Where-Object { $_ -and $_ -notin @('.', '..') })
    if ($parts.Count -eq 0) { return @("General") }
    return $parts
}

# The public search API returns no category, so the hierarchy is read from the page breadcrumbs.
# Only links of the form /help/<locale>/<id>-<slug> are categories; the KB root and the answer itself are skipped.
function Get-CategoryPathFromPage {
    param([string] $PageHtml)
    $nav = [regex]::Match($PageHtml, '(?s)<nav class="breadcrumbs">(.*?)</nav>')
    if (-not $nav.Success) { return @("General") }

    $crumbs = [regex]::Matches($nav.Groups[1].Value, '(?s)<a class="breadcrumb" href="([^"]*)".*?<span>(.*?)</span>')
    $parts  = @($crumbs |
        Where-Object { $_.Groups[1].Value -match '^/help/[^/]+/\d+[^/]*$' } |
        ForEach-Object { ConvertFrom-HtmlText $_.Groups[2].Value })

    if ($parts.Count -eq 0) { return @("General") }
    return $parts
}


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

    $t = $t -replace '(?s)<a[^>]*href=["'']([^"'']*)["''][^>]*>(.*?)</a>', '[$2]($1)'
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
    $t = [System.Net.WebUtility]::HtmlDecode($t) -replace "\u00a0", ' '
    $t = $t -replace '(\r?\n\s*){3,}', "`n`n"

    return $t.Trim()
}

function Get-ArticleHtml {
    param([string] $PageHtml)
    foreach ($pattern in @(
        '(?s)<article[^>]*>(.*?)</article>',
        '(?s)class="[^"]*answer[^"]*"[^>]*>(.*?)</section>',
        '(?s)<main[^>]*>(.*?)</main>'
    )) {
        $match = [regex]::Match($PageHtml, $pattern)
        if ($match.Success) { return $match.Groups[1].Value }
    }
    return $PageHtml
}

function Get-AttachmentLinks {
    param([string] $Html)
    return @([regex]::Matches($Html, '/api/v1/attachments/(\d+)') |
        ForEach-Object { [int] $_.Groups[1].Value } |
        Sort-Object -Unique |
        ForEach-Object {
            [PSCustomObject]@{
                RelativeUrl = "/api/v1/attachments/$_"
                Id          = $_
                FileName    = $null
            }
        })
}

# Some pages are only a short teaser in the article body; the search snippet is
# then at least as useful as the page content.
function Get-SnippetMarkdown {
    param($Detail, [string] $Url)
    $body = [string] $Detail.PSObject.Properties['body']?.Value
    return (ConvertTo-Markdown $body) + "`n`n> Preview only, see [full article]($Url)"
}


# Stops when a page brings no new results, which also guards against servers that ignore paging.
function Get-AllArticleDetails {
    $all  = [System.Collections.Generic.List[object]]::new()
    $seen = [System.Collections.Generic.HashSet[string]]::new()
    $page = 1

    while ($true) {
        $body = @{
            knowledge_base_id = $KBId
            locale            = $Language
            query             = "*"
            flavor            = $Flavor
            page              = $page
            per_page          = $PageSize
        } | ConvertTo-Json

        $result      = Invoke-RestMethod -Uri "$ZammadUrl/api/v1/knowledge_bases/search" -Method POST -Headers (Get-Headers) -Body $body
        $pageDetails = @($result.details | Where-Object { $_ -and $seen.Add([string] $_.url) })
        if ($pageDetails.Count -eq 0) { break }

        $all.AddRange($pageDetails)
        Write-Host "     Page ${page}: $($pageDetails.Count) articles" -ForegroundColor DarkGray

        if (@($result.details).Count -lt $PageSize) { break }
        $page++
    }

    return $all
}

function Get-PublicArticle {
    param($Detail, [string] $Url)
    $page     = (Invoke-WebRequest -Uri $Url -TimeoutSec 15).Content
    $html     = Get-ArticleHtml $page
    $markdown = ConvertTo-Markdown $html
    if ($markdown.Length -lt 30) { $markdown = Get-SnippetMarkdown $Detail $Url }

    return [PSCustomObject]@{
        CategoryPath = Get-CategoryPathFromPage $page
        Markdown     = $markdown
        Attachments  = if ($DownloadAttachments) { Get-AttachmentLinks $html } else { @() }
    }
}

# Uses only the first translation, which is the only one in single-locale knowledge bases.
function Get-AgentArticle {
    param($Detail, [int] $AnswerId)
    $answerUri = "$ZammadUrl/api/v1/knowledge_bases/$KBId/answers/$AnswerId"
    $answer    = (Invoke-RestMethod -Uri $answerUri -Headers (Get-Headers) -TimeoutSec 15).assets.KnowledgeBaseAnswer."$AnswerId"

    $translationId = @($answer.translation_ids)[0]
    if (-not $translationId) { throw "No translation found for answer $AnswerId" }

    $assets    = (Invoke-RestMethod -Uri "$answerUri`?include_contents=$translationId" -Headers (Get-Headers) -TimeoutSec 15).assets
    $contentId = $assets.KnowledgeBaseAnswerTranslation."$translationId".content_id
    $html      = $assets.KnowledgeBaseAnswerTranslationContent."$contentId".body

    $attachments = @()
    if ($DownloadAttachments) {
        $attachments = @(@($answer.attachments) | Where-Object { $_ } | ForEach-Object {
            [PSCustomObject]@{ RelativeUrl = $_.url; Id = $_.id; FileName = $_.filename }
        })
    }

    return [PSCustomObject]@{
        CategoryPath = Get-CategoryPathFromSubtitle ([string] $Detail.PSObject.Properties['subtitle']?.Value)
        Markdown     = ConvertTo-Markdown $html
        Attachments  = $attachments
    }
}

function Get-Article {
    param($Detail)
    $answerId = Get-AnswerId $Detail.url
    $url      = "$ZammadUrl$($Detail.url)"
    $content  = if ($Flavor -eq "agent") { Get-AgentArticle $Detail $answerId } else { Get-PublicArticle $Detail $url }

    return [PSCustomObject]@{
        Id           = $answerId
        Title        = ConvertFrom-HtmlText $Detail.title
        Url          = $url
        CategoryPath = @($content.CategoryPath)
        Markdown     = $content.Markdown
        Attachments  = @($content.Attachments)
    }
}


# In agent mode the API provides the file name; public downloads only carry it in Content-Disposition.
function Get-AttachmentFileName {
    param($Attachment, $Response)
    if ($Attachment.FileName) { return [string] $Attachment.FileName }

    $disposition = [string] ($Response.Headers["Content-Disposition"] | Select-Object -First 1)
    $match = [regex]::Match($disposition, 'filename\*?=(?:UTF-8'''')?(?:"([^"]+)"|([^\s;]+))')
    if ($match.Success) {
        $raw = if ($match.Groups[1].Value) { $match.Groups[1].Value } else { $match.Groups[2].Value }
        return [System.Uri]::UnescapeDataString($raw)
    }
    return "attachment_$($Attachment.Id)"
}

# Returns the Markdown list entry for the attachment. Names that collide within one
# article get the attachment ID appended so neither file is overwritten.
function Save-Attachment {
    param($Attachment, [string] $Folder, [System.Collections.Generic.HashSet[string]] $UsedNames)
    $request = @{ Uri = "$ZammadUrl$($Attachment.RelativeUrl)"; TimeoutSec = 60; MaximumRedirection = 5 }
    if ($Flavor -eq "agent") { $request["Headers"] = Get-Headers }
    $response = Invoke-WebRequest @request

    $name = Get-SafeName (Get-AttachmentFileName $Attachment $response) -KeepExtension
    if (-not $UsedNames.Add($name)) {
        $name = Get-SafeName ("{0} ({1}){2}" -f [System.IO.Path]::GetFileNameWithoutExtension($name), $Attachment.Id, [System.IO.Path]::GetExtension($name)) -KeepExtension
        [void] $UsedNames.Add($name)
    }

    [System.IO.Directory]::CreateDirectory($Folder) | Out-Null
    [System.IO.File]::WriteAllBytes((Join-Path $Folder $name), $response.RawContentStream.ToArray())
    Write-Host "     Attachment: $name" -ForegroundColor DarkGray

    $folderName = Split-Path $Folder -Leaf
    return "- [$name](./$([System.Uri]::EscapeDataString($folderName))/$([System.Uri]::EscapeDataString($name)))"
}

# YAML values are written as JSON strings, which YAML accepts and which escapes quotes safely.
function Get-ArticleDocument {
    param($Article, [string[]] $AttachmentLines)
    $frontMatter = @(
        "---"
        "title: $(ConvertTo-Json -InputObject $Article.Title)"
        "category: $(ConvertTo-Json -InputObject @($Article.CategoryPath) -Compress)"
        "source: $(ConvertTo-Json -InputObject $Article.Url)"
        "---"
    ) -join "`n"

    $sections = @($frontMatter, "# $($Article.Title)", $Article.Markdown)
    if ($AttachmentLines) { $sections += "## Attachments`n`n" + ($AttachmentLines -join "`n") }
    return ($sections | Where-Object { $_ }) -join "`n`n"
}

# Titles that collide within one folder get the answer ID appended. Articles are processed
# in answer ID order, so the resulting names are the same on every run.
function Get-ArticleBasePath {
    param($Article, [string] $Root, [System.Collections.Generic.HashSet[string]] $UsedPaths)
    $folder = $Root
    foreach ($segment in $Article.CategoryPath) { $folder = Join-Path $folder (Get-SafeName $segment) }

    $basePath = Join-Path $folder (Get-SafeName $Article.Title)
    if (-not $UsedPaths.Add($basePath)) {
        $basePath = "$basePath ($($Article.Id))"
        [void] $UsedPaths.Add($basePath)
    }
    return $basePath
}

function Write-Article {
    param($Article, [string] $Root, [System.Collections.Generic.HashSet[string]] $UsedPaths)
    $basePath = Get-ArticleBasePath $Article $Root $UsedPaths
    $usedNames = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    $attachmentLines = @($Article.Attachments | ForEach-Object { Save-Attachment $_ "${basePath}_attachments" $usedNames })

    [System.IO.Directory]::CreateDirectory((Split-Path $basePath -Parent)) | Out-Null
    [System.IO.File]::WriteAllText("$basePath.md", (Get-ArticleDocument $Article $attachmentLines) + "`n", $Utf8NoBom)
    return $attachmentLines.Count
}

function Remove-Folder {
    param([string] $Path)
    if (Test-Path -LiteralPath $Path) { Remove-Item -LiteralPath $Path -Recurse -Force }
}

# Swapping a completely written folder keeps sync tools from ever seeing a half-finished export.
function Publish-Export {
    param([string] $Staging, [string] $Target)
    $previous = "$Target.old"
    Remove-Folder $previous
    if (Test-Path -LiteralPath $Target) { Move-Item -LiteralPath $Target -Destination $previous }
    Move-Item -LiteralPath $Staging -Destination $Target
    Remove-Folder $previous
}


Write-Host ""
Write-Host "  Zammad KB -> Markdown Export" -ForegroundColor Cyan
Write-Host "  $ZammadUrl  ->  $OutputPath" -ForegroundColor DarkGray
Write-Host ""

Write-Host "[1/3] Fetching article list..." -ForegroundColor Yellow
$details = @(Get-AllArticleDetails | Where-Object { Get-AnswerId $_.url } | Sort-Object { Get-AnswerId $_.url })
if ($details.Count -eq 0) { throw "No articles found, previous export left untouched" }
Write-Host "     $($details.Count) articles found in total" -ForegroundColor Green

Write-Host "[2/3] Exporting articles..." -ForegroundColor Yellow
$staging = "$OutputPath.new"
Remove-Folder $staging
[System.IO.Directory]::CreateDirectory($staging) | Out-Null

$usedPaths       = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
$attachmentCount = 0
$counter         = 0

try {
    foreach ($detail in $details) {
        $counter++
        $article = Get-Article $detail
        Write-Host "   [$counter/$($details.Count)] $($article.Title)" -ForegroundColor DarkGray
        $attachmentCount += Write-Article $article $staging $usedPaths
    }
}
catch {
    Remove-Folder $staging
    Write-Host "  FAILED at article $counter, previous export left untouched" -ForegroundColor Red
    throw
}

Write-Host "[3/3] Replacing previous export..." -ForegroundColor Yellow
Publish-Export $staging $OutputPath

Write-Host ""
Write-Host "  Done" -ForegroundColor Cyan
Write-Host "  $($details.Count) articles exported to: $OutputPath"
if ($attachmentCount -gt 0) { Write-Host "  $attachmentCount attachments downloaded" -ForegroundColor Green }

Write-Host ""
Write-Host "  Folder overview:" -ForegroundColor DarkGray
@(Get-ChildItem -LiteralPath $OutputPath -Directory) | ForEach-Object {
    $count = @(Get-ChildItem -LiteralPath $_.FullName -Recurse -Filter "*.md").Count
    Write-Host "    $($_.Name)/ ($count articles)" -ForegroundColor DarkGray
}

Write-Host ""
