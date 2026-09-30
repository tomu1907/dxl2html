<#
.SYNOPSIS
    Lotus / HCL Notes Mails (DXL) als HTML anzeigen - nativ unter Windows.

.DESCRIPTION
    - Mail-Kopf (Von, An, Kopie, Betreff, Datum ...) und Rich-Text-Body
      (Formatierung, Listen, Tabellen, Sektionen, Links, eingebettete Bilder)
    - Anhaenge ($FILE) werden extrahiert und verlinkt (Oeffnen / Speichern);
      Vorschau fuer Bilder, PDF und Textdateien
    - Mehrere Mails pro DXL-Datei, Ordner als Eingabe, Index-Seite
    - Kein JavaScript, keine externen Ressourcen
    - Laeuft mit Windows PowerShell 5.1 (vorinstalliert) und PowerShell 7

.PARAMETER Path
    DXL-Datei(en) oder Ordner. Ohne Angabe erscheint ein Datei-Dialog.

.PARAMETER OutDir
    Ausgabeordner (Standard: .\dxl_html bzw. neben der gewaehlten Datei).

.PARAMETER Embed
    Anhaenge als data:-URI einbetten (eine einzige HTML-Datei, nur "Speichern").

.PARAMETER Open
    Ergebnis im Standardbrowser oeffnen.

.EXAMPLE
    .\dxl2html.ps1 mail.dxl -Open
    .\dxl2html.ps1 C:\Export -OutDir C:\Temp\mails
    .\dxl2html.ps1 mail.dxl -Embed
#>
[CmdletBinding()]
param(
    [Parameter(Position = 0, ValueFromRemainingArguments = $true)]
    [string[]]$Path,
    [string]$OutDir,
    [switch]$Embed,
    [switch]$Open
)

$ErrorActionPreference = 'Stop'

# ---------------------------------------------------------------------------
# Helfer
# ---------------------------------------------------------------------------
$script:Embed   = [bool]$Embed
$script:Files   = $null
$script:Pardefs = @{}
$script:FilesDir = $null
$script:Taken   = $null

$script:Col = @{
    'dkred' = '#800000'; 'dkgreen' = '#008000'; 'dkblue' = '#000080'
    'dkyellow' = '#808000'; 'dkmagenta' = '#800080'; 'dkcyan' = '#008080'
    'gray' = '#808080'; 'ltgray' = '#c0c0c0'; 'lightgray' = '#c0c0c0'; 'dkgray' = '#404040'
}
$script:Mime = @{
    '.pdf' = 'application/pdf'; '.png' = 'image/png'; '.jpg' = 'image/jpeg'; '.jpeg' = 'image/jpeg'
    '.gif' = 'image/gif'; '.bmp' = 'image/bmp'; '.webp' = 'image/webp'; '.svg' = 'image/svg+xml'
    '.txt' = 'text/plain'; '.csv' = 'text/csv'; '.xml' = 'application/xml'; '.json' = 'application/json'
    '.zip' = 'application/zip'
}
$script:ImgExt = @('.png', '.jpg', '.jpeg', '.gif', '.bmp', '.webp', '.svg')
$script:TxtExt = @('.txt', '.log', '.csv', '.json', '.xml', '.md', '.ini', '.cfg', '.yaml', '.yml', '.sql', '.tsv', '.properties')
$script:Icons = @{
    '.pdf' = '&#128213;'; '.doc' = '&#128221;'; '.docx' = '&#128221;'; '.xls' = '&#128202;'; '.xlsx' = '&#128202;'
    '.csv' = '&#128202;'; '.ppt' = '&#128253;'; '.pptx' = '&#128253;'; '.zip' = '&#128476;'; '.7z' = '&#128476;'
    '.rar' = '&#128476;'; '.msg' = '&#9993;'; '.eml' = '&#9993;'; '.txt' = '&#128196;'; '.ics' = '&#128197;'
}

function Esc([string]$s) {
    if ($null -eq $s) { return '' }
    return [System.Net.WebUtility]::HtmlEncode($s)
}

function Get-Attr($n, [string]$name) {
    if ($null -eq $n) { return '' }
    $attrs = $n.Attributes
    if ($null -eq $attrs) { return '' }
    $a = $attrs[$name]
    if ($null -eq $a) { return '' }
    return [string]$a.Value
}

function Get-Child($n, [string]$name) {
    foreach ($c in $n.ChildNodes) {
        if ($c.NodeType.ToString() -eq 'Element' -and $c.LocalName -eq $name) { return $c }
    }
    return $null
}

function Get-SafeName([string]$n) {
    if ([string]::IsNullOrEmpty($n)) { $n = 'attachment' }
    $n = $n.Replace('\', '/')
    $n = $n.Substring($n.LastIndexOf('/') + 1)
    $n = [regex]::Replace($n, '[<>:"/\\|?*\x00-\x1f]', '_').Trim(' ', '.')
    if ($n -eq '') { $n = 'attachment' }
    if ($n.Length -gt 150) { $n = $n.Substring(0, 150) }
    return $n
}

function Get-Human([long]$n) {
    if ($n -lt 1024) { return "$n B" }
    if ($n -lt 1048576) { return ('{0:N1} KB' -f ($n / 1KB)) }
    if ($n -lt 1073741824) { return ('{0:N1} MB' -f ($n / 1MB)) }
    return ('{0:N1} GB' -f ($n / 1GB))
}

function Format-Dt([string]$s) {
    $s = ([string]$s).Trim()
    if ($s -match '^(\d{4})(\d{2})(\d{2})(?:T(\d{2})(\d{2})(\d{2}))?') {
        if ($Matches[4]) { return "$($Matches[3]).$($Matches[2]).$($Matches[1]) $($Matches[4]):$($Matches[5]):$($Matches[6])" }
        return "$($Matches[3]).$($Matches[2]).$($Matches[1])"
    }
    return $s
}

function Get-PrettyName([string]$n) {
    if ($n -match '^CN=([^/]+)') { return $Matches[1] }
    return $n
}

function Get-NamesHtml($list) {
    $parts = @()
    foreach ($n in @($list)) {
        $parts += ('<span title="' + (Esc $n) + '">' + (Esc (Get-PrettyName $n)) + '</span>')
    }
    return ($parts -join ', ')
}

function Get-CssColor([string]$c) {
    if ([string]::IsNullOrWhiteSpace($c)) { return $null }
    $c = $c.Trim().ToLower()
    if ($script:Col.ContainsKey($c)) { return $script:Col[$c] }
    if ($c -match '^#[0-9a-f]{3,8}$' -or $c -match '^[a-z]+$') { return $c }
    return $null
}

# ---------------------------------------------------------------------------
# Zugriff auf das Notes-Dokument
# ---------------------------------------------------------------------------
function Get-ItemValues($item) {
    $out = New-Object System.Collections.Generic.List[string]
    foreach ($c in $item.ChildNodes) {
        if ($c.NodeType.ToString() -ne 'Element') { continue }
        $t = $c.LocalName
        if ($t -eq 'text') { $out.Add([string]$c.InnerText) }
        elseif ($t -eq 'number') { $out.Add($c.InnerText.Trim()) }
        elseif ($t -eq 'datetime') { $out.Add((Format-Dt $c.InnerText)) }
        elseif ($t -eq 'textlist' -or $t -eq 'numberlist' -or $t -eq 'datetimelist') {
            foreach ($x in $c.ChildNodes) {
                if ($x.NodeType.ToString() -ne 'Element') { continue }
                if ($x.LocalName -eq 'datetime') { $out.Add((Format-Dt $x.InnerText)) }
                else { $out.Add([string]$x.InnerText) }
            }
        }
    }
    return , $out.ToArray()
}

function Get-Items($doc) {
    $h = @{}
    foreach ($it in $doc.ChildNodes) {
        if ($it.NodeType.ToString() -eq 'Element' -and $it.LocalName -eq 'item') {
            $k = (Get-Attr $it 'name').ToLower()
            if (-not $h.ContainsKey($k)) { $h[$k] = New-Object System.Collections.ArrayList }
            [void]$h[$k].Add($it)
        }
    }
    return $h
}

function Get-Vals($items, [string]$name) {
    $res = New-Object System.Collections.Generic.List[string]
    $k = $name.ToLower()
    if ($items.ContainsKey($k)) {
        foreach ($it in $items[$k]) {
            foreach ($v in @(Get-ItemValues $it)) { if ($v -ne '') { $res.Add($v) } }
        }
    }
    return , $res.ToArray()
}

function Get-First($arr) {
    $a = @($arr)
    if ($a.Count -gt 0) { return [string]$a[0] }
    return ''
}

function Get-DocFiles($doc) {
    $files = [ordered]@{}
    foreach ($f in $doc.SelectNodes("//*[local-name()='file']")) {
        $name = Get-Attr $f 'name'
        $fd = $f.SelectSingleNode("*[local-name()='filedata']")
        if ($name -eq '' -or $null -eq $fd) { continue }
        try { $bytes = [Convert]::FromBase64String(($fd.InnerText -replace '\s', '')) } catch { continue }
        $mod = ''
        $m = $f.SelectSingleNode("*[local-name()='modified']")
        if ($null -ne $m) { $mod = Format-Dt $m.InnerText }
        if (-not $files.Contains($name)) {
            $files[$name] = [pscustomobject]@{
                Name = $name; Data = $bytes; Modified = $mod; Href = $null; DlName = (Get-SafeName $name)
            }
        }
    }
    return $files
}

# ---------------------------------------------------------------------------
# Anhaenge: Speichern / URL
# ---------------------------------------------------------------------------
function Get-Href($att) {
    if ($null -ne $att.Href) { return $att.Href }
    if ($script:Embed) {
        $ext = [System.IO.Path]::GetExtension($att.Name).ToLower()
        $mime = 'application/octet-stream'
        if ($script:Mime.ContainsKey($ext)) { $mime = $script:Mime[$ext] }
        $att.Href = 'data:' + $mime + ';base64,' + [Convert]::ToBase64String($att.Data)
    }
    else {
        if (-not (Test-Path -LiteralPath $script:FilesDir)) { New-Item -ItemType Directory -Path $script:FilesDir | Out-Null }
        $fn = Get-SafeName $att.Name
        $base = [System.IO.Path]::GetFileNameWithoutExtension($fn)
        $ext = [System.IO.Path]::GetExtension($fn)
        $k = 1
        while ($script:Taken.Contains($fn.ToLower())) { $k++; $fn = "${base}_$k$ext" }
        [void]$script:Taken.Add($fn.ToLower())
        [System.IO.File]::WriteAllBytes((Join-Path $script:FilesDir $fn), $att.Data)
        $att.DlName = $fn
        $dirName = Split-Path -Leaf $script:FilesDir
        $att.Href = [Uri]::EscapeDataString($dirName) + '/' + [Uri]::EscapeDataString($fn)
    }
    return $att.Href
}

function Get-LinkAttrs($att) {
    if ($script:Embed) { return 'download="' + (Esc $att.DlName) + '"' }
    return 'target="_blank" rel="noopener"'
}

# ---------------------------------------------------------------------------
# Rich-Text-Renderer
# ---------------------------------------------------------------------------
function Get-FontCss($f) {
    $st = (Get-Attr $f 'style').ToLower() -split '\s+'
    $css = @()
    if ($st -contains 'bold') { $css += 'font-weight:bold' }
    if ($st -contains 'italic') { $css += 'font-style:italic' }
    $deco = @()
    if ($st -contains 'underline') { $deco += 'underline' }
    if ($st -contains 'strikethrough') { $deco += 'line-through' }
    if ($deco.Count -gt 0) { $css += ('text-decoration:' + ($deco -join ' ')) }
    if ($st -contains 'superscript') { $css += 'vertical-align:super;font-size:smaller' }
    if ($st -contains 'subscript') { $css += 'vertical-align:sub;font-size:smaller' }
    $col = Get-CssColor (Get-Attr $f 'color')
    if ($col) { $css += "color:$col" }
    $sz = Get-Attr $f 'size'
    if ($sz -match '^\d+(\.\d+)?pt$') { $css += "font-size:$sz" }
    return ($css -join ';')
}

function Get-ParStyle($pd) {
    if ($null -eq $pd) { return '' }
    $css = @()
    $al = Get-Attr $pd 'align'
    if ($al -eq 'center' -or $al -eq 'right') { $css += "text-align:$al" }
    elseif ($al -eq 'full') { $css += 'text-align:justify' }
    $lm = Get-Attr $pd 'leftmargin'
    if ($lm -match '^([\d.]+)in') {
        $v = [double]::Parse($Matches[1], [Globalization.CultureInfo]::InvariantCulture) - 1.0
        if ($v -gt 0.05) { $css += ('margin-left:' + $v.ToString('0.00', [Globalization.CultureInfo]::InvariantCulture) + 'in') }
    }
    if ($css.Count -gt 0) { return ' style="' + ($css -join ';') + '"' }
    return ''
}

function Render-Picture($el) {
    foreach ($c in $el.ChildNodes) {
        if ($c.NodeType.ToString() -ne 'Element') { continue }
        $t = $c.LocalName
        if ($t -eq 'gif' -or $t -eq 'jpeg' -or $t -eq 'png') {
            $data = ($c.InnerText -replace '\s', '')
            if ($data -ne '') { return '<img class="pic" alt="" src="data:image/' + $t + ';base64,' + $data + '">' }
        }
        elseif ($t -eq 'imageref') {
            $nm = Get-Attr $c 'name'
            if ($script:Files.Contains($nm)) {
                $att = $script:Files[$nm]
                return '<img class="pic" alt="' + (Esc $att.Name) + '" src="' + (Get-Href $att) + '">'
            }
        }
        elseif ($t -eq 'notesbitmap') { return '<span class="miss">[Notes-Bitmap - nicht darstellbar]</span>' }
    }
    return ''
}

function Render-InlineElement($c, [string]$cur) {
    $t = $c.LocalName
    if ($t -eq 'break') { return '<br>' }
    if ($t -eq 'tab') { return '&emsp;' }
    if ($t -eq 'attachmentref') {
        $name = Get-Attr $c 'name'
        $disp = Get-Attr $c 'displayname'
        if ($disp -eq '') { $disp = $name }
        if ($script:Files.Contains($name)) {
            $att = $script:Files[$name]
            return '<a class="att" href="' + (Get-Href $att) + '" ' + (Get-LinkAttrs $att) + '>&#128206; ' + (Esc $disp) + '</a>'
        }
        return '<span class="att miss">&#128206; ' + (Esc $disp) + ' (Datei nicht in der DXL enthalten)</span>'
    }
    if ($t -eq 'picture') { return (Render-Picture $c) }
    if ($t -eq 'sharedimageref') { return '<span class="miss">[Gemeinsam genutztes Bild]</span>' }
    if ($t -eq 'doclink') {
        $d = Get-Attr $c 'description'
        if ($d -eq '') { $d = Get-Attr $c 'hint' }
        return '<span class="miss">[Doclink: ' + (Esc $d) + ']</span>'
    }
    if ($t -eq 'pardef' -or $t -eq 'title') { return '' }
    return (Render-Inline $c $cur)
}

function Render-Inline($el, [string]$style) {
    $sb = New-Object System.Text.StringBuilder
    $cur = $style
    foreach ($c in $el.ChildNodes) {
        $nt = $c.NodeType.ToString()
        if ($nt -eq 'Text' -or $nt -eq 'CDATA' -or $nt -eq 'SignificantWhitespace') {
            $txt = [string]$c.Value
            if ($txt.Trim() -eq '' -and $txt -match '[\r\n]') { continue }
            $s = Esc $txt
            if ($cur -ne '') { [void]$sb.Append('<span style="' + $cur + '">' + $s + '</span>') }
            else { [void]$sb.Append($s) }
        }
        elseif ($nt -eq 'Element') {
            $t = $c.LocalName
            if ($t -eq 'font') { $cur = Get-FontCss $c }
            elseif ($t -eq 'run') { [void]$sb.Append((Render-Inline $c $cur)) }
            elseif ($t -eq 'urllink') {
                $href = (Get-Attr $c 'href').Trim()
                $inner = Render-Inline $c $cur
                if ($href -match '^(https?:|mailto:|ftp:)') {
                    [void]$sb.Append('<a href="' + (Esc $href) + '" target="_blank" rel="noopener noreferrer">' + $inner + '</a>')
                }
                else { [void]$sb.Append($inner) }
            }
            else { [void]$sb.Append((Render-InlineElement $c $cur)) }
        }
    }
    return $sb.ToString()
}

function Render-Table($t) {
    $rows = New-Object System.Text.StringBuilder
    foreach ($r in $t.ChildNodes) {
        if ($r.NodeType.ToString() -ne 'Element' -or $r.LocalName -ne 'tablerow') { continue }
        [void]$rows.Append('<tr>')
        foreach ($c in $r.ChildNodes) {
            if ($c.NodeType.ToString() -ne 'Element' -or $c.LocalName -ne 'tablecell') { continue }
            $attrs = ''
            $cs = Get-Attr $c 'columnspan'; if ($cs -match '^\d+$') { $attrs += " colspan=`"$cs`"" }
            $rs = Get-Attr $c 'rowspan'; if ($rs -match '^\d+$') { $attrs += " rowspan=`"$rs`"" }
            $bg = Get-CssColor (Get-Attr $c 'bgcolor')
            if ($bg) { $attrs += " style=`"background:$bg`"" }
            [void]$rows.Append("<td$attrs>" + (Render-Blocks $c) + '</td>')
        }
        [void]$rows.Append('</tr>')
    }
    return '<div class="tw"><table class="nt">' + $rows.ToString() + '</table></div>'
}

function Render-Section($s) {
    $title = Get-Child $s 'title'
    $head = 'Abschnitt'
    if ($null -ne $title) { $h = Render-Inline $title ''; if ($h -ne '') { $head = $h } }
    return '<details open class="sec"><summary>' + $head + '</summary>' + (Render-Blocks $s) + '</details>'
}

function Render-Blocks($container) {
    $sb = New-Object System.Text.StringBuilder
    $lst = ''
    foreach ($c in $container.ChildNodes) {
        if ($c.NodeType.ToString() -ne 'Element') { continue }
        $t = $c.LocalName
        if ($t -eq 'pardef') { $script:Pardefs[(Get-Attr $c 'id')] = $c }
        elseif ($t -eq 'par') {
            $pd = $script:Pardefs[(Get-Attr $c 'def')]
            if ($null -ne $pd -and (Get-Attr $pd 'hide') -ne '') { continue }
            $content = Render-Inline $c ''
            $kind = Get-Attr $pd 'list'
            if ($kind -ne '') {
                $want = 'ul'
                if (@('number', 'uppernumber', 'lowernumber', 'alphaupper', 'alphalower') -contains $kind) { $want = 'ol' }
                if ($lst -ne $want) {
                    if ($lst -ne '') { [void]$sb.Append("</$lst>") }
                    [void]$sb.Append("<$want>")
                    $lst = $want
                }
                if ($content -eq '') { $content = '&nbsp;' }
                [void]$sb.Append("<li>$content</li>")
            }
            else {
                if ($lst -ne '') { [void]$sb.Append("</$lst>"); $lst = '' }
                if ($content -eq '') { $content = '<br>' }
                [void]$sb.Append('<p' + (Get-ParStyle $pd) + '>' + $content + '</p>')
            }
        }
        elseif ($t -eq 'table') {
            if ($lst -ne '') { [void]$sb.Append("</$lst>"); $lst = '' }
            [void]$sb.Append((Render-Table $c))
        }
        elseif ($t -eq 'section') {
            if ($lst -ne '') { [void]$sb.Append("</$lst>"); $lst = '' }
            [void]$sb.Append((Render-Section $c))
        }
        elseif ($t -eq 'attachmentref' -or $t -eq 'picture') {
            if ($lst -ne '') { [void]$sb.Append("</$lst>"); $lst = '' }
            [void]$sb.Append('<p>' + (Render-InlineElement $c '') + '</p>')
        }
    }
    if ($lst -ne '') { [void]$sb.Append("</$lst>") }
    return $sb.ToString()
}

# ---------------------------------------------------------------------------
# Mail -> HTML
# ---------------------------------------------------------------------------
$script:Css = @'
:root{color-scheme:light}
*{box-sizing:border-box}
body{margin:0;background:#eef0f3;font:14px/1.45 "Segoe UI",Arial,sans-serif;color:#1b1f24}
.wrap{max-width:960px;margin:24px auto;background:#fff;border:1px solid #d5d9df;border-radius:8px;overflow:hidden}
.hdr{padding:18px 22px;background:#f7f8fa;border-bottom:1px solid #e1e4e8}
.hdr h1{margin:0 0 10px;font-size:20px}
.hdr table{border-collapse:collapse}
.hdr th{color:#5b6570;font-weight:600;text-align:left;padding:2px 14px 2px 0;vertical-align:top;white-space:nowrap}
.hdr td{padding:2px 0;word-break:break-word}
.body{padding:20px 22px;overflow-wrap:anywhere}
.body p{margin:0;min-height:1.35em;white-space:pre-wrap}
.body ul,.body ol{margin:.2em 0 .2em 1.4em;padding:0}
.pic{max-width:100%;height:auto;vertical-align:middle}
.tw{overflow-x:auto;margin:.4em 0}
.nt{border-collapse:collapse}
.nt td{border:1px solid #cfd4da;padding:4px 8px;vertical-align:top}
.sec>summary{cursor:pointer;font-weight:600}
.miss{color:#9a5b00}
.att{white-space:nowrap}
.atts{padding:16px 22px 20px;border-top:1px solid #e1e4e8;background:#fbfbfc}
.atts h2{margin:0 0 10px;font-size:15px}
.a{border:1px solid #dde1e6;border-radius:6px;background:#fff;padding:8px 12px;margin-bottom:8px}
.a small{color:#6b7580;margin-left:6px}
.a .dl{margin-left:10px;font-size:12px}
.prev{max-width:100%;max-height:520px;display:block;margin-top:8px;border:1px solid #e1e4e8}
.pdf{width:100%;height:640px;border:1px solid #e1e4e8;margin-top:6px}
pre.txt{max-height:360px;overflow:auto;background:#f4f5f7;padding:8px;margin:8px 0 0;font-size:12px;white-space:pre-wrap}
a{color:#0b5cad}
.idx{width:100%;border-collapse:collapse}
.idx th,.idx td{padding:8px 12px;border-bottom:1px solid #e1e4e8;text-align:left}
@media print{body{background:#fff}.wrap{border:0;margin:0}.pdf{display:none}}
'@

function Get-FallbackBody($items) {
    if ($items.ContainsKey('body')) {
        $txt = ''
        foreach ($it in $items['body']) { $txt += (@(Get-ItemValues $it) -join "`n") }
        if ($txt.Trim() -ne '') { return '<pre class="txt" style="max-height:none">' + (Esc $txt) + '</pre>' }
    }
    return '<p class="miss">Kein darstellbarer Nachrichtentext gefunden.</p>'
}

function Get-AttachmentsHtml {
    if ($script:Files.Count -eq 0) { return '' }
    $sb = New-Object System.Text.StringBuilder
    foreach ($att in $script:Files.Values) {
        $href = Get-Href $att
        $ext = [System.IO.Path]::GetExtension($att.Name).ToLower()
        $meta = Get-Human $att.Data.Length
        if ($att.Modified -ne '') { $meta += ', ge&auml;ndert ' + (Esc $att.Modified) }
        $icon = '&#128206;'
        if ($script:Icons.ContainsKey($ext)) { $icon = $script:Icons[$ext] }
        elseif ($script:ImgExt -contains $ext) { $icon = '&#128444;' }
        [void]$sb.Append('<div class="a"><span>' + $icon + '</span> <a href="' + $href + '" ' + (Get-LinkAttrs $att) + '><b>' + (Esc $att.Name) + '</b></a>')
        [void]$sb.Append('<small>' + $meta + '</small><a class="dl" href="' + $href + '" download="' + (Esc $att.DlName) + '">Speichern</a>')
        if ($script:ImgExt -contains $ext) {
            [void]$sb.Append('<img class="prev" src="' + $href + '" alt="' + (Esc $att.Name) + '">')
        }
        elseif ($ext -eq '.pdf') {
            [void]$sb.Append('<details open><summary>PDF-Vorschau</summary><iframe class="pdf" src="' + $href + '"></iframe></details>')
        }
        elseif ($script:TxtExt -contains $ext) {
            $len = [Math]::Min($att.Data.Length, 150000)
            try { $txt = (New-Object System.Text.UTF8Encoding($false, $true)).GetString($att.Data, 0, $len) }
            catch { $txt = [System.Text.Encoding]::GetEncoding('iso-8859-1').GetString($att.Data, 0, $len) }
            if ($att.Data.Length -gt 150000) { $txt += "`n[... gekuerzt ...]" }
            [void]$sb.Append('<details><summary>Textvorschau</summary><pre class="txt">' + (Esc $txt) + '</pre></details>')
        }
        [void]$sb.Append('</div>')
    }
    return '<div class="atts"><h2>Anh&auml;nge (' + $script:Files.Count + ')</h2>' + $sb.ToString() + '</div>'
}

function Convert-Doc($doc, [string]$outdir, [string]$stem) {
    $items = Get-Items $doc
    $script:Files = Get-DocFiles $doc
    $script:Pardefs = @{}
    $script:Taken = New-Object 'System.Collections.Generic.HashSet[string]'
    $script:FilesDir = Join-Path $outdir ($stem + '_files')

    $subject = Get-First (Get-Vals $items 'Subject')
    if ($subject -eq '') { $subject = '(kein Betreff)' }
    $sender = Get-First (Get-Vals $items 'From')
    if ($sender -eq '') { $sender = Get-First (Get-Vals $items 'Principal') }
    if ($sender -eq '') { $sender = Get-First (Get-Vals $items '$AltFrom') }
    $date = Get-First (Get-Vals $items 'PostedDate')
    if ($date -eq '') { $date = Get-First (Get-Vals $items 'DeliveredDate') }
    if ($date -eq '') { $date = Get-First (Get-Vals $items '$Created') }
    if ($date -eq '') {
        $cr = $doc.SelectSingleNode("*[local-name()='noteinfo']/*[local-name()='created']")
        if ($null -ne $cr) { $date = Format-Dt $cr.InnerText }
    }

    $hdr = New-Object System.Text.StringBuilder
    $addRow = {
        param($label, $content)
        if ($content -ne '') { [void]$hdr.Append("<tr><th>$label</th><td>$content</td></tr>") }
    }
    if ($sender -ne '') { & $addRow 'Von' (Get-NamesHtml @($sender)) }
    & $addRow 'An' (Get-NamesHtml (Get-Vals $items 'SendTo'))
    & $addRow 'Kopie' (Get-NamesHtml (Get-Vals $items 'CopyTo'))
    & $addRow 'Blindkopie' (Get-NamesHtml (Get-Vals $items 'BlindCopyTo'))
    & $addRow 'Datum' (Esc $date)
    $imp = Get-First (Get-Vals $items 'Importance')
    if ($imp -eq '1') { & $addRow 'Priorit&auml;t' 'Hoch' }
    elseif ($imp -eq '3') { & $addRow 'Priorit&auml;t' 'Niedrig' }
    if ($script:Files.Count -gt 0) { & $addRow 'Anh&auml;nge' ("$($script:Files.Count) Datei(en)") }

    $body = ''
    if ($items.ContainsKey('body')) {
        foreach ($it in $items['body']) {
            $rt = Get-Child $it 'richtext'
            if ($null -ne $rt) { $body += (Render-Blocks $rt) }
        }
    }
    if ($body.Trim() -eq '') { $body = Get-FallbackBody $items }

    $atts = Get-AttachmentsHtml
    $page = @"
<!DOCTYPE html>
<html lang="de"><head><meta charset="utf-8">
<meta http-equiv="Content-Security-Policy" content="script-src 'none'">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>$(Esc $subject)</title><style>$($script:Css)</style></head>
<body><div class="wrap">
<div class="hdr"><h1>$(Esc $subject)</h1><table>$($hdr.ToString())</table></div>
<div class="body">$body</div>
$atts
</div></body></html>
"@
    $file = Join-Path $outdir ($stem + '.html')
    [System.IO.File]::WriteAllText($file, $page, (New-Object System.Text.UTF8Encoding($false)))
    return [pscustomobject]@{
        File = ($stem + '.html'); Subject = $subject; From = (Get-PrettyName $sender)
        Date = $date; Atts = $script:Files.Count
    }
}

function Write-Index([string]$outdir, $entries) {
    $rows = New-Object System.Text.StringBuilder
    foreach ($e in $entries) {
        $a = ''
        if ($e.Atts -gt 0) { $a = '&#128206; ' + $e.Atts }
        [void]$rows.Append('<tr><td><a href="' + [Uri]::EscapeDataString($e.File) + '">' + (Esc $e.Subject) + '</a></td><td>' +
            (Esc $e.From) + '</td><td>' + (Esc $e.Date) + '</td><td>' + $a + '</td></tr>')
    }
    $page = @"
<!DOCTYPE html><html lang="de"><head><meta charset="utf-8"><title>DXL-Mails</title><style>$($script:Css)</style></head>
<body><div class="wrap"><div class="hdr"><h1>$($entries.Count) Mails</h1></div>
<table class="idx"><tr><th>Betreff</th><th>Von</th><th>Datum</th><th>Anh&auml;nge</th></tr>$($rows.ToString())</table>
</div></body></html>
"@
    $p = Join-Path $outdir 'index.html'
    [System.IO.File]::WriteAllText($p, $page, (New-Object System.Text.UTF8Encoding($false)))
    return $p
}

# ---------------------------------------------------------------------------
# Hauptprogramm
# ---------------------------------------------------------------------------
$dialog = $false
if (-not $Path -or $Path.Count -eq 0) {
    Add-Type -AssemblyName System.Windows.Forms
    $dlg = New-Object System.Windows.Forms.OpenFileDialog
    $dlg.Title = 'DXL-Datei(en) waehlen'
    $dlg.Filter = 'DXL (*.dxl;*.xml)|*.dxl;*.xml|Alle Dateien (*.*)|*.*'
    $dlg.Multiselect = $true
    if ($dlg.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) { return }
    $Path = $dlg.FileNames
    $dialog = $true
}

$inputs = @()
foreach ($p in $Path) {
    if (Test-Path -LiteralPath $p -PathType Container) {
        $inputs += Get-ChildItem -LiteralPath $p -Recurse -File | Where-Object { $_.Extension -in '.dxl', '.xml' } | Sort-Object FullName
    }
    elseif (Test-Path -LiteralPath $p -PathType Leaf) { $inputs += Get-Item -LiteralPath $p }
    else { Write-Warning "Nicht gefunden: $p" }
}
if ($inputs.Count -eq 0) { Write-Error 'Keine DXL-Dateien gefunden.'; return }

if (-not $OutDir) {
    if ($dialog) { $OutDir = Join-Path $inputs[0].DirectoryName 'dxl_html' }
    else { $OutDir = Join-Path (Get-Location).Path 'dxl_html' }
}
if (-not (Test-Path -LiteralPath $OutDir)) { New-Item -ItemType Directory -Path $OutDir | Out-Null }
$OutDir = (Resolve-Path -LiteralPath $OutDir).Path

$entries = New-Object System.Collections.ArrayList
$used = New-Object 'System.Collections.Generic.HashSet[string]'

foreach ($f in $inputs) {
    $n = 0
    $settings = New-Object System.Xml.XmlReaderSettings
    $settings.DtdProcessing = [System.Xml.DtdProcessing]::Ignore
    $settings.XmlResolver = $null
    $settings.CheckCharacters = $false
    $reader = $null
    try {
        $reader = [System.Xml.XmlReader]::Create($f.FullName, $settings)
        while (-not $reader.EOF) {
            if ($reader.NodeType -eq [System.Xml.XmlNodeType]::Element -and $reader.LocalName -eq 'document') {
                $xml = $reader.ReadOuterXml()   # Reader steht danach schon auf dem naechsten Knoten
                $n++
                $dom = New-Object System.Xml.XmlDocument
                $dom.PreserveWhitespace = $true
                $dom.XmlResolver = $null
                $dom.LoadXml($xml)

                $stem = Get-SafeName ([System.IO.Path]::GetFileNameWithoutExtension($f.Name))
                if ($n -gt 1) { $stem = "${stem}_$n" }
                $base = $stem; $k = 1
                while ($used.Contains($stem.ToLower())) { $k++; $stem = "$base-$k" }
                [void]$used.Add($stem.ToLower())

                $e = Convert-Doc $dom.DocumentElement $OutDir $stem
                [void]$entries.Add($e)
                Write-Host ("OK  {0} -> {1}  ({2} Anhang/Anhaenge)" -f $f.Name, $e.File, $e.Atts)
            }
            else { [void]$reader.Read() }
        }
    }
    catch {
        Write-Warning ("Fehler bei {0}: {1}" -f $f.FullName, $_.Exception.Message)
    }
    finally { if ($null -ne $reader) { $reader.Close() } }
    if ($n -eq 0) { Write-Warning "$($f.Name) enthaelt kein <document>." }
}

if ($entries.Count -eq 0) { return }
$target = Join-Path $OutDir $entries[0].File
if ($entries.Count -gt 1) { $target = Write-Index $OutDir $entries }
Write-Host ""
Write-Host "Fertig: $target"
if ($Open -or $dialog) { Start-Process $target }
