$ErrorActionPreference = "Stop"
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem
Add-Type -AssemblyName System.Drawing

$outDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$markupPath = Join-Path $outDir "raport_markup.txt"
$shot = Join-Path $outDir "zrzut_strona_glowna.png"
$docx = Join-Path $outDir "Raport_Audytu_Greenwashingu_Hi_DESIGN.docx"
$work = Join-Path $outDir "_docx_build"
if (Test-Path $work) { Remove-Item $work -Recurse -Force }
New-Item -ItemType Directory -Path (Join-Path $work "word\_rels") | Out-Null
New-Item -ItemType Directory -Path (Join-Path $work "word\media") | Out-Null
New-Item -ItemType Directory -Path (Join-Path $work "_rels") | Out-Null
Copy-Item $shot (Join-Path $work "word\media\image1.png")

$img = [System.Drawing.Image]::FromFile($shot)
$cx = 5486400
$cy = [int]($cx * ($img.Height / [double]$img.Width))
$img.Dispose()

function X([string]$s) {
    if ($null -eq $s) { return "" }
    return ($s -replace '&','&amp;' -replace '<','&lt;' -replace '>','&gt;')
}

function Run([string]$text, [int]$sz, [bool]$bold, [bool]$italic, [string]$color) {
    $b = ""; if ($bold) { $b = "<w:b/>" }
    $i = ""; if ($italic) { $i = "<w:i/>" }
    return "<w:r><w:rPr><w:rFonts w:ascii=`"Arial`" w:hAnsi=`"Arial`" w:cs=`"Arial`"/>$b$i<w:sz w:val=`"$sz`"/><w:szCs w:val=`"$sz`"/><w:color w:val=`"$color`"/></w:rPr><w:t xml:space=`"preserve`">$(X $text)</w:t></w:r>"
}

function Para([string]$text, [int]$sz, [bool]$bold, [bool]$italic, [string]$color, [int]$before, [int]$after) {
    return "<w:p><w:pPr><w:spacing w:before=`"$before`" w:after=`"$after`" w:line=`"276`" w:lineRule=`"auto`"/></w:pPr>$(Run $text $sz $bold $italic $color)</w:p>"
}

function Cell([string]$innerXml, [string]$fill, [int]$width, [string]$border="E2E8F0", [int]$leftSz=4) {
    return @"
<w:tc>
  <w:tcPr>
    <w:tcW w:w="$width" w:type="dxa"/>
    <w:shd w:val="clear" w:color="auto" w:fill="$fill"/>
    <w:tcMar><w:top w:w="60" w:type="dxa"/><w:left w:w="80" w:type="dxa"/><w:bottom w:w="60" w:type="dxa"/><w:right w:w="80" w:type="dxa"/></w:tcMar>
    <w:tcBorders>
      <w:top w:val="single" w:sz="4" w:space="0" w:color="$border"/>
      <w:left w:val="single" w:sz="$leftSz" w:space="0" w:color="$border"/>
      <w:bottom w:val="single" w:sz="4" w:space="0" w:color="$border"/>
      <w:right w:val="single" w:sz="4" w:space="0" w:color="$border"/>
    </w:tcBorders>
    <w:vAlign w:val="center"/>
  </w:tcPr>
  $innerXml
</w:tc>
"@
}

function Tbl([string]$rowsXml) {
    return "<w:tbl><w:tblPr><w:tblW w:w=`"5000`" w:type=`"pct`"/><w:tblLayout w:type=`"fixed`"/></w:tblPr>$rowsXml</w:tbl><w:p/>"
}

function Callout([string]$title, [string[]]$lines, [string]$fill, [string]$border, [string]$titleColor) {
    $paras = Para $title 19 $true $false $titleColor 0 60
    foreach ($line in $lines) {
        if ($line.Trim().Length -gt 0) { $paras += Para $line 19 $false $false "16242C" 40 40 }
    }
    $cell = Cell $paras $fill 9360 $border 24
    return Tbl "<w:tr>$cell</w:tr>"
}

$raw = [System.IO.File]::ReadAllText($markupPath, [System.Text.UTF8Encoding]::new($false))
$raw = $raw -replace "`r`n","`n" -replace "`r","`n"
$parts = [regex]::Split($raw, "(?m)^::([A-Z0-9]+)::\s*$")
# parts[0] preamble, then tag, body, tag, body...
$blocks = @()
for ($i = 1; $i -lt $parts.Length; $i += 2) {
    $blocks += [pscustomobject]@{ Tag = $parts[$i]; Body = $parts[$i+1].Trim() }
}

$sb = New-Object System.Text.StringBuilder
foreach ($b in $blocks) {
    switch ($b.Tag) {
        "KICKER" { [void]$sb.Append((Para $b.Body 18 $true $false "2A8AB0" 0 40)) }
        "TITLE" { [void]$sb.Append((Para $b.Body 40 $true $false "08181F" 40 40)) }
        "SUB" { [void]$sb.Append((Para $b.Body 20 $false $true "5A6B77" 0 120)) }
        "SHOT" {
            $drawing = @"
<w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:drawing>
<wp:inline distT="0" distB="0" distL="0" distR="0">
<wp:extent cx="$cx" cy="$cy"/>
<wp:effectExtent l="0" t="0" r="0" b="0"/>
<wp:docPr id="1" name="zrzut"/>
<wp:cNvGraphicFramePr><a:graphicFrameLocks xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" noChangeAspect="1"/></wp:cNvGraphicFramePr>
<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
<pic:nvPicPr><pic:cNvPr id="0" name="zrzut.png"/><pic:cNvPicPr/></pic:nvPicPr>
<pic:blipFill><a:blip r:embed="rId4"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>
<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="$cx" cy="$cy"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>
</pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>
"@
            [void]$sb.Append($drawing)
        }
        "CAP" { [void]$sb.Append((Para $b.Body 17 $false $true "5A6B77" 40 160)) }
        "META" {
            $rows = ""
            $n = 0
            foreach ($line in ($b.Body -split "`n")) {
                if ($line -notmatch '\|\|\|') { continue }
                $k,$v = $line.Split(@('|||'),2,[System.StringSplitOptions]::None)
                $n++
                $fillR = if ($n -eq 6) { "FFF1F2" } else { "FFFFFF" }
                $col = if ($n -eq 6) { "BE123C" } else { "16242C" }
                $bold = ($n -eq 6)
                $left = Cell (Para $k.Trim() 17 $true $false "16242C" 0 0) "F8FAFC" 2200
                $right = Cell (Para $v.Trim() 17 $bold $false $col 0 0) $fillR 7160
                $rows += "<w:tr>$left$right</w:tr>"
            }
            [void]$sb.Append((Tbl $rows))
        }
        "VERDICT" {
            $lines = @($b.Body -split "`n" | Where-Object { $_.Trim().Length -gt 0 })
            $title = $lines[0]
            $rest = @()
            if ($lines.Length -gt 1) { $rest = $lines[1..($lines.Length-1)] }
            [void]$sb.Append((Callout ("■ " + $title) $rest "FFF1F2" "E11D48" "BE123C"))
        }
        "H" { [void]$sb.Append((Para $b.Body 28 $true $false "08181F" 280 80)) }
        "P" { [void]$sb.Append((Para $b.Body 21 $false $false "16242C" 0 120)) }
        "FINDINGS" {
            $heads = @("ID","Analizowane twierdzenie","Podstawa prawna","Ryzyko","Istota naruszenia")
            $widths = @(700,2800,2100,1100,2660)
            $hcells = ""
            for ($c=0; $c -lt 5; $c++) { $hcells += Cell (Para $heads[$c] 16 $true $false "FFFFFF" 0 0) "08181F" $widths[$c] "08181F" }
            $rows = "<w:tr>$hcells</w:tr>"
            $ri = 0
            foreach ($line in ($b.Body -split "`n")) {
                if ($line -notmatch '\|\|\|') { continue }
                $cols = $line.Split(@('|||'),5,[System.StringSplitOptions]::None)
                $fill = if ($ri % 2 -eq 0) { "FFFFFF" } else { "F8FAFC" }
                $cells = ""
                for ($c=0; $c -lt 5; $c++) {
                    $color = "16242C"; $bold = $false
                    if ($c -eq 0) { $bold = $true; $color = "08181F" }
                    if ($c -eq 3) {
                        $bold = $true
                        $color = if ($cols[3].Trim() -eq "SREDNIE" -or $cols[3].Trim() -eq "ŚREDNIE") { "B45309" } else { "BE123C" }
                    }
                    $cells += Cell (Para $cols[$c].Trim() 15 $bold $false $color 0 0) $fill $widths[$c]
                }
                $rows += "<w:tr>$cells</w:tr>"
                $ri++
            }
            [void]$sb.Append((Tbl $rows))
        }
        "CASE" { [void]$sb.Append((Para $b.Body 23 $true $false "08181F" 240 40)) }
        "META2" {
            $color = if ($b.Body -match "SREDNIE|ŚREDNIE") { "B45309" } else { "BE123C" }
            [void]$sb.Append((Para $b.Body 17 $true $false $color 0 80))
        }
        "LABEL" { [void]$sb.Append((Para $b.Body 20 $true $false "16242C" 80 40)) }
        "Q" { [void]$sb.Append((Para $b.Body 20 $false $true "16242C" 0 60)) }
        "CALLOK" {
            $lines = @($b.Body -split "`n" | Where-Object { $_.Trim().Length -gt 0 })
            [void]$sb.Append((Callout ("■ " + $lines[0]) @($lines[1..($lines.Length-1)]) "ECFDF5" "059669" "047857"))
        }
        "CALLNO" {
            $lines = @($b.Body -split "`n" | Where-Object { $_.Trim().Length -gt 0 })
            [void]$sb.Append((Callout ("■ " + $lines[0]) @($lines[1..($lines.Length-1)]) "F8FAFC" "94A3B8" "5A6B77"))
        }
        "COMP" {
            $heads = @("Twierdzenie podmiotu","Rzeczywisty stan prawno-srodowiskowy")
            # Polish header written from markup first line if present; default below replaced by markup header row starting with #
            $rows = ""
            $ri = 0
            $started = $false
            foreach ($line in ($b.Body -split "`n")) {
                if ($line.StartsWith("#")) {
                    $hh = $line.Substring(1).Split(@('|||'),2,[System.StringSplitOptions]::None)
                    $rows += "<w:tr>$(Cell (Para $hh[0].Trim() 16 $true $false 'FFFFFF' 0 0) '08181F' 4680 '08181F')$(Cell (Para $hh[1].Trim() 16 $true $false 'FFFFFF' 0 0) '08181F' 4680 '08181F')</w:tr>"
                    $started = $true
                    continue
                }
                if ($line -notmatch '\|\|\|') { continue }
                $cols = $line.Split(@('|||'),2,[System.StringSplitOptions]::None)
                $fill = if ($ri % 2 -eq 0) { "FFFFFF" } else { "F8FAFC" }
                $rows += "<w:tr>$(Cell (Para $cols[0].Trim() 16 $false $false '16242C' 0 0) $fill 4680)$(Cell (Para $cols[1].Trim() 16 $false $false '16242C' 0 0) $fill 4680)</w:tr>"
                $ri++
            }
            if (-not $started) {
                $rows = "<w:tr>$(Cell (Para 'Twierdzenie podmiotu' 16 $true $false 'FFFFFF' 0 0) '08181F' 4680 '08181F')$(Cell (Para 'Rzeczywisty stan' 16 $true $false 'FFFFFF' 0 0) '08181F' 4680 '08181F')</w:tr>" + $rows
            }
            [void]$sb.Append((Tbl $rows))
        }
        "RISKH" { [void]$sb.Append((Para $b.Body 21 $true $false "16242C" 160 40)) }
        "STAGE" {
            $lines = @($b.Body -split "`n" | Where-Object { $_.Trim().Length -gt 0 })
            $kind = $lines[0]
            $map = @{
                "RED" = @("FFF1F2","E11D48","BE123C")
                "AMBER" = @("FFFBEB","D97706","B45309")
                "BLUE" = @("F0F9FF","0284C7","2A8AB0")
                "GREEN" = @("ECFDF5","059669","047857")
            }
            $c = $map[$kind]
            [void]$sb.Append((Callout ("■ " + $lines[1]) @($lines[2..($lines.Length-1)]) $c[0] $c[1] $c[2]))
        }
    }
}

$document = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
<w:body>
$($sb.ToString())
<w:sectPr>
  <w:headerReference w:type="default" r:id="rId2"/>
  <w:footerReference w:type="default" r:id="rId3"/>
  <w:pgSz w:w="11906" w:h="16838"/>
  <w:pgMar w:top="1361" w:right="964" w:bottom="1134" w:left="964" w:header="284" w:footer="284" w:gutter="0"/>
</w:sectPr>
</w:body>
</w:document>
"@

$header = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:tbl>
<w:tblPr><w:tblW w:w="5000" w:type="pct"/></w:tblPr>
<w:tr>
<w:tc><w:tcPr><w:tcW w:w="4500" w:type="dxa"/><w:shd w:val="clear" w:color="auto" w:fill="08181F"/></w:tcPr>
<w:p><w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:b/><w:sz w:val="20"/><w:color w:val="FFFFFF"/></w:rPr><w:t>ESG INSTITUTE</w:t></w:r></w:p></w:tc>
<w:tc><w:tcPr><w:tcW w:w="4860" w:type="dxa"/><w:shd w:val="clear" w:color="auto" w:fill="08181F"/></w:tcPr>
<w:p><w:pPr><w:jc w:val="right"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:b/><w:sz w:val="16"/><w:color w:val="FFFFFF"/></w:rPr><w:t>AUDYT RYZYK GREENWASHINGU</w:t></w:r></w:p>
<w:p><w:pPr><w:jc w:val="right"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="16"/><w:color w:val="81C0DF"/></w:rPr><w:t>esginstitute.eu</w:t></w:r></w:p></w:tc>
</w:tr></w:tbl><w:p/>
</w:hdr>
"@

$footer = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:p>
<w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="16"/><w:color w:val="5A6B77"/></w:rPr><w:t xml:space="preserve">ESG Institute  ·  Audyt Ryzyk Greenwashingu  ·  Hi DESIGN  ·  str. </w:t></w:r>
<w:r><w:rPr><w:sz w:val="16"/><w:color w:val="5A6B77"/></w:rPr><w:fldChar w:fldCharType="begin"/></w:r>
<w:r><w:rPr><w:sz w:val="16"/></w:rPr><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r>
<w:r><w:fldChar w:fldCharType="end"/></w:r>
</w:p>
</w:ftr>
"@

$styles = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="21"/></w:rPr></w:rPrDefault></w:docDefaults>
</w:styles>
"@

$ctypes = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Default Extension="png" ContentType="image/png"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>
<Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>
</Types>
"@

$rels = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>
"@

$drels = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>
<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/>
<Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/>
</Relationships>
"@

$utf8 = New-Object System.Text.UTF8Encoding $false
[System.IO.File]::WriteAllText((Join-Path $work "[Content_Types].xml"), $ctypes, $utf8)
[System.IO.File]::WriteAllText((Join-Path $work "_rels\.rels"), $rels, $utf8)
[System.IO.File]::WriteAllText((Join-Path $work "word\document.xml"), $document, $utf8)
[System.IO.File]::WriteAllText((Join-Path $work "word\styles.xml"), $styles, $utf8)
[System.IO.File]::WriteAllText((Join-Path $work "word\header1.xml"), $header, $utf8)
[System.IO.File]::WriteAllText((Join-Path $work "word\footer1.xml"), $footer, $utf8)
[System.IO.File]::WriteAllText((Join-Path $work "word\_rels\document.xml.rels"), $drels, $utf8)

if (Test-Path $docx) { Remove-Item $docx -Force }
$zip = [System.IO.Compression.ZipFile]::Open($docx, [System.IO.Compression.ZipArchiveMode]::Create)
Get-ChildItem $work -Recurse -File | ForEach-Object {
    $rel = $_.FullName.Substring($work.Length + 1).Replace("\","/")
    [System.IO.Compression.ZipFileExtensions]::CreateEntryFromFile($zip, $_.FullName, $rel) | Out-Null
}
$zip.Dispose()
Remove-Item $work -Recurse -Force
Write-Host "SAVED" (Get-Item $docx).Length "blocks" $blocks.Count
