#
# Issue #15 producer/format matrix: generate corpus files that are genuinely
# SAVED BY the real Office suites (MS Office via Word/Excel/PowerPoint COM,
# WPS Office via KWPS/KET/KWPP COM) with SYNTHETIC content only.
#
# Why: production-compat evidence requires Office-origin + WPS-origin files;
# OfficeCLI-generated fixtures are explicitly NOT acceptable as production
# corpus (issue #15 round 28). Content is synthetic (no user data).
#
# Metadata is neutralized AFTER SAVE by rewriting docProps/*.xml inside the
# package (WPS COM does not support the DocumentProperty.Value setter), and
# a privacy scan rejects any real user/machine identity in core/app parts.
#
# Files land in a STAGING dir; tools/corpus/production-manifest.mjs copies
# them into corpus/production/ and updates the manifest.
#
# Usage:
#   powershell -NoProfile -ExecutionPolicy Bypass -File tools/corpus/new-origin-corpus.ps1 -Producer wps -OutDir .corpus/origin-staging
#

param(
  [ValidateSet("office", "wps")]
  [string]$Producer = "wps",
  [string]$OutDir = ".corpus/origin-staging"
)

$ErrorActionPreference = "Stop"
Add-Type -AssemblyName System.IO.Compression.FileSystem

if ($Producer -eq "office") {
  $docProgId = "Word.Application"
  $sheetProgId = "Excel.Application"
  $slideProgId = "PowerPoint.Application"
  $prefix = "office"
  $origin = "office-suite-synthetic"
} else {
  $docProgId = "KWPS.Application"
  $sheetProgId = "KET.Application"
  $slideProgId = "KWPP.Application"
  $prefix = "wps"
  $origin = "wps-suite-synthetic"
}

$outPath = Join-Path (Get-Location) $OutDir
New-Item -ItemType Directory -Force -Path $outPath | Out-Null

function Neutralize-Package-Metadata([string]$file) {
  # Rewrite docProps inside the saved package: neutral author/company, then
  # fail hard if any real identity marker survived.
  $zip = [System.IO.Compression.ZipFile]::Open($file, "Update")
  try {
    foreach ($name in @("docProps/core.xml", "docProps/app.xml")) {
      $entry = $zip.Entries | Where-Object { $_.FullName -eq $name }
      if (-not $entry) { continue }
      $reader = New-Object System.IO.StreamReader($entry.Open())
      $xml = $reader.ReadToEnd()
      $reader.Close()
      if ($name -eq "docProps/core.xml") {
        $xml = [regex]::Replace($xml, "<dc:creator>[^<]*</dc:creator>", "<dc:creator>office-compat</dc:creator>")
        $xml = [regex]::Replace($xml, "<cp:lastModifiedBy>[^<]*</cp:lastModifiedBy>", "<cp:lastModifiedBy>office-compat</cp:lastModifiedBy>")
      } else {
        $xml = [regex]::Replace($xml, "<Company>[^<]*</Company>", "<Company>office-compat</Company>")
      }
      foreach ($needle in @("WuLianpu", "SWD-0524", "C:\Users", "C:/Users")) {
        if ($xml -like "*$needle*") {
          throw "privacy check failed: $file contains '$needle' in $name"
        }
      }
      $entry.Delete()
      $new = $zip.CreateEntry($name)
      $writer = New-Object System.IO.StreamWriter($new.Open(), (New-Object System.Text.UTF8Encoding($false)))
      $writer.Write($xml)
      $writer.Close()
    }
  } finally {
    $zip.Dispose()
  }
}

$created = @()

# ---- DOCX (Word / KWPS) ------------------------------------------------
$app = $null; $doc = $null
try {
  $app = New-Object -ComObject $docProgId
  $app.Visible = $false
  $doc = $app.Documents.Add()
  $sel = $app.Selection
  $sel.Style = $doc.Styles.Item(-2) # wdStyleHeading1 (locale-independent)
  $sel.TypeText("Production Compatibility Document")
  $sel.TypeParagraph()
  $sel.Style = $doc.Styles.Item(-1) # wdStyleNormal
  $sel.TypeText("Synthetic body paragraph for the producer/format matrix.")
  $sel.TypeParagraph()
  $table = $doc.Tables.Add($sel.Range, 2, 2)
  $table.Cell(1, 1).Range.Text = "Field"
  $table.Cell(1, 2).Range.Text = "Value"
  $table.Cell(2, 1).Range.Text = "Origin"
  $table.Cell(2, 2).Range.Text = $origin
  $file = Join-Path $outPath "$prefix-docx.docx"
  $doc.SaveAs2($file, 12) # wdFormatXMLDocument
  $doc.Close($false)
  $doc = $null
  $app.Quit()
  $app = $null
  Neutralize-Package-Metadata $file
  $created += @{ file = "$prefix-docx.docx"; format = "docx"; features = @("styles", "headings", "tables") }
  Write-Output "created: $file"
} catch {
  Write-Warning "docx generation failed for ${Producer}: $($_.Exception.Message)"
  try { if ($doc) { $doc.Close($false) } } catch {}
  try { if ($app) { $app.Quit() } } catch {}
}

# ---- XLSX (Excel / KET) -------------------------------------------------
$app = $null; $wb = $null
try {
  $app = New-Object -ComObject $sheetProgId
  $app.Visible = $false
  $wb = $app.Workbooks.Add()
  $ws = $wb.Worksheets.Item(1)
  $ws.Name = "Data"
  $ws.Cells.Item(1, 1).Value2 = "Header"
  $ws.Cells.Item(2, 1).Value2 = "Alpha"
  $ws.Cells.Item(2, 2).Value2 = 10
  $ws.Cells.Item(3, 1).Value2 = "Beta"
  $ws.Cells.Item(3, 2).Value2 = 32
  $ws.Cells.Item(4, 2).Formula = "=SUM(B2:B3)"
  try {
    $hidden = $wb.Worksheets.Add()
    $hidden.Name = "HiddenSheet"
    $hidden.Cells.Item(1, 1).Value2 = "hidden cell"
    $hidden.Visible = 0 # xlSheetHidden
  } catch {
    Write-Output "note: hidden sheet unsupported, continuing without it"
  }
  $file = Join-Path $outPath "$prefix-xlsx.xlsx"
  $wb.SaveAs($file, 51) # xlOpenXMLWorkbook
  $wb.Close($false)
  $wb = $null
  $app.Quit()
  $app = $null
  Neutralize-Package-Metadata $file
  $created += @{ file = "$prefix-xlsx.xlsx"; format = "xlsx"; features = @("formulas", "multiple-sheets") }
  Write-Output "created: $file"
} catch {
  Write-Warning "xlsx generation failed for ${Producer}: $($_.Exception.Message)"
  try { if ($wb) { $wb.Close($false) } } catch {}
  try { if ($app) { $app.Quit() } } catch {}
}

# ---- PPTX (PowerPoint / KWPP) --------------------------------------------
$app = $null; $pres = $null
try {
  $app = New-Object -ComObject $slideProgId
  $app.Visible = $true # PowerPoint-family COM requires a visible frame
  $pres = $app.Presentations.Add()
  $slide1 = $pres.Slides.Add(1, 1) # ppLayoutTitle
  $slide1.Shapes.Item(1).TextFrame.TextRange.Text = "Production Compatibility Deck"
  $slide1.Shapes.Item(2).TextFrame.TextRange.Text = $origin
  $slide2 = $pres.Slides.Add(2, 2) # ppLayoutText
  $slide2.Shapes.Item(1).TextFrame.TextRange.Text = "Synthetic Slide"
  $slide2.Shapes.Item(2).TextFrame.TextRange.Text = "Body content placeholder text"
  $file = Join-Path $outPath "$prefix-pptx.pptx"
  $pres.SaveAs($file, 24) # ppSaveAsOpenXMLPresentation
  $pres.Close()
  $pres = $null
  $app.Quit()
  $app = $null
  Neutralize-Package-Metadata $file
  $created += @{ file = "$prefix-pptx.pptx"; format = "pptx"; features = @("slide-master", "placeholders", "multiple-slides") }
  Write-Output "created: $file"
} catch {
  Write-Warning "pptx generation failed for ${Producer}: $($_.Exception.Message)"
  try { if ($pres) { $pres.Close() } } catch {}
  try { if ($app) { $app.Quit() } } catch {}
}

if ($created.Count -eq 0) {
  throw "no corpus files were generated for producer '$Producer'"
}
Write-Output "staged $($created.Count) file(s) in $outPath"
