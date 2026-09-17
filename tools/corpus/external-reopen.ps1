#
# External reopen validator (issue #15): open a Runtime-saved Office file in
# the REAL WPS suite (KWPS/KET/KWPP COM — the actual application, not our
# Runtime, not the OfficeCLI engine) and assert the expected marker text is
# readable. This is the strongest automatable form of the "external
# Office/WPS reopen" acceptance — it cannot be satisfied by ZIP smoke.
#
# stdout: single-line JSON {opened, markerFound, app}; exit 0 on success.
#
# Usage:
#   powershell -NoProfile -ExecutionPolicy Bypass -File tools/corpus/external-reopen.ps1 -File <path> -Marker <text>
#

param(
  [Parameter(Mandatory = $true)][string]$File,
  [Parameter(Mandatory = $true)][string]$Marker
)

$ErrorActionPreference = "Stop"
$result = @{ opened = $false; markerFound = $false; app = $null }
$app = $null

function Emit([int]$code) {
  $result | ConvertTo-Json -Compress | Write-Output
  exit $code
}

try {
  $ext = [System.IO.Path]::GetExtension($File).ToLower()
  switch ($ext) {
    ".docx" { $app = New-Object -ComObject KWPS.Application; $result.app = "wps-word" }
    ".xlsx" { $app = New-Object -ComObject KET.Application; $result.app = "wps-excel" }
    ".pptx" { $app = New-Object -ComObject KWPP.Application; $result.app = "wps-ppt" }
    default { throw "unsupported extension: $ext" }
  }

  if ($ext -eq ".docx") {
    $app.Visible = $false
    $doc = $app.Documents.Open($File, $false, $true) # ConfirmConversions, ReadOnly
    $result.opened = $true
    $text = $doc.Content.Text
    $result.markerFound = $text -ne $null -and $text.Contains($Marker)
    $doc.Close($false)
  } elseif ($ext -eq ".xlsx") {
    $app.Visible = $false
    $wb = $app.Workbooks.Open($File, 0, $true) # UpdateLinks=0, ReadOnly
    $result.opened = $true
    foreach ($ws in $wb.Worksheets) {
      $value = $ws.Cells.Item(1, 1).Value2
      if ($value -eq $Marker) { $result.markerFound = $true; break }
    }
    $wb.Close($false)
  } else {
    $app.Visible = $true # PowerPoint-family COM requires a visible frame
    $pres = $app.Presentations.Open($File, $true, $false, $true) # ReadOnly, Untitled, WithWindow
    $result.opened = $true
    foreach ($slide in $pres.Slides) {
      foreach ($shape in $slide.Shapes) {
        try {
          $text = $shape.TextFrame.TextRange.Text
          if ($text -ne $null -and $text.Contains($Marker)) { $result.markerFound = $true; break }
        } catch { continue } # shapes without a text frame
      }
      if ($result.markerFound) { break }
    }
    $pres.Close()
  }
} catch {
  Write-Output ("error: " + $_.Exception.Message)
  Emit 2
} finally {
  try { if ($app) { $app.Quit() } } catch {}
}

if ($result.opened -and $result.markerFound) { Emit 0 } else { Emit 3 }
