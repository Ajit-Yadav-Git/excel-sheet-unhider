# Builds sample/sample-hidden-sheets.xlsx (raw OpenXML, no Excel/Python needed).
# Sheets: Summary (visible), Sales (hidden), Costs (hidden), Secrets (veryHidden), Notes (visible).
$ErrorActionPreference = "Stop"
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

$root = Split-Path -Parent $PSScriptRoot
$outDir = Join-Path $root "sample"
$out = Join-Path $outDir "sample-hidden-sheets.xlsx"
New-Item -ItemType Directory -Force $outDir | Out-Null
if (Test-Path $out) { Remove-Item $out -Force }

# name, state, rows (each row = array of cells; numbers stay numeric)
$sheets = @(
  @{ n = "Summary"; s = "visible";    r = @(@("Item","Value"), @("Total sales", 1250), @("Total costs", 830)) },
  @{ n = "Sales";   s = "hidden";     r = @(@("Region","Amount"), @("North", 400), @("South", 450), @("West", 400)) },
  @{ n = "Costs";   s = "hidden";     r = @(@("Category","Amount"), @("Ops", 500), @("Marketing", 330)) },
  @{ n = "Secrets"; s = "veryHidden"; r = @(@("Key","Note"), @("Only unhidable via code", "VBA/Office.js")) },
  @{ n = "Notes";   s = "visible";    r = @(@("Note"), @("Open this file with the add-in loaded; hidden sheets should appear.")) }
)

function Esc($t) { [System.Security.SecurityElement]::Escape([string]$t) }
function Col($i) { [string][char](65 + $i) }

function SheetXml($rows) {
  $sb = New-Object System.Text.StringBuilder
  [void]$sb.Append('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>')
  for ($ri = 0; $ri -lt $rows.Count; $ri++) {
    [void]$sb.Append("<row r=`"$($ri+1)`">")
    $cells = @($rows[$ri])
    for ($ci = 0; $ci -lt $cells.Count; $ci++) {
      $ref = "$(Col $ci)$($ri+1)"; $v = $cells[$ci]
      if ($v -is [int] -or $v -is [double]) { [void]$sb.Append("<c r=`"$ref`"><v>$v</v></c>") }
      else { [void]$sb.Append("<c r=`"$ref`" t=`"inlineStr`"><is><t>$(Esc $v)</t></is></c>") }
    }
    [void]$sb.Append("</row>")
  }
  [void]$sb.Append("</sheetData></worksheet>")
  $sb.ToString()
}

$parts = [ordered]@{}
$ct = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
$wb = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>'
$rels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
for ($i = 0; $i -lt $sheets.Count; $i++) {
  $id = $i + 1; $sh = $sheets[$i]
  $ct += "<Override PartName=`"/xl/worksheets/sheet$id.xml`" ContentType=`"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml`"/>"
  $state = if ($sh.s -eq "visible") { "" } else { " state=`"$($sh.s)`"" }
  $wb += "<sheet name=`"$($sh.n)`" sheetId=`"$id`"$state r:id=`"rId$id`"/>"
  $rels += "<Relationship Id=`"rId$id`" Type=`"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet`" Target=`"worksheets/sheet$id.xml`"/>"
  $parts["xl/worksheets/sheet$id.xml"] = SheetXml $sh.r
}
$parts["[Content_Types].xml"] = $ct + "</Types>"
$parts["_rels/.rels"] = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>'
$parts["xl/workbook.xml"] = $wb + "</sheets></workbook>"
$parts["xl/_rels/workbook.xml.rels"] = $rels + "</Relationships>"

$zip = [System.IO.Compression.ZipFile]::Open($out, "Create")
try {
  foreach ($k in $parts.Keys) {
    $e = $zip.CreateEntry($k)
    $w = New-Object System.IO.StreamWriter($e.Open(), (New-Object System.Text.UTF8Encoding($false)))
    $w.Write($parts[$k]); $w.Dispose()
  }
} finally { $zip.Dispose() }
Write-Host "Created $out"
