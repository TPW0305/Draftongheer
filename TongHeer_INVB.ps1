# TongHeer_INVB.ps1
# แปลง Export Summary Report (xlsx) ของ TONG HEER FASTENERS (THAILAND) เป็นไฟล์ INVB_*.xlsx สำหรับนำเข้า ECS
# เปิดใช้ผ่าน TongHeer_INVB.bat (ดับเบิลคลิก หรือลากไฟล์รายงานมาวางบน .bat)
param(
    [string]$InputFile,
    [string]$OutFile,
    [switch]$NoGui,
    [switch]$AddPrice,
    [string]$SelfTestPng
)
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem
$Inv = [Globalization.CultureInfo]::InvariantCulture
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$CfgPath = Join-Path $ScriptDir 'tongheer_map.json'

# ค่าเริ่มต้น: ถูกเขียนลง tongheer_map.json ครั้งแรกที่รัน หลังจากนั้นแก้ที่ไฟล์ json ได้เลย
$DefaultCfg = @'
{
  "constants": {
    "taxno": "0205547018706",
    "incoterm": "FOB",
    "brand": "\"NO BRAND\"",
    "tariff": "000073181590",
    "se": "50001",
    "statistic": "000",
    "privilege": "003",
    "importtarrif": "9PART3",
    "unit": "KGM",
    "currency": "USD",
    "pkgunit": "PX",
    "productyear": 2026,
    "origin": "TH",
    "partnamet": "สกรูทำด้วยสเตนเลส",
    "pdtgroup": "02",
    "permit1": "BOI0205547018706|0994000740964|05/10/2026",
    "itemsep": "N"
  },
  "prefixes": [
    { "prefix": "1THA", "branch": "0000", "boi": "1366(2)/2548" },
    { "prefix": "2THA", "branch": "0001", "boi": "66-1786-2-04-1-0" },
    { "prefix": "3THA", "branch": "0001", "boi": "61-1365-1-00-0-0" }
  ],
  "consignees": [
    { "mark": "WURTH INDUSTRY USA (REVCAR FASTENERS)", "code": "WURTH RE" },
    { "mark": "BBI", "code": "TONG WIN" },
    { "mark": "FASTCO", "code": "FASTENAL" },
    { "mark": "S.S.C", "code": "STAR" }
  ],
  "partnos": [
    { "prefix": "1THA", "partno": "304 HEX (AM)", "subgroup": "304 HEX (AM)", "subConfirmed": false },
    { "prefix": "1THA", "partno": "316 HE(AM)", "subgroup": "316 HE(AM)", "subConfirmed": false },
    { "prefix": "1THA", "partno": "304 HEX LAG A2  (AM)", "subgroup": "304 HEX LAG A2", "subConfirmed": true },
    { "prefix": "1THA", "partno": "304 DIN 931(AM)", "subgroup": "304 DIN 931(AM)", "subConfirmed": false },
    { "prefix": "1THA", "partno": "316 DIN 931(AM)", "subgroup": "316 DIN 931", "subConfirmed": true },
    { "prefix": "1THA", "partno": "304 DIN 933 (AM)", "subgroup": "304 DIN 933 (AM)", "subConfirmed": false },
    { "prefix": "1THA", "partno": "DIN 912 SOCKET 316(AM)", "subgroup": "DIN 912 SOCKET 316(AM)", "subConfirmed": false },
    { "prefix": "1THA", "partno": "DIN 912 SOCKET304 (AM)", "subgroup": "DIN 912 SOCKET304 (AM)", "subConfirmed": false },
    { "prefix": "1THA", "partno": "933 DIN316 (AM)", "subgroup": "933 DIN316 (AM)", "subConfirmed": false },
    { "prefix": "2THA", "partno": "304 SOCKET(PT)", "subgroup": "304 SOCKET(PT)", "subConfirmed": true },
    { "prefix": "2THA", "partno": "316 SOCKET(PT)", "subgroup": "316 SOCKET(PT)", "subConfirmed": true },
    { "prefix": "2THA", "partno": "304 BUTTON SOCKET(PT)", "subgroup": "304 BUTTON SOCKET(PT)", "subConfirmed": true },
    { "prefix": "2THA", "partno": "316 BUTTON SOCKET(PT)", "subgroup": "316 BUTTON SOCKET(PT)", "subConfirmed": true },
    { "prefix": "2THA", "partno": "304  SOCKET FLAT (PT)", "subgroup": "304 SOCKET FLAT(PT)", "subConfirmed": true },
    { "prefix": "2THA", "partno": "316  SOCKET FLAT (PT)", "subgroup": "316 SOCKET FLAT(PT)", "subConfirmed": true },
    { "prefix": "2THA", "partno": "304 HEX CAR  (PT)", "subgroup": "304 HEX CAR(PT)", "subConfirmed": true },
    { "prefix": "2THA", "partno": "316 HEX CAR  (PT)", "subgroup": "316 HEX CAR(PT)", "subConfirmed": true },
    { "prefix": "2THA", "partno": "304 HEX (PT)", "subgroup": "304 HEX (PT)", "subConfirmed": true },
    { "prefix": "2THA", "partno": "316 HE(PT)", "subgroup": "316 HE(PT)", "subConfirmed": true },
    { "prefix": "2THA", "partno": "304 DIN 933 (PT)", "subgroup": "304 DIN 933 (PT)", "subConfirmed": true },
    { "prefix": "2THA", "partno": "304 HEX FLANGE (PT)", "subgroup": "304 HEX FLANGE (PT)", "subConfirmed": true },
    { "prefix": "2THA", "partno": "316 (BOLT PT)", "subgroup": "316 (BOLT PT)", "subConfirmed": true },
    { "prefix": "2THA", "partno": "304 THREA(PT)", "subgroup": "304 THREA(PT)", "subConfirmed": true },
    { "prefix": "3THA", "partno": "304 HEX (PT3)", "subgroup": "304 HEX (PT3)", "subConfirmed": false }
  ],
  "products": [
    { "prefix": "2THA", "name": "HEX. SOCKET SCREWS STAINLESS STEEL 304", "model": "HEX SOCKET 304", "partno": "304 SOCKET(PT)" },
    { "prefix": "2THA", "name": "HEX. SOCKET SCREWS STAINLESS STEEL 316", "model": "HEX SOCKET 316", "partno": "316 SOCKET(PT)" },
    { "prefix": "2THA", "name": "HEX. BUTTON SOCKET SCREWS STAINLESS STEEL 304", "model": "HEX. BUTTON 304", "partno": "304 BUTTON SOCKET(PT)" },
    { "prefix": "2THA", "name": "HEX. CARRIAGE BOLTS STAINLESS STEEL 316", "model": "HEX CARRIAGE 316", "partno": "316 HEX CAR  (PT)" },
    { "prefix": "2THA", "name": "HEX BOLTS ( NO WASHER FACE ) STAINLESS STEEL 316", "model": "HEX BOLTS 316", "partno": "316 (BOLT PT)" },
    { "prefix": "2THA", "name": "HEX CAP SCREWS STAINLESS STEEL 304", "model": "HEX CAP 304", "partno": "304 HEX (PT)" },
    { "prefix": "2THA", "name": "HEX CAP SCREWS STAINLESS STEEL 316", "model": "HEX CAP 316", "partno": "316 HE(PT)" }
  ]
}
'@

# ---------- ตัวช่วยทั่วไป ----------
function Norm([string]$s) { if ($null -eq $s) { return '' }; (($s -replace '\s+', ' ').Trim()).ToUpperInvariant() }
function KeyOf([string]$s) { if ($null -eq $s) { return '' }; ($s -replace '[^A-Za-z0-9]', '').ToUpperInvariant() }
function ColNum([string]$s) { $n = 0; foreach ($ch in $s.ToCharArray()) { $n = $n * 26 + ([int]$ch - 64) }; $n }
function ColName([int]$n) { $s = ''; while ($n -gt 0) { $m = ($n - 1) % 26; $s = [char](65 + $m) + $s; $n = [int][math]::Floor(($n - 1) / 26) }; $s }
function ToNum($s) {
    $d = 0.0
    if ([double]::TryParse([string]$s, [Globalization.NumberStyles]::Float, $Inv, [ref]$d)) { return $d }
    return $null
}
function Fmt2($v) { if ($null -eq $v) { return '-' }; ([double]$v).ToString('#,##0.00', $Inv) }
function SerialToText($serial) { [DateTime]::FromOADate([double]$serial).ToString('dd/MM/yyyy', $Inv) }
function TextToSerial([string]$t) {
    $d = [DateTime]::MinValue
    if ([DateTime]::TryParseExact($t.Trim(), [string[]]@('dd/MM/yyyy', 'd/M/yyyy'), $Inv, [Globalization.DateTimeStyles]::None, [ref]$d)) { return [int]$d.ToOADate() }
    return $null
}

# ---------- config ----------
function Load-Config {
    $utf8 = New-Object Text.UTF8Encoding($true)
    if (-not (Test-Path -LiteralPath $CfgPath)) { [IO.File]::WriteAllText($CfgPath, $DefaultCfg, $utf8) }
    [IO.File]::ReadAllText($CfgPath, [Text.Encoding]::UTF8) | ConvertFrom-Json
}
function Save-Config($cfg) {
    $utf8 = New-Object Text.UTF8Encoding($true)
    [IO.File]::WriteAllText($CfgPath, ($cfg | ConvertTo-Json -Depth 8), $utf8)
}
function Find-Consignee($cfg, [string]$mark) {
    $k = Norm $mark
    foreach ($c in $cfg.consignees) { if ((Norm $c.mark) -eq $k) { return $c.code } }
    return ''
}
function Find-Product($cfg, [string]$prefix, [string]$name, [string]$model) {
    $n = Norm $name; $m = Norm $model
    foreach ($p in $cfg.products) { if ($p.prefix -eq $prefix -and (Norm $p.name) -eq $n -and (Norm $p.model) -eq $m) { return $p } }
    return $null
}
function Find-PartNo($cfg, [string]$prefix, [string]$partno) {
    foreach ($p in $cfg.partnos) { if ($p.prefix -eq $prefix -and $p.partno -ceq $partno) { return $p } }
    return $null
}
function Default-SubGroup($cfg, [string]$prefix, [string]$partno) {
    $p = Find-PartNo $cfg $prefix $partno
    if ($p) { return $p.subgroup }
    return $partno
}

# ---------- อ่าน xlsx (ไม่ต้องใช้ Excel, อ่านได้แม้ไฟล์เปิดค้างอยู่) ----------
function Read-ZipText($zip, [string]$name) {
    $e = $zip.GetEntry($name)
    if (-not $e) { return $null }
    $sr = New-Object IO.StreamReader($e.Open(), [Text.Encoding]::UTF8)
    try { return $sr.ReadToEnd() } finally { $sr.Close() }
}
function Read-Xlsx([string]$Path) {
    $cells = New-Object Collections.ArrayList
    $fs = New-Object IO.FileStream($Path, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::ReadWrite)
    try {
        $zip = New-Object IO.Compression.ZipArchive($fs, [IO.Compression.ZipArchiveMode]::Read)
        $ss = New-Object Collections.ArrayList
        $sst = Read-ZipText $zip 'xl/sharedStrings.xml'
        if ($sst) { foreach ($si in ([xml]$sst).sst.si) { [void]$ss.Add($si.InnerText) } }
        $sheetPath = 'xl/worksheets/sheet1.xml'
        $wbText = Read-ZipText $zip 'xl/workbook.xml'
        $relText = Read-ZipText $zip 'xl/_rels/workbook.xml.rels'
        if ($wbText -and $relText) {
            $sheet = @(([xml]$wbText).workbook.sheets.sheet)[0]
            $rid = $sheet.GetAttribute('id', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships')
            foreach ($r in ([xml]$relText).Relationships.Relationship) {
                if ($r.Id -eq $rid) {
                    $t = [string]$r.Target
                    if ($t.StartsWith('/')) { $sheetPath = $t.TrimStart('/') } else { $sheetPath = 'xl/' + $t }
                }
            }
        }
        $sheetText = Read-ZipText $zip $sheetPath
        if (-not $sheetText) { throw "ไม่พบชีตในไฟล์ $Path" }
        foreach ($row in ([xml]$sheetText).worksheet.sheetData.row) {
            foreach ($c in $row.c) {
                $ref = [string]$c.r
                if ($ref -notmatch '^([A-Z]+)(\d+)$') { continue }
                $col = $Matches[1]; $rn = [int]$Matches[2]
                $type = $c.GetAttribute('t')
                $val = $null
                if ($type -eq 'inlineStr') { $n = $c.SelectSingleNode('*[local-name()="is"]'); if ($n) { $val = $n.InnerText } }
                else {
                    $n = $c.SelectSingleNode('*[local-name()="v"]')
                    if ($n) { $val = $n.InnerText; if ($type -eq 's') { $val = [string]$ss[[int]$val] } }
                }
                if ($null -eq $val -or ([string]$val).Trim() -eq '') { continue }
                [void]$cells.Add([pscustomobject]@{ Col = $col; ColN = (ColNum $col); Row = $rn; Val = [string]$val; IsText = ($type -eq 's' -or $type -eq 'inlineStr' -or $type -eq 'str') })
            }
        }
        $zip.Dispose()
    } finally { $fs.Close() }
    return , $cells
}

# ---------- แกะรายงาน Export Summary Report ----------
function Parse-Report([string]$Path) {
    if ([IO.Path]::GetExtension($Path).ToLowerInvariant() -ne '.xlsx') { throw "รองรับเฉพาะไฟล์ .xlsx (ถ้าเป็น .xls ให้ Save As เป็น .xlsx ก่อน)" }
    $cells = Read-Xlsx $Path
    $idx = @{}
    foreach ($c in $cells) { $idx["$($c.ColN)|$($c.Row)"] = $c }
    $warn = New-Object Collections.Generic.List[string]

    $findKey = { param($k) foreach ($c in $cells) { if ($c.IsText -and (KeyOf $c.Val) -eq $k) { return $c } }; return $null }
    $findLike = { param($k) foreach ($c in $cells) { if ($c.IsText -and (KeyOf $c.Val) -like $k) { return $c } }; return $null }
    $right = { param($cell) if (-not $cell) { return $null }; $best = $null; foreach ($c in $cells) { if ($c.Row -eq $cell.Row -and $c.ColN -gt $cell.ColN) { if (-not $best -or $c.ColN -lt $best.ColN) { $best = $c } } }; return $best }

    if (-not (& $findLike 'EXPORTSUMMARYREPORT*')) { $warn.Add('ไม่พบหัวรายงาน "Export Summary Report" - ตรวจว่าเลือกไฟล์ถูกหรือไม่') }

    # --- เลขอินวอย: ใช้ตารางสรุปตามอินวอยด้านล่างถ้ามี ไม่งั้นใช้ช่อง Invoice# ---
    $rxInv = '^\s*([1-3]THA\d+)\s*$'
    $invoices = New-Object Collections.Generic.List[string]
    $invHdr = & $findKey 'INVOICENO'
    $totAmount = $null
    if ($invHdr) {
        $r = $invHdr.Row + 1
        while ($true) {
            $c = $idx["$($invHdr.ColN)|$r"]
            if ($c -and $c.Val -match $rxInv) { $invoices.Add($Matches[1]); $r++ } else { break }
        }
        $amtHdr = $null
        foreach ($c in $cells) { if ($c.Row -eq $invHdr.Row -and (KeyOf $c.Val) -eq 'AMOUNT') { $amtHdr = $c } }
        $totRow = $null
        foreach ($c in $cells) { if ($c.ColN -eq $invHdr.ColN -and $c.Row -gt $invHdr.Row -and (KeyOf $c.Val) -eq 'TOTAL') { $totRow = $c } }
        if ($amtHdr -and $totRow) { $t = $idx["$($amtHdr.ColN)|$($totRow.Row)"]; if ($t) { $totAmount = ToNum $t.Val } }
    }
    $invLbl = & $findLike 'INVOICE'
    $invCell = & $right $invLbl
    if ($invoices.Count -eq 0) {
        if ($invCell -and $invCell.Val -match $rxInv) { $invoices.Add($Matches[1]) }
        else { $warn.Add('ไม่พบเลขอินวอยในรายงาน - กรอก INV_NO เอง') }
    }

    # --- วันที่อินวอย / ETD ---
    $dateSerial = $null
    if ($invLbl) {
        foreach ($c in ($cells | Where-Object { $_.Row -eq $invLbl.Row -and $_.ColN -gt $invLbl.ColN } | Sort-Object ColN)) {
            if ($invCell -and $c.ColN -eq $invCell.ColN) { continue }
            $n = ToNum $c.Val
            if (-not $c.IsText -and $null -ne $n -and $n -gt 30000 -and $n -lt 70000) { $dateSerial = [int][math]::Floor($n); break }
            if ($c.Val -match '^\s*(\d{1,2})[-/](\d{1,2})[-/](\d{2,4})') {
                $y = [int]$Matches[3]; if ($y -lt 100) { $y += 2000 }
                try { $dateSerial = [int](New-Object DateTime($y, [int]$Matches[1], [int]$Matches[2])).ToOADate(); break } catch { }
            }
        }
    }
    $etdSerial = $null
    $etdCell = & $right (& $findKey 'ETD')
    if ($etdCell) { $n = ToNum $etdCell.Val; if ($null -ne $n -and $n -gt 30000 -and $n -lt 70000) { $etdSerial = [int][math]::Floor($n) } }
    if ($null -eq $dateSerial) { $warn.Add('ไม่พบวันที่อินวอยในรายงาน - กรอก INV_DATE เอง (dd/mm/yyyy)') }
    elseif ($null -ne $etdSerial) {
        $d = [DateTime]::FromOADate($dateSerial)
        $gap = $etdSerial - $dateSerial
        if ($gap -lt 0 -or $gap -gt 60) {
            $msg = "INV_DATE $(SerialToText $dateSerial) ไม่สอดคล้องกับ ETD $(SerialToText $etdSerial)"
            if ($d.Day -le 12 -and $d.Day -ne $d.Month) {
                $sw = [int](New-Object DateTime($d.Year, $d.Day, $d.Month)).ToOADate()
                $g2 = $etdSerial - $sw
                if ($g2 -ge 0 -and $g2 -le 60) { $msg += " - อาจสลับวัน/เดือน น่าจะเป็น $(SerialToText $sw)" }
            }
            $warn.Add($msg + ' (ตรวจและแก้ในช่อง INV_DATE)')
        }
    } else { $warn.Add('ไม่พบ ETD จึงตรวจการสลับวัน/เดือนของ INV_DATE ไม่ได้') }

    # --- รายการสินค้า (ตารางส่วนบน) ---
    $netH = & $findKey 'NETWEIGHT'
    if (-not $netH) { throw 'ไม่พบหัวคอลัมน์ NetWeight ในรายงาน' }
    $grossH = $null; $amtH = $null; $descH = $null
    foreach ($c in $cells) {
        if ($c.Row -ne $netH.Row) { continue }
        switch (KeyOf $c.Val) { 'GROSSWEIGHT' { $grossH = $c } 'AMOUNT' { $amtH = $c } 'SUBGROUP' { $descH = $c } }
    }
    if (-not $grossH -or -not $amtH) { throw 'ไม่พบหัวคอลัมน์ GrossWeight / Amount ในรายงาน' }
    $descCol = 3; if ($descH) { $descCol = $descH.ColN }
    $subRows = @($cells | Where-Object { $_.IsText -and (KeyOf $_.Val) -like 'SUBTOTALFOR*' } | Sort-Object Row | ForEach-Object { $_.Row })
    if ($subRows.Count -eq 0) { throw 'ไม่พบรายการสินค้า (บรรทัด "Sub Total For :") ในรายงาน' }
    $items = New-Object Collections.ArrayList
    $prev = $netH.Row
    foreach ($r in $subRows) {
        $lines = @($cells | Where-Object { $_.ColN -eq $descCol -and $_.Row -gt $prev -and $_.Row -lt $r } | Sort-Object Row | ForEach-Object { ($_.Val -replace '\s*[\r\n]+\s*', ' ').Trim() })
        $name = ''; $modelLine = ''; $raw = ''
        foreach ($l in $lines) {
            if ($l -match '^\(\s*MODEL\s*:\s*(.*?)\s*\)\s*$') { if (-not $modelLine) { $modelLine = $l; $raw = $Matches[1] } }
            elseif (-not $name) { $name = $l }
        }
        $n = $idx["$($netH.ColN)|$r"]; $g = $idx["$($grossH.ColN)|$r"]; $a = $idx["$($amtH.ColN)|$r"]
        $it = [pscustomobject]@{
            Name = $name; ModelLine = $modelLine; RawCode = $raw
            Net = $(if ($n) { ToNum $n.Val }); Gross = $(if ($g) { ToNum $g.Val }); Amount = $(if ($a) { ToNum $a.Val })
            PartNo = ''; SubGroup = ''
        }
        $no = $items.Count + 1
        if (-not $name) { $warn.Add("รายการ ${no}: ไม่พบชื่อสินค้า") }
        if (-not $modelLine) { $warn.Add("รายการ ${no}: ไม่พบบรรทัด ( MODEL : ... ) - RAWCODE จะว่าง") }
        if ($null -eq $it.Net -or $null -eq $it.Gross -or $null -eq $it.Amount) { throw "รายการ ${no} (แถว $r): อ่าน Net/Gross/Amount ไม่ได้" }
        $it.Net = [math]::Round($it.Net, 2); $it.Gross = [math]::Round($it.Gross, 2); $it.Amount = [math]::Round($it.Amount, 2)
        [void]$items.Add($it)
        $prev = $r
    }

    # --- ยอดรวม / พาเลท / marking ---
    $totNet = $null; $c = & $right (& $findLike 'TOTALNETWEIGHT*'); if ($c) { $totNet = ToNum $c.Val }
    $totGross = $null; $c = & $right (& $findLike 'TOTALGROSSWEIGHT*'); if ($c) { $totGross = ToNum $c.Val }
    if ($null -eq $totAmount) {
        # ไม่มีตารางสรุปตามอินวอย: ใช้แถวผลรวมใต้รายการสุดท้าย
        foreach ($r in ($prev + 1)..($prev + 3)) { $c = $idx["$($amtH.ColN)|$r"]; if ($c -and -not $c.IsText) { $totAmount = ToNum $c.Val; break } }
    }
    $pallets = $null; $c = & $right (& $findLike 'TOTALPALLET*'); if ($c) { $pallets = ToNum $c.Val }
    if ($null -eq $pallets) { $warn.Add('ไม่พบ Total Pallet - กรอก NUMPKG เอง') }
    $marking = ''; $c = & $right (& $findKey 'MARKING'); if ($c) { $marking = $c.Val.Trim() }
    if (-not $marking) { $warn.Add('ไม่พบ Marking ท้ายรายงาน - กรอก MARK1 เอง') }
    $customer = ''; $c = & $right (& $findKey 'CUSTOMER'); if ($c) { $customer = $c.Val.Trim() }

    [pscustomobject]@{
        Path = $Path; Invoices = @($invoices); DateSerial = $dateSerial; EtdSerial = $etdSerial
        Items = $items; TotNet = $totNet; TotGross = $totGross; TotAmount = $totAmount
        Pallets = $pallets; Marking = $marking; Customer = $customer; Warnings = $warn
    }
}

# ---------- ใส่กฎของ skill ลงบนรายงานที่แกะแล้ว ----------
function Resolve-Report($rep, $cfg) {
    $warn = New-Object Collections.Generic.List[string]
    foreach ($w in $rep.Warnings) { $warn.Add($w) }
    $invNo = ''; $prefix = ''; $branch = ''; $boi = ''
    if ($rep.Invoices.Count -gt 0) {
        $first = $rep.Invoices[0]
        $prefix = $first.Substring(0, 4)
        $invNo = $first
        if ($rep.Invoices.Count -eq 2) { $warn.Add("รายงานมี 2 อินวอย ($($rep.Invoices -join ', ')) - ใส่ INV_NO เลขแรกเลขเดียว และรวมเป็นไฟล์เดียว") }
        elseif ($rep.Invoices.Count -gt 2) {
            $rest = @($rep.Invoices | Select-Object -Skip 1 | ForEach-Object { $_.Substring($_.Length - 3) })
            $invNo = $first + ',' + ($rest -join ',')
            $warn.Add("รายงานมี $($rep.Invoices.Count) อินวอย - ใส่ INV_NO แบบย่อ $invNo (ยังสงสัยว่าเลขคั่นจุลภาคทำให้ ECS คิดราคาผิดหรือไม่)")
        }
        if (@($rep.Invoices | Where-Object { $_.Substring(0, 4) -ne $prefix }).Count -gt 0) { $warn.Add("อินวอยในรายงานมี prefix ไม่เหมือนกัน - ใช้กฎของ $prefix ทั้งไฟล์") }
        $pf = $cfg.prefixes | Where-Object { $_.prefix -eq $prefix } | Select-Object -First 1
        if ($pf) { $branch = $pf.branch; $boi = $pf.boi } else { $warn.Add("ไม่รู้จัก prefix $prefix (ไม่มี EXP_BRANCH / BOI)") }
    }
    $consignee = Find-Consignee $cfg $rep.Marking
    if ($rep.Marking -and -not $consignee) { $warn.Add("Marking `"$($rep.Marking)`" ยังไม่เคยกำหนด CONSIGNEE - กรอกรหัสตาม ECS (โปรแกรมจะจำไว้)") }
    $i = 0
    foreach ($it in $rep.Items) {
        $i++
        $p = Find-Product $cfg $prefix $it.Name $it.RawCode
        if ($p) {
            $it.PartNo = $p.partno
            $it.SubGroup = Default-SubGroup $cfg $prefix $p.partno
        } else {
            $it.PartNo = ''; $it.SubGroup = ''
            $warn.Add("รายการ ${i}: ยังไม่รู้ PARTNO ของ `"$($it.Name) $($it.ModelLine)`" ($prefix) - กรอกให้ตรง DATA_TONGHEER (โปรแกรมจะจำไว้)")
        }
    }
    [pscustomobject]@{ InvNo = $invNo; Prefix = $prefix; Branch = $branch; Boi = $boi; Consignee = $consignee; Warnings = $warn }
}
function Get-SubGroupWarnings($cfg, [string]$prefix, $items) {
    $out = New-Object Collections.Generic.List[string]
    $i = 0
    foreach ($it in $items) {
        $i++
        if (-not $it.PartNo) { continue }
        $p = Find-PartNo $cfg $prefix $it.PartNo
        if (-not $p -or -not $p.subConfirmed) { $out.Add("รายการ ${i}: SUBGROUP `"$($it.SubGroup)`" ยังไม่เคยยืนยันกับหน้าจอ ECS - เช็กก่อนใช้") }
    }
    return , $out
}
function Get-TotalsText($rep) {
    $sn = 0.0; $sg = 0.0; $sa = 0.0
    foreach ($it in $rep.Items) { $sn += $it.Net; $sg += $it.Gross; $sa += $it.Amount }
    $one = {
        param($label, $sum, $ref)
        if ($null -eq $ref) { return "$label $(Fmt2 $sum) (ไม่พบยอดในรายงาน)" }
        $diff = [math]::Round($sum - $ref, 2)
        if ([math]::Abs($diff) -lt 0.005) { return "$label $(Fmt2 $sum) = รายงาน" }
        return "$label $(Fmt2 $sum) / รายงาน $(Fmt2 $ref) (ต่าง $($diff.ToString('0.00', $Inv)))"
    }
    "ยอดรวม $($rep.Items.Count) รายการ:  " + (& $one 'Net' $sn $rep.TotNet) + '   |   ' + (& $one 'Gross' $sg $rep.TotGross) + '   |   ' + (& $one 'Amount' $sa $rep.TotAmount)
}

# ---------- เขียน INVB xlsx ----------
function Write-Invb($m, $cfg, [string]$Path, [bool]$WithPrice) {
    $k = $cfg.constants
    $esc = { param($s) [Security.SecurityElement]::Escape([string]$s) }
    $num = { param($v) ([double]$v).ToString('R', $Inv) }
    $headers = 'EXP_TAXNO,EXP_BRANCH,INV_NO,INV_DATE,INCOTERM,PAYTERM,CONSIGNEE,PONO,ITEM,PDTGROUP,PARTNO,PARTNAME,PARTNAME2,PARTNAME3,PARTNAMET,PARTNAMET2,PARTNAMET3,BRAND,MARK1,MARK2,MARK3,MARK4,TARIFF,SE,STATISTIC,PRIVILEGE,IMPORTTARRIF,QUANTITY,QUNIT,TARIFFQTY,TARIFFUNITY,GROSS,WEIGHT,WEIGHTU,INVCURR,AMOUNT,NUMPKG,UNTPKG,PRODUCTYEAR,FOC,FORMULAR1,F19,RAWCODE,BOI,BOILIC,FREEZONE,EPZ,IMPTAXINC,ORIGIN,DREMARK,MODELNO,MODELVERSION,MODELTAXNO,RTCPRODUCTCODE,PERMIT1,REFNO,ITEMSEP,SUBGROUP'.Split(',')
    if ($WithPrice) { $headers += 'PRICE' }
    # ชนิดเซลล์: T=ข้อความ(Text format) S=ข้อความ N=ตัวเลขทศนิยม I=จำนวนเต็ม D=วันที่ E=ว่าง X=ไม่เขียน P=ข้อความไม่มีสไตล์ G=SUBGROUP R=ตัวเลขเต็มความละเอียด
    $style = @{ T = 2; S = 4; N = 5; I = 4; D = 6; E = 4; G = 7; R = 4 }
    $sb = New-Object Text.StringBuilder
    $cell = {
        param($ref, $kind, $v)
        switch ($kind) {
            'X' { }
            'E' { [void]$sb.Append("<c r=`"$ref`" s=`"4`"/>") }
            'P' { [void]$sb.Append("<c r=`"$ref`" t=`"inlineStr`"><is><t xml:space=`"preserve`">$(& $esc $v)</t></is></c>") }
            { $_ -in 'T', 'S', 'G' } { [void]$sb.Append("<c r=`"$ref`" s=`"$($style[$kind])`" t=`"inlineStr`"><is><t xml:space=`"preserve`">$(& $esc $v)</t></is></c>") }
            default { [void]$sb.Append("<c r=`"$ref`" s=`"$($style[$kind])`"><v>$(& $num $v)</v></c>") }
        }
    }
    [void]$sb.Append('<row r="1">')
    for ($c = 0; $c -lt $headers.Count; $c++) { [void]$sb.Append("<c r=`"$(ColName ($c + 1))1`" s=`"1`" t=`"inlineStr`"><is><t>$($headers[$c])</t></is></c>") }
    [void]$sb.Append('</row>')
    $r = 1; $i = 0
    foreach ($it in $m.Items) {
        $r++; $i++
        $mark = '-DO-'; $pkg = 0
        if ($i -eq 1) { $mark = $m.Mark; $pkg = $m.Pallets }
        $row = @(
            @('T', $k.taxno), @('T', $m.Branch), @('T', $m.InvNo), @('D', $m.DateSerial), @('S', $k.incoterm), @('E'), @('T', $m.Consignee), @('E'),
            @('I', $i), @('T', $k.pdtgroup), @('T', $it.PartNo), @('S', $it.Name), @('S', $it.ModelLine), @('E'), @('S', $k.partnamet), @('E'), @('E'),
            @('S', $k.brand), @('S', $mark), @('E'), @('E'), @('E'), @('T', $k.tariff), @('T', $k.se), @('T', $k.statistic), @('T', $k.privilege), @('S', $k.importtarrif),
            @('N', $it.Net), @('S', $k.unit), @('N', $it.Net), @('S', $k.unit), @('N', $it.Gross), @('N', $it.Net), @('S', $k.unit), @('S', $k.currency), @('N', $it.Amount),
            @('I', $pkg), @('S', $k.pkgunit), @('I', $k.productyear), @('E'), @('E'), @('E'), @('T', $it.RawCode), @('T', $m.Boi), @('E'), @('E'), @('E'), @('E'),
            @('S', $k.origin), @('E'), @('E'), @('E'), @('E'), @('E'), @('T', $k.permit1), @('X'), @('P', $k.itemsep), @('G', $it.SubGroup)
        )
        if ($WithPrice) { $row += , @('R', ($it.Amount / $it.Net)) }
        [void]$sb.Append("<row r=`"$r`">")
        for ($c = 0; $c -lt $row.Count; $c++) {
            $spec = $row[$c]; $kind = $spec[0]; $v = $null
            if ($spec.Count -gt 1) { $v = $spec[1] }
            if ($kind -in 'T', 'S', 'G' -and [string]$v -eq '') { $kind = 'E' }
            & $cell "$(ColName ($c + 1))$r" $kind $v
        }
        [void]$sb.Append('</row>')
    }
    # ความกว้างคอลัมน์ตามไฟล์ INVB ตัวอย่าง
    $wide = @{ 3 = 26; 12 = 52; 13 = 30; 15 = 30; 19 = 40; 27 = 14; 31 = 13; 39 = 13; 52 = 14; 54 = 16 }
    $cols = New-Object Text.StringBuilder
    for ($c = 1; $c -le 55; $c++) { $w = 12; if ($wide.ContainsKey($c)) { $w = $wide[$c] }; [void]$cols.Append("<col width=`"$w`" customWidth=`"1`" min=`"$c`" max=`"$c`"/>") }
    $ns = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
    $hd = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + "`n"
    $sheetXml = $hd + "<worksheet xmlns=`"$ns`"><dimension ref=`"A1:$(ColName $headers.Count)$r`"/><sheetViews><sheetView workbookViewId=`"0`"><pane ySplit=`"1`" topLeftCell=`"A2`" activePane=`"bottomLeft`" state=`"frozen`"/><selection pane=`"bottomLeft`" activeCell=`"A1`" sqref=`"A1`"/></sheetView></sheetViews><sheetFormatPr baseColWidth=`"8`" defaultRowHeight=`"15`"/><cols>$cols</cols><sheetData>$sb</sheetData><pageMargins left=`"0.75`" right=`"0.75`" top=`"1`" bottom=`"1`" header=`"0.5`" footer=`"0.5`"/></worksheet>"
    $stylesXml = $hd + "<styleSheet xmlns=`"$ns`"><numFmts count=`"1`"><numFmt numFmtId=`"164`" formatCode=`"dd/mm/yyyy`"/></numFmts><fonts count=`"3`"><font><sz val=`"11`"/><name val=`"Calibri`"/><family val=`"2`"/></font><font><b/><sz val=`"10`"/><name val=`"Arial`"/></font><font><sz val=`"10`"/><name val=`"Arial`"/></font></fonts><fills count=`"2`"><fill><patternFill patternType=`"none`"/></fill><fill><patternFill patternType=`"gray125`"/></fill></fills><borders count=`"1`"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count=`"1`"><xf numFmtId=`"0`" fontId=`"0`" fillId=`"0`" borderId=`"0`"/></cellStyleXfs><cellXfs count=`"8`"><xf numFmtId=`"0`" fontId=`"0`" fillId=`"0`" borderId=`"0`" xfId=`"0`"/><xf numFmtId=`"0`" fontId=`"1`" fillId=`"0`" borderId=`"0`" xfId=`"0`" applyFont=`"1`"/><xf numFmtId=`"49`" fontId=`"2`" fillId=`"0`" borderId=`"0`" xfId=`"0`" applyNumberFormat=`"1`" applyFont=`"1`"/><xf numFmtId=`"164`" fontId=`"2`" fillId=`"0`" borderId=`"0`" xfId=`"0`" applyNumberFormat=`"1`" applyFont=`"1`"/><xf numFmtId=`"0`" fontId=`"2`" fillId=`"0`" borderId=`"0`" xfId=`"0`" applyFont=`"1`"/><xf numFmtId=`"4`" fontId=`"2`" fillId=`"0`" borderId=`"0`" xfId=`"0`" applyNumberFormat=`"1`" applyFont=`"1`"/><xf numFmtId=`"164`" fontId=`"2`" fillId=`"0`" borderId=`"0`" xfId=`"0`" applyNumberFormat=`"1`" applyFont=`"1`"/><xf numFmtId=`"49`" fontId=`"0`" fillId=`"0`" borderId=`"0`" xfId=`"0`" applyNumberFormat=`"1`"/></cellXfs><cellStyles count=`"1`"><cellStyle name=`"Normal`" xfId=`"0`" builtinId=`"0`"/></cellStyles></styleSheet>"
    $wbXml = $hd + "<workbook xmlns=`"$ns`" xmlns:r=`"http://schemas.openxmlformats.org/officeDocument/2006/relationships`"><sheets><sheet name=`"SHEET1`" sheetId=`"1`" r:id=`"rId1`"/></sheets></workbook>"
    $wbRels = $hd + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>'
    $rootRels = $hd + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>'
    $ctypes = $hd + '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>'
    $parts = [ordered]@{ '[Content_Types].xml' = $ctypes; '_rels/.rels' = $rootRels; 'xl/workbook.xml' = $wbXml; 'xl/_rels/workbook.xml.rels' = $wbRels; 'xl/styles.xml' = $stylesXml; 'xl/worksheets/sheet1.xml' = $sheetXml }
    $enc = New-Object Text.UTF8Encoding($false)
    $fs = New-Object IO.FileStream($Path, [IO.FileMode]::Create, [IO.FileAccess]::Write)
    try {
        $zip = New-Object IO.Compression.ZipArchive($fs, [IO.Compression.ZipArchiveMode]::Create)
        foreach ($name in $parts.Keys) {
            $e = $zip.CreateEntry($name)
            $sw = New-Object IO.StreamWriter($e.Open(), $enc)
            $sw.Write($parts[$name]); $sw.Close()
        }
        $zip.Dispose()
    } finally { $fs.Close() }
}

# ---------- จำค่าที่ผู้ใช้กรอก ----------
function Learn-Model($cfg, $m, [string]$prefix) {
    $changed = $false
    if ($m.Mark -and $m.Consignee) {
        $hit = $null
        foreach ($c in $cfg.consignees) { if ((Norm $c.mark) -eq (Norm $m.Mark)) { $hit = $c } }
        if (-not $hit) { $cfg.consignees += [pscustomobject]@{ mark = $m.Mark; code = $m.Consignee }; $changed = $true }
        elseif ($hit.code -cne $m.Consignee) { $hit.code = $m.Consignee; $changed = $true }
    }
    foreach ($it in $m.Items) {
        $p = Find-Product $cfg $prefix $it.Name $it.RawCode
        if (-not $p) { $cfg.products += [pscustomobject]@{ prefix = $prefix; name = $it.Name; model = $it.RawCode; partno = $it.PartNo }; $changed = $true }
        elseif ($p.partno -cne $it.PartNo) { $p.partno = $it.PartNo; $changed = $true }
        $pn = Find-PartNo $cfg $prefix $it.PartNo
        if (-not $pn) { $cfg.partnos += [pscustomobject]@{ prefix = $prefix; partno = $it.PartNo; subgroup = $it.SubGroup; subConfirmed = $true }; $changed = $true }
        elseif ($pn.subgroup -cne $it.SubGroup) { $pn.subgroup = $it.SubGroup; $pn.subConfirmed = $true; $changed = $true }
    }
    if ($changed) { Save-Config $cfg }
}
function Out-PathFor($rep, [string]$invNo) {
    $first = ($invNo -split ',')[0].Trim()
    Join-Path (Split-Path -Parent $rep.Path) "INVB_$first.xlsx"
}

# ================= โหมดคำสั่ง (ไม่มีหน้าต่าง) =================
if ($NoGui) {
    $cfg = Load-Config
    $rep = Parse-Report (Resolve-Path -LiteralPath $InputFile).Path
    $res = Resolve-Report $rep $cfg
    $missing = @($rep.Items | Where-Object { -not $_.PartNo })
    if ($missing.Count -gt 0 -or -not $res.Consignee -or -not $res.Branch -or $null -eq $rep.DateSerial -or $null -eq $rep.Pallets) {
        $res.Warnings | ForEach-Object { Write-Output "! $_" }
        throw 'ข้อมูลยังไม่ครบ - เปิดโปรแกรมแบบมีหน้าต่างเพื่อกรอกส่วนที่ขาด'
    }
    $m = [pscustomobject]@{ InvNo = $res.InvNo; DateSerial = $rep.DateSerial; Branch = $res.Branch; Boi = $res.Boi; Consignee = $res.Consignee; Mark = $rep.Marking; Pallets = [int]$rep.Pallets; Items = $rep.Items }
    if (-not $OutFile) { $OutFile = Out-PathFor $rep $res.InvNo }
    Write-Invb $m $cfg $OutFile ([bool]$AddPrice)
    Write-Output "สร้างไฟล์: $OutFile"
    Write-Output "INV_NO $($m.InvNo) | INV_DATE $(SerialToText $m.DateSerial) | BRANCH $($m.Branch) | BOI $($m.Boi) | CONSIGNEE $($m.Consignee) | MARK1 $($m.Mark) | NUMPKG $($m.Pallets)"
    Write-Output (Get-TotalsText $rep)
    $res.Warnings | ForEach-Object { Write-Output "! $_" }
    foreach ($w in (Get-SubGroupWarnings $cfg $res.Prefix $rep.Items)) { Write-Output "! $w" }
    return
}

# ================= หน้าต่างโปรแกรม =================
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[Windows.Forms.Application]::EnableVisualStyles()
$script:Cfg = $null; $script:Rep = $null; $script:Res = $null
$font = New-Object Drawing.Font('Tahoma', 9)

$form = New-Object Windows.Forms.Form
$form.Text = 'Tong Heer : Export Summary Report -> INVB (ECS)'
$form.ClientSize = New-Object Drawing.Size(1120, 720)
$form.MinimumSize = New-Object Drawing.Size(900, 600)
$form.StartPosition = 'CenterScreen'
$form.Font = $font
$form.AllowDrop = $true

$btnOpen = New-Object Windows.Forms.Button
$btnOpen.Text = 'เลือกไฟล์รายงาน (xlsx)...'
$btnOpen.SetBounds(10, 10, 190, 30)
$lblFile = New-Object Windows.Forms.Label
$lblFile.Text = 'เลือกไฟล์ หรือลากไฟล์ Export Summary Report มาวางในหน้าต่างนี้'
$lblFile.SetBounds(210, 17, 900, 20)
$lblFile.Anchor = 'Top,Left,Right'
$form.Controls.AddRange(@($btnOpen, $lblFile))

function New-Field([string]$text, [int]$x, [int]$w) {
    $l = New-Object Windows.Forms.Label
    $l.Text = $text; $l.SetBounds($x, 50, $w, 16)
    $t = New-Object Windows.Forms.TextBox
    $t.SetBounds($x, 68, $w, 22)
    $form.Controls.AddRange(@($l, $t))
    return $t
}
$txtInv = New-Field 'INV_NO' 10 230
$txtDate = New-Field 'INV_DATE (dd/mm/yyyy)' 250 140
$txtMark = New-Field 'MARK1 (Marking)' 400 290
$txtCons = New-Field 'CONSIGNEE (รหัสใน ECS)' 700 160
$txtPkg = New-Field 'NUMPKG (พาเลทรวม)' 870 120

$lblInfo = New-Object Windows.Forms.Label
$lblInfo.SetBounds(10, 100, 1100, 18)
$lblInfo.Anchor = 'Top,Left,Right'
$lblInfo.ForeColor = [Drawing.Color]::DimGray

$grid = New-Object Windows.Forms.DataGridView
$grid.SetBounds(10, 125, 860, 330)
$grid.Anchor = 'Top,Bottom,Left,Right'
$grid.AllowUserToAddRows = $false
$grid.AllowUserToDeleteRows = $false
$grid.RowHeadersVisible = $false
$grid.SelectionMode = 'CellSelect'
$grid.MultiSelect = $false
$grid.AutoSizeColumnsMode = 'Fill'
$grid.EditMode = 'EditOnEnter'
$colDefs = @(
    @('ITEM', 'ITEM', 30, $true), @('NAME', 'PARTNAME', 210, $true), @('MODEL', 'RAWCODE (MODEL)', 100, $true),
    @('NET', 'NET', 55, $true), @('GROSS', 'GROSS', 55, $true), @('AMOUNT', 'AMOUNT', 60, $true),
    @('PARTNO', 'PARTNO (แก้ได้)', 120, $false), @('SUBGROUP', 'SUBGROUP (แก้ได้)', 120, $false)
)
foreach ($d in $colDefs) {
    $c = New-Object Windows.Forms.DataGridViewTextBoxColumn
    $c.Name = $d[0]; $c.HeaderText = $d[1]; $c.FillWeight = $d[2]; $c.ReadOnly = $d[3]
    $c.SortMode = 'NotSortable'
    if ($d[3]) { $c.DefaultCellStyle.BackColor = [Drawing.Color]::WhiteSmoke }
    if ($d[0] -in 'NET', 'GROSS', 'AMOUNT') { $c.DefaultCellStyle.Alignment = 'MiddleRight' }
    [void]$grid.Columns.Add($c)
}

$lblList = New-Object Windows.Forms.Label
$lblList.Text = 'PARTNO ที่รู้จัก (ดับเบิลคลิก = ใส่ในแถวที่เลือก)'
$lblList.SetBounds(880, 125, 230, 32)
$lblList.Anchor = 'Top,Right'
$lst = New-Object Windows.Forms.ListBox
$lst.SetBounds(880, 160, 230, 295)
$lst.Anchor = 'Top,Bottom,Right'
$lst.IntegralHeight = $false

$lblTot = New-Object Windows.Forms.Label
$lblTot.SetBounds(10, 462, 1100, 20)
$lblTot.Anchor = 'Bottom,Left,Right'
$lblTot.Font = New-Object Drawing.Font('Tahoma', 9, [Drawing.FontStyle]::Bold)

$txtWarn = New-Object Windows.Forms.TextBox
$txtWarn.Multiline = $true; $txtWarn.ReadOnly = $true; $txtWarn.ScrollBars = 'Vertical'
$txtWarn.SetBounds(10, 488, 1100, 175)
$txtWarn.Anchor = 'Bottom,Left,Right'
$txtWarn.BackColor = [Drawing.Color]::LightYellow

$chkPrice = New-Object Windows.Forms.CheckBox
$chkPrice.Text = 'เพิ่มคอลัมน์ PRICE ท้ายตาราง (AMOUNT / QUANTITY ต่อรายการ - ใช้ทดสอบปัญหาราคาต่อหน่วยใน ECS)'
$chkPrice.SetBounds(10, 678, 640, 24)
$chkPrice.Anchor = 'Bottom,Left'
$chkPrice.Checked = [bool]$AddPrice

$btnGen = New-Object Windows.Forms.Button
$btnGen.Text = 'สร้างไฟล์ INVB'
$btnGen.SetBounds(900, 672, 210, 38)
$btnGen.Anchor = 'Bottom,Right'
$btnGen.Enabled = $false
$btnGen.Font = New-Object Drawing.Font('Tahoma', 10, [Drawing.FontStyle]::Bold)
$form.Controls.AddRange(@($lblInfo, $grid, $lblList, $lst, $lblTot, $txtWarn, $chkPrice, $btnGen))

function Gui-Paint {
    foreach ($row in $grid.Rows) {
        $color = [Drawing.Color]::White
        if (-not ([string]$row.Cells['PARTNO'].Value).Trim()) { $color = [Drawing.Color]::Khaki }
        $row.Cells['PARTNO'].Style.BackColor = $color
        $row.Cells['SUBGROUP'].Style.BackColor = $color
    }
    $txtCons.BackColor = $(if ($txtCons.Text.Trim()) { [Drawing.Color]::White } else { [Drawing.Color]::Khaki })
}
function Gui-Load([string]$path) {
    try {
        $script:Cfg = Load-Config
        $script:Rep = Parse-Report $path
        $script:Res = Resolve-Report $script:Rep $script:Cfg
    } catch {
        [void][Windows.Forms.MessageBox]::Show("อ่านไฟล์ไม่ได้:`r`n$($_.Exception.Message)", 'ผิดพลาด', 'OK', 'Error')
        return
    }
    $rep = $script:Rep; $res = $script:Res
    $lblFile.Text = $path
    $txtInv.Text = $res.InvNo
    $txtDate.Text = $(if ($null -ne $rep.DateSerial) { SerialToText $rep.DateSerial } else { '' })
    $txtMark.Text = $rep.Marking
    $txtCons.Text = $res.Consignee
    $txtPkg.Text = $(if ($null -ne $rep.Pallets) { [string][int]$rep.Pallets } else { '' })
    $etd = $(if ($null -ne $rep.EtdSerial) { SerialToText $rep.EtdSerial } else { '-' })
    $lblInfo.Text = "Customer: $($rep.Customer)   |   prefix $($res.Prefix)  ->  EXP_BRANCH $($res.Branch),  BOI $($res.Boi)   |   ETD $etd"
    $grid.Rows.Clear()
    $i = 0
    foreach ($it in $rep.Items) {
        $i++
        [void]$grid.Rows.Add([object[]]@($i, $it.Name, $it.RawCode, (Fmt2 $it.Net), (Fmt2 $it.Gross), (Fmt2 $it.Amount), $it.PartNo, $it.SubGroup))
    }
    $lst.Items.Clear()
    foreach ($p in $script:Cfg.partnos) { if ($p.prefix -eq $res.Prefix) { [void]$lst.Items.Add($p.partno) } }
    $lblTot.Text = Get-TotalsText $rep
    $w = New-Object Collections.Generic.List[string]
    foreach ($x in $res.Warnings) { $w.Add($x) }
    foreach ($x in (Get-SubGroupWarnings $script:Cfg $res.Prefix $rep.Items)) { $w.Add($x) }
    if ($w.Count -eq 0) { $txtWarn.Text = 'ไม่มีจุดที่ต้องตรวจเพิ่ม' } else { $txtWarn.Text = "จุดที่ต้องตรวจ:`r`n- " + ($w -join "`r`n- ") }
    $btnGen.Enabled = $true
    Gui-Paint
}
function Gui-Generate([string]$outOverride) {
    if (-not $script:Rep) { return $null }
    $grid.EndEdit() | Out-Null
    $rep = $script:Rep; $cfg = $script:Cfg; $res = $script:Res
    $err = New-Object Collections.Generic.List[string]
    $invNo = $txtInv.Text.Trim()
    if (-not $invNo) { $err.Add('ยังไม่ได้กรอก INV_NO') }
    $serial = TextToSerial $txtDate.Text
    if ($null -eq $serial) { $err.Add('INV_DATE ต้องเป็นรูปแบบ dd/mm/yyyy (ปี ค.ศ.)') }
    $mark = $txtMark.Text.Trim()
    if (-not $mark) { $err.Add('ยังไม่ได้กรอก MARK1') }
    $cons = $txtCons.Text.Trim()
    if (-not $cons) { $err.Add('ยังไม่ได้กรอก CONSIGNEE (รหัสตาม ECS)') }
    $pkg = 0
    if (-not [int]::TryParse($txtPkg.Text.Trim(), [ref]$pkg)) { $err.Add('NUMPKG ต้องเป็นจำนวนเต็ม') }
    if (-not $res.Branch) { $err.Add("ไม่รู้จัก prefix อินวอย `"$($res.Prefix)`" (ต้องเป็น 1THA / 2THA / 3THA)") }
    for ($i = 0; $i -lt $rep.Items.Count; $i++) {
        $pn = [string]$grid.Rows[$i].Cells['PARTNO'].Value
        $sg = [string]$grid.Rows[$i].Cells['SUBGROUP'].Value
        if (-not $pn.Trim()) { $err.Add("รายการ $($i + 1): ยังไม่ได้กรอก PARTNO") }
        if (-not $sg.Trim()) { $sg = $pn }
        $rep.Items[$i].PartNo = $pn; $rep.Items[$i].SubGroup = $sg
    }
    if ($err.Count -gt 0) {
        Gui-Paint
        if (-not $outOverride) { [void][Windows.Forms.MessageBox]::Show(($err -join "`r`n"), 'ข้อมูลยังไม่ครบ', 'OK', 'Warning') }
        return $null
    }
    $m = [pscustomobject]@{ InvNo = $invNo; DateSerial = $serial; Branch = $res.Branch; Boi = $res.Boi; Consignee = $cons; Mark = $mark; Pallets = $pkg; Items = $rep.Items }
    $out = $outOverride
    if (-not $out) {
        $out = Out-PathFor $rep $invNo
        if (Test-Path -LiteralPath $out) {
            $ans = [Windows.Forms.MessageBox]::Show("มีไฟล์ $([IO.Path]::GetFileName($out)) อยู่แล้ว ต้องการเขียนทับหรือไม่?", 'ยืนยัน', 'YesNo', 'Question')
            if ($ans -ne 'Yes') { return $null }
        }
    }
    try { Write-Invb $m $cfg $out $chkPrice.Checked }
    catch {
        [void][Windows.Forms.MessageBox]::Show("เขียนไฟล์ไม่ได้ (ถ้าไฟล์เปิดอยู่ใน Excel ให้ปิดก่อน):`r`n$($_.Exception.Message)", 'ผิดพลาด', 'OK', 'Error')
        return $null
    }
    if (-not $outOverride) { Learn-Model $cfg $m $res.Prefix }
    $sub = Get-SubGroupWarnings $cfg $res.Prefix $rep.Items
    $msg = "สร้างไฟล์แล้ว:`r`n$out`r`n`r`n$(Get-TotalsText $rep)"
    if ($sub.Count -gt 0) { $msg += "`r`n`r`n" + ($sub -join "`r`n") }
    if (-not $outOverride) { [void][Windows.Forms.MessageBox]::Show($msg, 'เสร็จแล้ว', 'OK', 'Information') }
    Gui-Paint
    return $out
}

$btnOpen.Add_Click({
    $dlg = New-Object Windows.Forms.OpenFileDialog
    $dlg.Filter = 'Excel (*.xlsx)|*.xlsx'
    $dlg.Title = 'เลือกไฟล์ Export Summary Report'
    if ($dlg.ShowDialog($form) -eq 'OK') { Gui-Load $dlg.FileName }
})
$btnGen.Add_Click({ [void](Gui-Generate '') })
$form.Add_DragEnter({ param($s, $e) if ($e.Data.GetDataPresent([Windows.Forms.DataFormats]::FileDrop)) { $e.Effect = 'Copy' } })
$form.Add_DragDrop({ param($s, $e) $f = @($e.Data.GetData([Windows.Forms.DataFormats]::FileDrop)); if ($f.Count -gt 0) { Gui-Load ([string]$f[0]) } })
$grid.Add_CellEndEdit({
    param($s, $e)
    if ($grid.Columns[$e.ColumnIndex].Name -eq 'PARTNO' -and $script:Res) {
        $pn = [string]$grid.Rows[$e.RowIndex].Cells['PARTNO'].Value
        $grid.Rows[$e.RowIndex].Cells['SUBGROUP'].Value = Default-SubGroup $script:Cfg $script:Res.Prefix $pn
    }
    Gui-Paint
})
$lst.Add_DoubleClick({
    if ($null -eq $lst.SelectedItem -or $null -eq $grid.CurrentCell -or -not $script:Res) { return }
    $grid.EndEdit() | Out-Null
    $row = $grid.Rows[$grid.CurrentCell.RowIndex]
    $pn = [string]$lst.SelectedItem
    $row.Cells['PARTNO'].Value = $pn
    $row.Cells['SUBGROUP'].Value = Default-SubGroup $script:Cfg $script:Res.Prefix $pn
    Gui-Paint
})
$txtCons.Add_TextChanged({ Gui-Paint })

if ($SelfTestPng) {
    # ทดสอบหน้าต่างโดยไม่ต้องคลิก: โหลดไฟล์ สร้างผลลัพธ์ แล้วเก็บภาพหน้าจอ
    $form.Show(); [Windows.Forms.Application]::DoEvents()
    Gui-Load (Resolve-Path -LiteralPath $InputFile).Path
    [Windows.Forms.Application]::DoEvents()
    $o = Gui-Generate $OutFile
    $bmp = New-Object Drawing.Bitmap($form.Width, $form.Height)
    $form.DrawToBitmap($bmp, (New-Object Drawing.Rectangle(0, 0, $form.Width, $form.Height)))
    $bmp.Save($SelfTestPng, [Drawing.Imaging.ImageFormat]::Png)
    $form.Close()
    Write-Output "selftest out: $o"
    return
}

try {
    if ($InputFile) { $form.Add_Shown({ Gui-Load (Resolve-Path -LiteralPath $InputFile).Path }) }
    [void]$form.ShowDialog()
} catch {
    [void][Windows.Forms.MessageBox]::Show($_.Exception.Message, 'ผิดพลาด', 'OK', 'Error')
}
