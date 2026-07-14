# ============================================================
#  견적생성 : 화주요청서 + 운임표 → 무브양식 견적서(건별 .xlsx)
#
#  [항상 지킬 4단계 — 순서 고정]
#   1. 요청서에서 출발·도착·사이즈·대수 읽기
#   2. 운임표에서 같은 구간 항목별 단가 찾기
#   3. 금액 = 단가 × 수량 (CNTR=대수, BL=1, USD는 ×환율)
#   4. 무브양식 빈칸 채워 건별 견적서로 저장
# ============================================================
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$OutDir    = Join-Path $ScriptDir "견적서_출력"
$RateFile  = Join-Path $ScriptDir "운임표.xlsx"
$ReqFile   = Join-Path $ScriptDir "화주요청서.xlsx"

function OLE($r,$g,$b){ return [int]($r + ($g*256) + ($b*65536)) }

Write-Host "=================================================" -ForegroundColor Cyan
Write-Host "  포워딩 견적 자동생성 시작" -ForegroundColor Cyan
Write-Host "=================================================" -ForegroundColor Cyan

foreach ($f in @($RateFile,$ReqFile)) {
    if (-not (Test-Path $f)) {
        Write-Host "`n  오류: $(Split-Path -Leaf $f) 없음. 먼저 '1_초기설정.bat' 실행하세요." -ForegroundColor Red
        try { Read-Host "`nEnter 로 닫기" } catch {}; exit 1
    }
}

$excel = New-Object -ComObject Excel.Application
$excel.Visible=$false; $excel.DisplayAlerts=$false

$saved=@(); $skipped=@()

try {
    # ── 운임표 로드 ───────────────────────────────────────
    Write-Host "`n  운임표 읽는 중..." -ForegroundColor Gray
    $rwb = $excel.Workbooks.Open($RateFile)

    # 운임DB → 구간별 항목 리스트
    $dbSheet = $null; $infSheet = $null
    foreach ($sh in $rwb.Sheets) {
        if ($sh.Name -eq "운임DB")   { $dbSheet = $sh }
        if ($sh.Name -eq "구간정보") { $infSheet = $sh }
    }
    if (-not $dbSheet)  { throw "운임표에 '운임DB' 시트가 없습니다." }
    if (-not $infSheet) { throw "운임표에 '구간정보' 시트가 없습니다." }

    $rateMap = @{}   # key "출발|도착" -> 항목 배열
    $lastRow = $dbSheet.UsedRange.Rows.Count
    for ($r=2;$r -le $lastRow;$r++){
        $from = $dbSheet.Cells.Item($r,1).Text.Trim()
        if (-not $from) { continue }
        $to   = $dbSheet.Cells.Item($r,2).Text.Trim()
        $key  = "$from|$to"
        if (-not $rateMap.ContainsKey($key)) { $rateMap[$key] = @() }
        $rateMap[$key] += [pscustomobject]@{
            cat  = $dbSheet.Cells.Item($r,4).Text.Trim()
            name = $dbSheet.Cells.Item($r,5).Text.Trim()
            r20  = [double]$dbSheet.Cells.Item($r,6).Value2
            c20  = $dbSheet.Cells.Item($r,7).Text.Trim()
            u20  = $dbSheet.Cells.Item($r,8).Text.Trim()
            r40  = [double]$dbSheet.Cells.Item($r,9).Value2
            c40  = $dbSheet.Cells.Item($r,10).Text.Trim()
            u40  = $dbSheet.Cells.Item($r,11).Text.Trim()
            note = $dbSheet.Cells.Item($r,12).Text.Trim()
        }
    }

    $infoMap = @{}   # key -> {fx,term,title,seg}
    $lastRowI = $infSheet.UsedRange.Rows.Count
    for ($r=2;$r -le $lastRowI;$r++){
        $from = $infSheet.Cells.Item($r,1).Text.Trim()
        if (-not $from) { continue }
        $to   = $infSheet.Cells.Item($r,2).Text.Trim()
        $infoMap["$from|$to"] = [pscustomobject]@{
            fx    = [double]$infSheet.Cells.Item($r,3).Value2
            term  = $infSheet.Cells.Item($r,4).Text.Trim()
            title = $infSheet.Cells.Item($r,5).Text.Trim()
            seg   = $infSheet.Cells.Item($r,6).Text.Trim()
        }
    }
    $rwb.Close($false)
    Write-Host "  → 구간 $($rateMap.Count)개 로드" -ForegroundColor Green

    # ── 1단계: 요청서 읽기 ───────────────────────────────
    Write-Host "`n[1단계] 요청서 읽는 중..." -ForegroundColor Yellow
    $qwb = $excel.Workbooks.Open($ReqFile)
    $qs  = $qwb.Sheets.Item(1)
    $requests=@()
    $lastRowR = $qs.UsedRange.Rows.Count
    for ($r=2;$r -le $lastRowR;$r++){
        $name = $qs.Cells.Item($r,1).Text.Trim()
        $from = $qs.Cells.Item($r,2).Text.Trim()
        if (-not $name -or -not $from) { continue }
        $requests += [pscustomobject]@{
            화주 = $name
            출발 = $from
            도착 = $qs.Cells.Item($r,3).Text.Trim()
            사이즈 = $qs.Cells.Item($r,4).Text.Trim().ToUpper()
            대수 = [double]$qs.Cells.Item($r,5).Value2
        }
    }
    $qwb.Close($false)
    Write-Host "  → 요청 $($requests.Count)건" -ForegroundColor Green

    if (-not (Test-Path $OutDir)) { New-Item -ItemType Directory -Path $OutDir | Out-Null }
    $today = (Get-Date).ToString("yyyy.MM.dd")

    # ── 건별 처리 ─────────────────────────────────────────
    foreach ($req in $requests) {
        Write-Host ("`n" + ("-"*50))
        Write-Host "  처리: $($req.화주) | $($req.출발)->$($req.도착) | $($req.사이즈) x$($req.대수)"

        $key = "$($req.출발)|$($req.도착)"
        if (-not $rateMap.ContainsKey($key)) {
            Write-Host "  ※ 운임표에 [$key] 구간 없음 → 건너뜀" -ForegroundColor Red
            $skipped += "$($req.화주)($key)"; continue
        }
        $items = $rateMap[$key]
        $info  = $infoMap[$key]
        $fx    = if ($info) { $info.fx } else { 0 }
        if ($fx -le 0) { $fx = 1 }

        $qty = [double]$req.대수
        if ($qty -le 0) { $qty = 1 }

        # ── 새 워크북 / 무브양식 레이아웃 ──
        $owb = $excel.Workbooks.Add()
        while ($owb.Sheets.Count -gt 1) { $owb.Sheets.Item($owb.Sheets.Count).Delete() }
        $ws = $owb.Sheets.Item(1); $ws.Name="견적서"

        # 열 너비
        $widths = @(20,34,15,8,7,16,15,8,7,16,30)
        for ($c=1;$c -le 11;$c++){ $ws.Columns.Item($c).ColumnWidth = $widths[$c-1] }

        function T($r,$c,$v){ $ws.Cells.Item($r,$c).Value2=[string]$v }
        function N($r,$c,$v){ $ws.Cells.Item($r,$c).Value2=[double]$v }
        function Merge($r1,$c1,$r2,$c2){ $ws.Range($ws.Cells.Item($r1,$c1),$ws.Cells.Item($r2,$c2)).Merge() }

        # 제목
        Merge 2 1 2 11
        T 2 1 "운 임 견 적 서"
        $tt=$ws.Cells.Item(2,1)
        $tt.Font.Size=20; $tt.Font.Bold=$true; $tt.Font.Color=OLE 31 78 121
        $tt.HorizontalAlignment=-4108

        $titleTxt = if ($info) { $info.title } else { "$($req.출발) - $($req.도착)" }
        $segTxt   = if ($info) { $info.seg } else { "" }
        $term     = if ($info) { $info.term } else { "CFR" }

        # 헤더 정보
        T 4 1 "수    신 :"; T 4 2 $req.화주
        T 5 1 "참    조 :"
        T 6 1 "제    목 :"; T 6 2 "$titleTxt 해상 FCL 수출 운임 견적서"
        T 7 1 "날    짜 :"; T 7 2 $today
        for ($r=4;$r -le 7;$r++){ $ws.Cells.Item($r,1).Font.Bold=$true }

        # 우측 정보 박스
        T 4 8 "Q'TY";           T 4 10 "$($req.대수) x $($req.사이즈)"
        T 5 8 "PAYMENT TERMS";  T 5 10 $term
        T 6 8 "EX-RATE (USD)";  N 6 10 $fx
        $ws.Cells.Item(6,10).NumberFormat = "#,##0.00"
        for ($r=4;$r -le 6;$r++){ $ws.Cells.Item($r,8).Font.Bold=$true; $ws.Range($ws.Cells.Item($r,8),$ws.Cells.Item($r,9)).Merge() }

        # 운송구간
        T 9 1 "운송구간"; $ws.Cells.Item(9,1).Font.Bold=$true; $ws.Cells.Item(9,1).Interior.Color=OLE 217 225 242
        Merge 9 2 9 11; T 9 2 $segTxt

        # ── 표 헤더 (2줄) ──
        $hRow1 = 10; $hRow2 = 11
        Merge $hRow1 1 $hRow2 1; T $hRow1 1 "카테고리"
        Merge $hRow1 2 $hRow2 2; T $hRow1 2 "항목"
        Merge $hRow1 3 $hRow1 6; T $hRow1 3 "20FT"
        Merge $hRow1 7 $hRow1 10; T $hRow1 7 "40FT"
        Merge $hRow1 11 $hRow2 11; T $hRow1 11 "REMARK"
        T $hRow2 3 "기본요율"; T $hRow2 4 "UOM"; T $hRow2 5 "Q'TY"; T $hRow2 6 "금액(KRW)"
        T $hRow2 7 "기본요율"; T $hRow2 8 "UOM"; T $hRow2 9 "Q'TY"; T $hRow2 10 "금액(KRW)"
        # 헤더 스타일
        $hdrRange = $ws.Range($ws.Cells.Item($hRow1,1),$ws.Cells.Item($hRow2,11))
        $hdrRange.Font.Bold=$true; $hdrRange.Font.Color=OLE 255 255 255
        $hdrRange.Interior.Color=OLE 31 78 121
        $hdrRange.HorizontalAlignment=-4108; $hdrRange.VerticalAlignment=-4108

        # ── 항목 행 채우기 (2·3단계) ──
        $startRow = 12
        $r = $startRow
        $sum20=0.0; $sum40=0.0
        $catStartRow=@{}   # 카테고리별 시작행
        $catEndRow=@{}     # 카테고리별 끝행
        $catOrder=@()

        foreach ($it in $items) {
            # 20FT
            $q20 = if ($it.u20 -eq "CNTR") { $qty } else { 1 }
            $base20 = if ($it.c20 -eq "USD") { $it.r20 * $fx } else { $it.r20 }
            $amt20 = $base20 * $q20
            # 40FT
            $q40 = if ($it.u40 -eq "CNTR") { $qty } else { 1 }
            $base40 = if ($it.c40 -eq "USD") { $it.r40 * $fx } else { $it.r40 }
            $amt40 = $base40 * $q40

            $sum20 += $amt20; $sum40 += $amt40

            T $r 1 $it.cat
            T $r 2 $it.name
            # 20FT 단가(통화별 표기)
            N $r 3 $it.r20
            $ws.Cells.Item($r,3).NumberFormat = if ($it.c20 -eq "USD") { '"USD" #,##0.00' } else { '"KRW" #,##0' }
            T $r 4 $it.u20; N $r 5 $q20
            N $r 6 $amt20; $ws.Cells.Item($r,6).NumberFormat='"KRW" #,##0'
            # 40FT
            N $r 7 $it.r40
            $ws.Cells.Item($r,7).NumberFormat = if ($it.c40 -eq "USD") { '"USD" #,##0.00' } else { '"KRW" #,##0' }
            T $r 8 $it.u40; N $r 9 $q40
            N $r 10 $amt40; $ws.Cells.Item($r,10).NumberFormat='"KRW" #,##0'
            T $r 11 $it.note

            $ws.Cells.Item($r,2).HorizontalAlignment=-4131  # 좌측
            foreach ($cc in @(4,5,8,9)){ $ws.Cells.Item($r,$cc).HorizontalAlignment=-4108 }

            if (-not $catStartRow.ContainsKey($it.cat)) { $catStartRow[$it.cat]=$r; $catOrder+=$it.cat }
            $catEndRow[$it.cat]=$r
            $r++
        }

        # 카테고리 세로 병합
        foreach ($cat in $catOrder) {
            $cs=$catStartRow[$cat]; $ce=$catEndRow[$cat]
            if ($ce -gt $cs) { $ws.Range($ws.Cells.Item($cs,1),$ws.Cells.Item($ce,1)).Merge() }
            $cell=$ws.Cells.Item($cs,1)
            $cell.HorizontalAlignment=-4108; $cell.VerticalAlignment=-4108
            $cell.Font.Bold=$true; $cell.Interior.Color=OLE 226 232 240
        }

        # ── GRAND TOTAL ──
        $tRow=$r
        Merge $tRow 1 $tRow 5; T $tRow 1 "$term GRAND TOTAL"
        N $tRow 6 $sum20; $ws.Cells.Item($tRow,6).NumberFormat='"KRW" #,##0'
        T $tRow 7 ""; T $tRow 8 ""; T $tRow 9 ""
        Merge $tRow 7 $tRow 9;
        N $tRow 10 $sum40; $ws.Cells.Item($tRow,10).NumberFormat='"KRW" #,##0'
        $trng=$ws.Range($ws.Cells.Item($tRow,1),$ws.Cells.Item($tRow,11))
        $trng.Font.Bold=$true; $trng.Interior.Color=OLE 255 242 204
        $ws.Cells.Item($tRow,1).HorizontalAlignment=-4108

        # 요청 사이즈 총액 강조
        $emphCol = if ($req.사이즈 -eq "40FT") {10} else {6}
        $ws.Cells.Item($tRow,$emphCol).Font.Color=OLE 192 0 0
        $ws.Cells.Item($tRow,$emphCol).Font.Size=12

        # ── 테두리 (표 전체) ──
        $tbl=$ws.Range($ws.Cells.Item($hRow1,1),$ws.Cells.Item($tRow,11))
        for ($b=7;$b -le 12;$b++){ try{ $tbl.Borders.Item($b).LineStyle=1; $tbl.Borders.Item($b).Weight=2 }catch{} }

        # 안내문
        $noteRow=$tRow+2
        T $noteRow 1 "※ 본 견적은 발행일로부터 7일간 유효하며, VAT 및 현지 부대비용은 별도입니다."
        $ws.Cells.Item($noteRow,1).Font.Size=9; $ws.Cells.Item($noteRow,1).Font.Color=OLE 120 120 120

        # 4단계: 저장
        $safe = ($req.화주 -replace '[\\/:*?"<>|]','_')
        $outName = "견적_${safe}_$($req.출발)-$($req.도착)_$($req.사이즈).xlsx"
        $outPath = Join-Path $OutDir $outName
        if (Test-Path $outPath) { Remove-Item $outPath -Force }
        $owb.SaveAs($outPath, 51)
        $owb.Close($false)

        $t20 = "{0:N0}" -f $sum20; $t40 = "{0:N0}" -f $sum40
        Write-Host "  [3단계] 합계 20FT=KRW $t20 / 40FT=KRW $t40"
        Write-Host "  [4단계] 저장 → $outName" -ForegroundColor Green
        $saved += $outName
    }

} finally {
    $excel.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
}

Write-Host "`n=================================================" -ForegroundColor Cyan
Write-Host "  완료: $($saved.Count)건 저장 / $($skipped.Count)건 미매칭" -ForegroundColor Cyan
if ($skipped.Count -gt 0){ Write-Host "  미매칭: $($skipped -join ', ')" -ForegroundColor Red }
Write-Host "  출력 폴더: $OutDir" -ForegroundColor Cyan
Write-Host "=================================================" -ForegroundColor Cyan
try { Start-Process explorer.exe $OutDir } catch {}
try { if ([Environment]::UserInteractive -and -not [Console]::IsInputRedirected){ Read-Host "`nEnter 로 닫기" } } catch {}
