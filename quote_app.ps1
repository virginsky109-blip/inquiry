# ============================================================
#  견적 프로그램 (클릭식 GUI)
#  수출/수입 · 구간(국가→항구) · 컨테이너타입 · 대수 · 상차지
#  버튼 클릭 → 무브양식 견적서(.xlsx) 자동 생성
#  실행:  견적프로그램.bat  (powershell -STA)
# ============================================================
param([switch]$SelfTest)

$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$OutDir    = Join-Path $ScriptDir "견적서_출력"
$RateFile  = Join-Path $ScriptDir "운임표.xlsx"

# ── 데이터 로드 (운임표) ─────────────────────────────────
function Load-Data {
    if (-not (Test-Path $RateFile)) { throw "운임표.xlsx 가 없습니다. 먼저 '1_초기설정.bat' 을 실행하세요." }
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible=$false; $excel.DisplayAlerts=$false
    try {
        $wb = $excel.Workbooks.Open($RateFile)
        $dbSheet=$null; $infSheet=$null
        foreach ($sh in $wb.Sheets) {
            if ($sh.Name -eq "운임DB")   { $dbSheet=$sh }
            if ($sh.Name -eq "구간정보") { $infSheet=$sh }
        }
        if (-not $dbSheet -or -not $infSheet) { throw "운임표 시트(운임DB/구간정보) 구성이 올바르지 않습니다." }

        # 운임DB
        $rateMap=@{}
        $lr=$dbSheet.UsedRange.Rows.Count
        for ($r=2;$r -le $lr;$r++){
            $from=$dbSheet.Cells.Item($r,1).Text.Trim()
            if (-not $from) { continue }
            $to=$dbSheet.Cells.Item($r,2).Text.Trim()
            $key="$from|$to"
            if (-not $rateMap.ContainsKey($key)) { $rateMap[$key]=@() }
            $rateMap[$key]+=[pscustomobject]@{
                cat =$dbSheet.Cells.Item($r,4).Text.Trim()
                name=$dbSheet.Cells.Item($r,5).Text.Trim()
                r20 =[double]$dbSheet.Cells.Item($r,6).Value2
                c20 =$dbSheet.Cells.Item($r,7).Text.Trim()
                u20 =$dbSheet.Cells.Item($r,8).Text.Trim()
                r40 =[double]$dbSheet.Cells.Item($r,9).Value2
                c40 =$dbSheet.Cells.Item($r,10).Text.Trim()
                u40 =$dbSheet.Cells.Item($r,11).Text.Trim()
                note=$dbSheet.Cells.Item($r,12).Text.Trim()
            }
        }

        # 구간정보
        $routes=@()
        $lri=$infSheet.UsedRange.Rows.Count
        for ($r=2;$r -le $lri;$r++){
            $dir=$infSheet.Cells.Item($r,1).Text.Trim()
            if (-not $dir) { continue }
            $routes+=[pscustomobject]@{
                dir  =$dir
                fc   =$infSheet.Cells.Item($r,2).Text.Trim()
                from =$infSheet.Cells.Item($r,3).Text.Trim()
                tc   =$infSheet.Cells.Item($r,4).Text.Trim()
                to   =$infSheet.Cells.Item($r,5).Text.Trim()
                load =$infSheet.Cells.Item($r,6).Text.Trim()
                fx   =[double]$infSheet.Cells.Item($r,7).Value2
                term =$infSheet.Cells.Item($r,8).Text.Trim()
                title=$infSheet.Cells.Item($r,9).Text.Trim()
                seg  =$infSheet.Cells.Item($r,10).Text.Trim()
            }
        }
        $wb.Close($false)
        return [pscustomobject]@{ rateMap=$rateMap; routes=$routes }
    } finally {
        $excel.Quit()
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    }
}

# ── 견적서 생성 (단일 컨테이너 타입) ─────────────────────
#  타입 → 요율열 매핑: 20FT→20열, 그 외(40FT/40HC/40DG)→40열, 20DG→20열
function Get-SizeCols($type) {
    switch -Wildcard ($type.ToUpper()) {
        "20*" { return @{ pick="20" } }
        default { return @{ pick="40" } }   # 40FT/40HC/40DG
    }
}

function New-Quote {
    param($Data,$화주,$방향,$출발항,$도착항,$타입,$대수,$상차지)

    $key="$출발항|$도착항"
    if (-not $Data.rateMap.ContainsKey($key)) { throw "운임표에 [$출발항 → $도착항] 구간이 없습니다." }
    $items=$Data.rateMap[$key]
    $route=$Data.routes | Where-Object { $_.from -eq $출발항 -and $_.to -eq $도착항 -and $_.dir -eq $방향 } | Select-Object -First 1
    if (-not $route) { $route=$Data.routes | Where-Object { $_.from -eq $출발항 -and $_.to -eq $도착항 } | Select-Object -First 1 }
    $fx=if ($route){ $route.fx } else { 1 }; if ($fx -le 0){ $fx=1 }
    $term=if ($route){ $route.term } else { "CFR" }
    $title=if ($route){ $route.title } else { "$출발항 - $도착항" }
    $seg=if ($route){ $route.seg } else { "" }
    if (-not $상차지 -and $route){ $상차지=$route.load }

    $qty=[double]$대수; if ($qty -le 0){ $qty=1 }
    $pick=(Get-SizeCols $타입).pick

    if (-not (Test-Path $OutDir)) { New-Item -ItemType Directory -Path $OutDir | Out-Null }
    $today=(Get-Date).ToString("yyyy.MM.dd")

    $excel=New-Object -ComObject Excel.Application
    $excel.Visible=$false; $excel.DisplayAlerts=$false
    try {
        $owb=$excel.Workbooks.Add()
        while ($owb.Sheets.Count -gt 1){ $owb.Sheets.Item($owb.Sheets.Count).Delete() }
        $ws=$owb.Sheets.Item(1); $ws.Name="견적서"

        function OLE($r,$g,$b){ return [int]($r+($g*256)+($b*65536)) }
        function T($r,$c,$v){ $ws.Cells.Item($r,$c).Value2=[string]$v }
        function N($r,$c,$v){ $ws.Cells.Item($r,$c).Value2=[double]$v }
        function MG($r1,$c1,$r2,$c2){ $ws.Range($ws.Cells.Item($r1,$c1),$ws.Cells.Item($r2,$c2)).Merge() }

        $widths=@(20,40,16,8,7,18,34)
        for ($c=1;$c -le 7;$c++){ $ws.Columns.Item($c).ColumnWidth=$widths[$c-1] }

        # 제목
        MG 2 1 2 7; T 2 1 "운 임 견 적 서"
        $tt=$ws.Cells.Item(2,1); $tt.Font.Size=20; $tt.Font.Bold=$true; $tt.Font.Color=OLE 31 78 121; $tt.HorizontalAlignment=-4108

        T 4 1 "수    신 :"; T 4 2 $화주
        T 5 1 "참    조 :"
        T 6 1 "제    목 :"; T 6 2 "$title 해상 FCL $방향 운임 견적서"
        T 7 1 "날    짜 :"; T 7 2 $today
        for ($r=4;$r -le 7;$r++){ $ws.Cells.Item($r,1).Font.Bold=$true }

        # 우측 정보
        T 4 5 "Q'TY";          MG 4 5 4 5; T 4 6 "$대수 x $타입"
        T 5 5 "PAYMENT TERMS"; T 5 6 $term
        T 6 5 "EX-RATE(USD)";  N 6 6 $fx; $ws.Cells.Item(6,6).NumberFormat="#,##0.00"
        $loadLbl=if($방향 -eq "수입"){"하차지"}else{"상차지"}
        T 7 5 $loadLbl; T 7 6 $상차지
        for ($r=4;$r -le 7;$r++){ $ws.Cells.Item($r,5).Font.Bold=$true }

        # 운송구간
        T 9 1 "운송구간"; $ws.Cells.Item(9,1).Font.Bold=$true; $ws.Cells.Item(9,1).Interior.Color=OLE 217 225 242
        MG 9 2 9 7; T 9 2 $seg

        # 표 헤더
        $hr=10
        T $hr 1 "카테고리"; T $hr 2 "항목"; T $hr 3 "기본요율($타입)"; T $hr 4 "UOM"; T $hr 5 "Q'TY"; T $hr 6 "금액(KRW)"; T $hr 7 "REMARK"
        $hrng=$ws.Range($ws.Cells.Item($hr,1),$ws.Cells.Item($hr,7))
        $hrng.Font.Bold=$true; $hrng.Font.Color=OLE 255 255 255; $hrng.Interior.Color=OLE 31 78 121
        $hrng.HorizontalAlignment=-4108

        # 항목
        $r=$hr+1; $startRow=$r; $sum=0.0
        $catStart=@{}; $catEnd=@{}; $catOrder=@()
        foreach ($it in $items) {
            if ($pick -eq "20") { $rate=$it.r20; $cur=$it.c20; $uom=$it.u20 }
            else                { $rate=$it.r40; $cur=$it.c40; $uom=$it.u40 }
            $q=if ($uom -eq "CNTR"){ $qty } else { 1 }
            $base=if ($cur -eq "USD"){ $rate*$fx } else { $rate }
            $amt=$base*$q; $sum+=$amt

            T $r 1 $it.cat
            T $r 2 $it.name; $ws.Cells.Item($r,2).HorizontalAlignment=-4131
            N $r 3 $rate; $ws.Cells.Item($r,3).NumberFormat=if($cur -eq "USD"){'"USD" #,##0.00'}else{'"KRW" #,##0'}
            T $r 4 $uom; $ws.Cells.Item($r,4).HorizontalAlignment=-4108
            N $r 5 $q;   $ws.Cells.Item($r,5).HorizontalAlignment=-4108
            N $r 6 $amt; $ws.Cells.Item($r,6).NumberFormat='"KRW" #,##0'
            T $r 7 $it.note

            if (-not $catStart.ContainsKey($it.cat)){ $catStart[$it.cat]=$r; $catOrder+=$it.cat }
            $catEnd[$it.cat]=$r
            $r++
        }
        foreach ($cat in $catOrder){
            $cs=$catStart[$cat]; $ce=$catEnd[$cat]
            if ($ce -gt $cs){ $ws.Range($ws.Cells.Item($cs,1),$ws.Cells.Item($ce,1)).Merge() }
            $cell=$ws.Cells.Item($cs,1); $cell.HorizontalAlignment=-4108; $cell.VerticalAlignment=-4108
            $cell.Font.Bold=$true; $cell.Interior.Color=OLE 226 232 240
        }

        # GRAND TOTAL
        $tr=$r
        MG $tr 1 $tr 5; T $tr 1 "$term GRAND TOTAL ($타입 x $대수)"
        N $tr 6 $sum; $ws.Cells.Item($tr,6).NumberFormat='"KRW" #,##0'
        $trng=$ws.Range($ws.Cells.Item($tr,1),$ws.Cells.Item($tr,7))
        $trng.Font.Bold=$true; $trng.Interior.Color=OLE 255 242 204
        $ws.Cells.Item($tr,1).HorizontalAlignment=-4108
        $ws.Cells.Item($tr,6).Font.Color=OLE 192 0 0; $ws.Cells.Item($tr,6).Font.Size=12

        # 테두리
        $tbl=$ws.Range($ws.Cells.Item($hr,1),$ws.Cells.Item($tr,7))
        for ($b=7;$b -le 12;$b++){ try{ $tbl.Borders.Item($b).LineStyle=1; $tbl.Borders.Item($b).Weight=2 }catch{} }

        # 안내
        T ($tr+2) 1 "※ 본 견적은 발행일로부터 7일간 유효하며, VAT 및 현지 부대비용은 별도입니다."
        $ws.Cells.Item($tr+2,1).Font.Size=9; $ws.Cells.Item($tr+2,1).Font.Color=OLE 120 120 120

        # DG(위험물) 안내 — 할증료 미확정
        if ($타입 -match "DG") {
            T ($tr+3) 1 "※ DG(위험물) 화물입니다. DG 할증료는 요율 미확정으로 본 견적에 미포함되었으며, 확정 시 별도 추가됩니다."
            $ws.Cells.Item($tr+3,1).Font.Size=9; $ws.Cells.Item($tr+3,1).Font.Bold=$true; $ws.Cells.Item($tr+3,1).Font.Color=OLE 192 0 0
        }

        $safe=($화주 -replace '[\\/:*?"<>|]','_')
        $outName="견적_${safe}_${출발항}-${도착항}_${타입}.xlsx"
        $outPath=Join-Path $OutDir $outName
        if (Test-Path $outPath){ Remove-Item $outPath -Force }
        $owb.SaveAs($outPath,51)
        $owb.Close($false)
        return [pscustomobject]@{ path=$outPath; total=$sum }
    } finally {
        $excel.Quit()
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    }
}

# ══════════════════════════════════════════════════════════
#  SelfTest 모드 (GUI 없이 엔진만 검증)
# ══════════════════════════════════════════════════════════
if ($SelfTest) {
    Write-Host "[SelfTest] 데이터 로드..."
    $data=Load-Data
    Write-Host "  구간 $($data.rateMap.Count)개 / 구간정보 $($data.routes.Count)행"
    $res=New-Quote -Data $data -화주 "테스트상사㈜" -방향 "수출" -출발항 "부산" -도착항 "첸나이" -타입 "40HC" -대수 2 -상차지 "평택 포승읍 (경기)"
    Write-Host "  [40HC x2] $($res.path)  → KRW $("{0:N0}" -f $res.total)"
    $res2=New-Quote -Data $data -화주 "테스트상사㈜" -방향 "수출" -출발항 "부산" -도착항 "람차방" -타입 "40DG" -대수 1 -상차지 "평택 포승읍 (경기)"
    Write-Host "  [40DG x1] $($res2.path)  → KRW $("{0:N0}" -f $res2.total)"
    exit 0
}

# ══════════════════════════════════════════════════════════
#  GUI
# ══════════════════════════════════════════════════════════
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

$data = $null
try { $data = Load-Data } catch {
    [System.Windows.Forms.MessageBox]::Show($_.Exception.Message,"오류",0,16) | Out-Null
    exit 1
}

$form=New-Object System.Windows.Forms.Form
$form.Text="포워딩 견적 프로그램"
$form.Size=New-Object System.Drawing.Size(520,560)
$form.StartPosition="CenterScreen"
$form.Font=New-Object System.Drawing.Font("맑은 고딕",10)
$form.BackColor=[System.Drawing.Color]::White

function AddLabel($text,$x,$y){
    $l=New-Object System.Windows.Forms.Label
    $l.Text=$text; $l.Location=New-Object System.Drawing.Point($x,$y)
    $l.AutoSize=$true; $l.Font=New-Object System.Drawing.Font("맑은 고딕",10,[System.Drawing.FontStyle]::Bold)
    $form.Controls.Add($l); return $l
}
function AddCombo($x,$y,$w){
    $c=New-Object System.Windows.Forms.ComboBox
    $c.Location=New-Object System.Drawing.Point($x,$y)
    $c.Size=New-Object System.Drawing.Size($w,26)
    $c.DropDownStyle="DropDownList"
    $form.Controls.Add($c); return $c
}

# 제목 배너
$banner=New-Object System.Windows.Forms.Label
$banner.Text="  운임 견적 생성"
$banner.Location=New-Object System.Drawing.Point(0,0)
$banner.Size=New-Object System.Drawing.Size(520,44)
$banner.Font=New-Object System.Drawing.Font("맑은 고딕",14,[System.Drawing.FontStyle]::Bold)
$banner.ForeColor=[System.Drawing.Color]::White
$banner.BackColor=[System.Drawing.Color]::FromArgb(31,78,121)
$banner.TextAlign="MiddleLeft"
$form.Controls.Add($banner)

$colL=30; $colF=170; $fw=290; $y=64; $dy=44

AddLabel "수출/수입" $colL $y | Out-Null
$cbDir=AddCombo $colF $y $fw
AddLabel "출발 국가" $colL ($y+$dy) | Out-Null
$cbFCountry=AddCombo $colF ($y+$dy) $fw
AddLabel "출발항" $colL ($y+$dy*2) | Out-Null
$cbFrom=AddCombo $colF ($y+$dy*2) $fw
AddLabel "도착 국가" $colL ($y+$dy*3) | Out-Null
$cbTCountry=AddCombo $colF ($y+$dy*3) $fw
AddLabel "도착항" $colL ($y+$dy*4) | Out-Null
$cbTo=AddCombo $colF ($y+$dy*4) $fw
AddLabel "컨테이너 타입" $colL ($y+$dy*5) | Out-Null
$cbType=AddCombo $colF ($y+$dy*5) $fw
AddLabel "대수" $colL ($y+$dy*6) | Out-Null
$numQty=New-Object System.Windows.Forms.NumericUpDown
$numQty.Location=New-Object System.Drawing.Point($colF,($y+$dy*6))
$numQty.Size=New-Object System.Drawing.Size(80,26)
$numQty.Minimum=1; $numQty.Maximum=999; $numQty.Value=1
$form.Controls.Add($numQty)

$lblLoad=AddLabel "상차지" $colL ($y+$dy*7)
$cbLoad=AddCombo $colF ($y+$dy*7) $fw

AddLabel "화주명" $colL ($y+$dy*8) | Out-Null
$txtName=New-Object System.Windows.Forms.TextBox
$txtName.Location=New-Object System.Drawing.Point($colF,($y+$dy*8))
$txtName.Size=New-Object System.Drawing.Size($fw,26)
$txtName.Text="우리산업㈜"
$form.Controls.Add($txtName)

# 컨테이너 타입 목록
@("20FT","40FT","40HC","20DG","40DG") | ForEach-Object { [void]$cbType.Items.Add($_) }
$cbType.SelectedIndex=0

# ── 콤보 채우기 로직 ──
function Fill-Directions {
    $cbDir.Items.Clear()
    $dirs=$data.routes | Select-Object -ExpandProperty dir -Unique
    if (-not ($dirs -contains "수출")){ }
    foreach ($d in @("수출","수입")){ [void]$cbDir.Items.Add($d) }
    $cbDir.SelectedItem="수출"
}
function Fill-FCountry {
    $cbFCountry.Items.Clear()
    $dir=$cbDir.SelectedItem
    $list=$data.routes | Where-Object { $_.dir -eq $dir } | Select-Object -ExpandProperty fc -Unique
    foreach ($x in $list){ if($x){ [void]$cbFCountry.Items.Add($x) } }
    if ($cbFCountry.Items.Count -gt 0){ $cbFCountry.SelectedIndex=0 }
}
function Fill-From {
    $cbFrom.Items.Clear()
    $dir=$cbDir.SelectedItem; $fc=$cbFCountry.SelectedItem
    $list=$data.routes | Where-Object { $_.dir -eq $dir -and $_.fc -eq $fc } | Select-Object -ExpandProperty from -Unique
    foreach ($x in $list){ if($x){ [void]$cbFrom.Items.Add($x) } }
    if ($cbFrom.Items.Count -gt 0){ $cbFrom.SelectedIndex=0 }
}
function Fill-TCountry {
    $cbTCountry.Items.Clear()
    $dir=$cbDir.SelectedItem; $fc=$cbFCountry.SelectedItem; $from=$cbFrom.SelectedItem
    $list=$data.routes | Where-Object { $_.dir -eq $dir -and $_.fc -eq $fc -and $_.from -eq $from } | Select-Object -ExpandProperty tc -Unique
    foreach ($x in $list){ if($x){ [void]$cbTCountry.Items.Add($x) } }
    if ($cbTCountry.Items.Count -gt 0){ $cbTCountry.SelectedIndex=0 }
}
function Fill-To {
    $cbTo.Items.Clear()
    $dir=$cbDir.SelectedItem; $fc=$cbFCountry.SelectedItem; $from=$cbFrom.SelectedItem; $tc=$cbTCountry.SelectedItem
    $list=$data.routes | Where-Object { $_.dir -eq $dir -and $_.fc -eq $fc -and $_.from -eq $from -and $_.tc -eq $tc } | Select-Object -ExpandProperty to -Unique
    foreach ($x in $list){ if($x){ [void]$cbTo.Items.Add($x) } }
    if ($cbTo.Items.Count -gt 0){ $cbTo.SelectedIndex=0 }
}
function Fill-Load {
    $cbLoad.Items.Clear()
    $dir=$cbDir.SelectedItem; $from=$cbFrom.SelectedItem; $to=$cbTo.SelectedItem
    $lbl=if ($dir -eq "수입"){ "하차지" } else { "상차지" }
    $lblLoad.Text=$lbl
    $list=$data.routes | Where-Object { $_.dir -eq $dir -and $_.from -eq $from -and $_.to -eq $to } | Select-Object -ExpandProperty load -Unique
    foreach ($x in $list){ if($x){ [void]$cbLoad.Items.Add($x) } }
    if ($cbLoad.Items.Count -eq 0){ [void]$cbLoad.Items.Add("(직접 입력/미지정)") }
    $cbLoad.SelectedIndex=0
}

$cbDir.Add_SelectedIndexChanged({ Fill-FCountry; Fill-From; Fill-TCountry; Fill-To; Fill-Load })
$cbFCountry.Add_SelectedIndexChanged({ Fill-From; Fill-TCountry; Fill-To; Fill-Load })
$cbFrom.Add_SelectedIndexChanged({ Fill-TCountry; Fill-To; Fill-Load })
$cbTCountry.Add_SelectedIndexChanged({ Fill-To; Fill-Load })
$cbTo.Add_SelectedIndexChanged({ Fill-Load })

# 견적 생성 버튼
$btn=New-Object System.Windows.Forms.Button
$btn.Text="견적서 생성"
$btn.Location=New-Object System.Drawing.Point($colF,($y+$dy*9+6))
$btn.Size=New-Object System.Drawing.Size(160,40)
$btn.BackColor=[System.Drawing.Color]::FromArgb(31,78,121)
$btn.ForeColor=[System.Drawing.Color]::White
$btn.Font=New-Object System.Drawing.Font("맑은 고딕",11,[System.Drawing.FontStyle]::Bold)
$btn.FlatStyle="Flat"
$form.Controls.Add($btn)

$btn.Add_Click({
    try {
        $방향=$cbDir.SelectedItem; $출발항=$cbFrom.SelectedItem; $도착항=$cbTo.SelectedItem
        $타입=$cbType.SelectedItem; $대수=[int]$numQty.Value; $상차지=$cbLoad.SelectedItem
        $화주=$txtName.Text.Trim()
        if (-not $출발항 -or -not $도착항){ throw "출발항/도착항을 선택하세요." }
        if (-not $화주){ throw "화주명을 입력하세요." }

        $key="$출발항|$도착항"
        if (-not $data.rateMap.ContainsKey($key)){
            throw "운임표에 [$출발항 → $도착항] 구간 운임이 없습니다.`n(운임표.xlsx 에 추가 후 1_초기설정 없이도 바로 반영됩니다)"
        }
        # DG 타입인데 요율이 40/20 그대로면 안내
        $res=New-Quote -Data $data -화주 $화주 -방향 $방향 -출발항 $출발항 -도착항 $도착항 -타입 $타입 -대수 $대수 -상차지 $상차지
        $msg="견적서가 생성되었습니다.`n`n$([System.IO.Path]::GetFileName($res.path))`n합계: KRW $("{0:N0}" -f $res.total)`n`n파일을 여시겠습니까?"
        $ans=[System.Windows.Forms.MessageBox]::Show($msg,"완료",4,64)
        if ($ans -eq "Yes"){ Start-Process $res.path }
    } catch {
        [System.Windows.Forms.MessageBox]::Show($_.Exception.Message,"오류",0,16) | Out-Null
    }
})

# 안내 라벨
$note=New-Object System.Windows.Forms.Label
$note.Text="※ 수입 구간·특수타입(DG) 요율은 운임표.xlsx 에 추가하면 자동 반영됩니다."
$note.Location=New-Object System.Drawing.Point($colL,($y+$dy*10+10))
$note.Size=New-Object System.Drawing.Size(460,40)
$note.ForeColor=[System.Drawing.Color]::Gray
$note.Font=New-Object System.Drawing.Font("맑은 고딕",8)
$form.Controls.Add($note)

# 초기 채우기
Fill-Directions; Fill-FCountry; Fill-From; Fill-TCountry; Fill-To; Fill-Load

[void]$form.ShowDialog()
