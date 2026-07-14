# ============================================================
#  초기설정 : 운임표.xlsx + 화주요청서.xlsx 생성
#  (처음 한 번만. 이후에는 운임표 숫자만 매달 갱신해서 재사용)
# ============================================================
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path

function OLE($r,$g,$b){ return [int]($r + ($g*256) + ($b*65536)) }
function SetTxt($ws,$r,$c,$v){ $ws.Cells.Item($r,$c).Value2 = [string]$v }
function SetNum($ws,$r,$c,$v){ $ws.Cells.Item($r,$c).Value2 = [double]$v }
function Hdr($cell,$rr,$gg,$bb){
    $cell.Font.Bold=$true; $cell.Font.Color=OLE 255 255 255
    $cell.Interior.Color=OLE $rr $gg $bb
    $cell.HorizontalAlignment=-4108; $cell.VerticalAlignment=-4108
}

Write-Host "초기설정: 운임표/요청서 생성 중..." -ForegroundColor Cyan

$excel = New-Object -ComObject Excel.Application
$excel.Visible=$false; $excel.DisplayAlerts=$false

try {
    # ══════════════════════════════════════════════════════
    #  운임표.xlsx  (시트 2개: 운임DB, 구간정보)
    # ══════════════════════════════════════════════════════
    $wb = $excel.Workbooks.Add()
    # 여분 시트 정리 후 2개 확보
    while ($wb.Sheets.Count -lt 2) { $wb.Sheets.Add() | Out-Null }
    while ($wb.Sheets.Count -gt 2) { $wb.Sheets.Item($wb.Sheets.Count).Delete() }

    # ---- 시트1 : 운임DB ----
    $db = $wb.Sheets.Item(1); $db.Name="운임DB"
    $dbHead = @("출발항","도착항","순번","카테고리","항목",
                "20FT_단가","20FT_통화","20FT_UOM",
                "40FT_단가","40FT_통화","40FT_UOM","비고")
    for ($i=1;$i -le $dbHead.Count;$i++){ SetTxt $db 1 $i $dbHead[$i-1]; Hdr $db.Cells.Item(1,$i) 31 78 121 }

    # 공통 부대비용 정의용 헬퍼 (항목,카테고리)
    $CAT_KR="한국발생 수출부대비용"; $CAT_SEA="해상운임"

    # 각 구간별 항목 배열 (순서 고정)
    # 필드: cat,name, r20,cur20,uom20, r40,cur40,uom40, note
    $routes = @()

    # 1) 부산 -> 첸나이 (인도향)
    $routes += [pscustomobject]@{ from="부산"; to="첸나이"; items=@(
        @{cat=$CAT_KR; name="TERMINAL HANDLING CHARGE(터미널 핸들링)"; r20=150000;c20="KRW";u20="CNTR"; r40=210000;c40="KRW";u40="CNTR"; note=""}
        @{cat=$CAT_KR; name="WHARFAGE(부두사용료)"; r20=4575;c20="KRW";u20="CNTR"; r40=9150;c40="KRW";u40="CNTR"; note=""}
        @{cat=$CAT_KR; name="DOC FEE(비엘비)"; r20=40000;c20="KRW";u20="BL"; r40=40000;c40="KRW";u40="BL"; note=""}
        @{cat=$CAT_KR; name="SEAL CHARGE(컨테이너 씰)"; r20=8000;c20="KRW";u20="CNTR"; r40=8000;c40="KRW";u40="CNTR"; note=""}
        @{cat=$CAT_KR; name="PSF(항만안전관리비)"; r20=259;c20="KRW";u20="CNTR"; r40=518;c40="KRW";u40="CNTR"; note=""}
        @{cat=$CAT_KR; name="TRUCKING CHARGE(내륙운송료)"; r20=697400;c20="KRW";u20="CNTR"; r40=780500;c40="KRW";u40="CNTR"; note="Vat 10% 별도, 평택 포승읍->부산신항"}
        @{cat=$CAT_KR; name="TRUCKING HANDLING CHARGE(운송 핸들링)"; r20=50000;c20="KRW";u20="CNTR"; r40=50000;c40="KRW";u40="CNTR"; note="Vat 10% 별도"}
        @{cat=$CAT_SEA; name="LSS(저유황유 할증료)"; r20=0;c20="USD";u20="CNTR"; r40=0;c40="USD";u40="CNTR"; note="2026년 3분기 유류할증료 운임 포함"}
        @{cat=$CAT_SEA; name="OCEAN FREIGHT(해상운임)"; r20=1500;c20="USD";u20="CNTR"; r40=1800;c40="USD";u40="CNTR"; note="선사 WHL, 부산항->첸나이 향"}
    )}

    # 2) 부산 -> 람차방
    $routes += [pscustomobject]@{ from="부산"; to="람차방"; items=@(
        @{cat=$CAT_KR; name="TERMINAL HANDLING CHARGE(터미널 핸들링)"; r20=150000;c20="KRW";u20="CNTR"; r40=210000;c40="KRW";u40="CNTR"; note=""}
        @{cat=$CAT_KR; name="WHARFAGE(부두사용료)"; r20=4575;c20="KRW";u20="CNTR"; r40=9150;c40="KRW";u40="CNTR"; note=""}
        @{cat=$CAT_KR; name="DOC FEE(비엘비)"; r20=40000;c20="KRW";u20="BL"; r40=40000;c40="KRW";u40="BL"; note=""}
        @{cat=$CAT_KR; name="SEAL CHARGE(컨테이너 씰)"; r20=8000;c20="KRW";u20="CNTR"; r40=8000;c40="KRW";u40="CNTR"; note=""}
        @{cat=$CAT_KR; name="PSF(항만안전관리비)"; r20=259;c20="KRW";u20="CNTR"; r40=518;c40="KRW";u40="CNTR"; note=""}
        @{cat=$CAT_KR; name="TRUCKING CHARGE(내륙운송료)"; r20=697400;c20="KRW";u20="CNTR"; r40=780500;c40="KRW";u40="CNTR"; note="평택 포승읍->부산항"}
        @{cat=$CAT_KR; name="TRUCKING HANDLING CHARGE(운송 핸들링)"; r20=20000;c20="KRW";u20="CNTR"; r40=20000;c40="KRW";u40="CNTR"; note=""}
        @{cat=$CAT_SEA; name="LSS(저유황유 할증료)"; r20=200;c20="USD";u20="CNTR"; r40=400;c40="USD";u40="CNTR"; note="2026년 3분기 저유황유 할증료 분리 청구"}
        @{cat=$CAT_SEA; name="OCEAN FREIGHT(해상운임)"; r20=350;c20="USD";u20="CNTR"; r40=550;c40="USD";u40="CNTR"; note="부산항->람차방, 태국향"}
    )}

    # 3) 인천 -> 람차방
    $routes += [pscustomobject]@{ from="인천"; to="람차방"; items=@(
        @{cat=$CAT_KR; name="TERMINAL HANDLING CHARGE(터미널 핸들링)"; r20=150000;c20="KRW";u20="CNTR"; r40=210000;c40="KRW";u40="CNTR"; note=""}
        @{cat=$CAT_KR; name="WHARFAGE(부두사용료)"; r20=4345;c20="KRW";u20="CNTR"; r40=8690;c40="KRW";u40="CNTR"; note=""}
        @{cat=$CAT_KR; name="DOC FEE(비엘비)"; r20=40000;c20="KRW";u20="BL"; r40=40000;c40="KRW";u40="BL"; note=""}
        @{cat=$CAT_KR; name="SEAL CHARGE(컨테이너 씰)"; r20=8000;c20="KRW";u20="CNTR"; r40=8000;c40="KRW";u40="CNTR"; note=""}
        @{cat=$CAT_KR; name="PSF(항만안전관리비)"; r20=259;c20="KRW";u20="CNTR"; r40=518;c40="KRW";u40="CNTR"; note=""}
        @{cat=$CAT_KR; name="TRUCKING CHARGE(내륙운송료)"; r20=362500;c20="KRW";u20="CNTR"; r40=407200;c40="KRW";u40="CNTR"; note="평택 포승읍->인천신항"}
        @{cat=$CAT_KR; name="TRUCKING HANDLING CHARGE(운송 핸들링)"; r20=20000;c20="KRW";u20="CNTR"; r40=20000;c40="KRW";u40="CNTR"; note=""}
        @{cat=$CAT_SEA; name="LSS(저유황유 할증료)"; r20=200;c20="USD";u20="CNTR"; r40=400;c40="USD";u40="CNTR"; note="2026년 3분기 저유황유 할증료 분리 청구"}
        @{cat=$CAT_SEA; name="OCEAN FREIGHT(해상운임)"; r20=350;c20="USD";u20="CNTR"; r40=550;c40="USD";u40="CNTR"; note="인천항->람차방, 태국향"}
    )}

    # 4) 평택 -> 람차방
    $routes += [pscustomobject]@{ from="평택"; to="람차방"; items=@(
        @{cat=$CAT_KR; name="TERMINAL HANDLING CHARGE(터미널 핸들링)"; r20=150000;c20="KRW";u20="CNTR"; r40=210000;c40="KRW";u40="CNTR"; note=""}
        @{cat=$CAT_KR; name="WHARFAGE(부두사용료)"; r20=2345;c20="KRW";u20="CNTR"; r40=4690;c40="KRW";u40="CNTR"; note=""}
        @{cat=$CAT_KR; name="DOC FEE(비엘비)"; r20=40000;c20="KRW";u20="BL"; r40=40000;c40="KRW";u40="BL"; note=""}
        @{cat=$CAT_KR; name="SEAL CHARGE(컨테이너 씰)"; r20=8000;c20="KRW";u20="CNTR"; r40=8000;c40="KRW";u40="CNTR"; note=""}
        @{cat=$CAT_KR; name="PSF(항만안전관리비)"; r20=259;c20="KRW";u20="CNTR"; r40=518;c40="KRW";u40="CNTR"; note=""}
        @{cat=$CAT_KR; name="TRUCKING CHARGE(내륙운송료)"; r20=170000;c20="KRW";u20="CNTR"; r40=194700;c40="KRW";u40="CNTR"; note="평택 포승읍->평택항"}
        @{cat=$CAT_KR; name="TRUCKING HANDLING CHARGE(운송 핸들링)"; r20=20000;c20="KRW";u20="CNTR"; r40=20000;c40="KRW";u40="CNTR"; note=""}
        @{cat=$CAT_SEA; name="LSS(저유황유 할증료)"; r20=200;c20="USD";u20="CNTR"; r40=400;c40="USD";u40="CNTR"; note="2026년 3분기 저유황유 할증료 분리 청구"}
        @{cat=$CAT_SEA; name="OCEAN FREIGHT(해상운임)"; r20=350;c20="USD";u20="CNTR"; r40=550;c40="USD";u40="CNTR"; note="평택항->람차방, 태국향"}
    )}

    # 운임DB 채우기
    $row = 2
    foreach ($rt in $routes) {
        $seq = 1
        foreach ($it in $rt.items) {
            SetTxt $db $row 1 $rt.from
            SetTxt $db $row 2 $rt.to
            SetNum $db $row 3 $seq
            SetTxt $db $row 4 $it.cat
            SetTxt $db $row 5 $it.name
            SetNum $db $row 6 $it.r20; SetTxt $db $row 7 $it.c20; SetTxt $db $row 8 $it.u20
            SetNum $db $row 9 $it.r40; SetTxt $db $row 10 $it.c40; SetTxt $db $row 11 $it.u40
            SetTxt $db $row 12 $it.note
            $row++; $seq++
        }
    }
    $db.Columns.AutoFit() | Out-Null
    $db.Rows.Item(1).Font.Bold = $true

    # ---- 시트2 : 구간정보 (방향/국가/항구/상차지/환율 + 표시 문구) ----
    $inf = $wb.Sheets.Item(2); $inf.Name="구간정보"
    $infHead = @("방향","출발국가","출발항","도착국가","도착항","상차지","적용환율","운송조건","제목(영문)","운송구간(영문)")
    for ($i=1;$i -le $infHead.Count;$i++){ SetTxt $inf 1 $i $infHead[$i-1]; Hdr $inf.Cells.Item(1,$i) 55 86 35 }

    $infData = @(
        @{dir="수출";fc="대한민국";from="부산";tc="인도";to="첸나이";load="평택 포승읍 (경기)";fx=1556.60;term="CFR";
          title="BUSAN, KOREA - KATTUPALLI/CHENNAI, INDIA";
          seg="FROM BUSAN TO KATTUPALLI/CHENNAI, INDIA (UNDER CFR PAYMENT TERM)"}
        @{dir="수출";fc="대한민국";from="부산";tc="태국";to="람차방";load="평택 포승읍 (경기)";fx=1556.60;term="CFR";
          title="BUSAN, KOREA - LAEMCHABANG, THAILAND";
          seg="FROM BUSAN TO LAEMCHABANG, THAILAND (UNDER CFR PAYMENT TERM)"}
        @{dir="수출";fc="대한민국";from="인천";tc="태국";to="람차방";load="평택 포승읍 (경기)";fx=1556.60;term="CFR";
          title="INCHEON, KOREA - LAEMCHABANG, THAILAND";
          seg="FROM INCHEON TO LAEMCHABANG, THAILAND (UNDER CFR PAYMENT TERM)"}
        @{dir="수출";fc="대한민국";from="평택";tc="태국";to="람차방";load="평택 포승읍 (경기)";fx=1556.60;term="CFR";
          title="PYEONGTAEK, KOREA - LAEMCHABANG, THAILAND";
          seg="FROM PYEONGTAEK TO LAEMCHABANG, THAILAND (UNDER CFR PAYMENT TERM)"}
    )
    $r=2
    foreach ($d in $infData) {
        SetTxt $inf $r 1 $d.dir;  SetTxt $inf $r 2 $d.fc;  SetTxt $inf $r 3 $d.from
        SetTxt $inf $r 4 $d.tc;   SetTxt $inf $r 5 $d.to;  SetTxt $inf $r 6 $d.load
        SetNum $inf $r 7 $d.fx;   SetTxt $inf $r 8 $d.term
        SetTxt $inf $r 9 $d.title; SetTxt $inf $r 10 $d.seg
        $r++
    }
    $inf.Columns.AutoFit() | Out-Null

    $p1 = Join-Path $ScriptDir "운임표.xlsx"
    if (Test-Path $p1) { Remove-Item $p1 -Force }
    $wb.SaveAs($p1, 51)
    $wb.Close($false)
    Write-Host "  생성: 운임표.xlsx  (운임DB + 구간정보)" -ForegroundColor Green

    # ══════════════════════════════════════════════════════
    #  화주요청서.xlsx
    # ══════════════════════════════════════════════════════
    $wb2 = $excel.Workbooks.Add()
    while ($wb2.Sheets.Count -gt 1) { $wb2.Sheets.Item($wb2.Sheets.Count).Delete() }
    $rq = $wb2.Sheets.Item(1); $rq.Name="요청서"

    $rqHead = @("화주명","출발항","도착항","사이즈","대수")
    for ($i=1;$i -le $rqHead.Count;$i++){ SetTxt $rq 1 $i $rqHead[$i-1]; Hdr $rq.Cells.Item(1,$i) 112 48 160 }

    # 샘플 요청 4건
    $samples = @(
        @("우리산업㈜","부산","첸나이","20FT",1),
        @("우리산업㈜","부산","람차방","40FT",2),
        @("우리산업㈜","인천","람차방","20FT",3),
        @("우리산업㈜","평택","람차방","40FT",1)
    )
    for ($s=0;$s -lt $samples.Count;$s++){
        $row = $s+2; $d=$samples[$s]
        SetTxt $rq $row 1 $d[0]; SetTxt $rq $row 2 $d[1]
        SetTxt $rq $row 3 $d[2]; SetTxt $rq $row 4 $d[3]; SetNum $rq $row 5 $d[4]
        for ($c=1;$c -le 5;$c++){ $rq.Cells.Item($row,$c).HorizontalAlignment=-4108 }
    }
    # 안내 메모
    SetTxt $rq 8 1 "※ 사이즈는 20FT 또는 40FT 로 입력. 대수는 컨테이너 수량."
    $rq.Cells.Item(8,1).Font.Italic=$true
    $rq.Cells.Item(8,1).Font.Color=OLE 120 120 120
    SetTxt $rq 9 1 "※ 출발항/도착항 이름은 운임표(구간정보)와 똑같이 입력해야 조회됩니다."
    $rq.Cells.Item(9,1).Font.Italic=$true
    $rq.Cells.Item(9,1).Font.Color=OLE 120 120 120

    $rq.Columns.AutoFit() | Out-Null
    $rq.Columns.Item(1).ColumnWidth = 16

    $p2 = Join-Path $ScriptDir "화주요청서.xlsx"
    if (Test-Path $p2) { Remove-Item $p2 -Force }
    $wb2.SaveAs($p2, 51)
    $wb2.Close($false)
    Write-Host "  생성: 화주요청서.xlsx" -ForegroundColor Green

} finally {
    $excel.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
}

Write-Host "`n완료! '2_견적생성.bat' 을 실행하면 견적서가 만들어집니다." -ForegroundColor Cyan
try { if ([Environment]::UserInteractive -and -not [Console]::IsInputRedirected){ Read-Host "Enter 키로 닫기" } } catch {}
