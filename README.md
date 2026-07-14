# Forwarding Quote System V2

`index.html` 하나로 실행하는 포워딩 견적서 출력기입니다.

## 실행

1. `index.html`을 Chrome 또는 Edge에서 엽니다.
2. **요율 관리**에서 담당자별 운임표와 안전운송운임표를 업로드합니다.
3. 운송 방식, Incoterms, POL/POD, 화물 정보를 선택합니다.
4. 자동으로 만들어진 부대비용을 확인·수정한 뒤 견적서를 저장하거나 PDF/Excel로 내보냅니다.

## AI로 운임표 읽기 (개인 PC용)

1. `AI_Setup.bat`을 실행하고 본인 OpenAI API 키를 한 번 입력합니다.
2. `AI_Helper_Start.bat`을 실행합니다. 실행 창은 견적 작업 중 열어 둡니다.
3. `index.html`의 **AI로 요율표 읽기**에서 Excel/CSV/JSON 파일을 선택합니다.
4. AI가 열 이름과 최대 30개 샘플 행만 보고 자료 종류와 열 연결을 제안합니다. 적용 전 결과 건수와 주의사항을 반드시 확인합니다.

키는 HTML, Git 저장소 또는 업로드한 견적서에 저장되지 않습니다. 현재 Windows 사용자만 읽을 수 있도록 암호화하여 `%LOCALAPPDATA%\ForwardingQuote`에 저장합니다. AI는 금액을 추측하지 않고, 파일에 있는 열을 표준 요율표 열에 연결하는 역할만 합니다.

## 지원 기능

- 단일 화면: FCL / LCL / AIR 및 EXW~DDP
- FCL: 컨테이너 타입·대수 / LCL: CBM·TON·RT / AIR: G/W·CBM·C/W
- 운송 조건에 따라 국내발생비용·해외발생비용·운임 항목 자동 구성
- 개인 브라우저별 요율 업로드 및 수정: CSV / JSON / Excel(.xlsx)
- 안전운송운임만 적용하는 내륙운송료 연동: 상·하차지 + 국내항 + 20/40 규격
- 사용자 직접 수정, 고객별 견적 저장/불러오기, 비고 입력
- 비용 규칙(Charge Rule) 업로드: 운송 방식·Incoterms·수출입별 표시 항목을 담당자가 변경
- PDF 인쇄 저장 및 Excel(.xlsx) 저장

## 운임표 양식

요율 관리 화면에서 양식 파일을 내려받을 수 있습니다.

`rate_master` 필수 열:

`chargeId, mode, pol, pod, container, amount, currency`

`safe_rate` 필수 열:

`place, port, size, amount`

`charge_rule` 필수 열:

`id, name, section, modes, terms, directions, unit, currency`

`modes`, `terms`, `directions`에 여러 조건을 넣을 때는 `|`로 구분합니다. 예: `FCL|LCL`, `FOB|CFR|CIF`, `export|import`.

안전운송운임에는 **안전운송운임(화주 청구 하한)** 금액만 넣습니다. 운수사업자간운임과 안전위탁운임은 넣지 않습니다. 기존 `안전운임_조회기.html`을 요율 관리 화면에서 그대로 업로드해도, 그중 안전운송운임 데이터만 자동으로 읽습니다.

## 데이터 보관 방식

업로드한 요율과 저장 견적은 해당 사용자의 브라우저에 저장됩니다. 서로 다른 담당자가 같은 HTML을 실행하면 각자 자신의 요율과 견적을 관리합니다. 여러 사람이 하나의 공용 요율을 동시에 관리하려면 추후 서버·로그인 기능을 붙여야 합니다.
