# PPA 계약관리 — 정적 HTML 대시보드 (임시 · 서버 불필요)

`backend`/`frontend`의 편집 가능한 웹앱을 올릴 사내망 서버가 아직 없는 동안,
**서버 없이 지금 바로 쓸 수 있는 조회 전용 대시보드**입니다. 사내에서 이미
검증된 패턴(정산 대시보드: `ppa_core.py`/`ppa_pipeline.py`/`dashboard_render.py`)을
그대로 따랐습니다 — 데이터를 JSON으로 HTML 안에 통째로 넣고, 탭/검색/정렬을
JS로 처리하는 단일 파일. Python 스크립트를 실행할 수 있는 PC/VDI 어디서든
바로 결과 파일을 만들 수 있고, 그 파일은 사내망 공유폴더나 메일, Teams로
전달하면 누구나 브라우저로 바로 엽니다.

## ⚠️ 이건 조회 전용입니다

이 대시보드에는 값을 고쳐서 저장하는 기능이 없습니다. **데이터 수정은 계속
엑셀에서** 하고, 팀원들에게 최신 상태를 보여주고 싶을 때 아래 스크립트를 다시
실행해서 HTML을 새로 만들면 됩니다. 여러 명이 동시에 실시간으로 같이
편집하는 것까지 필요하다면, `backend`/`frontend`로 만든 편집 가능한 웹앱을
사내망 서버에 배포해서 쓰세요 (README 최상단 참고).

## 파일 구성

| 파일 | 역할 |
|---|---|
| `ppa_schema.py` | 6개 표 스키마 정의 + 검증 엔진(PK공란/중복, FK공란/참조, 조합중복). 웹앱 백엔드의 `tableDefs.ts`/`validation.ts`와 동일한 내용 |
| `ppa_loader.py` | xlsm(openpyxl) 또는 CSV 폴더에서 데이터 읽기 |
| `ppa_dashboard_render.py` | 탭형 대시보드 HTML 렌더(표별 탭 + 검증 탭, 검색/정렬, 셀 단위 오류 하이라이트) |
| `build_dashboard.py` | 실제 파일로 `PPA현황.html` 생성하는 실행 스크립트 |
| `make_demo.py` | 화면 확인용 데모(의도적 오류 케이스 포함) 생성기 |
| `vendor/` | **동봉된 openpyxl** — pip install도, 인터넷 연결도 필요 없습니다 (아래 참고) |

## pip도, 인터넷도 필요 없습니다 (openpyxl 동봉됨)

`vendor/` 폴더에 openpyxl을 미리 넣어뒀습니다(순수 Python 코드라 컴파일된
바이너리 없이 폴더째로 복사만 하면 어디서든 동작). `ppa_loader.py`가 자동으로
`vendor/`를 먼저 찾아 쓰므로, **`pip install openpyxl`을 못 하는 환경에서도
xlsm 방식이 그대로 됩니다.** (직접 테스트: 시스템에 openpyxl이 전혀 없는
상태를 만들어 실행해도 정상 동작하는 것을 확인했습니다.) 이 폴더를
통째로(=`vendor/` 포함) 옮기기만 하면 됩니다 — 일부 파일만 옮기면 안 됩니다.

## Python 자체가 안 깔려있다면

VDI에 Python도 설치되어 있지 않다면, 인터넷이 되는 다른 PC에서 미리 준비해서
가져가면 됩니다:

1. 인터넷 되는 PC에서 <https://www.python.org/downloads/windows/> 접속
2. "Windows embeddable package (64-bit)" zip 다운로드 (예: `python-3.12.x-embed-amd64.zip`,
   공식 배포판, 설치 없이 압축만 풀면 바로 실행되는 휴대용 버전)
3. 압축을 풀어서 나온 내용물을 이 `static-dashboard` 폴더 옆에
   `python-embed/` 같은 폴더로 같이 넣기
4. VDI 안에서는 `python` 대신 그 폴더의 `python.exe`를 직접 실행:
   ```powershell
   .\python-embed\python.exe build_dashboard.py --xlsm=...\PPA파일.xlsm
   ```

이렇게 하면 Python 인터프리터 + openpyxl + 스크립트가 전부 폴더 하나에
들어있어서, VDI에 아무것도 설치되어 있지 않아도(그리고 인터넷이 전혀 안
되어도) 그대로 실행됩니다.

## 지금 화면 먼저 보기

```bash
python3 make_demo.py
```

`PPA현황_데모.html`을 브라우저로 엽니다. 표별 탭(발전소·구매계약·수요기업·
판매계약·전기사용지·수급매칭) + 검증 탭, 검색/정렬, 빨간 셀로 표시되는 검증
오류(일부러 PK중복/발전소ID참조 실패/조합중복 사례를 넣어뒀습니다)를 확인할
수 있습니다.

## 실제 데이터로 생성하기

**방법 A — xlsm 파일 직접 사용 (동봉된 openpyxl 사용, pip 불필요)**
```bash
python3 build_dashboard.py --xlsm=/path/to/PPA파일.xlsm
```

**방법 B — CSV 폴더 사용 (vendor/도 필요 없는 가장 가벼운 방식)**

엑셀에서 각 시트를 `다른 이름으로 저장 → CSV UTF-8`로 내보내서, 시트 이름과
똑같은 파일명(`T_발전소.csv`, `T_구매계약.csv`, `T_수요기업.csv`,
`T_판매계약.csv`, `T_전기사용지.csv`, `T_수급매칭.csv`)으로 한 폴더에 모은 뒤:

```bash
python3 build_dashboard.py --csv-dir=/path/to/csv폴더
```

- 표 이름과 헤더 텍스트로 자동 매칭합니다. PK중복/PK공란/*참조/조합중복 같은
  검증 열은 자동으로 건너뜁니다(실행 로그에 "인식 안 된 헤더"로 표시되는 게
  정상입니다).
- `--out=파일명.html`로 출력 파일명을 바꿀 수 있습니다 (기본값 `PPA현황.html`).
- 데이터가 바뀌면 스크립트를 다시 실행해서 파일을 새로 만들고 다시
  공유하면 됩니다.

## 검증됨

합성 데이터로 6개 표 전체 스키마, PK 중복, FK 참조 실패, 조합중복 케이스를
실제로 만들어 브라우저로 열어 확인했습니다 — 검증 탭 집계와 표별 화면의 셀
하이라이트가 정확히 일치합니다. xlsm 입력과 CSV 폴더 입력 두 경로 모두
검증 열을 올바르게 인식/제외하는 것도 확인했습니다.
