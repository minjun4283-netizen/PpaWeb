# PPA 계약관리 웹앱

사내망의 매크로 포함 엑셀(xlsm)로 관리하던 PPA(전력구매계약) 데이터를, 팀원들이
iPad를 포함한 브라우저로 함께 보고 편집할 수 있도록 옮긴 웹 애플리케이션입니다.
원본 VBA 매크로(검증실행, 변경이력, Access DB 연동, 수급매칭 툴팁 등)의 기능을
그대로 웹으로 포팅했습니다.

## ⚠️ 스키마는 추측입니다 — 실사용 전 필수 확인

원본 xlsm의 VBA 코드에는 PK/FK 컬럼명과 검증 로직에 쓰인 일부 컬럼만 드러나 있고,
나머지 실제 업무 컬럼(계약기간, 단가, 용량 등)은 알 수 없었습니다. 그래서 6개 표
(`T_발전소`, `T_구매계약`, `T_수급매칭`, `T_전기사용지`, `T_판매계약`, `T_수요기업`)에
재생에너지 PPA 도메인 지식으로 합리적인 컬럼을 추측해서 채워두었습니다
(`backend/src/schema/tableDefs.ts` 참고).

**실제 운영에 쓰기 전에** 로그인 후 관리자 계정으로 **컬럼관리** 화면
(`/admin`)에서 실제 엑셀 표의 컬럼명과 비교해 빠진 컬럼을 추가해주세요. 컬럼은
언제든 추가할 수 있고, 기존 데이터는 유지됩니다.

## 원본 매크로 기능 → 웹 기능 매핑

| VBA 매크로 | 웹 기능 |
|---|---|
| `검증실행` (PK공란/PK중복/참조/조합중복) | 검증 리포트 페이지 — `POST /api/validate` |
| `변경기준_저장` + `변경이력_비교실행` (스냅샷 비교) | **실시간 변경이력** — 모든 저장/수정/삭제가 즉시 `change_log`에 기록됩니다. 사람이 미리 스냅샷을 눌러둬야 하는 원본 방식보다 안전해서, 스냅샷 개념은 없앴습니다. |
| `추가수정_파일생성` | 검증 리포트 화면의 "추가/수정 파일 내보내기" — 기준 시각 이후 변경분을 xlsx로 다운로드 |
| `엑셀시트별_Access테이블전송` | 관리자 페이지의 "Access DB 동기화" — **Windows + ACE OLEDB 드라이버 환경에서만 동작** (아래 참고) |
| T_수급매칭 하이퍼링크 툴팁(UserForm) | 표 안의 ⓘ 버튼 — 탭하면 연결된 발전소/수요기업 정보 팝업 표시 (iPad에서는 hover가 없으므로 탭 방식으로 변경) |

## 구조

```
backend/   Node.js + Express + TypeScript + SQLite(better-sqlite3)
frontend/  React + TypeScript + Vite
```

데이터는 SQLite 파일(`backend/data/ppaweb.db`) 하나에 저장됩니다. 별도 DB 서버가
필요 없어 사내망의 작은 서버 하나로 충분합니다.

## 로컬 개발 실행

```bash
# 백엔드
cd backend
npm install
npm run dev        # http://localhost:4000

# 프론트엔드 (다른 터미널)
cd frontend
npm install
npm run dev         # http://localhost:5173 (API는 /api 로 프록시됨)
```

최초 실행 시 관리자 계정이 자동 생성되고 콘솔에 비밀번호가 출력됩니다
(`admin` / `changeme123`, `ADMIN_PASSWORD` 환경변수로 변경 가능). **로그인 후
반드시 비밀번호를 바꾸세요.**

## 사내망 배포 (Docker)

서버 환경이 아직 정해지지 않았다는 전제로, 어떤 리눅스 서버에서도 바로 띄울 수
있도록 Docker로 구성했습니다.

```bash
docker compose up -d --build
```

- 접속: `http://<사내서버 IP>:4000`
- 데이터는 Docker volume(`ppaweb-data`)에 영구 저장됩니다.
- `docker-compose.yml`의 `JWT_SECRET`, `ADMIN_PASSWORD`는 배포 전에 반드시 바꿔주세요.
- Windows Server/IIS 환경이라면 Docker 대신 `backend`를 `npm run build && npm start`로,
  `frontend`를 `npm run build`로 빌드해 정적 파일을 IIS에 올리는 방식도 가능합니다
  (백엔드가 `backend/dist/public`에 프론트 빌드 결과가 있으면 그것도 함께 서빙합니다).

## Access DB 동기화 (선택 기능)

원본 매크로처럼 `PPA계약DB_전기사용지 반영.accdb`로 데이터를 계속 내보내야 한다면:

- 이 기능은 **Windows + Microsoft Access Database Engine(ACE OLEDB) 드라이버**가
  설치된 환경에서만 동작합니다 (원본 매크로와 동일한 제약).
- `ACCESS_DB_PATH` 환경변수에 accdb 파일 경로(또는 서버가 접근 가능한 네트워크
  경로)를 지정하세요.
- 사내망 서버가 Linux라면, 이 웹앱 자체가 아니라 accdb 파일을 볼 수 있는 별도
  Windows 머신(또는 예약 작업)에서 백엔드의 `syncTableToAccess` 로직을 재사용해
  주기적으로 동기화하는 구조를 권장합니다.
- 이 환경에서 사용할 수 없을 때는 관리자 페이지에 "Windows 전용 기능" 안내만
  표시되고, 나머지 기능은 정상 동작합니다.

## 인증 / 팀원 계정

간단한 아이디/비밀번호 로그인 (JWT, httpOnly 쿠키)입니다. 사내망 전용 배포를
전제로 별도 SSO 연동은 넣지 않았습니다. 관리자 계정으로 로그인 후
**컬럼관리 / Access 동기화** 페이지 상단의 **팀원 계정 관리**에서 팀원 계정을
추가/삭제할 수 있습니다.

## 알려진 제한사항

- 컬럼 스키마는 추측값입니다 (위 경고 참고).
- Access DB 동기화는 Windows 전용이며, 이 저장소의 개발/테스트 환경(Linux)에서는
  "사용 불가" 응답만 검증했습니다.
- Docker 이미지 빌드는 이 개발 환경에 Docker 데몬이 없어 실제 빌드까지는
  검증하지 못했습니다 (Dockerfile 구성 자체는 표준적인 멀티스테이지 빌드입니다).
