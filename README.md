# 한양이엔지 관제 시스템

한양이엔지 사내 PC에서 **S5D / DDWORKS** 프로그램의 실행·작업 현황을 수집하고,
구역(라인)별 실시간 접속 현황을 관제 대시보드로 보여주는 시스템입니다.

---

## 현재 운영 기준

이제 운영/수정 기준 파일은 아래 통합본 하나입니다.

```text
monitor/hanyang_monitor.py
```

`Client.py`와 `viewer_v4.py`는 `hanyang_monitor.py`로 통합되었고, 구버전 원본 보관 폴더(`legacy/`)는 삭제했습니다.

---

## 폴더 구조

```text
.
├── monitor/                 # 통합 독립실행 프로그램
│   ├── hanyang_monitor.py   # 백그라운드 모니터링 + 관제 대시보드 + 트레이 상주
│   └── requirements.txt
├── server/                  # 관제 수집 서버
│   ├── server.py
│   └── requirements.txt
├── admin/                   # 접속 제한 인원 설정 GUI
│   ├── admin.py
│   └── requirements.txt
├── .gitignore
└── README.md
```

---

## 시스템 구성

```text
[사내 PC들]                         [관제 서버]                     [관리자]
┌─────────────────────────┐        ┌──────────────┐
│ monitor/hanyang_monitor │  POST  │              │   GET 사원명부/로그
│  · 백그라운드 모니터링   │ ─────► │  server.py   │ ◄───────────────┐
│    (S5D/DDWORKS 감지)    │ /receive_status      │                  │
│  · 관제 대시보드(GUI)    │ ◄───── │  (Flask)     │                  │
│  · 트레이 상주           │  로그/명부            │                  │
└─────────────────────────┘        │  · report_*.csv          ┌──────────────┐
                                    │  · 사원명부.xlsx          │  admin.py    │
                                    │  · zone_limits.json ◄──── │ 접속제한 설정 │
                                    └──────────────┘   POST     └──────────────┘
```

---

## 통합 프로그램: `monitor/hanyang_monitor.py`

`hanyang_monitor.py`는 기존 `Client.py`와 `viewer_v4.py` 역할을 하나로 합친 독립 실행 프로그램입니다.

| 기능 | 설명 |
|------|------|
| 백그라운드 모니터링 | 이 PC의 S5D/DDWORKS 실행·작업 현황을 서버로 전송 |
| 관제 대시보드 | 전체 구역의 접속 현황을 막대그래프/명단으로 표시 |
| 시스템 트레이 상주 | 창을 닫아도 모니터링 지속, 트레이에서 다시 열기/종료 |
| 단일 인스턴스 | Windows 뮤텍스로 중복 실행 방지 |
| 부팅 시 자동 실행 | `--minimized` 인자로 트레이 모드 자동 시작 등록 |

### 실행

```bat
cd monitor
pip install -r requirements.txt
python hanyang_monitor.py
```

트레이 숨김 모드로 실행하려면 아래처럼 실행합니다.

```bat
python hanyang_monitor.py --minimized
```

### PyInstaller 빌드

```bat
cd monitor
pip install pyinstaller
pyinstaller --noconsole --onefile --name HanyangMonitor hanyang_monitor.py
```

결과물:

```text
dist/HanyangMonitor.exe
```

---

## 서버: `server/server.py`

현재 레포에 포함된 기본 수집 서버입니다.

```bat
cd server
pip install -r requirements.txt
python server.py
```

`사원명부.xlsx`는 `이름 | IP | PCNAME | 팀 | 공종 | 권한` 헤더를 기준으로 읽으며(pandas), 서버 실행 시 `server/` 폴더에서 실행해야 리포트/설정 파일이 같은 폴더에 생성됩니다.

현재 기본 서버 구조는 다음 파일을 사용합니다.

| 파일 | 설명 |
|------|------|
| `사원명부.xlsx` | 인가 IP/사원 정보 원본 |
| `report_YYYY-MM-DD.csv` | 일자별 수신 로그 |
| `execution_log.csv` | 프로그램 실행 인가/거부 이력 |
| `zone_limits.json` | 구역별 접속 제한 인원 |

---

## 주요 API

| Method | Path | 용도 |
|--------|------|------|
| POST | `/receive_status` | 모니터가 보내는 프로세스 상태 수신 |
| GET | `/api/employees` | 사원명부 수신 |
| GET | `/api/logs` | 오늘자 리포트 CSV 수신 |
| GET/POST | `/api/zone_limits` | 구역별 접속 제한 인원 조회/저장 |
| POST | `/api/verify_execution` | 범용 프로그램 실행 인가 검증 |
| GET | `/api/execution_logs`, `/api/denied_summary` | 실행 인가 로그/거부 요약 |
| GET | `/api/server_status` | 서버 PC CPU/RAM/디스크 사용률 (psutil) |
| GET | `/api/active_clients` | 최근 5분 내 활성 클라이언트/토폴로지 상태 |
| POST | `/api/bom_auth`, `/bom/start_token`, `/bom/verify_token` | BOM 매크로 실행 인가 (토큰 방식) |

---

## 사원명부 형식

기본 서버를 사용할 경우 `server/사원명부.xlsx`는 아래 순서를 따릅니다.

| 열 | 내용 | 예시 |
|----|------|------|
| A | 이름 | 홍길동 |
| B | IP 주소 | 12.26.204.51 |
| C | PC 이름 | PC-DESIGN-01 |
| D | 팀 | 설계1팀 |
| E | 공종 | S.GAS |

---

## 공통 설정

서버 주소와 인증키는 각 파일 상단 상수로 관리합니다.
다운로드 후 수정할 때는 우선 `monitor/hanyang_monitor.py`의 아래 값을 확인하면 됩니다.

```python
SERVER_URL = "http://12.26.204.100:5000"
API_KEY = "HanyangENG-Monitor-2026!"
```

`SERVER_URL`은 운영 서버 주소에 맞게 수정해야 합니다.

---

## Git 관리 기준

- 운영/수정 기준은 `monitor/hanyang_monitor.py`입니다.
- 구버전 `Client.py`, `viewer_v4.py`, `legacy/` 폴더는 삭제되었습니다.
- 런타임 생성 데이터(`report_*.csv`, `사원명부.xlsx`, 로그, 설정 파일)는 `.gitignore`로 저장소에서 제외합니다.
