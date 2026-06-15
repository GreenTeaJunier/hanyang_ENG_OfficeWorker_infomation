# 한양이엔지 관제 시스템 (hanyang_ENG_OfficeWorker_infomation)

한양이엔지 사내 PC에서 **S5D / DDWORKS** 프로그램의 실행·작업 현황을 수집하고,
구역(라인)별 실시간 접속 현황을 관제 대시보드로 보여주는 시스템입니다.

---

## 📁 폴더 구조

```
.
├── monitor/                 # ★ 통합 독립실행 프로그램 (구 Client.py + viewer_v4.py)
│   ├── hanyang_monitor.py   #   - 백그라운드 모니터링 + 관제 대시보드 + 트레이 상주
│   └── requirements.txt
├── server/                  # 관제 수집 서버 (Flask)
│   ├── server.py
│   └── requirements.txt
├── admin/                   # 최고관리자용 접속 제한 인원 설정 GUI
│   ├── admin.py
│   └── requirements.txt
├── legacy/                  # 통합 전 원본 보존 (삭제하지 않음)
│   ├── Client.py            #   - 구 모니터링 에이전트 원본
│   └── viewer_v4.py         #   - 구 관제 대시보드 원본
├── .gitignore
└── README.md
```

> 기존의 `Client.py` 와 `viewer_v4.py` 는 **`monitor/hanyang_monitor.py` 하나로 통합**되었습니다.
> 단, 원본은 삭제하지 않고 **`legacy/` 폴더에 그대로 보존**해 두었습니다. (참고/롤백용)

---

## 🧩 시스템 구성도

```
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

- **monitor**: 각 PC에서 자기 자신의 S5D/DDWORKS 실행 현황을 서버로 전송하고(구 Client),
  동시에 전체 구역의 접속 현황을 차트로 보여줍니다(구 Viewer).
- **server**: 상태를 받아 일자별 CSV로 기록하고, 사원명부/접속제한 정보를 제공합니다.
- **admin**: 구역별 접속 제한 인원을 서버에 저장합니다. (초과 시 대시보드에서 빨간색 경고)

---

## ⭐ 통합 프로그램: `monitor/hanyang_monitor.py`

`Client.py`(백그라운드 모니터링 에이전트)와 `viewer_v4.py`(관제 대시보드)를
**하나의 독립 실행 프로그램**으로 합쳤습니다.

실행하면 한 프로세스 안에서 다음이 동시에 동작합니다.

| 기능 | 설명 | 출처 |
|------|------|------|
| 백그라운드 모니터링 | 이 PC의 S5D/DDWORKS 실행·작업 현황을 10초마다 서버로 전송 | 구 `Client.py` |
| 관제 대시보드(GUI) | 전체 구역의 접속 현황을 막대그래프/명단으로 표시 | 구 `viewer_v4.py` |
| 시스템 트레이 상주 | 창을 닫아도 모니터링은 계속, 트레이에서 다시 열기/종료 | 구 `Client.py` |
| 단일 인스턴스 | 중복 실행 방지(Windows 뮤텍스) | 구 `Client.py` |
| 부팅 시 자동 실행 | Windows 시작 시 **트레이 모드**로 자동 실행 등록 | 구 `Client.py` |

### 동작 방식
- **창의 X 버튼** → 프로그램이 종료되지 않고 **트레이로 숨겨집니다**(모니터링 계속).
- **트레이 아이콘 더블클릭** → 관제 화면 다시 열기.
- **트레이 메뉴 → ❌ 완전 종료** → 모니터링까지 완전히 종료.
- 부팅 자동 실행 시에는 `--minimized` 인자로 시작되어 **창 없이 트레이에서 상주**합니다.
- 상태바의 `🖥 내 PC:` 표시로 내 PC의 모니터링/전송 상태를 실시간 확인할 수 있습니다.

### 실행
```bash
cd monitor
pip install -r requirements.txt
python hanyang_monitor.py            # 창을 띄우고 실행
python hanyang_monitor.py --minimized  # 트레이에 숨긴 채 백그라운드로 실행
```

### .exe 빌드 (PyInstaller)
```bash
cd monitor
pip install pyinstaller
pyinstaller --noconsole --onefile --name HanyangMonitor hanyang_monitor.py
# 결과물: dist/HanyangMonitor.exe
```
> `--noconsole` 로 콘솔 창 없이 트레이/GUI만 표시됩니다.
> 빌드된 `.exe` 와 같은 폴더에 `viewer_settings.json`, `monitor_debug.log` 가 생성됩니다.

---

## 🖥 서버: `server/server.py`

상태 수집 및 데이터 제공을 담당하는 Flask 서버입니다. (포트 `5000`)

```bash
cd server
pip install -r requirements.txt
python server.py
```

- 서버 폴더에 **`사원명부.xlsx`** 가 있어야 인가된 IP를 인식합니다. (아래 형식 참고)
- 실행 위치와 무관하게 `server.py` 가 있는 폴더를 기준으로 데이터 파일을 읽고 씁니다.

### 생성/사용 파일 (서버 폴더 기준)
| 파일 | 설명 |
|------|------|
| `사원명부.xlsx` | 인가 IP/사원 정보 원본 (직접 준비) |
| `report_YYYY-MM-DD.csv` | 일자별 수신 로그 |
| `execution_log.csv` | 프로그램 실행 인가/거부 이력 |
| `zone_limits.json` | 구역별 접속 제한 인원 (admin이 저장) |

### 주요 API
| 메서드 | 경로 | 용도 |
|--------|------|------|
| POST | `/receive_status` | 모니터가 보내는 프로세스 상태 수신 |
| GET  | `/api/employees` | 사원명부.xlsx 내려주기 (대시보드용) |
| GET  | `/api/logs` | 오늘자 리포트 CSV 내려주기 |
| GET/POST | `/api/zone_limits` | 구역별 접속 제한 인원 조회/저장 |
| POST | `/api/verify_execution` | 범용 프로그램 실행 인가 검증 |
| GET  | `/api/execution_logs`, `/api/denied_summary` | 실행 인가 로그/거부 요약 |

---

## 🔒 관리자: `admin/admin.py`

구역별 **접속 제한 인원**을 설정해 서버(`zone_limits.json`)에 저장합니다.
저장하면 모든 대시보드에 자동 반영되어, 제한 초과 구역은 빨간 막대로 강조됩니다.

```bash
cd admin
pip install -r requirements.txt
python admin.py
```

---

## 📋 `사원명부.xlsx` 형식

첫 행은 헤더(`이름`, `IP` …)이며, 다음 열 순서를 따릅니다.

| 열 | 내용 | 예시 |
|----|------|------|
| A | 이름 | 홍길동 |
| B | IP 주소 | 12.26.204.51 |
| C | PC 이름 | PC-DESIGN-01 |
| D | 팀 | 설계1팀 |
| E | 공종 | S.GAS |

---

## ⚙️ 공통 설정

서버 주소와 인증키는 각 파일 상단 상수로 관리합니다. 배포 환경에 맞게 수정하세요.

```python
SERVER_URL = "http://12.26.204.100:5000"
API_KEY    = "HanyangENG-Monitor-2026!"
```

- `monitor/hanyang_monitor.py`, `server/server.py`, `admin/admin.py` 의 `API_KEY` 는 **모두 동일**해야 합니다.
- 모니터의 감지 대상 프로세스는 `TARGET_PROCESSES`, 구역 구성은 `REGION_CONFIG` 에서 조정합니다.

---

## 🔐 보안 참고
- API 키와 서버 IP가 소스에 하드코딩되어 있습니다. 실제 운영 시에는 환경변수/설정파일로 분리하는 것을 권장합니다.
- 런타임 생성 데이터(`report_*.csv`, `사원명부.xlsx`, 각종 로그/설정)는 `.gitignore` 로 저장소에서 제외됩니다.
