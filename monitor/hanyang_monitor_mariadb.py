# -*- coding: utf-8 -*-
"""
한양이엔지 통합 관제 프로그램 - MariaDB JSON 전환 실행 파일

이 파일은 기존 monitor/hanyang_monitor.py 원본을 직접 뜯어고치지 않고,
실행 시점에 아래 기능만 패치해서 실행합니다.

1. SERVER_URL / STATUS_REPORT_URL을 HYserver 8080 기준으로 변경
2. /api/employees 응답을 Excel 파일이 아니라 JSON으로 파싱
3. NMS(평택 5D 관제) 콜렉터 연동
   - heartbeat: agent.last_seen 갱신용
   - ingest: CPU/RAM/디스크/네트워크/프로세스/최근 로그 적재용
   - MariaDB NOW() 기준과 맞추기 위해 NMS 전송 timestamp는 로컬 시간(naive ISO) 사용

직원 PC에 배포할 때는 이 파일을 PyInstaller로 패키징하세요.
"""

from __future__ import annotations

import hashlib
import json
import logging
import os
import socket
import sys
import threading
from datetime import datetime
from pathlib import Path

import psutil

import hanyang_monitor as hm


# ============================================================
# 1. 운영 서버 설정
# ============================================================
# 환경변수가 있으면 환경변수 우선, 없으면 기본값 사용.
# 직원 PC 배포 전 서버 IP가 다르면 여기 기본값만 수정하면 됩니다.
hm.SERVER_URL = os.getenv("HANYANG_MONITOR_SERVER_URL", "http://12.26.204.100:8080").rstrip("/")
hm.API_KEY = os.getenv("HANYANG_MONITOR_API_KEY", "HanyangENG-Monitor-2026!")
hm.STATUS_REPORT_URL = f"{hm.SERVER_URL}/receive_status"


# ============================================================
# 2. /api/employees JSON 파싱 함수
# ============================================================
def parse_employee_payload(self, payload: dict) -> None:
    """MariaDB 기반 /api/employees JSON 응답을 기존 뷰어 내부 딕셔너리에 반영합니다."""
    self.emp_by_ip.clear()
    self.emp_by_pc.clear()
    self.emp_info.clear()

    employees = payload.get("employees", [])
    if not isinstance(employees, list):
        employees = []

    for row in employees:
        if not isinstance(row, dict):
            continue

        name = str(row.get("name", "")).strip()
        ip = str(row.get("ip", "")).strip()
        pc_name = str(row.get("pc_name", row.get("pc", ""))).strip()
        team = str(row.get("team", "미지정")).strip() or "미지정"
        gongjong = str(row.get("gongjong", "미지정")).strip() or "미지정"

        if not name:
            continue

        if ip:
            self.emp_by_ip[ip] = name
        if pc_name:
            self.emp_by_pc[pc_name] = name

        self.emp_info[name] = {
            "team": team,
            "gongjong": gongjong,
            "ip": ip,
            "pc": pc_name,
            "division": str(row.get("division", "")).strip(),
            "modeler_tool": str(row.get("modeler_tool", "")).strip(),
        }


# ============================================================
# 3. 기존 auto_connect_and_start 교체
# ============================================================
def auto_connect_and_start_json(self):
    self.lbl_status.config(text="⬇️ 서버에서 MariaDB 사원 목록 연동 중...", fg=self.colors["primary"])
    self.root.update()

    try:
        headers = {"X-API-Key": hm.API_KEY}
        res = hm.requests.get(
            f"{hm.SERVER_URL}/api/employees",
            headers=headers,
            timeout=5,
            proxies={"http": None, "https": None},
        )

        if res.status_code == 200:
            try:
                payload = res.json()
            except ValueError:
                self.lbl_status.config(
                    text="⚠️ 사원 목록 응답이 JSON이 아닙니다. HYserver_mariadb.py 실행 여부 확인",
                    fg=self.colors["status_offline"],
                )
                return

            self.parse_employee_payload(payload)
            self.fetch_zone_limits()
            self.debug_log(
                f"서버 연동 완료(JSON): IP {len(self.emp_by_ip)}건, "
                f"PC {len(self.emp_by_pc)}건, 사원정보 {len(self.emp_info)}건"
            )
            self.lbl_status.config(
                text="🟢 서버 연동 완료 (MariaDB JSON / 실시간 관제 중)",
                fg=self.colors["status_online"],
            )

            default_region = self.settings.get("default_region", "__favorites__")
            if default_region == "__favorites__":
                self.show_favorites()
            else:
                default_label = self._get_button_label(default_region)
                self.on_region_click(default_label)

            self.manual_refresh()

        elif res.status_code == 401:
            self.lbl_status.config(text="⚠️ 인증 실패: API KEY를 확인하세요.", fg=self.colors["status_offline"])
        else:
            self.lbl_status.config(text=f"⚠️ 사원 목록 연동 실패: HTTP {res.status_code}", fg=self.colors["status_offline"])

    except hm.requests.exceptions.RequestException as exc:
        self.lbl_status.config(text=f"❌ 연결 실패: {str(exc)[:30]}...", fg=self.colors["status_offline"])


# 기존 클래스에 JSON 방식 메서드 주입
hm.MonitorViewer.parse_employee_payload = parse_employee_payload
hm.MonitorViewer.auto_connect_and_start = auto_connect_and_start_json


# ============================================================
# 4. NMS(평택 5D 관제) 콜렉터 연동
# ============================================================
# 전송 규약(monitoring/common/models.py 기준):
#   POST /api/v1/heartbeat : {agent_id, hostname, ip, ts, agent_version, agent_hash}
#   POST /api/v1/ingest    : {agent_id, hostname, ip, sent_at, metrics[], processes[], logs[]}
#   인증 헤더: X-Agent-Token: <원본 토큰> (콜렉터는 sha256 해시로 대조)
#
# 설정(우선순위: 환경변수 > nms_agent_config.json > 비활성):
#   NMS_COLLECTOR_URL    예) http://12.26.204.100:9000
#   NMS_AGENT_TOKEN      NMS 서버에서 발급한 '원본 토큰'
#   NMS_AGENT_ID         미지정 시 hostname 사용
#   NMS_COLLECT_INTERVAL 보고 주기(초, 기본 30)
#   NMS_VERIFY_TLS       TLS 검증 여부(기본 true, 평문 http면 무관)
#
# ※ MariaDB/Dashboard가 NOW() 기준으로 조회하므로 ts/sent_at은 UTC가 아니라 로컬 시간으로 보냅니다.
# ※ URL과 토큰이 모두 설정돼 있을 때만 동작합니다.

NMS_AGENT_VERSION = "hanyang-monitor-mariadb/1.1"
NMS_PROCESS_LABEL = "HanyangMonitor"
NMS_CONFIG_FILE = os.path.join(hm.APP_DIR, "nms_agent_config.json")


def _load_nms_config() -> dict:
    """NMS 연동 설정을 환경변수 > nms_agent_config.json 순으로 읽어 반환합니다."""
    file_conf: dict = {}
    try:
        with open(NMS_CONFIG_FILE, encoding="utf-8-sig") as f:
            loaded = json.load(f)
            if isinstance(loaded, dict):
                file_conf = loaded
    except (OSError, ValueError):
        file_conf = {}

    def pick(env_key: str, file_key: str, default):
        env_val = os.getenv(env_key)
        if env_val not in (None, ""):
            return env_val
        if file_conf.get(file_key) not in (None, ""):
            return file_conf[file_key]
        return default

    hostname = socket.gethostname()
    try:
        interval = int(pick("NMS_COLLECT_INTERVAL", "collect_interval_sec", 30))
    except (TypeError, ValueError):
        interval = 30

    verify_raw = pick("NMS_VERIFY_TLS", "verify_tls", "true")
    return {
        "collector_url": str(pick("NMS_COLLECTOR_URL", "collector_url", "")).rstrip("/"),
        "agent_token": str(pick("NMS_AGENT_TOKEN", "agent_token", "")),
        "agent_id": str(pick("NMS_AGENT_ID", "agent_id", "") or hostname),
        "hostname": hostname,
        "interval_sec": max(interval, 5),
        "verify_tls": str(verify_raw).strip().lower() not in ("0", "false", "no", "off"),
    }


class NmsReporter:
    """이 모니터 PC를 NMS 콜렉터에 push-only 로 보고하는 경량 에이전트."""

    def __init__(self, config: dict, log_callback=None) -> None:
        self._cfg = config
        self._log = log_callback or (lambda msg: None)
        self._stop = threading.Event()
        self._thread = None
        self._last_disk = None
        self._last_net = None
        self._last_ts = None
        self._sent_log_keys: set[str] = set()

    # ── 공개 API ──
    def start(self) -> None:
        """연동이 설정돼 있으면 데몬 스레드를 띄웁니다. 아니면 조용히 통과합니다."""
        if not self._cfg.get("collector_url") or not self._cfg.get("agent_token"):
            self._log("[NMS] 콜렉터 연동 비활성화 (NMS_COLLECTOR_URL / NMS_AGENT_TOKEN 미설정)")
            return
        self._thread = threading.Thread(target=self._run, name="nms-reporter", daemon=True)
        self._thread.start()
        self._log(
            f"[NMS] 콜렉터 연동 시작 → {self._cfg['collector_url']} "
            f"(agent_id={self._cfg['agent_id']}, {self._cfg['interval_sec']}초 주기)"
        )

    def stop(self) -> None:
        """다음 주기 대기에서 루프를 종료시킵니다."""
        self._stop.set()

    # ── 내부 루프 ──
    def _run(self) -> None:
        # psutil CPU 사용률은 첫 호출이 기준점이라 한 번 버려 워밍업한다.
        try:
            psutil.cpu_percent(interval=None)
            self._last_disk = psutil.disk_io_counters()
            self._last_net = psutil.net_io_counters()
            self._last_ts = datetime.now()
        except Exception:
            pass

        agent_hash = self._self_hash()
        while not self._stop.is_set():
            try:
                self._send_heartbeat(agent_hash)
                self._send_ingest()
            except Exception as exc:  # 어떤 경우에도 모니터를 죽이지 않는다
                self._log(f"[NMS] 보고 중 예외 무시: {exc}")
            self._stop.wait(self._cfg["interval_sec"])

    # ── 전송 ──
    def _post(self, path: str, payload: dict) -> int:
        """콜렉터로 JSON POST. 사내 프록시 간섭을 피하려고 프록시를 끈다."""
        res = hm.requests.post(
            self._cfg["collector_url"] + path,
            json=payload,
            headers={
                "X-Agent-Token": self._cfg["agent_token"],
                "Content-Type": "application/json",
            },
            timeout=5,
            verify=self._cfg["verify_tls"],
            proxies={"http": None, "https": None},
        )

        if res.status_code in (401, 403):
            self._log(
                f"[NMS] {path} 인증 거부(HTTP {res.status_code}). "
                "이 PC의 NMS 등록 여부·토큰·화이트리스트 IP를 확인하세요."
            )
        elif not (200 <= res.status_code < 300):
            body = (res.text or "").replace("\n", " ")[:500]
            self._log(f"[NMS] {path} 전송 실패(HTTP {res.status_code}): {body}")
        return res.status_code

    def _send_heartbeat(self, agent_hash: str) -> None:
        self._post("/api/v1/heartbeat", {
            "agent_id": self._cfg["agent_id"],
            "hostname": self._cfg["hostname"],
            "ip": self._local_ip(),
            "ts": self._now_iso(),
            "agent_version": NMS_AGENT_VERSION,
            "agent_hash": agent_hash,
        })

    def _send_ingest(self) -> None:
        self._post("/api/v1/ingest", {
            "agent_id": self._cfg["agent_id"],
            "hostname": self._cfg["hostname"],
            "ip": self._local_ip(),
            "sent_at": self._now_iso(),
            "metrics": self._collect_metrics(),
            "processes": self._collect_processes(),
            "logs": self._collect_logs(),
        })

    # ── 수집 (monitoring/agent/collectors.py 와 동일한 와이어 스키마) ──
    def _collect_metrics(self) -> list:
        """CPU/RAM 사용률과 직전 호출 이후의 디스크/네트워크 처리량(KB/s)을 한 건 샘플링."""
        now = datetime.now()
        try:
            cpu_pct = psutil.cpu_percent(interval=None)
            mem = psutil.virtual_memory()
        except Exception:
            return []

        disk_read_kb = disk_write_kb = net_sent_kb = net_recv_kb = 0.0
        try:
            disk = psutil.disk_io_counters()
            net = psutil.net_io_counters()
            if self._last_ts is not None and self._last_disk is not None and self._last_net is not None:
                elapsed = max((now - self._last_ts).total_seconds(), 1e-6)
                disk_read_kb = (disk.read_bytes - self._last_disk.read_bytes) / 1024 / elapsed
                disk_write_kb = (disk.write_bytes - self._last_disk.write_bytes) / 1024 / elapsed
                net_sent_kb = (net.bytes_sent - self._last_net.bytes_sent) / 1024 / elapsed
                net_recv_kb = (net.bytes_recv - self._last_net.bytes_recv) / 1024 / elapsed
            self._last_disk, self._last_net, self._last_ts = disk, net, now
        except Exception:
            pass

        return [{
            "ts": now.isoformat(),
            "cpu_pct": round(cpu_pct, 1),
            "ram_pct": round(mem.percent, 1),
            "ram_used_mb": round(mem.used / 1024 / 1024, 1),
            "disk_read_kb": round(max(disk_read_kb, 0.0), 1),
            "disk_write_kb": round(max(disk_write_kb, 0.0), 1),
            "net_sent_kb": round(max(net_sent_kb, 0.0), 1),
            "net_recv_kb": round(max(net_recv_kb, 0.0), 1),
        }]

    def _collect_processes(self) -> list:
        """이 모니터 자신과(고정 라벨), 감시 대상(S5D/DDWORKS 등) 프로세스를 보고."""
        now_iso = self._now_iso()
        samples: list = []

        # (1) 모니터 자신 — NMS 가 항상 같은 이름으로 탐지하도록 고정 라벨 사용
        try:
            me = psutil.Process(os.getpid())
            samples.append({
                "ts": now_iso,
                "name": NMS_PROCESS_LABEL,
                "pid": me.pid,
                "cpu_pct": 0.0,
                "ram_mb": round(me.memory_info().rss / 1024 / 1024, 1),
            })
        except Exception:
            pass

        # (2) 모니터의 감시 대상이 떠 있으면 함께 보고 (있을 때만)
        targets = {str(p).lower() for p in getattr(hm, "TARGET_PROCESSES", [])}
        if targets:
            try:
                for proc in psutil.process_iter(["pid", "name", "memory_info"]):
                    try:
                        info = proc.info
                        pname = info.get("name") or ""
                        if pname.lower() in targets:
                            mem_info = info.get("memory_info")
                            ram_mb = (mem_info.rss / 1024 / 1024) if mem_info else 0.0
                            samples.append({
                                "ts": now_iso,
                                "name": pname,
                                "pid": info.get("pid") or 0,
                                "cpu_pct": 0.0,
                                "ram_mb": round(ram_mb, 1),
                            })
                    except (psutil.NoSuchProcess, psutil.AccessDenied):
                        continue
            except Exception:
                pass

        return samples

    def _collect_logs(self) -> list:
        """최근 monitor_debug.log 일부를 NMS app_log 로 전송합니다.

        기존 클라이언트는 logs=[]만 보내서 NMS 최근 로그가 비어 있었다.
        여기서는 마지막 80줄 중 아직 보낸 적 없는 줄만 최대 20건 전송한다.
        """
        log_path = Path(getattr(hm, "LOG_FILE", "") or "")
        if not log_path.exists():
            return []

        try:
            lines = log_path.read_text(encoding="utf-8", errors="replace").splitlines()[-80:]
        except Exception:
            return []

        rows = []
        for raw in lines:
            line = raw.strip()
            if not line:
                continue
            key = hashlib.sha1(line.encode("utf-8", errors="ignore")).hexdigest()
            if key in self._sent_log_keys:
                continue
            self._sent_log_keys.add(key)
            level = "ERROR" if any(word in line.upper() for word in ("ERROR", "실패", "예외", "거부")) else "INFO"
            rows.append({
                "ts": self._now_iso(),
                "level": level,
                "source": "HanyangMonitor",
                "message": line[:1000],
            })
            if len(rows) >= 20:
                break

        # 장시간 실행 시 set 무한 증가 방지
        if len(self._sent_log_keys) > 1000:
            self._sent_log_keys = set(list(self._sent_log_keys)[-500:])

        return rows

    # ── 보조 ──
    @staticmethod
    def _now_iso() -> str:
        """MariaDB NOW()와 Dashboard 조회 조건에 맞추기 위해 로컬 시간으로 전송."""
        return datetime.now().replace(microsecond=0).isoformat()

    @staticmethod
    def _local_ip() -> str:
        """외부로 향하는 소켓의 로컬 주소를 구해 이 PC의 실제 LAN IP를 반환."""
        try:
            s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
            s.connect(("8.8.8.8", 80))
            ip = s.getsockname()[0]
            s.close()
            return ip
        except Exception:
            try:
                return socket.gethostbyname(socket.gethostname())
            except Exception:
                return ""

    @staticmethod
    def _self_hash() -> str:
        """실행 엔트리포인트(프리징 EXE면 실행파일)의 sha256 — 캐주얼 변조 탐지용."""
        try:
            if getattr(sys, "frozen", False):
                target = sys.executable
            else:
                target = os.path.abspath(sys.argv[0]) if sys.argv and sys.argv[0] else ""
            if not target or not os.path.exists(target):
                return ""
            with open(target, "rb") as handle:
                return hashlib.sha256(handle.read()).hexdigest()
        except Exception:
            return ""


def start_nms_collector_reporter() -> "NmsReporter":
    """NMS 콜렉터 연동 리포터를 생성·기동하고 인스턴스를 반환합니다.

    설정이 없으면 비활성 상태로 반환됩니다(예외를 던지지 않습니다).
    """
    config = _load_nms_config()

    def _log(msg: str) -> None:
        # 모니터의 기존 디버그 로그 파일과 콘솔에 함께 남긴다.
        try:
            print(msg)
            with open(hm.LOG_FILE, "a", encoding="utf-8") as f:
                f.write(msg + "\n")
        except Exception:
            pass

    reporter = NmsReporter(config, log_callback=_log)
    reporter.start()
    return reporter


# ============================================================
# 5. 실행
# ============================================================
if __name__ == "__main__":
    # NMS 콜렉터(run_collector.py)가 이 PC를 탐지하도록 리포터를 먼저 띄운다.
    # 설정이 없으면 조용히 비활성화되므로 기존 모니터 동작에는 영향이 없다.
    try:
        start_nms_collector_reporter()
    except Exception:
        logging.basicConfig(
            filename=os.path.join(hm.APP_DIR, "monitor_mariadb_error.log"),
            level=logging.ERROR,
        )
        logging.error("NMS 콜렉터 리포터 시작 실패(모니터는 계속 실행)", exc_info=True)

    try:
        hm.main()
    except Exception:
        logging.basicConfig(
            filename=os.path.join(hm.APP_DIR, "monitor_mariadb_error.log"),
            level=logging.ERROR,
        )
        logging.error("MariaDB 전환 모니터 치명적 오류", exc_info=True)
        raise
