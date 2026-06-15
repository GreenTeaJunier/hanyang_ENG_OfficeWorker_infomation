# -*- coding: utf-8 -*-
"""
한양이엔지 통합 관제 프로그램 - MariaDB JSON 전환 실행 파일

이 파일은 기존 monitor/hanyang_monitor.py 원본을 직접 뜯어고치지 않고,
실행 시점에 아래 두 가지만 패치해서 실행합니다.

1. SERVER_URL / STATUS_REPORT_URL을 HYserver 8080 기준으로 변경
2. /api/employees 응답을 Excel 파일이 아니라 JSON으로 파싱

직원 PC에 배포할 때는 이 파일을 PyInstaller로 패키징하세요.
"""

import os
import sys
import logging

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
                f"서버 연동 완료(JSON): IP {len(self.emp_by_ip)}건, PC {len(self.emp_by_pc)}건, 사원정보 {len(self.emp_info)}건"
            )
            self.lbl_status.config(text="🟢 서버 연동 완료 (MariaDB JSON / 실시간 관제 중)", fg=self.colors["status_online"])

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
# 4. 실행
# ============================================================
if __name__ == "__main__":
    try:
        hm.main()
    except Exception:
        logging.basicConfig(
            filename=os.path.join(hm.APP_DIR, "monitor_mariadb_error.log"),
            level=logging.ERROR,
        )
        logging.error("MariaDB 전환 모니터 치명적 오류", exc_info=True)
        raise
