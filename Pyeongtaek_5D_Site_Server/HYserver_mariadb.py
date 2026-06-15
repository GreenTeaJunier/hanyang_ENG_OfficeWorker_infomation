# -*- coding: utf-8 -*-
"""
HYserver MariaDB 전환 실행 파일

기존 HYserver.py 원본은 보존하고, 실행 시점에 아래 항목만 MariaDB 방식으로 교체합니다.

1. load_authorized_ips()  : 사원명부.xlsx 대신 MariaDB에서 AUTHORIZED_IPS / EMPLOYEE_INFO 로드
2. /api/employees         : Excel 파일 다운로드 대신 JSON 사원 목록 반환

서버 PC에서는 최종적으로 아래 파일을 실행하세요.

    python HYserver_mariadb.py
"""

from __future__ import annotations

import os
import sys
from typing import Any

# HYserver.py는 Pyeongtaek_5D_Site_Server 폴더 안에 있고,
# office_worker_monitor는 레포 최상위 폴더에 있으므로 부모 경로를 sys.path에 추가한다.
CURRENT_DIR = os.path.dirname(os.path.abspath(__file__))
REPO_ROOT = os.path.dirname(CURRENT_DIR)
if REPO_ROOT not in sys.path:
    sys.path.insert(0, REPO_ROOT)
if CURRENT_DIR not in sys.path:
    sys.path.insert(0, CURRENT_DIR)

from waitress import serve

import HYserver as base
from office_worker_monitor.server.employee_mariadb import load_employees_from_mariadb


def load_authorized_ips_mariadb() -> None:
    """MariaDB에서 사원 목록을 읽어 HYserver 전역 캐시에 반영합니다."""
    try:
        authorized_ips, employee_info, employees = load_employees_from_mariadb()
        base.AUTHORIZED_IPS = authorized_ips
        base.EMPLOYEE_INFO = employee_info
        print(
            f"[EMPLOYEE_DB_OK] MariaDB 사원 목록 로드 완료: "
            f"허용 IP {len(base.AUTHORIZED_IPS)}건, 사원정보 {len(base.EMPLOYEE_INFO)}건, JSON {len(employees)}건"
        )
    except Exception as exc:
        # 보안 우선: DB 로드 실패 시 전체 미인가 처리
        base.AUTHORIZED_IPS = {}
        base.EMPLOYEE_INFO = {}
        print(f"[EMPLOYEE_DB_ERROR] MariaDB 사원 목록 로드 실패: {exc}")


def _employee_sort_key(item: tuple[str, dict[str, Any]]) -> tuple[str, str, str, str]:
    ip, info = item
    return (
        str(info.get("team", "")),
        str(info.get("gongjong", "")),
        str(info.get("name", "")),
        str(ip),
    )


def get_employees_json():
    """기존 /api/employees를 MariaDB JSON 반환 방식으로 대체합니다."""
    if not base.check_api_key():
        return base.jsonify({"message": "인증 실패"}), 401

    employees = []
    for ip, info in sorted(base.EMPLOYEE_INFO.items(), key=_employee_sort_key):
        employees.append({
            "name": info.get("name", ""),
            "ip": ip,
            "pc_name": info.get("pc_name", ""),
            "team": info.get("team", "미지정") or "미지정",
            "gongjong": info.get("gongjong", "미지정") or "미지정",
            "permission": info.get("permission", ""),
            "division": info.get("division", ""),
            "modeler_tool": info.get("modeler_tool", ""),
        })

    return base.jsonify({
        "source": "mariadb",
        "count": len(employees),
        "employees": employees,
    }), 200


# ============================================================
# HYserver.py 전역 함수/라우트 패치
# ============================================================
base.load_authorized_ips = load_authorized_ips_mariadb
# 기존 @app.route('/api/employees') endpoint 이름은 get_employees 이므로 view 함수만 교체한다.
base.app.view_functions["get_employees"] = get_employees_json


# ============================================================
# 실행
# ============================================================
if __name__ == "__main__":
    base.load_authorized_ips()

    print(f"🚀 한양이엔지 HYserver MariaDB 모드 시작 (포트: {base.HYSERVER_PORT})")
    print(f" └ 사원 목록 소스: MariaDB")
    print(f" └ 인가 IP: {len(base.AUTHORIZED_IPS)}건 | 상세 정보: {len(base.EMPLOYEE_INFO)}건")
    print(f" └ /api/employees: JSON 반환")
    print(f" └ /receive_status: 클라이언트 상태 수신")
    print(f" └ /reload_employees: MariaDB 재로드")
    print(f" └ psutil: {'✅ 활성' if base.HAS_PSUTIL else '❌ 미설치 (pip install psutil)'}")

    serve(base.app, host="0.0.0.0", port=base.HYSERVER_PORT)
