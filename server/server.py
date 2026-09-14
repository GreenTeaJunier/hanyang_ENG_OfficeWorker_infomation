from flask import Flask, request, jsonify, send_file
import csv
import json
import os
import secrets
import threading
from datetime import datetime, timedelta
from waitress import serve
import xlwings as xw  # 💡 openpyxl 대신 xlwings 사용
import pandas as pd
        
# True: request.remote_addr만 신뢰합니다. VPN/NAT 환경에서는 차단될 수 있습니다.
# False: request.remote_addr + 클라이언트가 보낸 client_ip 둘 다 확인합니다.
BOM_AUTH_STRICT_REMOTE_IP = False

# psutil: 서버 PC 하드웨어 모니터링 (대시보드 서버 성능 그래프용)
try:
    import psutil
    HAS_PSUTIL = True
except ImportError:
    HAS_PSUTIL = False
    print("[경고] psutil 미설치 - /api/server_status 비활성. pip install psutil")

app = Flask(__name__)

# 💡 API 인증키 (클라이언트/뷰어/관리자앱 모두 동일)
API_KEY = "HanyangENG-Monitor-2026!"

# 💡 접속 제한 인원 설정 파일 (관리자앱에서 저장, 뷰어에서 읽기)
ZONE_LIMITS_FILE = "zone_limits.json"

# 💡 [Phase 2] 인가된 IP 목록 (사원명부.xlsx 기반)
EMPLOYEE_FILE = "사원명부.xlsx"
AUTHORIZED_IPS = {}  # { "IP주소": "이름", ... } ← 기존 호환 유지

# 💡 [Phase 2 확장] 전체 사원 정보 (이름, 팀, 공종, PC이름)
EMPLOYEE_INFO = {}  # { "IP주소": {"name": "홍길동", "team": "설계팀", "gongjong": "S.GAS", "pc_name": "PC01"}, ... }

# 💡 [Phase 2] 프로그램 실행 인가 로그
EXECUTION_LOG_FILE = "execution_log.csv"


# ==========================================
# 💡 사원명부 로드 - xlwings + pandas DataFrame 방식
# ==========================================
def load_authorized_ips():
    """
    사원명부.xlsx에서 인가된 IP 목록과 사원 상세 정보를 메모리에 로드한다.

    현재 회사 표준 헤더:
        이름 | IP | PCNAME | 팀 | 공종 | 권한

    서버 내부 매핑:
        이름   -> name
        IP     -> ip
        PCNAME -> pc_name
        팀     -> team
        공종   -> gongjong
        권한   -> permission

    기존 호환 구조:
        AUTHORIZED_IPS = {"IP": "이름"}
        EMPLOYEE_INFO = {"IP": {상세 정보 dict}}
    """
    global AUTHORIZED_IPS, EMPLOYEE_INFO

    # HYserver.py 상단에 BASE_DIR을 아직 추가하지 않은 경우를 방어한다.
    base_dir = globals().get("BASE_DIR", os.path.dirname(os.path.abspath(__file__)))
    employee_file = globals().get("EMPLOYEE_FILE", os.path.join(base_dir, "사원명부.xlsx"))

    if not os.path.isabs(employee_file):
        employee_file = os.path.join(base_dir, employee_file)

    print(f"[BOM_AUTH][EMPLOYEE_PATH] {employee_file}")

    if not os.path.exists(employee_file):
        print(f"[BOM_AUTH][EMPLOYEE_ERROR] 사원명부 파일 없음: {employee_file}")
        AUTHORIZED_IPS = {}
        EMPLOYEE_INFO = {}
        return

    app_xw = None
    wb = None

    try:
        app_xw = xw.App(visible=False)
        app_xw.display_alerts = False
        app_xw.screen_updating = False

        wb = app_xw.books.open(employee_file)
        ws = wb.sheets[0]
        data = ws.used_range.value

        if not data:
            print("[BOM_AUTH][EMPLOYEE_ERROR] 사원명부에 데이터가 없습니다.")
            AUTHORIZED_IPS = {}
            EMPLOYEE_INFO = {}
            return

        # used_range가 1행만 반환하는 경우 구조 보정
        if data and not isinstance(data[0], list):
            data = [data]

        if len(data) < 2:
            print("[BOM_AUTH][EMPLOYEE_ERROR] 사원명부에 헤더 외 데이터가 없습니다.")
            AUTHORIZED_IPS = {}
            EMPLOYEE_INFO = {}
            return

        raw_headers = data[0]
        headers = [str(h).strip().upper() if h is not None else "" for h in raw_headers]

        # 헤더 별칭. 현재 표준은 이름 | IP | PCNAME | 팀 | 공종 | 권한 이지만,
        # 향후 약간의 표기 차이를 흡수할 수 있게 둔다.
        # [Phase 2 평택] 대분류팀(출도/CAD/출도지원) 컬럼이 추가되면 함께 인식한다.
        aliases = {
            "name": ["이름", "성명", "사원명", "NAME"],
            "ip": ["IP", "IP주소", "아이피", "IP ADDRESS", "IP_ADDRESS"],
            "pc_name": ["PCNAME", "PC_NAME", "PC NAME", "PC이름", "PC명", "컴퓨터이름", "HOSTNAME"],
            "team": ["팀", "부서", "소속", "TEAM"],
            "gongjong": ["공종", "공정", "업무", "GONGJONG"],
            "permission": ["권한", "사용권한", "허용여부", "PERMISSION", "AUTH"],
        }

        def find_col(key):
            for alias in aliases.get(key, []):
                alias_upper = alias.strip().upper()
                if alias_upper in headers:
                    return headers.index(alias_upper)
            return None

        idx_name = find_col("name")
        idx_ip = find_col("ip")
        idx_pc = find_col("pc_name")
        idx_team = find_col("team")
        idx_gongjong = find_col("gongjong")
        idx_permission = find_col("permission")

        if idx_ip is None:
            print(f"[BOM_AUTH][EMPLOYEE_ERROR] IP 헤더를 찾지 못했습니다. 현재 헤더: {raw_headers}")
            AUTHORIZED_IPS = {}
            EMPLOYEE_INFO = {}
            return

        def cell(row, idx):
            if idx is None:
                return ""
            if idx >= len(row):
                return ""
            value = row[idx]
            return "" if value is None else str(value).strip()

        def is_permission_denied(permission_text):
            p = str(permission_text or "").strip().upper()
            # 빈 값은 기존 운영 호환을 위해 허용으로 본다.
            if not p:
                return False
            deny_values = {"N", "NO", "FALSE", "0", "차단", "미허용", "사용안함", "비활성", "DISABLED", "DENY"}
            return p in deny_values

        new_ips = {}
        new_info = {}
        skipped_empty_ip = 0
        skipped_denied = 0

        for row in data[1:]:
            if not row:
                continue
            if not isinstance(row, list):
                row = [row]

            ip_val = cell(row, idx_ip)
            if not ip_val:
                skipped_empty_ip += 1
                continue

            permission_val = cell(row, idx_permission)
            if is_permission_denied(permission_val):
                skipped_denied += 1
                continue

            name_val = cell(row, idx_name)
            pc_name_val = cell(row, idx_pc)
            team_val = cell(row, idx_team)
            gongjong_val = cell(row, idx_gongjong)

            # 기존 호환 유지
            new_ips[ip_val] = name_val

            # 토큰 인증 확장 정보 (Phase 2 평택 사업부 조직 필드 포함)
            new_info[ip_val] = {
                "name": name_val,
                "team": team_val,
                "gongjong": gongjong_val,
                "pc_name": pc_name_val,
                "permission": permission_val,
            }

        AUTHORIZED_IPS = new_ips
        EMPLOYEE_INFO = new_info

        print(
            f"[BOM_AUTH][EMPLOYEE_OK] 사원명부 로드 완료: "
            f"허용 IP {len(AUTHORIZED_IPS)}개, IP공란 제외 {skipped_empty_ip}행, 권한차단 제외 {skipped_denied}행"
        )

    except Exception as e:
        print(f"[BOM_AUTH][EMPLOYEE_ERROR] 사원명부 로드 실패: {e}")
        # 보안 우선: 로드 실패 시 전체 차단
        AUTHORIZED_IPS = {}
        EMPLOYEE_INFO = {}

    finally:
        try:
            if wb is not None:
                wb.close()
        except Exception:
            pass

        try:
            if app_xw is not None:
                app_xw.quit()
        except Exception:
            try:
                app_xw.kill()
            except Exception:
                pass

def check_api_key():
    """API 키 인증 확인"""
    key = request.headers.get("X-API-Key", "")
    return key == API_KEY

# ==========================================
# 💡 [Phase 2] 프로그램 실행 인가 로그 기록
# ==========================================
def log_execution_attempt(ip, program_id, allowed, user_info=None, reason=""):
    """프로그램 실행 인가/거부 이력을 CSV에 기록"""
    now = datetime.now()
    time_str = now.strftime("%Y-%m-%d %H:%M:%S")

    name = user_info.get("name", "") if user_info else ""
    team = user_info.get("team", "") if user_info else ""
    result_str = "허용" if allowed else "거부"

    file_exists = os.path.isfile(EXECUTION_LOG_FILE)
    try:
        with open(EXECUTION_LOG_FILE, 'a', newline='', encoding='utf-8-sig') as f:
            writer = csv.writer(f)
            if not file_exists:
                writer.writerow(["시간", "IP", "사원명", "팀", "프로그램ID", "결과", "사유"])
            writer.writerow([time_str, ip, name, team, program_id, result_str, reason])
    except Exception as e:
        print(f"[오류] 실행 인가 로그 기록 실패: {e}")


# ==========================================
# 기존 API
# ==========================================
@app.route('/', methods=['GET'])
def health_check():
    return "<h1>✅ 한양이엔지 관제 서버가 정상 작동 중입니다!</h1>"


# ==========================================
# 💡 [Phase 2 기존] IP 인가 검증 엔드포인트 (하위 호환 유지)
# ==========================================
@app.route('/verify_ip', methods=['POST'])
def verify_ip():
    """클라이언트 IP가 사원명부에 등록되어 있는지 검증 (기존 Client.py 호환)"""
    if not check_api_key():
        return jsonify({"message": "인증 실패"}), 401

    data = request.json or {}
    client_ip = data.get('ip', '').strip()

    if not client_ip:
        return jsonify({"allowed": False, "message": "IP가 전달되지 않았습니다."}), 400

    if client_ip in AUTHORIZED_IPS:
        name = AUTHORIZED_IPS[client_ip]
        now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        print(f"[{now}] IP 인가 승인 | IP: {client_ip} | 사원: {name}")
        return jsonify({"allowed": True, "name": name}), 200
    else:
        now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        print(f"[{now}] ⛔ IP 인가 거부 | IP: {client_ip}")
        return jsonify({"allowed": False}), 200


# ==========================================
# 🆕 [Phase 2] 범용 프로그램 실행 인가 API
# ==========================================
@app.route('/api/verify_execution', methods=['POST'])
def verify_execution():
    """범용 프로그램 실행 인가 API"""
    if not check_api_key():
        return jsonify({"message": "인증 실패"}), 401

    data = request.json or {}
    client_ip = data.get('ip', '').strip()
    program_id = data.get('program_id', 'unknown').strip()
    hostname = data.get('hostname', '')

    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    if not client_ip:
        log_execution_attempt("", program_id, False, reason="IP 미전달")
        return jsonify({"allowed": False, "message": "IP가 전달되지 않았습니다."}), 400

    if client_ip not in EMPLOYEE_INFO:
        print(f"[{now_str}] ⛔ 실행 거부 | 프로그램: {program_id} | IP: {client_ip} | 사유: 미등록 PC")
        log_execution_attempt(client_ip, program_id, False, reason="미등록 PC")
        return jsonify({"allowed": False, "message": "미등록 PC"}), 200

    user_info = EMPLOYEE_INFO[client_ip]
    print(f"[{now_str}] ✅ 실행 허용 | 프로그램: {program_id} | IP: {client_ip} | 사원: {user_info['name']} ({user_info['team']})")
    log_execution_attempt(client_ip, program_id, True, user_info=user_info)

    return jsonify({
        "allowed": True,
        "user_info": {
            "name": user_info["name"],
            "team": user_info["team"],
            "gongjong": user_info["gongjong"],
            "pc_name": user_info["pc_name"],
        }
    }), 200


# ==========================================
# 🆕 [Phase 2] 실행 인가 로그 조회 (관리자용)
# ==========================================
@app.route('/api/execution_logs', methods=['GET'])
def get_execution_logs():
    """관리자가 실행 인가/거부 이력을 확인할 수 있는 API"""
    if not check_api_key():
        return jsonify({"message": "인증 실패"}), 401

    if not os.path.exists(EXECUTION_LOG_FILE):
        return jsonify({"logs": [], "count": 0}), 200

    filter_type = request.args.get('filter', '')
    filter_program = request.args.get('program_id', '')

    try:
        logs = []
        with open(EXECUTION_LOG_FILE, 'r', encoding='utf-8-sig') as f:
            reader = csv.DictReader(f)
            for row in reader:
                if filter_type == "denied" and row.get("결과") != "거부":
                    continue
                if filter_program and row.get("프로그램ID") != filter_program:
                    continue
                logs.append(row)

        recent_logs = logs[-500:]
        return jsonify({"logs": recent_logs, "count": len(recent_logs)}), 200

    except Exception as e:
        return jsonify({"message": f"로그 읽기 오류: {str(e)}"}), 500


# ==========================================
# 🆕 [Phase 2] 최근 비인가 접근 시도 요약 (관리자 알림용)
# ==========================================
@app.route('/api/denied_summary', methods=['GET'])
def denied_summary():
    """최근 24시간 내 거부된 접근 시도 요약 반환 (관리자 대시보드/알림용)"""
    if not check_api_key():
        return jsonify({"message": "인증 실패"}), 401

    if not os.path.exists(EXECUTION_LOG_FILE):
        return jsonify({"denied_count": 0, "denied_ips": []}), 200

    try:
        cutoff = datetime.now() - timedelta(hours=24)

        denied_entries = []
        with open(EXECUTION_LOG_FILE, 'r', encoding='utf-8-sig') as f:
            reader = csv.DictReader(f)
            for row in reader:
                if row.get("결과") != "거부":
                    continue
                try:
                    log_time = datetime.strptime(row["시간"], "%Y-%m-%d %H:%M:%S")
                    if log_time >= cutoff:
                        denied_entries.append({
                            "time": row["시간"],
                            "ip": row.get("IP", ""),
                            "program_id": row.get("프로그램ID", ""),
                            "reason": row.get("사유", ""),
                        })
                except ValueError:
                    continue

        return jsonify({
            "denied_count": len(denied_entries),
            "denied_ips": denied_entries[-100:]
        }), 200

    except Exception as e:
        return jsonify({"message": f"요약 조회 오류: {str(e)}"}), 500


# ==========================================
@app.route('/reload_employees', methods=['POST'])
def reload_employees():
    """관리자가 사원명부를 갱신했을 때 수동으로 IP 목록을 다시 로드"""
    if not check_api_key():
        return jsonify({"message": "인증 실패"}), 401
    load_authorized_ips()
    return jsonify({"message": "리로드 완료", "count": len(AUTHORIZED_IPS)}), 200


# ==========================================
@app.route('/receive_status', methods=['POST'])
def receive_status():
    """클라이언트에서 주기적으로 보내는 프로세스 상태 수신 및 로그 기록"""
    try:
        if not check_api_key():
            return jsonify({"message": "인증 실패"}), 401

        data = request.json
        if not data:
            return jsonify({"message": "데이터가 없습니다."}), 400

        ip = data.get('ip', '')

        # 💡 [Phase 2] 비인가 IP 차단
        if ip and ip.strip() not in AUTHORIZED_IPS:
            now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            print(f"[{now}] ⛔ 비인가 IP 로그 거부 | IP: {ip}")
            return jsonify({"message": "등록되지 않은 IP입니다. 접근이 거부되었습니다."}), 403

        hostname = data.get('hostname', '')
        process_name = data.get('process_name', '')
        is_running = "O (실행 중)" if data.get('is_running') else "X (미실행)"
        details = data.get('details', '')

        now = datetime.now()
        time_str = now.strftime("%Y-%m-%d %H:%M:%S")
        date_str = now.strftime("%Y-%m-%d")

        csv_file = f"report_{date_str}.csv"

        print(f"[{time_str}] 수신 | IP: {ip} | 프로세스: {process_name} | 상태: {is_running}")

        file_exists = os.path.isfile(csv_file)
        with open(csv_file, 'a', newline='', encoding='utf-8-sig') as f:
            writer = csv.writer(f)
            if not file_exists:
                writer.writerow(["시간", "IP Address", "PC 이름", "프로세스", "상태", "상세(창제목)"])
            writer.writerow([time_str, ip, hostname, process_name, is_running, details])

        return jsonify({"message": "수신 성공"}), 200

    except Exception as e:
        print(f"\n[서버 내부 오류 (500)] {str(e)}")
        import traceback
        traceback.print_exc()
        return jsonify({"message": f"서버 내부 오류: {str(e)}"}), 500


@app.route('/api/employees', methods=['GET'])
def get_employees():
    if not check_api_key():
        return jsonify({"message": "인증 실패"}), 401

    emp_file = "사원명부.xlsx"
    if os.path.exists(emp_file):
        return send_file(emp_file, as_attachment=True)
    return "사원명부 파일이 없습니다.", 404


@app.route('/api/logs', methods=['GET'])
def get_logs():
    if not check_api_key():
        return jsonify({"message": "인증 실패"}), 401

    date_str = datetime.now().strftime("%Y-%m-%d")
    csv_file = f"report_{date_str}.csv"
    if os.path.exists(csv_file):
        return send_file(csv_file, as_attachment=True)
    return "오늘자 로그가 아직 없습니다.", 404


# ==========================================
# 💡 [추가] 접속 제한 인원 API
# ==========================================
@app.route('/api/zone_limits', methods=['GET'])
def get_zone_limits():
    """뷰어에서 호출 — 현재 설정된 구역별 접속 제한 인원 반환"""
    if not check_api_key():
        return jsonify({"message": "인증 실패"}), 401

    try:
        if os.path.exists(ZONE_LIMITS_FILE):
            with open(ZONE_LIMITS_FILE, 'r', encoding='utf-8') as f:
                data = json.load(f)
            return jsonify(data), 200
        else:
            return jsonify({}), 200
    except Exception as e:
        return jsonify({"message": f"파일 읽기 오류: {str(e)}"}), 500


@app.route('/api/zone_limits', methods=['POST'])
def set_zone_limits():
    """관리자앱에서 호출 — 구역별 접속 제한 인원 저장"""
    if not check_api_key():
        return jsonify({"message": "인증 실패"}), 401

    data = request.json
    if data is None:
        return jsonify({"message": "데이터가 없습니다."}), 400

    try:
        with open(ZONE_LIMITS_FILE, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=2)

        now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        print(f"[{now}] 접속 제한 인원 업데이트 완료: {data}")

        return jsonify({"message": "저장 성공"}), 200
    except Exception as e:
        return jsonify({"message": f"저장 오류: {str(e)}"}), 500


# ==========================================
# 🆕 [NMS] 서버 PC 하드웨어 성능 지표
#     대시보드 탭2 "서버 CPU/RAM" 그래프가 이 API를 사용
# ==========================================
@app.route('/api/server_status', methods=['GET'])
def server_status():
    """서버 PC 자체의 CPU/RAM/디스크 사용률 반환 (psutil 기반)"""
    if not check_api_key():
        return jsonify({"message": "인증 실패"}), 401

    if not HAS_PSUTIL:
        return jsonify({"message": "psutil 미설치. pip install psutil"}), 500

    try:
        cpu_percent = psutil.cpu_percent(interval=0.5)
        mem = psutil.virtual_memory()
        disk = psutil.disk_usage('/')

        return jsonify({
            "cpu_percent": cpu_percent,
            "ram_percent": mem.percent,
            "ram_used_gb": round(mem.used / (1024**3), 1),
            "ram_total_gb": round(mem.total / (1024**3), 1),
            "disk_percent": disk.percent,
            "disk_used_gb": round(disk.used / (1024**3), 1),
            "disk_total_gb": round(disk.total / (1024**3), 1),
            "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        }), 200

    except Exception as e:
        return jsonify({"message": f"서버 상태 조회 오류: {str(e)}"}), 500


# ==========================================
# 🆕 [NMS] 활성 클라이언트 & 토폴로지 상태 조회
#     대시보드 탭2 "토폴로지 맵" + 탭1 "상태 카드"가 이 API를 사용
# ==========================================
"""
Server.py의 def active_clients() 함수를 아래로 통째로 교체하세요.

수정 내용:
  1. CSV 읽기 시 utf-8-sig → cp949 → utf-8(errors=replace) 순서로 fallback
  2. IP .strip() 정규화
  3. 등록된 PC는 차단 표시 안 함 (활성 우선)
"""

def _safe_read_csv(filepath):
    """CSV 파일을 여러 인코딩으로 시도하여 읽기 (한글 윈도우 호환)"""
    for enc in ['utf-8-sig', 'cp949', 'euc-kr']:
        try:
            with open(filepath, 'r', encoding=enc) as f:
                return list(csv.DictReader(f))
        except (UnicodeDecodeError, UnicodeError):
            continue
    # 최후 수단: 깨진 문자 무시
    with open(filepath, 'r', encoding='utf-8', errors='replace') as f:
        return list(csv.DictReader(f))


@app.route('/api/active_clients', methods=['GET'])
def active_clients():
    if not check_api_key():
        return jsonify({"message": "인증 실패"}), 401

    try:
        now = datetime.now()
        date_str = now.strftime("%Y-%m-%d")
        csv_file = f"report_{date_str}.csv"
        cutoff_5min = now - timedelta(minutes=5)

        # ── report CSV에서 마지막 활동 시각 추출 ──
        last_seen = {}
        running_progs = set()

        if os.path.exists(csv_file):
            rows = _safe_read_csv(csv_file)
            for row in rows:
                ip = (row.get("IP Address") or "").strip()
                time_str = (row.get("시간") or "").strip()
                status = row.get("상태") or ""
                proc = row.get("프로세스") or ""

                if not ip or not time_str:
                    continue

                try:
                    log_time = datetime.strptime(time_str, "%Y-%m-%d %H:%M:%S")
                    if ip not in last_seen or log_time > last_seen[ip]:
                        last_seen[ip] = log_time
                except ValueError:
                    continue

                if "실행 중" in status or "O" in status:
                    running_progs.add(f"{ip}:{proc}")

        # ── 오늘 차단된 IP ──
        denied_ips = set()
        if os.path.exists(EXECUTION_LOG_FILE):
            rows = _safe_read_csv(EXECUTION_LOG_FILE)
            for row in rows:
                if (row.get("결과") or "") != "거부":
                    continue
                denied_ip = (row.get("IP") or "").strip()
                time_str = (row.get("시간") or "").strip()
                if denied_ip and time_str:
                    try:
                        log_time = datetime.strptime(time_str, "%Y-%m-%d %H:%M:%S")
                        if log_time.date() == now.date():
                            denied_ips.add(denied_ip)
                    except ValueError:
                        continue

        # ── 사원명부 기반 노드 상태 ──
        nodes = []
        active_count = 0

        for ip, info in EMPLOYEE_INFO.items():
            ip_clean = ip.strip()
            ls = last_seen.get(ip_clean)

            if ls and ls >= cutoff_5min:
                status = "active"
                active_count += 1
            else:
                status = "inactive"

            nodes.append({
                "ip": ip_clean,
                "name": info.get("name", ""),
                "pc_name": info.get("pc_name", ""),
                "team": info.get("team", ""),
                "status": status,
                "last_seen": ls.strftime("%H:%M:%S") if ls else None
            })

        # 미등록 IP 중 차단된 것만 추가
        for denied_ip in denied_ips:
            if denied_ip not in EMPLOYEE_INFO:
                nodes.append({
                    "ip": denied_ip,
                    "name": "미등록",
                    "pc_name": "미등록 PC",
                    "team": "-",
                    "status": "denied",
                    "last_seen": None
                })

        return jsonify({
            "nodes": nodes,
            "active_count": active_count,
            "total_registered": len(EMPLOYEE_INFO),
            "running_programs": len(running_progs),
            "denied_ip_count": len(denied_ips - set(EMPLOYEE_INFO.keys())),
            "timestamp": now.strftime("%Y-%m-%d %H:%M:%S")
        }), 200

    except Exception as e:
        return jsonify({"message": f"활성 클라이언트 조회 오류: {str(e)}"}), 500

# 로그 모드:
# - "deny_only"              : 차단 기록만 저장. 로그량 최소. 권장.
# - "allow_and_deny"         : 승인/차단 모두 저장.
# - "first_allow_per_ip_day" : 차단은 항상 저장, 승인은 IP별 하루 첫 1회만 저장.
BOM_AUTH_LOG_MODE = "deny_only"
_BOM_ALLOW_LOG_CACHE = set()


def _bom_should_write_log(result: str, request_ip: str, program: str) -> bool:
    mode = str(globals().get("BOM_AUTH_LOG_MODE", "deny_only")).lower().strip()
    result = str(result or "").upper().strip()

    if mode == "allow_and_deny":
        return True

    # 차단/오류는 항상 저장합니다.
    if result != "ALLOW":
        return True

    if mode == "first_allow_per_ip_day":
        today = datetime.now().strftime("%Y-%m-%d")
        key = (today, request_ip, program)
        if key in _BOM_ALLOW_LOG_CACHE:
            return False
        _BOM_ALLOW_LOG_CACHE.add(key)
        return True

    # 기본값 deny_only: 승인 로그는 저장하지 않습니다.
    return False


@app.route("/api/bom_auth", methods=["POST"])
def api_bom_auth():
    payload = request.get_json(silent=True) or {}
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    request_ip = _bom_get_request_ip()
    program = str(payload.get("program", "HYENG_AUTO_BOM")).strip()
    pc_name = str(payload.get("pc_name", "")).strip()
    windows_user = str(payload.get("windows_user", "")).strip()

    base_log = {
        "timestamp": now,
        "program": program,
        "request_ip": request_ip,
        "pc_name": pc_name,
        "windows_user": windows_user,
    }

    if payload.get("api_key") != API_KEY:
        _bom_write_auth_log({**base_log, "result": "DENY", "reason": "INVALID_API_KEY"})
        return jsonify({
            "ok": False,
            "authorized": False,
            "reason": "INVALID_API_KEY",
            "message": "API 인증키가 일치하지 않습니다.",
        }), 401

    if not AUTHORIZED_IPS:
        load_authorized_ips()

    matched_ip = request_ip if request_ip in AUTHORIZED_IPS else ""

    if not matched_ip:
        _bom_write_auth_log({**base_log, "result": "DENY", "reason": "UNAUTHORIZED_IP", "matched_ip": ""})
        return jsonify({
            "ok": True,
            "authorized": False,
            "reason": "UNAUTHORIZED_IP",
            "message": "사원명부에 등록되지 않은 IP입니다.",
            "request_ip": request_ip,
        }), 403

    employee = _bom_employee_record(matched_ip)
    _bom_write_auth_log({
        **base_log,
        "result": "ALLOW",
        "reason": "AUTHORIZED",
        "matched_ip": matched_ip,
        "employee_name": employee.get("name", ""),
        "team": employee.get("team", ""),
        "gongjong": employee.get("gongjong", ""),
    })

    return jsonify({
        "ok": True,
        "authorized": True,
        "reason": "AUTHORIZED",
        "message": "BOM 매크로 실행이 승인되었습니다.",
        "matched_ip": matched_ip,
        "request_ip": request_ip,
        "employee": employee,
    }), 200

#=========================================토큰 보안 구역==========================================

# =============================================================================
# [D] BOM 1회용 토큰 인증 API - serve(app, ...) 바로 위에 삽입
# =============================================================================
# 아래 코드는 HYserver.py에 app, API_KEY, AUTHORIZED_IPS, EMPLOYEE_INFO,
# EXECUTION_LOG_FILE, load_authorized_ips가 이미 정의되어 있다는 전제로 동작한다.

# 운영 정책
BOM_TOKEN_TTL_SEC = 90
BOM_REQUIRED_VERSION = "BOM-HYBRID-V1"
BOM_REQUIRE_PC_NAME_MATCH = True
BOM_REQUIRE_USERNAME_MATCH = False
BOM_TOKEN_CLEANUP_INTERVAL_SEC = 300

# 메모리 토큰 저장소
# 실운영 장기화 시 SQLite/MariaDB 토큰 테이블로 이전 권장
BOM_TOKEN_STORE = {}
BOM_TOKEN_LOCK = threading.Lock()
BOM_LAST_TOKEN_CLEANUP_AT = None


def _bom_now():
    return datetime.now()


def _bom_now_text():
    return _bom_now().strftime("%Y-%m-%d %H:%M:%S")


def _bom_get_request_ip():
    """프록시 헤더를 신뢰하지 않고 실제 TCP 접속 IP만 사용한다."""
    return request.remote_addr or ""


def _bom_get_log_file():
    """execution_log.csv를 HYserver.py와 같은 폴더에 고정한다."""
    base_dir = globals().get("BASE_DIR", os.path.dirname(os.path.abspath(__file__)))
    log_file = globals().get("EXECUTION_LOG_FILE", os.path.join(base_dir, "execution_log.csv"))
    if not os.path.isabs(log_file):
        log_file = os.path.join(base_dir, log_file)
    return log_file


def _bom_write_auth_log(row):
    """
    기존 execution_log.csv와 최대한 호환되는 컬럼으로 기록한다.
    로그 기록 실패가 인증 API 500으로 번지지 않도록 방어한다.
    BOM_AUTH_LOG_MODE 설정에 따라 승인(ALLOW) 로그 저장 여부를 결정한다.
    """
    if not _bom_should_write_log(row.get("result", ""), row.get("request_ip", ""), row.get("program", "")):
        return

    fieldnames = [
        "timestamp",
        "program",
        "result",
        "reason",
        "request_ip",
        "declared_ip",
        "matched_ip",
        "employee_name",
        "team",
        "gongjong",
        "pc_name",
        "windows_user",
        "workbook_name",
    ]

    log_file = _bom_get_log_file()
    try:
        os.makedirs(os.path.dirname(log_file), exist_ok=True)
        file_exists = os.path.exists(log_file)
        with open(log_file, "a", newline="", encoding="utf-8-sig") as f:
            writer = csv.DictWriter(f, fieldnames=fieldnames)
            if not file_exists:
                writer.writeheader()
            writer.writerow({key: row.get(key, "") for key in fieldnames})
    except Exception as e:
        print(f"[BOM_AUTH][LOG_ERROR] execution_log.csv 기록 실패: {e}")
        print(f"[BOM_AUTH][LOG_PATH] {log_file}")


def _bom_employee_record(ip_addr):
    if ip_addr in EMPLOYEE_INFO:
        return EMPLOYEE_INFO.get(ip_addr) or {}
    if ip_addr in AUTHORIZED_IPS:
        return {"name": AUTHORIZED_IPS.get(ip_addr, "")}
    return {}


def _bom_cleanup_tokens():
    """만료된 토큰을 주기적으로 정리한다."""
    global BOM_LAST_TOKEN_CLEANUP_AT

    now = _bom_now()
    if BOM_LAST_TOKEN_CLEANUP_AT is not None:
        elapsed = (now - BOM_LAST_TOKEN_CLEANUP_AT).total_seconds()
        if elapsed < BOM_TOKEN_CLEANUP_INTERVAL_SEC:
            return

    with BOM_TOKEN_LOCK:
        expired_keys = []
        for token, rec in BOM_TOKEN_STORE.items():
            expires_at = rec.get("expires_at")
            if expires_at is None or expires_at < now:
                expired_keys.append(token)

        for token in expired_keys:
            BOM_TOKEN_STORE.pop(token, None)

        BOM_LAST_TOKEN_CLEANUP_AT = now


def _bom_json_error(reason, message, status_code=500, **extra):
    payload = {
        "ok": False,
        "authorized": False,
        "reason": reason,
        "message": message,
    }
    payload.update(extra)
    return jsonify(payload), status_code


@app.route("/bom/start_token", methods=["POST"])
def bom_start_token():
    """
    Python EXE가 호출한다.
    서버가 현재 접속 IP와 사원명부를 검증한 뒤 1회용 토큰을 발급한다.
    """
    try:
        _bom_cleanup_tokens()

        payload = request.get_json(silent=True) or {}
        request_ip = _bom_get_request_ip()
        declared_ip = str(payload.get("client_ip", "")).strip()
        program = str(payload.get("program", "HYENG_AUTO_BOM")).strip()
        username = str(payload.get("username", payload.get("windows_user", ""))).strip()
        pc_name = str(payload.get("pc_name", "")).strip()
        version = str(payload.get("version", "")).strip()
        workbook_name = str(payload.get("workbook_name", "TEMP_XLSM_BY_EXE")).strip()

        base_log = {
            "timestamp": _bom_now_text(),
            "program": program,
            "request_ip": request_ip,
            "declared_ip": declared_ip,
            "pc_name": pc_name,
            "windows_user": username,
            "workbook_name": workbook_name,
        }

        if payload.get("api_key") != API_KEY:
            _bom_write_auth_log({**base_log, "result": "DENY", "reason": "INVALID_API_KEY"})
            return _bom_json_error("INVALID_API_KEY", "API 인증키가 일치하지 않습니다.", 401)

        if version != BOM_REQUIRED_VERSION:
            _bom_write_auth_log({**base_log, "result": "DENY", "reason": "VERSION_MISMATCH"})
            return _bom_json_error("VERSION_MISMATCH", "BOM 실행기 버전이 서버 정책과 일치하지 않습니다.", 403)

        if not AUTHORIZED_IPS:
            load_authorized_ips()

        matched_ip = request_ip if request_ip in AUTHORIZED_IPS else ""
        if not matched_ip:
            _bom_write_auth_log({**base_log, "result": "DENY", "reason": "UNAUTHORIZED_IP"})
            return _bom_json_error(
                "UNAUTHORIZED_IP",
                "사원명부에 등록되지 않은 IP입니다.",
                403,
                request_ip=request_ip,
                declared_ip=declared_ip,
            )

        employee = _bom_employee_record(matched_ip)

        expected_pc_name = str(employee.get("pc_name", "")).strip().upper()
        received_pc_name = pc_name.strip().upper()
        if BOM_REQUIRE_PC_NAME_MATCH and expected_pc_name:
            if expected_pc_name != received_pc_name:
                _bom_write_auth_log({
                    **base_log,
                    "result": "DENY",
                    "reason": "PC_NAME_MISMATCH",
                    "matched_ip": matched_ip,
                    "employee_name": employee.get("name", ""),
                    "team": employee.get("team", ""),
                    "gongjong": employee.get("gongjong", ""),
                })
                return _bom_json_error(
                    "PC_NAME_MISMATCH",
                    "등록된 PCNAME과 현재 PC 이름이 일치하지 않습니다.",
                    403,
                    expected_pc_name=expected_pc_name,
                    received_pc_name=received_pc_name,
                )

        expected_username = str(employee.get("username", "")).strip().upper()
        received_username = username.strip().upper()
        if BOM_REQUIRE_USERNAME_MATCH and expected_username:
            if expected_username != received_username:
                _bom_write_auth_log({
                    **base_log,
                    "result": "DENY",
                    "reason": "USERNAME_MISMATCH",
                    "matched_ip": matched_ip,
                    "employee_name": employee.get("name", ""),
                    "team": employee.get("team", ""),
                    "gongjong": employee.get("gongjong", ""),
                })
                return _bom_json_error("USERNAME_MISMATCH", "등록된 사용자명과 현재 Windows 사용자명이 일치하지 않습니다.", 403)

        token = secrets.token_urlsafe(32)
        expires_at = _bom_now() + timedelta(seconds=BOM_TOKEN_TTL_SEC)

        token_record = {
            "token": token,
            "used": False,
            "issued_at": _bom_now(),
            "expires_at": expires_at,
            "request_ip": request_ip,
            "matched_ip": matched_ip,
            "username": username,
            "pc_name": pc_name,
            "version": version,
            "workbook_name": workbook_name,
            "employee": employee,
        }

        with BOM_TOKEN_LOCK:
            BOM_TOKEN_STORE[token] = token_record

        _bom_write_auth_log({
            **base_log,
            "result": "ALLOW",
            "reason": "TOKEN_ISSUED",
            "matched_ip": matched_ip,
            "employee_name": employee.get("name", ""),
            "team": employee.get("team", ""),
            "gongjong": employee.get("gongjong", ""),
        })

        return jsonify({
            "ok": True,
            "authorized": True,
            "reason": "TOKEN_ISSUED",
            "message": "BOM 실행 토큰이 발급되었습니다.",
            "token": token,
            "expires_at": expires_at.strftime("%Y-%m-%d %H:%M:%S"),
            "ttl_sec": BOM_TOKEN_TTL_SEC,
            "matched_ip": matched_ip,
            "employee": employee,
        }), 200

    except Exception as e:
        print(f"[BOM_AUTH][START_TOKEN_ERROR] {e}")
        return _bom_json_error("SERVER_INTERNAL_ERROR", f"서버 내부 오류가 발생했습니다: {e}", 500)


@app.route("/bom/verify_token", methods=["POST"])
def bom_verify_token():
    """
    Excel Workbook_Open이 호출한다.
    토큰을 검증하고 성공 즉시 used=True 처리한다.
    """
    try:
        _bom_cleanup_tokens()

        payload = request.get_json(silent=True) or {}
        request_ip = _bom_get_request_ip()
        declared_ip = str(payload.get("client_ip", "")).strip()
        program = str(payload.get("program", "BOM_XLSM_RUNTIME")).strip()
        token = str(payload.get("token", "")).strip()
        username = str(payload.get("username", payload.get("windows_user", ""))).strip()
        pc_name = str(payload.get("pc_name", "")).strip()
        version = str(payload.get("version", "")).strip()
        workbook_name = str(payload.get("workbook_name", "")).strip()

        base_log = {
            "timestamp": _bom_now_text(),
            "program": program,
            "request_ip": request_ip,
            "declared_ip": declared_ip,
            "pc_name": pc_name,
            "windows_user": username,
            "workbook_name": workbook_name,
        }

        if payload.get("api_key") != API_KEY:
            _bom_write_auth_log({**base_log, "result": "DENY", "reason": "INVALID_API_KEY"})
            return _bom_json_error("INVALID_API_KEY", "API 인증키가 일치하지 않습니다.", 401)

        if not token:
            _bom_write_auth_log({**base_log, "result": "DENY", "reason": "TOKEN_EMPTY"})
            return _bom_json_error("TOKEN_EMPTY", "검증할 토큰이 없습니다.", 403)

        with BOM_TOKEN_LOCK:
            rec = BOM_TOKEN_STORE.get(token)
            if rec is None:
                _bom_write_auth_log({**base_log, "result": "DENY", "reason": "TOKEN_NOT_FOUND"})
                return _bom_json_error("TOKEN_NOT_FOUND", "토큰이 존재하지 않거나 서버 재시작으로 초기화되었습니다.", 403)

            employee = rec.get("employee") or {}
            matched_ip = rec.get("matched_ip", "")

            if rec.get("used"):
                _bom_write_auth_log({
                    **base_log,
                    "result": "DENY",
                    "reason": "TOKEN_ALREADY_USED",
                    "matched_ip": matched_ip,
                    "employee_name": employee.get("name", ""),
                    "team": employee.get("team", ""),
                    "gongjong": employee.get("gongjong", ""),
                })
                return _bom_json_error("TOKEN_ALREADY_USED", "이미 사용된 토큰입니다.", 403)

            if rec.get("expires_at") < _bom_now():
                BOM_TOKEN_STORE.pop(token, None)
                _bom_write_auth_log({
                    **base_log,
                    "result": "DENY",
                    "reason": "TOKEN_EXPIRED",
                    "matched_ip": matched_ip,
                    "employee_name": employee.get("name", ""),
                    "team": employee.get("team", ""),
                    "gongjong": employee.get("gongjong", ""),
                })
                return _bom_json_error("TOKEN_EXPIRED", "토큰 유효 시간이 만료되었습니다.", 403)

            if request_ip != rec.get("request_ip"):
                _bom_write_auth_log({
                    **base_log,
                    "result": "DENY",
                    "reason": "REQUEST_IP_MISMATCH",
                    "matched_ip": matched_ip,
                    "employee_name": employee.get("name", ""),
                    "team": employee.get("team", ""),
                    "gongjong": employee.get("gongjong", ""),
                })
                return _bom_json_error("REQUEST_IP_MISMATCH", "토큰 발급 IP와 검증 IP가 다릅니다.", 403)

            if username != rec.get("username"):
                _bom_write_auth_log({
                    **base_log,
                    "result": "DENY",
                    "reason": "USERNAME_MISMATCH",
                    "matched_ip": matched_ip,
                    "employee_name": employee.get("name", ""),
                    "team": employee.get("team", ""),
                    "gongjong": employee.get("gongjong", ""),
                })
                return _bom_json_error("USERNAME_MISMATCH", "토큰 발급 사용자와 검증 사용자가 다릅니다.", 403)

            if BOM_REQUIRE_PC_NAME_MATCH:
                if str(pc_name).strip().upper() != str(rec.get("pc_name", "")).strip().upper():
                    _bom_write_auth_log({
                        **base_log,
                        "result": "DENY",
                        "reason": "PC_NAME_MISMATCH",
                        "matched_ip": matched_ip,
                        "employee_name": employee.get("name", ""),
                        "team": employee.get("team", ""),
                        "gongjong": employee.get("gongjong", ""),
                    })
                    return _bom_json_error(
                        "PC_NAME_MISMATCH",
                        "토큰 발급 PC 와 검증 PC 가 다릅니다.",
                        403
                    )

            if version != rec.get("version"):
                _bom_write_auth_log({
                    **base_log,
                    "result": "DENY",
                    "reason": "VERSION_MISMATCH",
                    "matched_ip": matched_ip,
                    "employee_name": employee.get("name", ""),
                    "team": employee.get("team", ""),
                    "gongjong": employee.get("gongjong", ""),
                })
                return _bom_json_error("VERSION_MISMATCH", "토큰 발급 버전과 검증 버전이 다릅니다.", 403)

            # 여기까지 통과하면 1회용 토큰 사용 처리
            rec["used"] = True
            rec["used_at"] = _bom_now()

        _bom_write_auth_log({
            **base_log,
            "result": "ALLOW",
            "reason": "TOKEN_VERIFIED",
            "matched_ip": matched_ip,
            "employee_name": employee.get("name", ""),
            "team": employee.get("team", ""),
            "gongjong": employee.get("gongjong", ""),
        })

        return jsonify({
            "ok": True,
            "authorized": True,
            "reason": "TOKEN_VERIFIED",
            "message": "BOM 실행 토큰 검증이 완료되었습니다.",
            "matched_ip": matched_ip,
            "employee": employee,
        }), 200

    except Exception as e:
        print(f"[BOM_AUTH][VERIFY_TOKEN_ERROR] {e}")
        return _bom_json_error("SERVER_INTERNAL_ERROR", f"서버 내부 오류가 발생했습니다: {e}", 500)

@app.route("/bom/debug_employee", methods=["GET"])
def bom_debug_employee():
    """
    서버 관리용 점검 API.
    운영 중 외부 공개가 우려되면 제거하거나 관리자망에서만 접근하도록 제한한다.
    """
    try:
        if not AUTHORIZED_IPS:
            load_authorized_ips()

        return jsonify({
            "ok": True,
            "count": len(AUTHORIZED_IPS),
            "employee_file": globals().get("EMPLOYEE_FILE", ""),
            "log_file": _bom_get_log_file(),
            "ips": list(AUTHORIZED_IPS.keys()),
        }), 200
    except Exception as e:
        return _bom_json_error("DEBUG_ERROR", str(e), 500)

# ==========================================
# 서버 시작
# ==========================================
if __name__ == '__main__':
    load_authorized_ips()
    print(f"🚀 한양이엔지 PLM Phase 2 - HYserver (포트: 5000 )")
    print(f" └ 인가 IP: {len(AUTHORIZED_IPS)}건 | 상세 정보: {len(EMPLOYEE_INFO)}건")
    print(f" └ psutil: {'✅ 활성' if HAS_PSUTIL else '❌ 미설치 (pip install psutil)'}")

    # 💡 디버그용 IP 목록 출력 (배포 시 제거해도 무방)
    print(f" └ [디버그] 현재 서버가 기억하는 IP 목록: {list(AUTHORIZED_IPS.keys())}")

    print(f" └ 기존 API:")
    print(f"     POST /receive_status        - 클라이언트 상태 수신")
    print(f"     POST /verify_ip             - IP 인가 검증")
    print(f"     POST /api/verify_execution  - 범용 실행 인가")
    print(f"     GET  /api/execution_logs    - 실행 인가 로그")
    print(f"     GET  /api/denied_summary    - 비인가 차단 요약")
    print(f"     GET  /api/zone_limits       - 구역 제한 조회")
    print(f"     POST /api/zone_limits       - 구역 제한 저장")
    print(f"     POST /reload_employees      - 사원명부 갱신")
    print(f" └ NMS 추가:")
    print(f"     GET  /api/server_status     - 서버 CPU/RAM/디스크")
    print(f"     GET  /api/active_clients    - 활성 클라이언트/토폴로지")
    serve(app, host='0.0.0.0', port=5000)