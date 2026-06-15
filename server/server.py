from flask import Flask, request, jsonify, send_file
import csv
import json
import os
from datetime import datetime
from waitress import serve
import xlwings as xw  # 💡 openpyxl 대신 xlwings 사용

app = Flask(__name__)

# 💡 API 인증키 (클라이언트/뷰어/관리자앱 모두 동일)
API_KEY = "HanyangENG-Monitor-2026!"

# 💡 접속 제한 인원 설정 파일 (관리자앱에서 저장, 뷰어에서 읽기)
ZONE_LIMITS_FILE = "zone_limits.json"

# 💡 [Phase 2] 인가된 IP 목록 (사원명부.xlsx 기반)
EMPLOYEE_FILE = "사원명부.xlsx"
AUTHORIZED_IPS = {}       # { "IP주소": "이름", ... }  ← 기존 호환 유지

# 💡 [Phase 2 확장] 전체 사원 정보 (이름, 팀, 공종, PC이름)
EMPLOYEE_INFO = {}        # { "IP주소": {"name": "홍길동", "team": "설계팀", "gongjong": "S.GAS", "pc_name": "PC01"}, ... }

# 💡 [Phase 2] 프로그램 실행 인가 로그
EXECUTION_LOG_FILE = "execution_log.csv"

# ==========================================
# 💡 [Phase 2] 사원명부 로드 (xlwings 모듈 사용)
# ==========================================
def load_authorized_ips():
    """사원명부.xlsx에서 인가된 IP 목록 + 상세 정보를 메모리에 로드 (xlwings)"""
    global AUTHORIZED_IPS, EMPLOYEE_INFO
    
    if not os.path.exists(EMPLOYEE_FILE):
        print(f"[경고] 서버와 같은 폴더에 '{EMPLOYEE_FILE}' 파일이 없습니다!")
        return

    app_xw = None
    try:
        # 💡 엑셀을 백그라운드(숨김) 모드로 실행
        app_xw = xw.App(visible=False)
        wb = app_xw.books.open(EMPLOYEE_FILE)
        ws = wb.sheets[0]
        
        # 💡 사용 중인 전체 셀 데이터를 2차원 리스트 형태로 한 번에 가져옴
        data = ws.used_range.value
        
        if not data:
            print("[경고] 엑셀 파일에 데이터가 없습니다.")
            return

        # 데이터가 1줄일 경우를 대비해 리스트 구조 통일
        if not isinstance(data[0], list):
            data = [data]

        new_ips = {}
        new_info = {}

        for row in data:
            if not row or len(row) < 2: 
                continue
            
            name_val = str(row[0]).strip() if row[0] is not None else ""
            ip_val = str(row[1]).strip() if row[1] is not None else ""
            
            # 헤더(첫 줄)이거나 IP가 비어있으면 건너뛰기
            if name_val == "이름" or ip_val.lower() == "ip": 
                continue
            if not ip_val:
                continue

            # C열: PC이름, D열: 팀, E열: 공종
            pc_name = str(row[2]).strip() if len(row) > 2 and row[2] is not None else ""
            team = str(row[3]).strip() if len(row) > 3 and row[3] is not None else "미지정"
            gongjong = str(row[4]).strip() if len(row) > 4 and row[4] is not None else "미지정"

            new_ips[ip_val] = name_val
            new_info[ip_val] = {
                "name": name_val,
                "team": team,
                "gongjong": gongjong,
                "pc_name": pc_name
            }

        AUTHORIZED_IPS = new_ips
        EMPLOYEE_INFO = new_info

        now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        print(f"[{now}] ✅ 인가 IP 목록 로드 완료: {len(AUTHORIZED_IPS)}건 (xlwings 사용)")

    except Exception as e:
        print(f"[오류] 사원명부 로드 실패 (xlwings): {e}")
    finally:
        # 💡 에러가 나든 성공하든 무조건 백그라운드 엑셀 프로세스를 종료해야 함
        if app_xw:
            try:
                wb.close()
            except:
                pass
            app_xw.quit()

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

    # ── IP 누락 ──
    if not client_ip:
        log_execution_attempt("", program_id, False, reason="IP 미전달")
        return jsonify({"allowed": False, "message": "IP가 전달되지 않았습니다."}), 400

    # ── IP 미등록 → 거부 ──
    if client_ip not in EMPLOYEE_INFO:
        print(f"[{now_str}] ⛔ 실행 거부 | 프로그램: {program_id} | IP: {client_ip} | 사유: 미등록 PC")
        log_execution_attempt(client_ip, program_id, False, reason="미등록 PC")
        return jsonify({"allowed": False, "message": "미등록 PC"}), 200

    # ── IP 등록됨 → 사원 정보 반환 ──
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
        from datetime import timedelta
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
    if not check_api_key():
        return jsonify({"message": "인증 실패"}), 401

    data = request.json
    if not data:
        return jsonify({"message": "데이터가 없습니다."}), 400

    ip = data.get('ip')

    # 💡 [Phase 2] 비인가 IP 차단
    if ip and ip.strip() not in AUTHORIZED_IPS:
        now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        print(f"[{now}] ⛔ 비인가 IP 로그 거부 | IP: {ip}")
        return jsonify({"message": "등록되지 않은 IP입니다. 접근이 거부되었습니다."}), 403

    hostname = data.get('hostname')
    process_name = data.get('process_name')
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

if __name__ == '__main__':
    # 💡 데이터 파일(사원명부/리포트/로그)을 server.py와 같은 폴더 기준으로 처리
    #    (어느 위치에서 실행하든 경로가 일관되도록 작업 디렉터리 고정)
    try:
        os.chdir(os.path.dirname(os.path.abspath(__file__)))
    except Exception:
        pass

    # 💡 [Phase 2] 서버 시작 시 사원명부에서 인가 IP 로드
    load_authorized_ips()
    print("🚀 실무용 무중단 수집 서버가 가동되었습니다... (포트: 5000)")
    print(f"   └ 인가 IP: {len(AUTHORIZED_IPS)}건 | 상세 정보: {len(EMPLOYEE_INFO)}건")
    print(f"   └ 범용 인증 API: /api/verify_execution (POST)")
    print(f"   └ 관리자 로그:   /api/execution_logs (GET)")
    print(f"   └ 거부 요약:     /api/denied_summary (GET)")
    serve(app, host='0.0.0.0', port=5000)
