import psutil
import requests
import socket
import time
import ctypes
import sys
import os
import winreg
import threading
import logging
import json # json형태로 로그를 예브게 출력하기 위해추가
from PIL import Image, ImageDraw
import pystray
from datetime import datetime

# ==========================================
# 🚨 [설정] 서버 IP 및 인증키
# ==========================================
SERVER_URL = "http://12.26.204.100:5000/receive_status"
API_KEY = "HanyangENG-Monitor-2026!"
TARGET_PROCESSES = ["S5D.exe", "dinno.hu3d.wpf.hookupdesigner.exe"]
CHECK_INTERVAL = 10
AUTOSTART_NAME = "HanyangENG_Monitor"

# 💡 로그 파일 경로 (.exe와 같은 폴더)
APP_DIR = os.path.dirname(sys.executable) if getattr(sys.'frozen', False) else os.path.dirname(os.path.abspath(__file__))
LOG_FILE = os.path.join(APP_DIR, "client_debug.log")
# ==========================================

# ==========================================
# 💡 디버그 로깅 설정
# ==========================================
debug_log_lines = []  # 메모리에 최근 로그 보관 (디버그 창 표시용)

def log(msg):
    """콘솔 + 파일 + 메모리에 로그 기록"""
    timestamp = datetime.now().strftime("%H:%M:%S")
    line = f"[{timestamp}] {msg}"
    print(line)
    debug_log_lines.append(line)
    if len(debug_log_lines) > 500: # 로그가 많아질 수 있으므로 보관 개수 증가
        debug_log_lines.pop(0)
    try:
        with open(LOG_FILE, 'a', encoding='utf-8') as f:
            f.write(line + "\n")
    except:
        pass

# ==========================================
# 💡 0. 중복 실행 시 기존 프로세스 강제 종료
# ==========================================

def kill_existion_insance():
  current_pid = os.getpid()
  current_proc = psutil.Process(current_pid)
  my_script_name = os.path.basename(__file__)

  killed_count = 0
  for proc in psutil.process_iter(['pid','name','cmdline']):
    try:
      if proc.info['pid'] == curret_pid:
        continue

      if proc.info['name'] and proc.info['name'].lower() == my_name.lower():
        if is_frozen:
          proc.terminate()
          proc.wait(timeout=3)
          killed_count += 1
        else:
          cmdline = proc.info.get('cmdline')
          if cmdline and any(my_script_name in cmd_part for cmd_part in cmdline):
            proc.terminate()
            proc.wait(timeout=3)
            killed_count += 1
        excc ept (psutil.NosuchProcess, psutil.AccessDenied, psutil.TimeoutExpired):
          pass

    if killed_count > 0:
      log(f"기존에 실행 중이던 프로세스 {killed_count}개를 강제 종료 했습니다.")

# ==========================================
# 💡 Windows 시작 시 자동 실행 등록
# ==========================================
def get_executable_path():
    if getattr(sys, 'frozen', False):
        return f'"{sys.executable}"'
    else:
        return f'"{sys.executable}" "{os.path.abspath(__file__)}"'

def register_autostart():
    try:
        exe_path = get_executable_path()
        reg_key = winreg.OpenKey(
            winreg.HKEY_CURRENT_USER,
            r"Software\Microsoft\Windows\CurrentVersion\Run",
            0, winreg.KEY_READ | winreg.KEY_WRITE
        )
        try:
            current_val, _ = winreg.QueryValueEx(reg_key, AUTOSTART_NAME)
            if current_val == exe_path:
                winreg.CloseKey(reg_key)
                return
        except FileNotFoundError:
            pass

        winreg.SetValueEx(reg_key, AUTOSTART_NAME, 0, winreg.REG_SZ, exe_path)
        winreg.CloseKey(reg_key)
        log(f"자동 시작 등록 완료: {exe_path}")
    except Exception as e:
        log(f"자동 시작 등록 실패 (무시됨): {e}")

# ==========================================
# 네트워크 / 프로세스 함수
# ==========================================

failed_queue = []
stop_event = threading.Evnt()

def get_local_ip():
    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        s.connect(("8.8.8.8", 80))
        ip = s.getsockname()[0]
        s.close()
        return ip
    except Exception:
        return socket.gethostbyname(socket.gethostname())

def get_detailed_window_title(process_name):
    target_pids = []
    for proc in psutil.process_iter(['name', 'pid']):
        try:
            if proc.info['name'] and proc.info['name'].lower() == process_name.lower():
                target_pids.append(proc.info['pid'])
        except (psutil.NoSuchProcess, psutil.AccessDenied):
            pass

    if not target_pids: return None

    EnumWindows = ctypes.windll.user32.EnumWindows
    EnumWindowsProc = ctypes.WINFUNCTYPE(ctypes.c_bool, ctypes.POINTER(ctypes.c_int), ctypes.POINTER(ctypes.c_int))
    GetWindowThreadProcessId = ctypes.windll.user32.GetWindowThreadProcessId
    GetWindowText = ctypes.windll.user32.GetWindowTextW
    GetWindowTextLength = ctypes.windll.user32.GetWindowTextLengthW
    IsWindowVisible = ctypes.windll.user32.IsWindowVisible

    found_titles = []
  
    def foreach_window(hwnd, lParam):
        if IsWindowVisible(hwnd):
            pid = ctypes.c_ulong()
            GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
            if pid.value in target_pids:
                length = GetWindowTextLength(hwnd)
                if length > 0:
                    buff = ctypes.create_unicode_buffer(length + 1)
                    GetWindowText(hwnd, buff, length + 1)
                    found_titles.append(buff.value)
        return True

    EnumWindows(EnumWindowsProc(foreach_window), 0)
    return " / ".join(found_titles) if found_titles else "창 제목 없음 (백그라운드)"

# 💡 S5D 시작 감지용 상태 저장 (프로세스별 최초 감지 시각)
def send_status():
    global failed_queue
    hostname = socket.gethostname()
    ip_address = get_local_ip()

    log("_" * 30)
    log("프로세스 상태 확인 및 전송 준비중...")

    for proc_name in TARGET_PROCESSES:
      window_details = get_detailed_window_title(proc_name)
      is_running = bool(window_details is not None)

      payload = 
      {
            "ip": ip_address,
            "hostname": hostname,
            "process_name": proc_name,
            "is_running": True,
            "details": window_details if window_details else "창 제목 없음"
      }
      failed_queue.append(payload)
      
    remaining_queue = []
    headers = {"X-API-Key": API_KEY}

        if not running:
            # 프로세스 꺼짐 → 감지 기록 초기화
            if proc_name in _process_first_seen:
                del _process_first_seen[proc_name]
                log(f"{proc_name} 종료 감지")
            payload = {
                "ip": ip_address, "hostname": hostname,
                "process_name": proc_name, "is_running": False, "details": ""
            }
            failed_queue.append(payload)          
            continue

        # 프로세스 실행 중 → 최초 감지 시각 기록
        if proc_name not in _process_first_seen:
            _process_first_seen[proc_name] = now_ts
            log(f"{proc_name} 새로 감지 — {S5D_STARTUP_GRACE_SEC}초 유예 후 창 제목 조회")

        # 💡 2단계: S5D는 시작 후 15초 대기 (Tiara 보안 모듈 초기화 회피)
        elapsed = now_ts - _process_first_seen[proc_name]
        if proc_name.lower() == "s5d.exe" and elapsed < S5D_STARTUP_GRACE_SEC:
            # 유예 기간 중 — 창 제목 조회 스킵, 실행 중만 보고
            payload = {
                "ip": ip_address, "hostname": hostname,
                "process_name": proc_name, "is_running": True,
                "details": f"초기화 대기 중 ({int(S5D_STARTUP_GRACE_SEC - elapsed)}초)"
            }
            failed_queue.append(payload)
            continue

        # 💡 3단계: 유예 완료 → 창 제목 조회
        window_details = get_detailed_window_title(proc_name)

        # 💡 4단계: DDWORKS는 Hookup Designer 창이 열려야 유효
        # [구역명] 패턴이 없으면 아직 작업 시작 전 → 미실행으로 처리
        if proc_name.lower() == "dinno.hu3d.wpf.hookupdesigner.exe":
            if not window_details or "[" not in window_details:
                payload = {
                    "ip": ip_address, "hostname": hostname,
                    "process_name": proc_name, "is_running": False, "details": ""
                }
                failed_queue.append(payload)
                continue



    for p in failed_queue:
        try:
            res = requests.post(SERVER_URL, json=p, headers=headers, timeout=5, proxies={"http": None, "https": None})
            if res.status_code == 403:
                log(f"서버에서 미인가 거부 (403) — 데이터 폐기")
            elif res.status_code != 200:
                remaining_queue.append(p)
        except requests.exceptions.RequestException:
            remaining_queue.append(p)

    failed_queue = remaining_queue[-50:]

def monitoring_task():
    log("모니터링 스레드 시작")
    while not stop_event.is_set():
        try:
            send_status()
        except Exception as e:
            log(f"send_status 오류: {e}")
        for _ in range(CHECK_INTERVAL):
            if stop_event.is_set(): break
            time.sleep(1)
    log("모니터링 스레드 종료")

# ==========================================
# 💡 시작 알림 — Win32 API 사용 (tkinter 충돌 완전 제거)
# ==========================================
def show_startup_notification():
    """💡 tkinter 대신 Win32 풍선 알림 또는 간단한 메시지로 시작 알림
    pystray의 notify를 사용하면 트레이 아이콘 생성 후에 가능하므로,
    여기서는 콘솔/로그에만 기록하고 트레이 생성 후 notify로 표시
    """
    # 서버 연결 테스트
    try:
        server_base = SERVER_URL.rsplit('/', 1)[0]
        res = requests.get(server_base + "/", timeout=3,
                           proxies={"http": None, "https": None})
        if res.status_code == 200:
            log("서버 연결 성공")
            return True
        else:
            log(f"서버 응답 이상 (코드: {res.status_code})")
            return False
    except:
        log("서버 연결 실패 — 백그라운드에서 재시도합니다.")
        return False

# ==========================================
# 💡 트레이 아이콘 + 메뉴
# ==========================================
def create_tray_icon():
    image = Image.new('RGBA', (64, 64), color=(0, 0, 0, 0))
    draw = ImageDraw.Draw(image)
    draw.ellipse((4, 4, 60, 60), fill="#00479A", outline="white", width=3)
    return image

def on_exit_clicked(icon, item):
    log("사용자가 종료 클릭")
    stop_event.set()
    icon.stop()

def on_debug_clicked(icon, item):
    """💡 디버그 창 열기 — 별도 스레드에서 tkinter 실행"""
    threading.Thread(target=_show_debug_window, daemon=True).start()

def _show_debug_window():
    """💡 디버그 로그 창 (독립 스레드에서 실행)"""
    import tkinter as tk
    from tkinter import scrolledtext

    debug_root = tk.Tk()
    debug_root.title("🔍 클라이언트 디버그 로그")
    debug_root.geometry("600x400")
    debug_root.attributes("-topmost", True)

    txt = scrolledtext.ScrolledText(debug_root, font=("Consolas", 9), bg="#1E1E1E", fg="#D4D4D4",
                                     insertbackground="white", wrap=tk.WORD)
    txt.pack(fill=tk.BOTH, expand=True)

    # 현재까지 쌓인 로그 표시
    for line in debug_log_lines:
        txt.insert(tk.END, line + "\n")
    txt.see(tk.END)

    # 실시간 갱신
    def refresh_log():
        current_count = int(txt.index('end-1c').split('.')[0])
        if len(debug_log_lines) > current_count - 1:
            for line in debug_log_lines[current_count - 1:]:
                txt.insert(tk.END, line + "\n")
            txt.see(tk.END)
        if debug_root.winfo_exists():
            debug_root.after(1000, refresh_log)

    debug_root.after(1000, refresh_log)

    # 상태 정보 표시
    status_frame = tk.Frame(debug_root, bg="#333333", pady=3)
    status_frame.pack(fill=tk.X)
    tk.Label(status_frame, text=f"서버: {SERVER_URL} | 주기: {CHECK_INTERVAL}초 | 큐: {len(failed_queue)}건",
             font=("맑은 고딕", 9), bg="#333333", fg="#AAAAAA").pack(side=tk.LEFT, padx=10)

    debug_root.mainloop()

# ==========================================

if __name__ == '__main__':
    try:
        log("=" * 40)
        log("클라이언트 시작")

        # 💡 1단계: Windows 시작 시 자동 실행 등록
        register_autostart()

        # 💡 2단계: 서버 연결 테스트 (tkinter 없이)
        server_ok = show_startup_notification()

        # 💡 3단계: 모니터링 스레드 시작
        monitor_thread = threading.Thread(target=monitoring_task, daemon=True)
        monitor_thread.start()
        log("모니터링 스레드 시작됨")

        # 💡 4단계: 트레이 아이콘 실행
        menu = pystray.Menu(
            pystray.MenuItem('🟢 관제 클라이언트 작동 중', lambda: None, enabled=False),
            pystray.Menu.SEPARATOR,
            pystray.MenuItem('🔍 디버그 로그', on_debug_clicked),
            pystray.MenuItem('❌ 종료', on_exit_clicked)
        )

        tray_icon = pystray.Icon("HanyangMonitor", create_tray_icon(), "한양이엔지 모니터링", menu)

        # 💡 트레이 생성 후 풍선 알림 표시
        def on_tray_ready(icon):
            if server_ok:
                icon.notify("서버 연결 성공! 관제가 시작되었습니다.", "한양이엔지 관제")
            else:
                icon.notify("서버 연결 실패 — 백그라운드에서 재시도합니다.", "한양이엔지 관제")

        tray_icon.run(setup=on_tray_ready)

    except Exception as e:
        log(f"치명적 오류: {e}")
        logging.basicConfig(filename=os.path.join(APP_DIR, "client_error.log"), level=logging.ERROR)
        logging.error(f"치명적 오류", exc_info=True)
