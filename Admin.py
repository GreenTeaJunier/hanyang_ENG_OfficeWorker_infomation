import tkinter as tk
from tkinter import ttk, messagebox
import json
import requests
import sys
import os

# ==========================================
# 🔧 [설정]
# ==========================================
SERVER_URL = "http://12.26.204.100:5000"
API_KEY = "HanyangENG-Monitor-2026!"

# 💡 구역명 화면 표시 변환 테이블 (뷰어와 동일하게 유지)
ZONE_DISPLAY_MAP = {
    "LSI": "S3",
    "P4_UW": "P4-UW",
    "P4_UE": "P4-UE",
    "P4_I:s5d.exe": "P4-I (S5D)",
    "P4_I:dinno.hu3d.wpf.hookupdesigner.exe": "P4-I (DDWORKS)",
}

# 💡 지역 → 구역 목록 (뷰어 REGION_CONFIG와 동일하게 유지)
REGION_CONFIG = {
    "평택": {
        "grouped": True,
        "groups": {
            "P1": ["P1_1", "P1_2", "P1_3"],
            "P2": ["P2_1", "P2_2", "P2_3", "P2_4"],
            "P3": ["P3_1", "P3_2", "P3_3", "P3_4", "P3_I"],
            "P4": ["P4_1", "P4_I:s5d.exe", "P4_I:dinno.hu3d.wpf.hookupdesigner.exe", "P4_UW", "P4_UE"],
        }
    },
    "기흥•화성": {"grouped": False, "zones": ["V1", "LSI", "17L"]},
    "천안": {"grouped": False, "zones": ["C"]},
    "해외": {"grouped": False, "zones": ["T", "X2"]}
}

def get_display_name(zone_name):
    if zone_name in ZONE_DISPLAY_MAP:
        return ZONE_DISPLAY_MAP[zone_name]
    return zone_name.replace("_", "-")

def get_all_zones():
    """모든 구역 목록: [(내부명, 표시명, 지역/그룹), ...]"""
    result = []
    for region, config in REGION_CONFIG.items():
        if config.get("grouped"):
            for group_name, sub_zones in config.get("groups", {}).items():
                for zone in sub_zones:
                    result.append((zone, get_display_name(zone), f"{region} > {group_name}"))
        else:
            for zone in config.get("zones", []):
                result.append((zone, get_display_name(zone), region))
    return result

# ==========================================

class AdminApp:
    def __init__(self, root):
        self.root = root
        self.root.title("🔒 한양이엔지 관제 시스템 — 최고관리자 설정")
        self.root.geometry("550x650")
        self.root.resizable(False, False)

        self.colors = {
            "primary": "#00479A",
            "header_text": "#FFFFFF",
            "card_bg": "#FFFFFF",
            "background": "#F5F7FA",
            "status_online": "#4CAF50",
            "status_offline": "#F44336",
        }

        self.root.configure(bg=self.colors["background"])
        self.limit_entries = {}   # zone_internal → Entry
        self.server_limits = {}   # 서버에서 받아온 현재 값

        self.setup_ui()
        self.root.after(300, self.load_from_server)

    def setup_ui(self):
        # ── 헤더 ──
        frame_header = tk.Frame(self.root, bg=self.colors["primary"], pady=10)
        frame_header.pack(fill=tk.X)

        tk.Label(frame_header, text="🔒 접속 제한 인원 관리 (최고관리자 전용)",
                 font=("맑은 고딕", 13, "bold"),
                 fg=self.colors["header_text"],
                 bg=self.colors["primary"]).pack(side=tk.LEFT, padx=15)

        # ── 상태바 ──
        frame_status = tk.Frame(self.root, bg="#E0F7FA", pady=6, padx=10)
        frame_status.pack(fill=tk.X)

        self.lbl_status = tk.Label(frame_status, text="⏳ 서버 연결 준비 중...",
                                    font=("맑은 고딕", 10, "bold"),
                                    bg="#E0F7FA", fg=self.colors["primary"])
        self.lbl_status.pack(side=tk.LEFT, padx=10)

        # ── 안내 문구 ──
        frame_info = tk.Frame(self.root, bg=self.colors["background"], pady=8)
        frame_info.pack(fill=tk.X)

        tk.Label(frame_info,
                 text="구역별 접속 제한 인원을 입력하세요. (0 또는 빈칸 = 제한 없음)\n"
                      "저장하면 모든 뷰어에 자동으로 반영됩니다.",
                 font=("맑은 고딕", 9), fg="#666666",
                 bg=self.colors["background"], justify=tk.LEFT).pack(anchor=tk.W, padx=20)

        # ── 구역 목록 테이블 ──
        frame_table = tk.Frame(self.root, bg=self.colors["card_bg"], relief=tk.RIDGE, bd=1)
        frame_table.pack(fill=tk.BOTH, expand=True, padx=15, pady=(5, 10))

        # 헤더
        frame_hdr = tk.Frame(frame_table, bg="#E8EAF6")
        frame_hdr.pack(fill=tk.X, padx=2, pady=(2, 0))
        tk.Label(frame_hdr, text="구역", width=18, font=("맑은 고딕", 10, "bold"),
                 bg="#E8EAF6", fg=self.colors["primary"]).pack(side=tk.LEFT, padx=5)
        tk.Label(frame_hdr, text="지역", width=16, font=("맑은 고딕", 10, "bold"),
                 bg="#E8EAF6", fg=self.colors["primary"]).pack(side=tk.LEFT, padx=5)
        tk.Label(frame_hdr, text="제한 인원", width=10, font=("맑은 고딕", 10, "bold"),
                 bg="#E8EAF6", fg=self.colors["primary"]).pack(side=tk.LEFT, padx=5)

        # 스크롤 영역
        frame_scroll = tk.Frame(frame_table, bg=self.colors["card_bg"])
        frame_scroll.pack(fill=tk.BOTH, expand=True)

        canvas = tk.Canvas(frame_scroll, bg=self.colors["card_bg"], highlightthickness=0)
        scrollbar = ttk.Scrollbar(frame_scroll, orient=tk.VERTICAL, command=canvas.yview)
        frame_inner = tk.Frame(canvas, bg=self.colors["card_bg"])

        frame_inner.bind("<Configure>",
                         lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=frame_inner, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # 구역별 입력 행 생성
        all_zones = get_all_zones()
        current_region = ""

        for zone_internal, display_name, region_label in all_zones:
            # 지역 구분선
            top_region = region_label.split(" > ")[0]
            if top_region != current_region:
                current_region = top_region
                sep = tk.Frame(frame_inner, bg="#DDDDDD", height=1)
                sep.pack(fill=tk.X, padx=5, pady=(8, 2))
                tk.Label(frame_inner, text=f"── {top_region} ──",
                         font=("맑은 고딕", 9, "bold"), fg="#888888",
                         bg=self.colors["card_bg"]).pack(anchor=tk.W, padx=10)

            frame_row = tk.Frame(frame_inner, bg=self.colors["card_bg"])
            frame_row.pack(fill=tk.X, padx=5, pady=2)

            tk.Label(frame_row, text=display_name, width=18, anchor=tk.W,
                     font=("맑은 고딕", 10), bg=self.colors["card_bg"]).pack(side=tk.LEFT, padx=5)
            tk.Label(frame_row, text=region_label, width=16, anchor=tk.W,
                     font=("맑은 고딕", 9), bg=self.colors["card_bg"], fg="#888888").pack(side=tk.LEFT, padx=5)

            entry = tk.Entry(frame_row, width=10, font=("맑은 고딕", 10), justify=tk.CENTER)
            entry.pack(side=tk.LEFT, padx=5)
            self.limit_entries[zone_internal] = entry

        # ── 하단 버튼 ──
        frame_buttons = tk.Frame(self.root, bg=self.colors["background"], pady=10)
        frame_buttons.pack(fill=tk.X)

        tk.Button(frame_buttons, text="🔄 서버에서 다시 불러오기",
                  font=("맑은 고딕", 10),
                  bg="#E8EAF6", fg=self.colors["primary"],
                  relief=tk.FLAT, padx=15, pady=5, cursor="hand2",
                  command=self.load_from_server).pack(side=tk.LEFT, padx=(20, 10))

        tk.Button(frame_buttons, text="💾 서버에 저장",
                  font=("맑은 고딕", 11, "bold"),
                  bg=self.colors["primary"], fg="white",
                  relief=tk.FLAT, padx=25, pady=5, cursor="hand2",
                  command=self.save_to_server).pack(side=tk.RIGHT, padx=(10, 20))

    # ==========================================
    # 🌐 서버 통신
    # ==========================================
    def load_from_server(self):
        """서버에서 현재 제한 인원 불러오기"""
        self.lbl_status.config(text="⬇️ 서버에서 불러오는 중...", fg=self.colors["primary"])
        self.root.update()

        try:
            headers = {"X-API-Key": API_KEY}
            res = requests.get(f"{SERVER_URL}/api/zone_limits", headers=headers,
                               timeout=5, proxies={"http": None, "https": None})

            if res.status_code == 200:
                self.server_limits = res.json()

                # Entry에 값 채우기
                for zone_internal, entry in self.limit_entries.items():
                    entry.delete(0, tk.END)
                    if zone_internal in self.server_limits:
                        entry.insert(0, str(self.server_limits[zone_internal]))

                self.lbl_status.config(text="🟢 서버 연동 완료 — 현재 설정값 표시 중",
                                        fg=self.colors["status_online"])
            elif res.status_code == 401:
                self.lbl_status.config(text="⚠️ 인증 실패: API KEY를 확인하세요.",
                                        fg=self.colors["status_offline"])
            else:
                self.lbl_status.config(text=f"⚠️ 서버 응답 오류 (코드: {res.status_code})",
                                        fg=self.colors["status_offline"])

        except requests.exceptions.RequestException as e:
            self.lbl_status.config(text=f"❌ 연결 실패: {str(e)[:30]}...",
                                    fg=self.colors["status_offline"])

    def save_to_server(self):
        """현재 입력값을 서버에 저장"""
        new_limits = {}
        for zone_internal, entry in self.limit_entries.items():
            val = entry.get().strip()
            if val:
                try:
                    num = int(val)
                    if num > 0:
                        new_limits[zone_internal] = num
                except ValueError:
                    display = get_display_name(zone_internal)
                    messagebox.showwarning("입력 오류", f"'{display}' 항목에 숫자만 입력해주세요.")
                    return

        self.lbl_status.config(text="⬆️ 서버에 저장 중...", fg=self.colors["primary"])
        self.root.update()

        try:
            headers = {"X-API-Key": API_KEY, "Content-Type": "application/json"}
            res = requests.post(f"{SERVER_URL}/api/zone_limits",
                                json=new_limits, headers=headers,
                                timeout=5, proxies={"http": None, "https": None})

            if res.status_code == 200:
                self.server_limits = new_limits
                self.lbl_status.config(text="✅ 저장 완료! 모든 뷰어에 자동 반영됩니다.",
                                        fg=self.colors["status_online"])
            elif res.status_code == 401:
                self.lbl_status.config(text="⚠️ 인증 실패: API KEY를 확인하세요.",
                                        fg=self.colors["status_offline"])
            else:
                self.lbl_status.config(text=f"⚠️ 저장 실패 (코드: {res.status_code})",
                                        fg=self.colors["status_offline"])

        except requests.exceptions.RequestException as e:
            self.lbl_status.config(text=f"❌ 연결 실패: {str(e)[:30]}...",
                                    fg=self.colors["status_offline"])


if __name__ == "__main__":
    root = tk.Tk()
    app = AdminApp(root)
    root.mainloop()
