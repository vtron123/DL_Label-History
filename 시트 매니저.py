import requests
import os
import sys
import subprocess
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import datetime
import re
import tkinter as tk
from tkinter import messagebox, ttk

# --- [1. 버전 및 업데이트 설정] ---
CURRENT_VERSION = "1.0.5"  # 버전 업데이트

VERSION_URL = "https://raw.githubusercontent.com/vtron123/DL_Label-History/main/version.txt"
EXE_URL = "https://github.com/vtron123/DL_Label-History/raw/main/시트%20매니저.exe"

# --- [2. 구글 시트 설정] ---
JSON_KEY_FILE = "service_account_key.json"
SHEET_NAME = "DL 트레이닝"

# 장비 매핑 정보 (LGES NB E62B 추가 완료)
MACHINE_MAP = {
    "ACC": ["ACC"],
    "EVB-CTL Verkor Pilot": ["VERKOR PILOT", "CTL PILOT"],
    "EVB-CTL Verkor GF": ["VERKOR GF", "CTL GF"],
    "7000BN": ["7000BN"],
    "HMC(현대차)": ["HMC", "현대차"],
    "Master Jig": ["MASTER JIG"],
    "LGES NA LV15": ["LV15", "NA LV15"],
    "LGES HG E81C": ["E81C"],
    "LGES WA15 E81B": ["E81B"],
    "LGES NB E62B": ["E62B", "NB E62B"],  # ✨ 새로 추가됨
    "PNT PFP-100E": ["PFP", "100E"],
    "EVB-CTS-C(원통형) 2170": ["2170"],
    "EVB-CTS-C(원통형) 4680": ["4680"],
    "EVB-XFP-A(HJV)": ["HJV", "XFP"],
    "1호기": ["1호기"],
    "2호기": ["2호기"],
    "JH3": ["JH3"],
    "선행검증라인CTA N32S2": ["N32S2", "선행검증", "CTA"]
}


# --- [3. 자동 업데이트 로직] ---
def check_update():
    try:
        response = requests.get(VERSION_URL, timeout=5)
        latest_version = response.text.strip()
        if latest_version > CURRENT_VERSION:
            if messagebox.askyesno("업데이트 알림", f"새로운 버전({latest_version})이 있습니다.\n지금 업데이트하시겠습니까?"):
                download_and_restart()
    except Exception as e:
        print(f"업데이트 확인 불가: {e}")


def download_and_restart():
    try:
        new_file = "시트_매니저_new.exe"
        target_file = "시트 매니저.exe"
        r = requests.get(EXE_URL, stream=True)
        with open(new_file, 'wb') as f:
            for chunk in r.iter_content(chunk_size=1024):
                if chunk: f.write(chunk)
        with open("update.bat", "w", encoding='cp949') as f:
            f.write(f'@echo off\n')
            f.write(f'timeout /t 3 /nobreak > nul\n')
            f.write(f'taskkill /f /im "{target_file}" > nul 2>&1\n')
            f.write(f'del "{target_file}"\n')
            f.write(f'ren "{new_file}" "{target_file}"\n')
            f.write(f'start "" "{target_file}"\n')
            f.write(f'del "%~f0"\n')
        subprocess.Popen("update.bat", shell=True)
        sys.exit()
    except Exception as e:
        messagebox.showerror("오류", f"다운로드 실패: {e}")


# --- [4. 시트 및 검색 기능] ---
def get_sheet():
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    creds = ServiceAccountCredentials.from_json_keyfile_name(JSON_KEY_FILE, scope)
    client = gspread.authorize(creds)
    return client.open(SHEET_NAME).get_worksheet(0)


class FindDialog(tk.Toplevel):
    def __init__(self, parent, text_widget):
        super().__init__(parent)
        self.title("단어 찾기")
        self.geometry("300x100")
        self.text_widget = text_widget
        self.attributes("-topmost", True)
        tk.Label(self, text="찾을 키워드:").pack(pady=5)
        self.edit = tk.Entry(self, width=30)
        self.edit.pack(pady=5)
        self.edit.focus_set()
        self.edit.bind("<Return>", lambda e: self.find())
        tk.Button(self, text="다음 찾기", command=self.find).pack()

    def find(self):
        self.text_widget.tag_remove('found', '1.0', tk.END)
        s = self.edit.get()
        if s:
            idx = '1.0'
            while True:
                idx = self.text_widget.search(s, idx, nocase=True, stopindex=tk.END)
                if not idx: break
                lastidx = f'{idx}+{len(s)}c'
                self.text_widget.tag_add('found', idx, lastidx)
                idx = lastidx
            self.text_widget.tag_config('found', background='yellow')


def perform_search(user_input):
    try:
        sheet = get_sheet()
        all_rows = sheet.get_all_values()
        total, history_log = 0, []
        clean_input = re.sub(r'(찾아줘|내역|적어줘|알려줘|보여줘|확인|얼마|언제|총|학습|장수|장|수|\?)', ' ', user_input).strip()
        search_keywords = [kw.upper() for kw in clean_input.split() if len(kw) > 0]
        if not search_keywords:
            messagebox.showwarning("주의", "조회할 키워드를 입력해주세요.")
            return
        for row in all_rows[1:]:
            row_content = " ".join(row).upper()
            if any(kw in row_content for kw in search_keywords):
                try:
                    date, m_name, count_val, memo = row[0], row[1], str(row[3]), row[4]
                    count_val = count_val.replace(',', '').strip()
                    if count_val.isdigit():
                        val = int(count_val)
                        total += val
                        history_log.append(f"• [{date}] {m_name} : {memo} ({val}장)")
                except:
                    continue
        if history_log:
            res_msg = f"🔍 검색어: {' / '.join(search_keywords)}\n📊 총 누적: {total:,}장\n" + "-" * 45 + "\n"
            for entry in history_log: res_msg += f"{entry}\n"
            result_win = tk.Toplevel(root)
            result_win.title("조회 결과")
            txt = tk.Text(result_win, padx=15, pady=15, font=("돋움", 10))
            txt.pack(fill="both", expand=True)
            txt.insert(tk.END, res_msg)
            result_win.bind("<Control-f>", lambda e: FindDialog(result_win, txt))
        else:
            messagebox.showwarning("실패", "기록이 없습니다.")
    except Exception as e:
        messagebox.showerror("오류", f"조회 에러: {e}")


def record_tab_process():
    user_input = record_entry.get().strip()
    if not user_input: return
    if not any(c.isdigit() for c in user_input) and any(sk in user_input for sk in ["찾아", "내역", "얼마"]):
        perform_search(user_input)
    else:
        try:
            sheet = get_sheet()
            today = datetime.datetime.now().strftime("%Y-%m-%d %H:%M")
            machine_found = "개별 장비"
            upper_input = user_input.upper()

            sorted_map = sorted(MACHINE_MAP.items(), key=lambda x: len(max(x[1], key=len)), reverse=True)
            for full_name, keywords in sorted_map:
                if any(kw in upper_input for kw in keywords):
                    machine_found = full_name
                    break

            count_match = re.search(r'(\d+)\s*(장|매|개|건)', user_input)
            count = count_match.group(1) if count_match else "0"
            summary = re.sub(r'(기록|해주세요|했어|요청|전달|완료|적어줘|입력).*', '', user_input).strip()

            # ✨ [개선된 기록 로직] 빈 행 건너뛰지 않고 마지막 데이터 바로 아래에 추가
            all_dates = sheet.col_values(1)
            next_row = len(all_dates) + 1
            new_data = [today, machine_found, "데이터 기록", count, summary]
            sheet.insert_row(new_data, index=next_row)  # append_row 대신 index 지정 삽입

            messagebox.showinfo("성공", f"✅ 기록 완료!\n장비: {machine_found}\n내용: {summary}\n수량: {count}장")
        except Exception as e:
            messagebox.showerror("오류", f"기록 실패: {e}")
    record_entry.delete(0, tk.END)


# --- [5. GUI 구성] ---
root = tk.Tk()
root.title(f"DL 트레이닝 매니저 v{CURRENT_VERSION}")
root.geometry("520x260")

notebook = ttk.Notebook(root)
notebook.pack(fill="both", expand=True, padx=10, pady=10)

record_frame = tk.Frame(notebook)
notebook.add(record_frame, text=" 📝 내역 기록 ")
tk.Label(record_frame, text="내용을 입력하세요 (예: E62B 100장 / Pilot 200장)", font=("맑은 고딕", 10)).pack(pady=20)
record_entry = tk.Entry(record_frame, width=60)
record_entry.pack(pady=5)
record_entry.focus_set()
record_entry.bind("<Return>", lambda e: record_tab_process())
tk.Button(record_frame, text="데이터 전송", command=record_tab_process, bg="#2E7D32", fg="white", width=20).pack(pady=15)

search_frame = tk.Frame(notebook)
notebook.add(search_frame, text=" 🔍 데이터 조회 ")
tk.Label(search_frame, text="키워드로 조회하세요", font=("맑은 고딕", 10)).pack(pady=20)
search_entry = tk.Entry(search_frame, width=60)
search_entry.pack(pady=5)
search_entry.bind("<Return>", lambda e: perform_search(search_entry.get().strip()))
tk.Button(search_frame, text="내역 찾아보기", command=lambda: perform_search(search_entry.get().strip()), bg="#1976D2",
          fg="white", width=20).pack(pady=15)

root.after(1000, check_update)
root.mainloop()