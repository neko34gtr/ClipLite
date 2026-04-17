import time
import threading
import win32clipboard
import tkinter as tk
from tkinter import filedialog
from PIL import Image, ImageGrab, ImageTk
from io import BytesIO
import ctypes
import windnd
from queue import Queue
import sys
import json
import os
import pystray
import base64
from datetime import datetime
from pystray import MenuItem as item
import win32com.client
import subprocess
import webbrowser  # メールURL起動用
import urllib.parse  # URLエンコード用
import requests  # [ADD] GitHub APIアクセス用

# --- Windowsタスクバー用 ID設定 (Pythonロゴ化を防止) ---
MY_APP_ID = 'tsai.cliplite.pro.v2' # 任意のユニークな文字列
try:
    ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(MY_APP_ID)
except:
    pass

# --- パス解決用関数 (EXE化対応) ---
def resource_path(relative_path):
    """ PyInstallerの1ファイルモードと通常実行の両方でパスを通す """
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

# [ADD] 外部ファイル .key からトークンを読み込む関数を新規追加 2026/04/17
def load_github_token():
    """プロジェクトルートの .key ファイルからトークンを読み込む"""
    # 実行ファイルと同じ階層の .key ファイルを参照
    key_path = os.path.join(os.path.dirname(__file__), '.key')
    if os.path.exists(key_path):
        try:
            with open(key_path, 'r', encoding='utf-8') as f:
                for line in f:
                    if line.startswith('GITHUB_TOKEN='):
                        parts = line.strip().split('=', 1)
                        if len(parts) > 1:
                            return parts[1]
        except Exception:
            return None
    return None

# --- ClipLite 定数・初期設定 ---
VERSION = "2.4.2"
AUTHOR_INFO = "neko52tsai@gmail.com"

# --- Git定数設定 ---
GITHUB_USER = "neko34gtr"
GITHUB_REPO = "ClipLite"
API_URL = f"https://api.github.com/repos/{GITHUB_USER}/{GITHUB_REPO}/releases/latest"
API_URL_DEV = f"https://api.github.com/repos/{GITHUB_USER}/{GITHUB_REPO}/releases"
# [ADD] 読み込んだトークンを保持する変数を定義
GITHUB_TOKEN = load_github_token()

# メールの件名と本文を定義
mail_subject = f"【問合せ】ClipLite Pro {VERSION} について"
mail_body = (
    "ご意見、ご質問など。\n\n"
    "お疲れ様です。\n"
    "ClipLiteProの動作に関して、以下の通り問い合わせいたします。\n\n"
    "【内容】\n"
    "・\n"
    "・\n"
    "・\n\n"
    "ご確認のほど、よろしくお願いいたします。"
)
# AUTHOR_INFO や件名、本文をURLエンコードして結合
SUPPORT_URL = (
    f"https://mail.google.com/mail/?view=cm&to={AUTHOR_INFO}"
    f"&su={urllib.parse.quote(mail_subject)}"
    f"&body={urllib.parse.quote(mail_body)}"
)

MAX_WIDTH = 1200
MAX_WIDTH_4K = 1920  # [ADD] 4K以上のソース時に許容する幅（フルHDサイズ）
WEBP_QUALITY = 82    # [ADD] 画質設定をわずかに向上（デフォルト80→82）
WEBP_METHOD = 6      # [ADD] 最高品質の圧縮アルゴリズム（画質とファイルサイズのバランスを最適化）

DARK_BG = "#1e1e1e"
DARK_FG = "#ffffff"
ACCENT_COLOR = "#0078d4"
STATUS_BG = "#2d2d2d"
CONFIG_FILE = "cl_config.json"
ICON_FILE = "ClipLite.ico" # ★作成したアイコンファイル名

# --- GUIアイコンデータ (アプリIconは外部ファイル優先) ---
iCON_HELP="iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAYAAAAf8/9hAAAA90lEQVR4AbRTuw6CQBAEQyj8A0srtbTSytARXn6BDZWJP6GtX6CljdpaQOhMLAylrf8CCc6aHEEyZ2Ei2cntzs3OLUfoGJrH9/1JGIY58EQ+08gMrYFpmueqqqbAAPktiqIuM6EGnueNIe4DdZRlKVxdq4QaZFn2gGAHqDiAu6uiuVIDEaRpusLoC7xCjDwWjkFrIOIkSY62bV9wiWupGb4aoHFZFEWOSTasWThqEATBHLiicQ/RENAGNYB6C3x8BdQ0qAEubQT8bkCP0pB0Ao2W0v8zcF231zyyXas9OgH+QMeyrJMSydquhRNQA9kAHKAZ7fq99wIAAP//LkSD0wAAAAZJREFUAwBP0jqleA1LRQAAAABJRU5ErkJggg=="
ICON_SETTING="iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAYAAAAf8/9hAAABPElEQVR4AZySXW6DMBCEwcdJLpE+BugjXCfhOvBYfh6bSyTHgc632JaFglo12pHHs7Mbr43LDn5FUdzKslyB+P3AlsUG1+v1tDPm74rw4A05a4DgnOvyPOdfO5kq8VMwwdF0GvPgpYa8NZCh0eYsELX2XyK1EGKvneWhZhthmiZm7IP7D2vva7YGvuBdAzTgLXGJmtNsd6BUeuRMR/wYx7EBcOXTqKkBTskbUDZt8BiG4SHNwvO4l8idWJ1dooR/h1vXtQXqEOcSv1RVddFq4XncS+ypAY7bBIhCDCW/eXcAj4mN2CtQl46Q3sFmyzI0EPZhjZo14DaViaL4b2GvgMka6IidNk+BYL5PkfRO9trT12wf0jzPr2VZGokt767ZBvGXmljA0ciJt3ipIWkngCDIxCfN9hB48AbDDwAAAP//gTvHWgAAAAZJREFUAwAUSKid//CuxgAAAABJRU5ErkJggg=="
ICON_POWER = "iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAYAAAAf8/9hAAABtklEQVR4AXySO0tDQRCFbzaV1qKVEDAaQbAWO5uYN4hEBC2EYCNWClaCP0AxoMRCbISohV3IAxTURsRCKxNIZS12CXYh8Ztld0FyY5iTOXNm7tnHvcob8ItGo6PJZPJTIHzAmDfQIBgMzvNQSGA4tD+cQSaTCbPabiwWk4f6J1Gkx8xWPB6foNThDLrd7gnKEas9k32D3iONglLqkKxDG6RSqe1erxfXiufdmeyXqkZcSiQSG8JVNpsdYvU9KUAL7typ/wS9AkILeIFA4CCdTg+rdrs9STEuIrvI12q1hnA/SE9mTC/U6XQiinOFjeBxtifLMf3x4xi8Wp35iEKYtUK5XHYG1Wr1HpMFgXA7wy5qlvPstGLg3Qq8HnuRWhJDgS7MH68xa6jcQ12O0LQCW1q2/J+8aHtcal2VSiUxeDNijhXcgNFcMr2cCOz8geM09HfAWY5FNLhkcNVwl/hWNimKwMa+EG3AJV1TVIDEGH83mDRBEVyADxY5Rx8BErfczYsQbSCkUqmk2Na6cIMp8hqQLc+QbeSZXbGFMxAB1ytWmoOfAXml3+QvjOXzPqUX5eEdNBe/AAAA//9tUdN3AAAABklEQVQDAHkIoROd5Bm2AAAAAElFTkSuQmCC"

class ToolTip:
    def __init__(self, widget, text):
        self.widget = widget
        self.text = text
        self.tip_window = None
        self.widget.bind("<Enter>", self.show_tip)
        self.widget.bind("<Leave>", self.hide_tip)

    def show_tip(self, event=None):
        if self.tip_window or not self.text: return
        x, y, _cx, cy = self.widget.bbox("insert")
        x = x + self.widget.winfo_rootx() + 25
        y = y + cy + self.widget.winfo_rooty() + 25
        self.tip_window = tw = tk.Toplevel(self.widget)
        tw.wm_overrideredirect(True)
        tw.wm_geometry(f"+{x}+{y}")
        label = tk.Label(tw, text=self.text, justify='left',
                         background="#333333", foreground="#ffffff", relief='flat', borderwidth=1,
                         font=("Segoe UI", "9", "normal"), padx=5, pady=2)
        label.pack(ipadx=1)

    def hide_tip(self, event=None):
        tw = self.tip_window
        self.tip_window = None
        if tw: tw.destroy()

class ClipLiteApp:
    def __init__(self, root):
        self.root = root
        self.root.title(f"ClipLite Pro v{VERSION}")
        self.latest_version_cached = None # [ADD] アップデートチェックの結果を保持する変数
        
        # アイコンの適用 (ウィンドウ　& タスクバー)
        try:
            # 1. 小・中アイコン用設定
            self.root.iconbitmap(resource_path(ICON_FILE))
            # 2. タスクバー・タスク切替用 (iconphoto)
            icon_img = ImageTk.PhotoImage(Image.open(resource_path(ICON_FILE)))
            self.root.iconphoto(True, icon_img)
        except Exception as e:
            print(f"Icon Load Error: {e}")
            pass

        # 設定の読み込み
        self.config = self.load_config()

        # デフォルト保存先設定
        default_pictures_path = os.path.join(os.path.expandvars("%USERPROFILE%"), "Pictures", "ClipLite_Exports")
        
        self.save_mode = tk.StringVar(value=self.config.get("save_mode", "local"))
        self.local_path = tk.StringVar(value=self.config.get("local_path", default_pictures_path))
        self.gdrive_path = tk.StringVar(value=self.config.get("gdrive_path", "G:/マイドライブ/images"))
        self.start_hidden = tk.BooleanVar(value=self.config.get("start_hidden", False))
        self.auto_fallback = tk.BooleanVar(value=self.config.get("auto_fallback", True))
        # 重複抑制設定（デフォルト5秒）
        self.save_interval = tk.IntVar(value=self.config.get("save_interval", 5))
        self.original_size_mode = tk.BooleanVar(value=self.config.get("original_size_mode", False)) # [ADD] オリジナルサイズ優先フラグ
        # [ADD] アップデートダイアログの管理用
        # ※これはGUI部品に紐付けないので、通常の辞書操作（self.config）だけでも動きますが、
        # 構造を明示的にするためにここで初期値を確定させておくと安全です。
        if "last_update_dialog_date" not in self.config:
            self.config["last_update_dialog_date"] = ""

        # ウィンドウ初期設定
        self.window_width = 320
        self.window_height = 180
        pos_x = self.config.get("pos_x", 100)
        pos_y = self.config.get("pos_y", 100)
        self.root.geometry(f"{self.window_width}x{self.window_height}+{pos_x}+{pos_y}")
        self.root.configure(bg=DARK_BG)
        self.set_dark_title_bar(self.root)

        # ✕ボタンは「隠す」
        self.root.protocol("WM_DELETE_WINDOW", self.hide_window)

        # --- 画像リソースの準備と着色 ---
        def get_colored_icon(b64_data, color_rgb):
            img_bytes = base64.b64decode(b64_data)
            img = Image.open(BytesIO(img_bytes)).convert("RGBA")
            # data = img.getdata() 2027/10に廃止予定のため下記コードへ修正
            # 新しい関数があれば使い、なければ古い方を使う
            data = img.get_flattened_data() if hasattr(img, 'get_flattened_data') else img.getdata()
            new_data = []
            for item in data:
                if item[3] > 0:
                    new_data.append((color_rgb[0], color_rgb[1], color_rgb[2], item[3]))
                else:
                    new_data.append(item)
            img.putdata(new_data)
            return ImageTk.PhotoImage(img)

        self.help_img = get_colored_icon(iCON_HELP, (120, 220, 120))     # 緑
        self.setting_img = get_colored_icon(ICON_SETTING, (180, 180, 180)) # 銀
        self.exit_img = get_colored_icon(ICON_POWER, (255, 107, 107))      # 朱

        # --- UIレイアウト ---
        # 操作ボタンコンテナ
        self.btn_container = tk.Frame(root, bg=DARK_BG)
        self.btn_container.place(relx=1.0, y=5, anchor="ne", x=-5)

        # 1. ヘルプアイコン（❔）
        self.help_btn = tk.Button(self.btn_container, image=self.help_img,
                                   bg=DARK_BG, activebackground=STATUS_BG, 
                                   relief="flat", bd=0, command=self.show_help, cursor="hand2")
        self.help_btn.pack(side="left", padx=5)
        ToolTip(self.help_btn, "アプリの使いかた")

        # 2. 設定アイコン（⚙）
        self.settings_btn = tk.Button(self.btn_container, image=self.setting_img,
                                      bg=DARK_BG, activebackground=STATUS_BG, 
                                      relief="flat", bd=0, command=self.open_options, cursor="hand2")
        self.settings_btn.pack(side="left", padx=5)
        ToolTip(self.settings_btn, "保存先・自動起動の設定")

        # 3. 常駐解除アイコン（⏻）
        self.exit_btn = tk.Button(self.btn_container, image=self.exit_img,
                                   bg=DARK_BG, activebackground=STATUS_BG, 
                                   relief="flat", bd=0, command=self.quit_app, cursor="hand2")
        self.exit_btn.pack(side="left", padx=5)
        ToolTip(self.exit_btn, "監視を完全に終了")

        self.label = tk.Label(root, text="ClipLite Optimizer", font=("Segoe UI", 12, "bold"), bg=DARK_BG, fg=DARK_FG)
        self.label.pack(pady=(35, 0))

        self.info_label = tk.Label(root, text="System Ready", font=("Segoe UI", 9), bg=DARK_BG, fg="#888888")
        self.info_label.pack(pady=(0, 5))

        self.status_frame = tk.Frame(root, bg=STATUS_BG, bd=1, relief="flat")
        self.status_frame.pack(fill="both", expand=True, padx=15, pady=(5, 15))

        self.status_label = tk.Label(self.status_frame, text="Monitoring Active", font=("Segoe UI", 9), bg=STATUS_BG, fg="#aaaaaa")
        self.status_label.pack(expand=True)

        # 内部処理用
        self.last_hash = None
        self.last_save_time = 0  # 重複抑制用タイマー
        self.processing_lock = threading.Lock()
        self.task_queue = Queue()

        windnd.hook_dropfiles(self.root, func=self.on_drop)
        
        threading.Thread(target=self.monitor_loop, daemon=True).start()
        threading.Thread(target=self.worker_loop, daemon=True).start()

        # 起動プロセスの最後にチェックを開始
        self.setup_tray()
        # アップデートを確認
        self.check_for_updates() # [ADD]

        if self.start_hidden.get():
            self.root.withdraw()

    # --- 自動アップデートチェック関数 pre-releaseにも対応出来るようにURLを分けて実装 2026/04/17 ---
    def check_for_updates(self):
        """GitHub Releasesから最新バージョンを確認する"""
        def _check():
            try:
                # [ADD] 認証ヘッダーの準備 2026/04/17
                headers = {}
                if GITHUB_TOKEN:
                    headers["Authorization"] = f"token {GITHUB_TOKEN}"

                # [ADD] プレリリース許可設定に応じてURLを切り替え 2026/04/17
                target_url = API_URL_DEV if self.allow_prerelease.get() else API_URL

                response = requests.get(target_url, headers=headers, timeout=10)
                if response.status_code == 200:
                    data = response.json()
                    
                    # [MOD] URLによってレスポンス形式が異なる（DEVはリスト、Latestは辞書）
                    latest_release = data[0] if isinstance(data, list) else data
                    v = latest_release.get("tag_name", "").replace("v", "") 
                    self.latest_version_cached = v

                    if v > VERSION:
                        last_check_str = self.config.get("last_update_dialog_date", "")
                        today_str = datetime.now().strftime("%Y-%m-%d")
                        if last_check_str != today_str:
                            # メインスレッドでダイアログを表示
                            self.root.after(0, lambda: self.ask_update_dialog(v, today_str))
                else:
                    self.latest_version_cached = "Offline"
            except Exception as e:
                print(f"[DEBUG] Update check failed: {e}")
                self.latest_version_cached = "Error"

        threading.Thread(target=_check, daemon=True).start()

    # --- 自動更新ロジック関数
    def perform_update(self):
        """最新のEXEをダウンロードして置換を実行する"""
        try:
            # [ADD] 認証ヘッダーの準備 2026/04/17
            headers = {}
            if GITHUB_TOKEN:
                headers["Authorization"] = f"token {GITHUB_TOKEN}"

            # 1. 最新リリースの情報を再取得
            # [MOD] 更新実行時も現在の設定に従って最新情報を取得するようにURLを切り替え 2026/04/17
            target_url = API_URL_DEV if self.allow_prerelease.get() else API_URL
            response = requests.get(target_url, headers=headers, timeout=10)
            response.raise_for_status()
            data = response.json()
            # [MOD] URLによってレスポンス形式が異なる（DEVはリスト、Latestは辞書）
            latest_release = data[0] if isinstance(data, list) else data
        
            # 2. Assetsの中から「.exe」ファイルを探す
            download_url = None
            for asset in latest_release.get("assets", []):
                if asset["name"].endswith(".exe"):
                    download_url = asset["browser_download_url"]
                    break

            if not download_url:
                tk.messagebox.showerror("Error", "リリースにEXEファイルが見つかりません。")
                return

            # 3. 一時フォルダにダウンロード
            # [MOD] EXEファイルのダウンロード本体に headers を追加 2026/04/17
            temp_exe = os.path.join(os.environ['TEMP'], "ClipLite_new.exe")
            exe_data = requests.get(download_url, headers=headers, timeout=30)
            with open(temp_exe, "wb") as f:
                f.write(exe_data.content)

            # 4. 現在実行中のパス
            dest_exe = sys.executable 

            # 5. 自己消滅 & 置換バッチの作成 (既存ロジックを活用)
            batch_path = os.path.join(os.environ['TEMP'], "cliplite_updater.bat")
            with open(batch_path, "w", encoding="shift-jis") as f:
                f.write(f'@echo off\n')
                f.write(f'timeout /t 2 /nobreak > nul\n')
                # temp_exe を dest_exe に上書きコピー
                f.write(f'copy /y "{temp_exe}" "{dest_exe}"\n')
                f.write(f'start "" "{dest_exe}"\n')
                f.write(f'del "{temp_exe}"\n') # 一時ファイル削除
                f.write(f'del "%~f0"\n')

            subprocess.Popen([batch_path], shell=True)
            self.quit_app()

        except Exception as e:
            tk.messagebox.showerror("Error", f"アップデート中にエラーが発生しました: {e}")

    # アップデートトリガー(ダイアログ版)
    def ask_update_dialog(self, new_ver, today_str):
        if tk.messagebox.askyesno("Update Available", f"最新版 v{new_ver} が見つかりました。\n今すぐアップデートしますか？"):
            self.perform_update()
        else:
            # 「いいえ」の場合、今日の日付を記録して今日はもう出さないようにする
            self.config["last_update_dialog_date"] = today_str
            self.save_config()

    def move_mouse_to_widget(self, widget):
        self.root.update_idletasks()
        x = widget.winfo_rootx() + (widget.winfo_width() // 2)
        y = widget.winfo_rooty() + (widget.winfo_height() // 2)
        ctypes.windll.user32.SetCursorPos(x, y)


    def center_window(self, window, width, height):
        """ウィンドウをスクリーン中央に配置"""
        screen_width = window.winfo_screenwidth()
        screen_height = window.winfo_screenheight()
        x = (screen_width // 2) - (width // 2)
        y = (screen_height // 2) - (height // 2)
        window.geometry(f"{width}x{height}+{x}+{y}")

    def load_config(self):
        if os.path.exists(CONFIG_FILE):
            try:
                with open(CONFIG_FILE, "r", encoding="utf-8") as f: return json.load(f)
            except: return {}
        return {}

    def save_config(self):
        try:
            geo = self.root.geometry()
            parts = geo.replace('x', '+').split('+')
            self.config.update({
                "pos_x": int(parts[2]),
                "pos_y": int(parts[3]),
                "save_mode": self.save_mode.get(),
                "local_path": self.local_path.get(),
                "gdrive_path": self.gdrive_path.get(),
                "start_hidden": self.start_hidden.get(),
                "auto_fallback": self.auto_fallback.get(),
                "save_interval": self.save_interval.get(),
                "original_size_mode": self.original_size_mode.get(), # [ADD]
                "allow_prerelease": self.allow_prerelease.get(), # [ADD]
                # [ADD] 最後にアップデートダイアログを出した日付を保存
                "last_update_dialog_date": self.config.get("last_update_dialog_date", "")
            })
            with open(CONFIG_FILE, "w", encoding="utf-8") as f:
                json.dump(self.config, f, ensure_ascii=False, indent=4)
        except: pass

    def set_dark_title_bar(self, window):
        window.update()
        DWMWA_USE_IMMERSIVE_DARK_MODE = 20
        hwnd = ctypes.windll.user32.GetParent(window.winfo_id())
        rendering_policy = ctypes.c_int(1)
        ctypes.windll.dwmapi.DwmSetWindowAttribute(hwnd, DWMWA_USE_IMMERSIVE_DARK_MODE, ctypes.byref(rendering_policy), ctypes.sizeof(rendering_policy))

    def get_startup_path(self):
        shell = win32com.client.Dispatch("WScript.Shell")
        return shell.SpecialFolders("Startup")

    def is_startup_registered(self):
        path = os.path.join(self.get_startup_path(), "ClipLitePro.lnk")
        return os.path.exists(path)

    def toggle_startup(self):
        startup_path = os.path.join(self.get_startup_path(), "ClipLitePro.lnk")
        if self.is_startup_registered():
            try:
                os.remove(startup_path)
                return False
            except: return True
        else:
            try:
                shell = win32com.client.Dispatch("WScript.Shell")
                shortcut = shell.CreateShortCut(startup_path)
                target = sys.executable if getattr(sys, 'frozen', False) else os.path.abspath(sys.argv[0])
                shortcut.Targetpath = target
                shortcut.WorkingDirectory = os.path.dirname(target)
                shortcut.save()
                return True
            except: return False

    def open_current_storage(self):
        """現在有効な保存先フォルダをエクスプローラーで開く"""
        mode = self.save_mode.get()
        path = self.local_path.get() if mode == "local" else self.gdrive_path.get()
        if mode == "gdrive" and not os.path.exists(path) and self.auto_fallback.get():
            path = self.local_path.get()
        if os.path.exists(path):
            subprocess.Popen(f'explorer "{os.path.normpath(path)}"')

    def open_options(self):
        # プレリリース項目の追加に伴い高さを微調整 (460 -> 490)
        w, h = 450, 490
        opt_win = tk.Toplevel(self.root)
        opt_win.title("ClipLite Options")
        self.center_window(opt_win, w, h)
        opt_win.configure(bg=DARK_BG)
        opt_win.lift()
        opt_win.focus_force()
        self.set_dark_title_bar(opt_win)

        tk.Label(opt_win, text="システム・起動設定", font=("Segoe UI", 10, "bold"), bg=DARK_BG, fg=DARK_FG).pack(pady=(15,5))

        startup_var = tk.StringVar(value="登録解除" if self.is_startup_registered() else "Windows起動時に実行する")
        def handle_startup():
            res = self.toggle_startup()
            startup_var.set("登録解除" if res else "Windows起動時に実行する")
            btn_startup.config(fg=ACCENT_COLOR if res else DARK_FG)

        btn_startup = tk.Button(opt_win, textvariable=startup_var, command=handle_startup, 
                                bg=STATUS_BG, fg=ACCENT_COLOR if self.is_startup_registered() else DARK_FG, 
                                relief="flat", padx=10, pady=5)
        btn_startup.pack(pady=5)

        tk.Checkbutton(opt_win, text="起動時にウィンドウを表示しない", variable=self.start_hidden, 
                       bg=DARK_BG, fg=DARK_FG, selectcolor=STATUS_BG, activebackground=DARK_BG).pack(pady=5)

        # [ADD] リサイズ無効化オプションの追加
        tk.Checkbutton(opt_win, text="オリジナルサイズで保存（リサイズを無効化）", variable=self.original_size_mode, 
                       bg=DARK_BG, fg=DARK_FG, selectcolor=STATUS_BG, activebackground=DARK_BG).pack(pady=5)

        tk.Checkbutton(opt_win, text="Drive未接続時に自動ローカル保存する", variable=self.auto_fallback, 
                       bg=DARK_BG, fg=DARK_FG, selectcolor=STATUS_BG, activebackground=DARK_BG).pack(pady=2)

        # [ADD] プレリリース許可のチェックボックス 2026/04/17
        tk.Checkbutton(opt_win, text="プレリリース版の更新通知を受け取る", variable=self.allow_prerelease, 
                       bg=DARK_BG, fg=DARK_FG, selectcolor=STATUS_BG, activebackground=DARK_BG).pack(pady=2)

        f_interval = tk.Frame(opt_win, bg=DARK_BG)
        f_interval.pack(pady=10)
        tk.Label(f_interval, text="重複保存を抑制する秒数:", bg=DARK_BG, fg=DARK_FG).pack(side="left")
        sp_interval = tk.Spinbox(f_interval, from_=0, to=60, textvariable=self.save_interval, width=5, 
                                 bg=STATUS_BG, fg=DARK_FG, insertbackground=DARK_FG, relief="flat")
        sp_interval.pack(side="left", padx=10)
        tk.Label(f_interval, text="秒", bg=DARK_BG, fg=DARK_FG).pack(side="left")

        tk.Label(opt_win, text="優先保存先ターゲット:", bg=DARK_BG, fg="#888888").pack(anchor="w", padx=30, pady=(10,0))
        f_mode = tk.Frame(opt_win, bg=DARK_BG)
        f_mode.pack(fill="x", padx=30)
        tk.Radiobutton(f_mode, text="ローカル", variable=self.save_mode, value="local", bg=DARK_BG, fg=DARK_FG, selectcolor=STATUS_BG).pack(side="left")
        tk.Radiobutton(f_mode, text="Google Drive", variable=self.save_mode, value="gdrive", bg=DARK_BG, fg=DARK_FG, selectcolor=STATUS_BG).pack(side="left", padx=20)

        for label, var in [("ローカルアーカイブ:", self.local_path), ("Google Drive同期パス:", self.gdrive_path)]:
            tk.Label(opt_win, text=label, bg=DARK_BG, fg="#888888").pack(anchor="w", padx=30, pady=(10,0))
            f = tk.Frame(opt_win, bg=DARK_BG)
            f.pack(fill="x", padx=30)
            tk.Entry(f, textvariable=var, bg=STATUS_BG, fg=DARK_FG, insertbackground=DARK_FG, relief="flat").pack(side="left", fill="x", expand=True)
            tk.Button(f, text="...", command=lambda v=var: self.select_dir(v), bg=STATUS_BG, fg=DARK_FG, relief="flat").pack(side="left", padx=5)

        btn_save = tk.Button(opt_win, text="設定を保存して閉じる",
                  command=lambda: [
                      self.save_config(),
                      self.check_for_updates(), # 設定変更後に再チェック
                      opt_win.destroy(),
                      self.hide_window() if self.start_hidden.get() else None
                  ], 
                  bg=ACCENT_COLOR, fg=DARK_FG, relief="flat", pady=10, width=25)
        btn_save.pack(pady=20)
        opt_win.after(100, lambda: self.move_mouse_to_widget(btn_save))

    def select_dir(self, var):
        current_path = var.get()
        initial_dir = current_path if os.path.exists(current_path) else None
        path = filedialog.askdirectory(initialdir=initial_dir)
        if path: var.set(path)


    def setup_tray(self):
        # トレイアイコン画像に作成したアイコンを適用
        try:
            # PILで開き、サイズをリサイズしてトレイ用に最適化して読み込む
            icon_raw = Image.open(resource_path(ICON_FILE))
            icon_img = icon_raw.resize((64, 64), Image.Resampling.LANCZOS)
        except:
            # 失敗時のフォールバック
            icon_img = Image.new('RGB', (64, 64), color=(0, 120, 212))

        menu = (
            item('表示', self.show_window), 
            item('保存場所を開く', self.open_current_storage),
            item('設定', self.open_options), 
            item('ヘルプ', self.show_help), 
            item('終了', self.quit_app)
        )
        self.tray_icon = pystray.Icon("ClipLite", icon_img, "ClipLite Pro", menu)
        threading.Thread(target=self.tray_icon.run, daemon=True).start()

    def show_help(self):
        # バージョン表示を追加するため、高さを少し広げる (310 -> 330)
        w, h = 350, 330 # [MOD]
        help_win = tk.Toplevel(self.root)
        help_win.title("About")
        self.center_window(help_win, w, h)
        help_win.configure(bg=DARK_BG)
        help_win.attributes("-topmost", True)
        self.set_dark_title_bar(help_win)

        # 1. アプリ名 (中央) pady=(20, 15)) 15→0
        tk.Label(help_win, text=f"ClipLite Pro v{VERSION}", font=("Segoe UI", 10, "bold"), 
                 bg=DARK_BG, fg=DARK_FG).pack(pady=(20, 0))

        # --- [ADD] アップデートステータスの表示 ---
        status_text = "Checking for updates..."
        status_fg = "#555555" # [MOD] 少し暗めのグレーにして目立たなくする #888888→#555555
        
        if self.latest_version_cached:
            if self.latest_version_cached in ["Error", "Offline"]:
                status_text = "Update check failed (Offline)"
                status_fg = "#888888"
            elif self.latest_version_cached and self.latest_version_cached > VERSION:
                # アップデートがある場合のみ「更新ボタン」を表示
                btn_update = tk.Button(help_win, text=f"v{self.latest_version_cached} へ更新", 
                                   command=self.perform_update,
                                   bg=ACCENT_COLOR, fg=DARK_FG, relief="flat", padx=10)
                btn_update.pack(pady=(0, 10))
                status_text = f"Update available: v{self.latest_version_cached}"
                status_fg = ACCENT_COLOR
            else:
                status_text = "You are using the latest version"
                status_fg = "#888888"
                
        tk.Label(help_win, text=status_text, font=("Segoe UI", 8), 
                 bg=DARK_BG, fg=status_fg).pack(pady=(2, 10))
        # ----------------------------------------

        # 2. 機能説明 (中央寄りの左揃え / padx=40 で位置を調整)
        help_text = (
            "【主な機能】\n"
            "● 自動最適化・軽量化\n"
            "   画像を1200pxにリサイズし、256色へ軽量化。\n"
            "● 自動WebPアーカイブ\n"
            "   履歴ファイルをGoogle Drive等へ自動保存。\n"
            "● 重複保存抑制機能\n"
            f"   {self.save_interval.get()}秒以内の同一保存をスキップします。\n"
            "● 自動フォールバック保存\n"
            "   Drive未接続時はローカルへ自動保存します。"
        )
        tk.Label(help_win, text=help_text, font=("Segoe UI", 9), 
                 bg=DARK_BG, fg=DARK_FG, justify="left").pack(padx=40, anchor="w")

        # 3. 開発者情報の横並びフレーム (左端の位置を本文と合わせて padx=40)
        author_frame = tk.Frame(help_win, bg=DARK_BG)
        author_frame.pack(anchor="w", padx=40, pady=(20, 0))

        tk.Label(author_frame, text="開発者：", font=("Segoe UI", 9), 
                 bg=DARK_BG, fg=DARK_FG).pack(side="left")

        # クリッカブルなメールアドレス
        email_label = tk.Label(author_frame, text=AUTHOR_INFO, font=("Segoe UI", 9, "underline"), 
                               bg=DARK_BG, fg=ACCENT_COLOR, cursor="hand2")
        email_label.pack(side="left")
        # webbrowser モジュールを使用したURL起動 
        email_label.bind("<Button-1>", lambda e: webbrowser.open(SUPPORT_URL))

        # 4. 閉じるボタン
        btn_close = tk.Button(help_win, text="閉じる", command=help_win.destroy, 
                              bg=STATUS_BG, fg=DARK_FG, relief="flat", padx=20)
        btn_close.pack(pady=20)
        
        # マウスカーソルをボタンへ自動移動
        help_win.after(100, lambda: self.move_mouse_to_widget(btn_close))

    def hide_window(self):
        self.save_config()
        self.root.withdraw()

    def show_window(self):
        self.root.after(0, self.root.deiconify)
        self.root.after(10, self.root.focus_force)

    def quit_app(self):
        # 設定を保存
        self.save_config()
        # トレイアイコンの停止
        if hasattr(self, 'tray_icon'):
            self.tray_icon.stop()
        # メインループ(mainloop)を終了させる
        self.root.after(0, self.root.quit)

    # WebPの保存パラメータに、最高品質の圧縮アルゴリズム（method=6）を指定します。
    def save_webp_file(self, img):
        now = datetime.now()
        current_time_val = time.time()
        
        # 重複抑制チェック
        if current_time_val - self.last_save_time < self.save_interval.get():
            return None, " (Suppressed)"

        date_folder = now.strftime("%Y%m%d")

        # --- [MOD] ファイル名生成ロジックの変更 --- 2024/04/17
        if "original_filename" in img.info:
            # ドラッグ&ドロップ時：元ファイル名 + mmddhhmmss.webp
            time_suffix = now.strftime("%m%d%H%M%S")
            file_name = f"{img.info['original_filename']}_{time_suffix}.webp"
        else:
            # クリップボード監視時：clip_+mmddhhmmss.webp
            file_name = now.strftime("clip_%m%d%H%M%S.webp")
        # -----------------------------------------

        mode = self.save_mode.get()
        base_dir = self.local_path.get() if mode == "local" else self.gdrive_path.get()
        
        fallback_msg = ""
        if mode == "gdrive" and not os.path.exists(base_dir):
            if self.auto_fallback.get():
                base_dir = self.local_path.get()
                fallback_msg = " (Fallback to Local)"
            else:
                raise FileNotFoundError("Drive not mounted")

        target_dir = os.path.join(base_dir, date_folder)
        if not os.path.exists(target_dir): os.makedirs(target_dir, exist_ok=True)
        save_path = os.path.join(target_dir, file_name)

        # [ADD] カラープロファイルの抽出
        # クリップボードデータにプロファイルが含まれている場合はそれを保持する
        icc = img.info.get("icc_profile")

        # [MOD] 画質重視のエンコード設定に変更（プロファイル対応版）
        # プロファイルが存在する場合のみ icc_profile を指定して色味を維持
        if icc:
            img.save(save_path, "WEBP", quality=WEBP_QUALITY, method=WEBP_METHOD, icc_profile=icc) # [MOD]
        else:
            img.save(save_path, "WEBP", quality=WEBP_QUALITY, method=WEBP_METHOD) # [MOD]

        self.last_save_time = current_time_val
        return save_path, fallback_msg

    def worker_loop(self):
        while True:
            img = self.task_queue.get()
            if img is None: break
            with self.processing_lock:
                try:
                    # [MOD] リサイズ判定ロジック
                    # オリジナルサイズモードがOFFの場合のみリサイズを実行
                    if not self.original_size_mode.get(): # [ADD] 分岐追加
                        # 4K(3840px)以上のソース時はしきい値を1920pxに引き上げ、それ以外は1200pxとする
                        current_limit = 1920 if img.width >= 3840 else 1200 # [MOD]
                        
                        if img.width > current_limit:
                            ratio = current_limit / img.width
                            img = img.resize((current_limit, int(img.height * ratio)), Image.Resampling.LANCZOS)
                    # [ADD] else時（ONの時）は、ここをスキップしてオリジナルのまま次へ進む

                    saved_path, fallback_msg = self.save_webp_file(img)

                    # --- [ADD] 保存処理が終わったら、名前情報を破棄して初期化する --- 2026/04/17
                    if "original_filename" in img.info:
                        del img.info["original_filename"]
                    # ----------------------------------------------------------

                    # クリップボード書き戻し処理
                    img_p = img.convert("P", palette=Image.ADAPTIVE, colors=256).convert("RGB")
                    output = BytesIO()
                    img_p.save(output, "BMP")
                    data = output.getvalue()[14:]
                    output.close()
                    
                    win32clipboard.OpenClipboard()
                    try:
                        win32clipboard.EmptyClipboard()
                        win32clipboard.SetClipboardData(win32clipboard.CF_DIB, data)
                    finally: win32clipboard.CloseClipboard()

                    if saved_path:
                        msg = f"Archived in {self.save_mode.get().upper()}{fallback_msg}"
                    else:
                        msg = f"Optimized only{fallback_msg}"
                    
                    self.root.after(0, lambda m=msg: self.update_ui_success(m))
                    # ハッシュ値を更新
                    self.last_hash = hash(img_p.tobytes()[:1024])
                except Exception as e:
                    error_txt = "Drive Error" if "Drive not mounted" in str(e) else "Processing Error"
                    self.root.after(0, lambda t=error_txt: self.update_ui_error(t))
            self.task_queue.task_done()

    def update_ui_success(self, msg):
        # メインウィンドウの表示更新
        self.status_frame.configure(bg=ACCENT_COLOR)
        self.status_label.configure(bg=ACCENT_COLOR, fg=DARK_FG, text="COMPLETE")
        self.info_label.config(text=msg, fg=ACCENT_COLOR)

        # --- ウィンドウが非表示の時のみ通知を表示 ---
        if self.start_hidden.get() and hasattr(self, 'tray_icon'):
            # スレッドセーフに通知を投げる
            # 通知の中身を msg (保存先情報) だけにしてシンプルに
            self.tray_icon.notify(
                msg,
                title="ClipLite Pro: Optimized"
            )
        # ----------------------------------------------

        self.root.after(2000, self.reset_status)

    def update_ui_error(self, msg):
        self.status_frame.configure(bg="#ff6b6b")
        self.status_label.configure(bg="#ff6b6b", fg=DARK_FG, text=msg)
        self.root.after(2000, self.reset_status)

    def reset_status(self):
        self.status_frame.configure(bg=STATUS_BG)
        self.status_label.configure(bg=STATUS_BG, fg="#aaaaaa", text="Monitoring Active")

    # --- ドロップファイル処理関数 ---
    def on_drop(self, filenames):
        for fname in filenames:
            path = fname.decode('utf-8', errors='ignore')
            if path.lower().endswith(('.png', '.jpg', '.jpeg', '.bmp', '.webp')):
                try:
                    # [MOD] copy()はメタデータが落ちる可能性があるため、そのまま開いて渡す
                    # img = Image.open(path).copy() [ORG]
                    img = Image.open(path)
                    img.load()  # [ADD] ファイルを閉じる前にピクセルデータをメモリにロード

                    # --- [ADD] 元のファイル名を保持 --- 2026/04/17
                    base_name = os.path.splitext(os.path.basename(path))[0]
                    img.info["original_filename"] = base_name
                    # ----------------------------------

                    # ★修正ポイント：ハッシュ値をリセットして確実に変換を走らせる
                    self.last_hash = None
                    self.task_queue.put(img) # [MOD] ロード済みオブジェクトを渡す
                except Exception as e:
                    pass
                break

    def monitor_loop(self):
        while True:
            try:
                # キューが空の時のみ監視を行う
                if self.task_queue.empty():
                    # クリップボードが開けるかチェック
                    try:
                        # win32clipboardを使って直接DIB形式を確認する
                        win32clipboard.OpenClipboard()
                        is_image = win32clipboard.IsClipboardFormatAvailable(win32clipboard.CF_DIB)
                        win32clipboard.CloseClipboard()
                    except:
                        # 他のアプリ（フォト等）がロックしている場合はスキップして次へ
                        is_image = False

                    if is_image:
                        # フォトやSnipping Toolの書き込み完了、通知オフによる遅延を考慮し、リトライを10回(最大1秒間)に強化
                        img = None
                        for _ in range(10):
                            # ImageGrabはCF_DIBがあれば画像として取得を試みる
                            img = ImageGrab.grabclipboard()
                            if img is not None:
                                break
                            time.sleep(0.1) # 0.1秒待って再試行

                        if isinstance(img, Image.Image):
                            # アルファチャンネル(RGBA)があるとハッシュ計算や変換で
                            # 稀にコケるため、RGBに変換して扱うのが安全
                            if img.mode != 'RGB':
                                img = img.convert('RGB')
                            # --- [ADD] クリップボード由来であることを保証するため、名前情報を初期化 ---
                            if hasattr(img, "info") and "original_filename" in img.info:
                                del img.info["original_filename"]
                            # ------------------------------------------------------------------
                                                            
                            curr_hash = hash(img.tobytes()[:1024])
                            if curr_hash != self.last_hash:
                                self.task_queue.put(img.copy())
                                # last_hashの更新はworker_loopに任せるかここで行う
            except Exception as e:
                pass
            # 監視頻度を1.2秒→1.0秒に。Snipping Toolの着弾を逃さない設定
            time.sleep(1.0)

def is_already_running():
    kernel32 = ctypes.windll.kernel32
    mutex = kernel32.CreateMutexW(None, False, "Local\\ClipLite_Pro_Mutex_CenterDiag_999")
    if kernel32.GetLastError() == 183:
        d = tk.Tk(); d.withdraw()
        ctypes.windll.user32.MessageBoxW(0, "既にバックグラウンドで起動しています。", "ClipLite Pro", 0x40 | 0x1000)
        d.destroy()
        return True
    return False

if __name__ == "__main__":
    if is_already_running(): sys.exit(0)
    root = tk.Tk()
    app = ClipLiteApp(root)
    root.mainloop()

    # mainloop終了後にクリーンにプロセスを落とす
    sys.exit(0)
