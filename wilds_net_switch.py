# -*- coding: utf-8 -*-
"""
Wilds Net Switch - Monster Hunter Wilds の通信をワンクリックでON/OFF
- 緑=通信許可(オンライン) / 赤=通信ブロック(オフライン)
- クリック後はローディング表示＆連続クリック防止（完了時に必ず解除）
- exeパスは「変更」ボタンで差し替え（JSON保存）
"""

import sys, os, json, subprocess, threading, locale, tkinter as tk
from tkinter import messagebox, filedialog, ttk
import ctypes, keyboard
from typing import Callable

# ===== 設定 =====
DEFAULT_GAME_EXE = r""
DEFAULT_HOTKEY = ""
RULE_NAME = "BlockWilds_WildsOutbound"
WINDOW_TITLE = "Wilds Net Switch"
MAX_PATH_CHARS = 56
# ===============

def _app_dir() -> str:
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))

CONFIG_PATH = os.path.join(_app_dir(), "WildsNetSwitch.config.json")

def load_config() -> dict:
    try:
        if not os.path.exists(CONFIG_PATH): return {}
        with open(CONFIG_PATH, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}

def save_config(cfg: dict) -> None:
    try:
        full_cfg = load_config()
        full_cfg.update(cfg)
        with open(CONFIG_PATH, "w", encoding="utf-8") as f:
            json.dump(full_cfg, f, ensure_ascii=False, indent=2)
    except Exception:
        pass

config = load_config()
GAME_EXE = config.get("game_exe", DEFAULT_GAME_EXE)
HOTKEY = config.get("hotkey", DEFAULT_HOTKEY)

# ---- PowerShell / Firewall ----
def is_admin() -> bool:
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except Exception:
        return False

def relaunch_as_admin():
    params = " ".join([f'"{arg}"' for arg in sys.argv])
    ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, params, None, 1)
    sys.exit(0)

def _decode_bytes(b: bytes) -> str:
    if b is None:
        return ""
    try:
        return b.decode("utf-8")
    except Exception:
        try:
            return b.decode(locale.getpreferredencoding(False), errors="replace")
        except Exception:
            return b.decode("utf-8", errors="replace")

def _run_ps(cmd: str) -> subprocess.CompletedProcess:
    ps_prefix = "[Console]::OutputEncoding=[Text.Encoding]::UTF8; $OutputEncoding=[Text.Encoding]::UTF8; "
    try:
        proc = subprocess.run(
            ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", ps_prefix + cmd],
            stdout=subprocess.PIPE, stderr=subprocess.PIPE,
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0)
        )
    except FileNotFoundError as e:
        raise RuntimeError("PowerShell が見つかりません。") from e
    proc.stdout = _decode_bytes(proc.stdout)
    proc.stderr = _decode_bytes(proc.stderr)
    return proc

def rule_exists() -> bool:
    ps = f"Get-NetFirewallRule -DisplayName '{RULE_NAME}' -ErrorAction SilentlyContinue | Select-Object -First 1 | ForEach-Object {{ 'YES' }}"
    return "YES" in (_run_ps(ps).stdout or "").strip()

def rule_enabled() -> bool:
    ps = (
        f"$r = Get-NetFirewallRule -DisplayName '{RULE_NAME}' -ErrorAction SilentlyContinue; "
        f"if ($r) {{ if ($r.Enabled -eq 'True') {{ 'ENABLED' }} else {{ 'DISABLED' }} }} else {{ 'NO_RULE' }}"
    )
    out = (_run_ps(ps).stdout or "").strip().upper()
    if out == "NO_RULE": return False
    return out == "ENABLED"

def create_or_replace_block_rule() -> bool:
    delete_rule()
    ps = (
        f"New-NetFirewallRule -DisplayName '{RULE_NAME}' "
        f"-Direction Outbound -Program '{GAME_EXE}' -Action Block -Enabled True"
    )
    return _run_ps(ps).returncode == 0

def allow_online() -> bool:
    # 復帰は規則を削除して確実に戻す
    return delete_rule()

def delete_rule() -> bool:
    return _run_ps(
        f"Get-NetFirewallRule -DisplayName '{RULE_NAME}' -ErrorAction SilentlyContinue | Remove-NetFirewallRule"
    ).returncode == 0

# ---- 小物（中央省略 & ツールチップ） ----
def ellipsize_middle(text: str, max_chars: int) -> str:
    if len(text) <= max_chars: return text
    keep = max_chars - 3
    left = keep // 2
    right = keep - left
    return text[:left] + "..." + text[-right:]

class Tooltip:
    def __init__(self, widget, text: str, delay_ms: int = 500):
        self.widget, self.text, self.delay = widget, text, delay_ms
        self.tip, self._id = None, None
        widget.bind("<Enter>", self._schedule); widget.bind("<Leave>", self._hide); widget.bind("<ButtonPress>", self._hide)
    def set_text(self, text: str): self.text = text
    def _schedule(self, _): self._cancel(); self._id = self.widget.after(self.delay, self._show)
    def _show(self):
        if self.tip: return
        x = self.widget.winfo_rootx() + 10; y = self.widget.winfo_rooty() + self.widget.winfo_height() + 6
        self.tip = tw = tk.Toplevel(self.widget); tw.wm_overrideredirect(True); tw.wm_geometry(f"+{x}+{y}")
        tk.Label(tw, text=self.text, justify="left", background="#FFFFE0",
                 relief="solid", borderwidth=1, font=("Segoe UI", 9)).pack(ipadx=6, ipady=3)
    def _hide(self, _=None): self._cancel()
    def _cancel(self):
        if self._id: self.widget.after_cancel(self._id); self._id = None
        if self.tip: self.tip.destroy(); self.tip = None

# ---- スイッチUI ----
class SwitchButton(tk.Canvas):
    """
    True = ON(緑: 許可/オンライン)
    False= OFF(赤: ブロック/オフライン)
    """
    def __init__(self, master, width=132, height=70, command: Callable[[bool], None] | None = None, initial=False):
        super().__init__(master, width=width, height=height, highlightthickness=0, bg=master["bg"])
        self.w, self.h, self.r = width, height, height // 2
        self.command, self.state = command, initial
        self.on_bg = "#34C759"; self.off_bg = "#FF4D4F"
        self.knob_color = "#FFFFFF"; self.shadow = "#D0D3D8"
        self.track_items = []; self.knob = None; self.knob_shadow = None
        self._anim_running = False; self._target_x = 0.0; self._disabled = False
        self.bind("<Button-1>", self._on_click)
        self._redraw(initial_layout=True)

    def set_state(self, on: bool, animate=False):
        self.state = on
        self._redraw()
        if animate: self._animate_to(self._knob_x_for_state(on))

    def set_disabled(self, disabled: bool):
        self._disabled = disabled
        try:
            if self.knob: self.itemconfig(self.knob, outline=("#B0B0B0" if disabled else "#E6E6E6"))
        except Exception: pass

    def _clear(self):
        for it in self.track_items: self.delete(it)
        self.track_items.clear()
        if self.knob_shadow: self.delete(self.knob_shadow); self.knob_shadow = None
        if self.knob: self.delete(self.knob); self.knob = None

    def _rounded_track(self, fill):
        r, w = self.r, self.w
        return [
            self.create_oval(2, 2, 2+2*r, 2+2*r, fill=fill, outline=fill),
            self.create_oval(w-(2+2*r), 2, w-2, 2+2*r, fill=fill, outline=fill),
            self.create_rectangle(2+r, 2, w-(2+r), 2+2*r, fill=fill, outline=fill),
        ]

    def _knob_x_for_state(self, on: bool) -> float:
        left = 2 + self.r; right = self.w - (2 + self.r)
        return right if on else left

    def _redraw(self, initial_layout=False):
        self._clear()
        bg = self.on_bg if self.state else self.off_bg
        self.track_items = self._rounded_track(bg)
        kr = self.r - 4; cx = self._knob_x_for_state(self.state); cy = self.h / 2
        self.knob_shadow = self.create_oval(cx-kr, cy-kr+2, cx+kr, cy+kr+2, fill=self.shadow, outline=self.shadow)
        self.knob = self.create_oval(cx-kr, cy-kr, cx+kr, cy+kr, fill=self.knob_color, outline="#E6E6E6")
        if initial_layout: self.update_idletasks()

    def _animate_to(self, target_x: float):
        if self._anim_running: return
        self._anim_running = True; self._target_x = target_x; self._step_animation()

    def _step_animation(self):
        try:
            if not self.knob: self._anim_running = False; return
            x0, _, x1, _ = self.coords(self.knob); cx = (x0+x1)/2; dx = self._target_x - cx
            if abs(dx) < 0.5: self._move_knob(self._target_x - cx); self._anim_running = False; return
            self._move_knob(dx * 0.35); self.after(16, self._step_animation)
        except Exception: self._anim_running = False

    def _move_knob(self, dx: float):
        self.move(self.knob, dx, 0)
        if self.knob_shadow: self.move(self.knob_shadow, dx, 0)

    def _on_click(self, _):
        if self._disabled: return
        if callable(self.command): self.command(not self.state)

# ---- メインアプリ ----
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(WINDOW_TITLE); self.resizable(False, False)
        w, h = 460, 350
        self.geometry(f"{w}x{h}+{(self.winfo_screenwidth()-w)//2}+{(self.winfo_screenheight()-h)//2}")
        self.configure(bg="#FFFFFF")

        tk.Label(self, text="Monster Hunter Wilds 通信スイッチ", font=("Segoe UI", 12, "bold"), bg="#FFFFFF").pack(pady=(16,6))
        self.label_status = tk.Label(self, text="状態取得中...", font=("Segoe UI", 16), bg="#FFFFFF")
        self.label_status.pack(pady=(6,10))

        allowed = True
        try:
            blocked = rule_exists() and rule_enabled(); allowed = not blocked
        except Exception as e:
            self.label_status.config(text=f"状態取得エラー: {e}", fg="#d32f2f")

        self.switch = SwitchButton(self, width=132, height=70, command=self.on_switch_clicked, initial=allowed)
        self.switch.pack(pady=8)

        self.loading_var = tk.StringVar(value="")
        tk.Label(self, textvariable=self.loading_var, font=("Segoe UI", 9), fg="#666", bg="#FFFFFF").pack(pady=(4,2))
        self.progress = ttk.Progressbar(self, mode="indeterminate", length=180)

        pf = tk.Frame(self, bg="#FFFFFF"); pf.pack(pady=(10,6), padx=14, fill="x")
        tk.Label(pf, text="対象:", font=("Segoe UI", 9, "bold"), fg="#333", bg="#FFFFFF").pack(side="left")
        self.label_path = tk.Label(pf, text=ellipsize_middle(GAME_EXE, MAX_PATH_CHARS),
                                   font=("Segoe UI", 9), fg="#555", bg="#FFFFFF", anchor="w")
        self.label_path.pack(side="left", fill="x", expand=True, padx=(4,0))
        self.path_tip = Tooltip(self.label_path, GAME_EXE)
        tk.Button(pf, text="変更", font=("Segoe UI", 9), command=self.on_change_path).pack(side="right", padx=(6,0))

        # --- ホットキー設定 ---
        hf = tk.Frame(self, bg="#FFFFFF")
        hf.pack(pady=(4,0), padx=14, fill="x")
        tk.Label(hf, text="ショートカット:", font=("Segoe UI", 9, "bold"), fg="#333", bg="#FFFFFF").pack(side="left")
        self.label_hotkey = tk.Label(hf, text=HOTKEY or "未設定", font=("Segoe UI", 9), fg="#555", bg="#FFFFFF")
        self.label_hotkey.pack(side="left", padx=4)
        self.btn_change_hotkey = tk.Button(hf, text="変更", font=("Segoe UI", 9), command=self.on_change_hotkey)
        self.btn_change_hotkey.pack(side="left", padx=6)
        # --------------------

        tk.Label(self, text="緑＝通信許可（オンライン） / 赤＝通信ブロック（オフライン）\nスペースキーでも切替できます",
                 font=("Segoe UI", 9), fg="#666", bg="#FFFFFF", justify="center").pack(pady=(2,8))

        self.bind("<space>", lambda e: self.on_switch_clicked(not self.switch.state))
        self._set_ui_text(allowed=allowed)
        self._busy = False  # ← 連打防止フラグ

        self.protocol("WM_DELETE_WINDOW", self.on_closing)
        self.start_hotkey_listener()

    def on_closing(self):
        keyboard.remove_all_hotkeys()
        self.destroy()

    def toggle_switch(self):
        # UIスレッドで安全に実行するために after を使う
        if self._busy: return
        self.after(0, self.on_switch_clicked, not self.switch.state)

    def start_hotkey_listener(self):
        global HOTKEY
        keyboard.remove_all_hotkeys()
        if HOTKEY:
            try:
                keyboard.add_hotkey(HOTKEY, self.toggle_switch)
                print(f"[*] ホットキー '{HOTKEY}' を登録しました。")
            except Exception as e:
                print(f"[!] ホッ​​トキーの登録に失敗しました: {e}")

    def _set_ui_text(self, allowed: bool, note: str = ""):
        if allowed:
            self.label_status.config(text="通信：許可中（オンライン）" + (f" / {note}" if note else ""), fg="#2e7d32")
        else:
            self.label_status.config(text="通信：ブロック中（オフライン）" + (f" / {note}" if note else ""), fg="#d32f2f")

    def _start_loading(self, msg: str):
        self._busy = True
        self.switch.set_disabled(True)
        self.loading_var.set(msg)
        self.progress.pack(pady=(0,4))
        self.progress.start(10)
        self.config(cursor="watch")

    def _stop_loading(self):
        # ★重要：ロック解除をここで必ず実施
        self._busy = False
        self.switch.set_disabled(False)
        self.progress.stop()
        self.progress.pack_forget()
        self.loading_var.set("")
        self.config(cursor="")

    def _refresh_path_label(self):
        self.label_path.config(text=ellipsize_middle(GAME_EXE, MAX_PATH_CHARS))
        self.path_tip.set_text(GAME_EXE)

    def on_change_path(self):
        global GAME_EXE
        path = filedialog.askopenfilename(title="MonsterHunterWilds.exe を選択",
                                          filetypes=[("実行ファイル","*.exe"), ("すべてのファイル","*.*")])
        if not path: return
        GAME_EXE = path; save_config({"game_exe": GAME_EXE}); self._refresh_path_label()
        try:
            if rule_exists() and rule_enabled():
                if not create_or_replace_block_rule():
                    messagebox.showwarning("注意", "新しいパスでブロックルールの再作成に失敗しました。")
        except Exception as e:
            messagebox.showwarning("注意", f"ルール更新中にエラーが発生しました。\n{e}")

    def on_change_hotkey(self):
        global HOTKEY
        
        self.btn_change_hotkey.config(state="disabled")

        win = tk.Toplevel(self)
        win.title("キー設定")
        win.transient(self)
        win.grab_set()
        win.geometry(f"300x100+{self.winfo_x()+80}+{self.winfo_y()+100}")
        win.resizable(False, False)
        win.configure(bg="#FAFAFA")

        lbl = tk.Label(win, text="設定したいショートカットキーを押してください...\n(Escでキャンセル)", font=("Segoe UI", 10), bg="#FAFAFA")
        lbl.pack(pady=20, padx=20)

        def cleanup():
            self.btn_change_hotkey.config(state="normal")
            if win.winfo_exists():
                win.grab_release()
                win.destroy()

        def record_key():
            try:
                new_hotkey = keyboard.read_hotkey(suppress=False)
                print(f"[*] キーを検出: {new_hotkey}")
                
                if new_hotkey == 'esc':
                    cleanup()
                    return

                HOTKEY = new_hotkey
                save_config({"hotkey": HOTKEY})
                self.label_hotkey.config(text=HOTKEY or "未設定")
                self.start_hotkey_listener()
            except Exception as e:
                print(f"[!] キーの読み取りに失敗: {e}")
            finally:
                cleanup()

        win.protocol("WM_DELETE_WINDOW", cleanup)
        win.after(200, record_key)


    def on_switch_clicked(self, want_allowed: bool):
        if self._busy: return
        if not GAME_EXE.lower().endswith(".exe"):
            messagebox.showerror("設定エラー", "対象の実行ファイルが .exe ではありません。『変更』から選択してください。")
            return

        self._start_loading("処理中…数秒お待ちください")

        def worker():
            ok, err = True, None
            try:
                if want_allowed:
                    ok = allow_online()            # 復帰は削除で確実に
                else:
                    ok = create_or_replace_block_rule()
            except Exception as e:
                ok, err = False, e

            def finish():
                # どんな結果でも必ずロック解除
                self._stop_loading()
                # 実状態で確定表示
                try:
                    real_blocked = rule_exists() and rule_enabled()
                    allowed_now = not real_blocked
                except Exception as e2:
                    messagebox.showwarning("注意", f"状態取得に失敗しました。\n{e2}")
                    allowed_now = want_allowed
                if not ok:
                    messagebox.showerror("エラー", f"切替に失敗しました。\n{err}" if err else "切替に失敗しました。")
                self.switch.set_state(allowed_now, animate=True)
                self._set_ui_text(allowed=allowed_now)

            self.after(0, finish)

        threading.Thread(target=worker, daemon=True).start()

# ---- 起動 ----
def main():
    try:
        if not is_admin():
            relaunch_as_admin(); return
    except Exception:
        return
    try:
        subprocess.run(
            ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", "$PSVersionTable.PSVersion.Major"],
            stdout=subprocess.PIPE, stderr=subprocess.PIPE,
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0)
        )
    except FileNotFoundError:
        messagebox.showerror("エラー", "PowerShell が見つかりません。"); return
    App().mainloop()

if __name__ == "__main__":
    main()
