import customtkinter as ctk
import win32api
import win32con
import threading
import time
import ctypes
import json
import os
import sys
import subprocess
import re
import webbrowser
import requests
from tkinter import messagebox
from dataclasses import dataclass, asdict, field
from typing import List, Optional, Tuple, Dict

from PIL import Image, ImageDraw
import pystray
from pystray import MenuItem as item

# ============================================================
# Omen Hz Controller Pro (AC/Battery: Hz + Power Plan + CPU Policy)
# - Plan seçimi GUID ile
# - CPU policy: EPP / Boost / Min/Max CPU / Core parking / Cooling
# - Planları yenile, standart planları geri yükle, plan kopyala oluştur
# - Tray: Göster / Şimdi Uygula / Çıkış
# ============================================================

APP_NAME = "Omen Hz Controller Pro"
MONITOR_INTERVAL_SEC = 3
APP_VERSION = "1.0"
PROGRAMS_URL = "https://www.saydut.com/static/programs.json"
PROGRAM_ID = "pc-performans-ayarlayici"

CONFIG_DIR = os.path.join(os.environ.get("APPDATA", os.path.expanduser("~")), "OmenHzController")
CONFIG_PATH = os.path.join(CONFIG_DIR, "config.json")


def ensure_config_dir() -> None:
    os.makedirs(CONFIG_DIR, exist_ok=True)


def _semver_tuple(v: str) -> tuple[int, int, int]:
    v = (v or "").strip()
    if v.startswith("v"):
        v = v[1:]
    parts = v.split(".")
    out = []
    for i in range(3):
        try:
            out.append(int(parts[i]))
        except Exception:
            out.append(0)
    return tuple(out)  # type: ignore


def _http_get_json(url: str, timeout: int = 15) -> dict:
    response = requests.get(url, timeout=timeout)
    response.raise_for_status()
    return response.json()


# ============================================================
# Windows Battery Status
# ============================================================
class PowerStatus(ctypes.Structure):
    _fields_ = [
        ("ACLineStatus", ctypes.c_byte),
        ("BatteryFlag", ctypes.c_byte),
        ("BatteryLifePercent", ctypes.c_byte),
        ("Reserved1", ctypes.c_byte),
        ("BatteryLifeTime", ctypes.c_int),
        ("BatteryFullLifeTime", ctypes.c_int),
    ]


def is_plugged_in() -> Optional[bool]:
    """
    True  -> prizde
    False -> bataryada
    None  -> okunamadı
    """
    try:
        status = PowerStatus()
        ok = ctypes.windll.kernel32.GetSystemPowerStatus(ctypes.byref(status))
        if not ok:
            return None
        if status.ACLineStatus == 1:
            return True
        if status.ACLineStatus == 0:
            return False
        return None
    except Exception:
        return None


# ============================================================
# Display / Hz
# ============================================================
def get_primary_device_name() -> Optional[str]:
    try:
        device = win32api.EnumDisplayDevices(None, 0)
        return device.DeviceName
    except Exception:
        return None


def get_current_settings():
    try:
        dev = get_primary_device_name()
        if not dev:
            return None
        return win32api.EnumDisplaySettings(dev, win32con.ENUM_CURRENT_SETTINGS)
    except Exception:
        return None


def get_current_hz() -> Optional[int]:
    s = get_current_settings()
    if not s:
        return None
    try:
        return int(s.DisplayFrequency)
    except Exception:
        return None


def list_supported_hz_for_current_mode() -> List[int]:
    """
    Mevcut çözünürlük + bpp için desteklenen Hz enumerates.
    Strict boş dönerse loose fallback uygular (bazı driver'larda şart).
    """
    dev = get_primary_device_name()
    cur = get_current_settings()
    if not dev or not cur:
        return []

    target_w, target_h = int(cur.PelsWidth), int(cur.PelsHeight)
    target_bpp = int(cur.BitsPerPel)

    freqs_strict = set()
    freqs_loose = set()

    i = 0
    while True:
        try:
            s = win32api.EnumDisplaySettings(dev, i)
        except Exception:
            break

        try:
            w, h, bpp, f = int(s.PelsWidth), int(s.PelsHeight), int(s.BitsPerPel), int(s.DisplayFrequency)
            if 20 <= f <= 500:
                if w == target_w and h == target_h:
                    freqs_loose.add(f)
                    if bpp == target_bpp:
                        freqs_strict.add(f)
        except Exception:
            pass

        i += 1

    out = sorted(freqs_strict) if freqs_strict else sorted(freqs_loose)
    return out


def set_hz(hz: int) -> Tuple[bool, str]:
    try:
        dev = get_primary_device_name()
        if not dev:
            return False, "Display device bulunamadı."

        cur = win32api.EnumDisplaySettings(dev, win32con.ENUM_CURRENT_SETTINGS)
        current_hz = int(cur.DisplayFrequency)

        hz = int(hz)
        if current_hz == hz:
            return True, "Zaten bu Hz'de."

        supported = list_supported_hz_for_current_mode()
        if supported and hz not in supported:
            return False, f"{hz} Hz desteklenmiyor. Desteklenen: {supported}"

        cur.DisplayFrequency = hz
        win32api.ChangeDisplaySettings(cur, 0)
        return True, "OK"
    except Exception as e:
        return False, f"Hata: {e}"


# ============================================================
# powercfg helpers
# ============================================================
def _run_powercfg(args: List[str]) -> Tuple[int, str]:
    """
    returns: (returncode, combined_output)
    """
    try:
        p = subprocess.run(
            ["powercfg", *args],
            check=False,
            capture_output=True,
            text=True,
            creationflags=subprocess.CREATE_NO_WINDOW if hasattr(subprocess, "CREATE_NO_WINDOW") else 0,
        )
        out = (p.stdout or "") + "\n" + (p.stderr or "")
        return p.returncode, out.strip()
    except Exception as e:
        return 1, str(e)


def list_power_schemes() -> List[Tuple[str, str, bool]]:
    """
    returns: [(guid, name, is_active), ...]
    """
    rc, txt = _run_powercfg(["/list"])
    if rc != 0:
        return []

    schemes: List[Tuple[str, str, bool]] = []
    guid_re = re.compile(r"Power Scheme GUID:\s*([0-9a-fA-F\-]{36})\s*\((.*?)\)\s*(\*)?")
    for line in txt.splitlines():
        m = guid_re.search(line)
        if not m:
            continue
        guid = m.group(1).strip()
        name = m.group(2).strip()
        is_active = bool(m.group(3))
        schemes.append((guid, name, is_active))
    return schemes


def get_active_power_scheme() -> Tuple[Optional[str], Optional[str]]:
    schemes = list_power_schemes()
    for guid, name, active in schemes:
        if active:
            return guid, name
    return None, None


def set_power_scheme_by_guid(guid: str) -> Tuple[bool, str]:
    if not guid:
        return False, "GUID boş."
    rc, out = _run_powercfg(["/setactive", guid])
    if rc != 0:
        return False, out or "powercfg /setactive başarısız."
    active_guid, _ = get_active_power_scheme()
    if (active_guid or "").lower() == guid.lower():
        return True, "OK"
    return False, "Plan değişti doğrulanamadı (OEM kısıtı olabilir)."


def restore_default_power_schemes() -> Tuple[bool, str]:
    rc, out = _run_powercfg(["-restoredefaultschemes"])
    if rc != 0:
        return False, out or "restoredefaultschemes başarısız (Admin gerekli olabilir)."
    return True, "Standart planlar geri yüklendi."


def duplicate_scheme(base_guid: str) -> Tuple[Optional[str], str]:
    """
    powercfg -duplicatescheme <guid> -> çıktıda yeni GUID olur.
    """
    if not base_guid:
        return None, "Base GUID boş."
    rc, out = _run_powercfg(["-duplicatescheme", base_guid])
    if rc != 0:
        return None, out or "duplicatescheme başarısız (Admin gerekli olabilir)."
    m = re.search(r"([0-9a-fA-F\-]{36})", out)
    if not m:
        return None, f"Yeni GUID parse edilemedi: {out}"
    return m.group(1), "OK"


def change_scheme_name(guid: str, name: str, description: str = "") -> Tuple[bool, str]:
    if not guid:
        return False, "GUID boş."
    rc, out = _run_powercfg(["-changename", guid, name, description])
    if rc != 0:
        return False, out or "changename başarısız."
    return True, "OK"


# ============================================================
# CPU power settings GUIDs
# ============================================================
SUB_PROCESSOR = "54533251-82be-4824-96c1-47b60b740d00"

# EPP (0..100)
PERFEPP = "36687f9e-e3a5-4dbf-b1dc-15eb381c6863"

# Boost mode (0 disabled, 1 enabled, 2 aggressive)
PERFBOOSTMODE = "be337238-0d82-4146-a960-4f3749d470c7"

# Min/Max CPU state (%)
PROCTHROTTLEMIN = "893dee8e-2bef-41e0-89c6-b55d0929964c"
PROCTHROTTLEMAX = "bc5038f7-23e0-4960-96da-33abaf5935ec"

# Core parking min cores (%)
CPMINCORES = "0cc5b647-c1df-4637-891a-dec35c318583"

# Cooling policy (0 passive, 1 active)
SYSTEMCOOLINGPOLICY = "94d3a615-a899-4ac5-ae2b-e4d8f634367f"


def _set_value_index(ac: bool, scheme: str, subgroup: str, setting: str, value: int) -> Tuple[bool, str]:
    cmd = "/setacvalueindex" if ac else "/setdcvalueindex"
    rc, out = _run_powercfg([cmd, scheme, subgroup, setting, str(int(value))])
    if rc != 0:
        return False, out or f"{cmd} başarısız."
    return True, "OK"


def apply_cpu_policy_to_scheme(
    scheme_guid: str,
    plugged: bool,
    epp: int,
    boost_mode: int,
    cpu_min: int,
    cpu_max: int,
    core_parking_min: int,
    cooling_policy: int,
) -> Tuple[bool, str]:
    """
    plugged=True  -> AC values
    plugged=False -> DC values
    """
    if not scheme_guid:
        return False, "Plan GUID seçili değil."

    # clamp
    epp = max(0, min(100, int(epp)))
    boost_mode = max(0, min(2, int(boost_mode)))
    cpu_min = max(0, min(100, int(cpu_min)))
    cpu_max = max(1, min(100, int(cpu_max)))
    if cpu_min > cpu_max:
        cpu_min = cpu_max

    core_parking_min = max(0, min(100, int(core_parking_min)))
    cooling_policy = 1 if int(cooling_policy) == 1 else 0

    is_ac = bool(plugged)

    ok, msg = _set_value_index(is_ac, scheme_guid, SUB_PROCESSOR, PERFEPP, epp)
    if not ok:
        return False, f"EPP ayarlanamadı: {msg}"

    # Boost (bazı cihazlarda yok)
    _set_value_index(is_ac, scheme_guid, SUB_PROCESSOR, PERFBOOSTMODE, boost_mode)

    ok, msg = _set_value_index(is_ac, scheme_guid, SUB_PROCESSOR, PROCTHROTTLEMIN, cpu_min)
    if not ok:
        return False, f"Min CPU ayarlanamadı: {msg}"

    ok, msg = _set_value_index(is_ac, scheme_guid, SUB_PROCESSOR, PROCTHROTTLEMAX, cpu_max)
    if not ok:
        return False, f"Max CPU ayarlanamadı: {msg}"

    # Core parking (bazı cihazlarda yok)
    _set_value_index(is_ac, scheme_guid, SUB_PROCESSOR, CPMINCORES, core_parking_min)

    # Cooling policy (bazı cihazlarda yok)
    _set_value_index(is_ac, scheme_guid, SUB_PROCESSOR, SYSTEMCOOLINGPOLICY, cooling_policy)

    # uygula
    _run_powercfg(["/setactive", scheme_guid])
    return True, "OK"


# ============================================================
# Config
# ============================================================
@dataclass
class CpuPolicy:
    epp: int = 0                 # 0 perf, 100 saving
    boost_mode: int = 1          # 0 off, 1 on, 2 aggressive
    cpu_min: int = 5             # %
    cpu_max: int = 100           # %
    core_parking_min: int = 100  # %
    cooling_policy: int = 1      # 1 active, 0 passive


def default_cpu_ac() -> CpuPolicy:
    return CpuPolicy(epp=0, boost_mode=2, cpu_min=5, cpu_max=100, core_parking_min=100, cooling_policy=1)


def default_cpu_bat() -> CpuPolicy:
    return CpuPolicy(epp=90, boost_mode=0, cpu_min=5, cpu_max=70, core_parking_min=20, cooling_policy=0)


@dataclass
class AppConfig:
    auto_mode: bool = True
    set_power_plan: bool = True
    set_cpu_policy: bool = True

    ac_hz: Optional[int] = None
    battery_hz: Optional[int] = None

    ac_plan_guid: Optional[str] = None
    battery_plan_guid: Optional[str] = None

    cpu_ac: CpuPolicy = field(default_factory=default_cpu_ac)      # FIX: default_factory
    cpu_bat: CpuPolicy = field(default_factory=default_cpu_bat)    # FIX: default_factory


def _dict_to_cpu_policy(d: dict, fallback: CpuPolicy) -> CpuPolicy:
    try:
        return CpuPolicy(
            epp=int(d.get("epp", fallback.epp)),
            boost_mode=int(d.get("boost_mode", fallback.boost_mode)),
            cpu_min=int(d.get("cpu_min", fallback.cpu_min)),
            cpu_max=int(d.get("cpu_max", fallback.cpu_max)),
            core_parking_min=int(d.get("core_parking_min", fallback.core_parking_min)),
            cooling_policy=int(d.get("cooling_policy", fallback.cooling_policy)),
        )
    except Exception:
        return fallback


def load_config() -> AppConfig:
    ensure_config_dir()
    if not os.path.exists(CONFIG_PATH):
        return AppConfig()

    try:
        with open(CONFIG_PATH, "r", encoding="utf-8") as f:
            raw = json.load(f)

        cfg = AppConfig(
            auto_mode=bool(raw.get("auto_mode", True)),
            set_power_plan=bool(raw.get("set_power_plan", True)),
            set_cpu_policy=bool(raw.get("set_cpu_policy", True)),
            ac_hz=raw.get("ac_hz", None),
            battery_hz=raw.get("battery_hz", None),
            ac_plan_guid=raw.get("ac_plan_guid", None),
            battery_plan_guid=raw.get("battery_plan_guid", None),
        )
        cfg.cpu_ac = _dict_to_cpu_policy(raw.get("cpu_ac", {}) or {}, default_cpu_ac())
        cfg.cpu_bat = _dict_to_cpu_policy(raw.get("cpu_bat", {}) or {}, default_cpu_bat())
        return cfg
    except Exception:
        return AppConfig()


def save_config(cfg: AppConfig) -> None:
    ensure_config_dir()
    payload = asdict(cfg)
    with open(CONFIG_PATH, "w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False, indent=2)


# ============================================================
# UI helpers
# ============================================================
ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("blue")

BOOST_LABELS = {
    0: "Disabled (Tasarruf)",
    1: "Enabled",
    2: "Aggressive (Performans)",
}

COOLING_LABELS = {
    1: "Active (Fan ile)",
    0: "Passive (Kısarak)",
}


def boost_label_to_value(label: str) -> int:
    for k, v in BOOST_LABELS.items():
        if v == label:
            return k
    return 1


def cooling_label_to_value(label: str) -> int:
    for k, v in COOLING_LABELS.items():
        if v == label:
            return k
    return 1


# Optional: psutil for live CPU info
try:
    import psutil  # type: ignore
except Exception:
    psutil = None


class HzApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.cfg = load_config()

        self.title(APP_NAME)
        self.geometry("560x720")
        self.protocol("WM_DELETE_WINDOW", self.hide_window)
        self.resizable(True, True)

        # Scrollable main container (ekrana sığmayan içerikler için)
        self.container = ctk.CTkScrollableFrame(self, corner_radius=0)
        self.container.pack(fill="both", expand=True, padx=0, pady=0)

        self.icon = None
        self._last_plug_state: Optional[bool] = None
        self._apply_lock = threading.Lock()
        self._update_prompted = False

        

        self._live_after_id = None  # after() job id for live CPU panel
# Hz
        self.supported_hz = list_supported_hz_for_current_mode()
        if not self.supported_hz:
            self.supported_hz = [60, 75, 120, 144, 165, 240]

        # Power plans
        self.schemes: List[Tuple[str, str, bool]] = []
        self.scheme_display_list: List[str] = []
        self.display_to_guid: Dict[str, str] = {}
        self.guid_to_display: Dict[str, str] = {}
        self.refresh_power_plans(initial=True)

        # Defaults Hz
        cur_hz = get_current_hz() or (self.supported_hz[-1] if self.supported_hz else 60)
        if self.cfg.ac_hz is None:
            self.cfg.ac_hz = cur_hz
        if self.cfg.battery_hz is None:
            self.cfg.battery_hz = min(self.supported_hz) if self.supported_hz else 60
        save_config(self.cfg)

        # ----------------- UI -----------------
        self.header = ctk.CTkLabel(self.container, text="Ekran + Güç Kontrol Merkezi (Pro)", font=("Roboto", 22, "bold"))
        self.header.pack(pady=(18, 6))

        self.status_label = ctk.CTkLabel(self.container, text=self._status_text(), text_color="#00ffcc", font=("Roboto", 13))
        self.status_label.pack(pady=(0, 10))

        # Live info panel
        self.live_label = ctk.CTkLabel(self.container, text="CPU: (pencere açıkken güncellenir)", font=("Roboto", 12))
        self.live_label.pack(pady=(0, 12))

        self.after(1500, self.check_launcher_update_background)

        # Toggles
        self.tog_frame = ctk.CTkFrame(self.container, corner_radius=14)
        self.tog_frame.pack(padx=16, pady=10, fill="x")

        self.switch_auto = ctk.CTkSwitch(self.tog_frame, text="Otomatik Geçiş (Priz/Batarya)", command=self.on_toggle)
        self.switch_auto.select() if self.cfg.auto_mode else self.switch_auto.deselect()
        self.switch_auto.pack(pady=(14, 8), padx=14, anchor="w")

        self.switch_power = ctk.CTkSwitch(self.tog_frame, text="Power Plan Değiştir", command=self.on_toggle)
        self.switch_power.select() if self.cfg.set_power_plan else self.switch_power.deselect()
        self.switch_power.pack(pady=(0, 8), padx=14, anchor="w")

        self.switch_cpu = ctk.CTkSwitch(self.tog_frame, text="CPU Policy Uygula (EPP/Boost/%)", command=self.on_toggle)
        self.switch_cpu.select() if self.cfg.set_cpu_policy else self.switch_cpu.deselect()
        self.switch_cpu.pack(pady=(0, 14), padx=14, anchor="w")

        # Hz frame
        self.hz_frame = ctk.CTkFrame(self.container, corner_radius=14)
        self.hz_frame.pack(padx=16, pady=10, fill="x")

        ctk.CTkLabel(self.hz_frame, text="Prizde (AC) Hz:").pack(padx=14, pady=(14, 4), anchor="w")
        self.ac_hz_menu = ctk.CTkOptionMenu(self.hz_frame, values=[str(x) for x in self.supported_hz], command=self.on_ac_hz_selected)
        self.ac_hz_menu.set(str(self.cfg.ac_hz))
        self.ac_hz_menu.pack(padx=14, pady=(0, 10), fill="x")

        ctk.CTkLabel(self.hz_frame, text="Bataryada Hz:").pack(padx=14, pady=(0, 4), anchor="w")
        self.bat_hz_menu = ctk.CTkOptionMenu(self.hz_frame, values=[str(x) for x in self.supported_hz], command=self.on_bat_hz_selected)
        self.bat_hz_menu.set(str(self.cfg.battery_hz))
        self.bat_hz_menu.pack(padx=14, pady=(0, 14), fill="x")

        # Plan frame
        self.plan_frame = ctk.CTkFrame(self.container, corner_radius=14)
        self.plan_frame.pack(padx=16, pady=10, fill="x")

        top_row = ctk.CTkFrame(self.plan_frame, corner_radius=0, fg_color="transparent")
        top_row.pack(fill="x", padx=14, pady=(14, 6))
        ctk.CTkLabel(top_row, text="Power Plan", font=("Roboto", 14, "bold")).pack(side="left")

        self.btn_refresh_plans = ctk.CTkButton(top_row, text="Planları Yenile", width=120, command=self.ui_refresh_plans)
        self.btn_refresh_plans.pack(side="right")

        ctk.CTkLabel(self.plan_frame, text="Prizde (AC) Power Plan:").pack(padx=14, pady=(6, 4), anchor="w")
        self.ac_plan_menu = ctk.CTkOptionMenu(self.plan_frame, values=self.scheme_display_list, command=self.on_ac_plan_selected)
        self.ac_plan_menu.pack(padx=14, pady=(0, 10), fill="x")

        ctk.CTkLabel(self.plan_frame, text="Bataryada Power Plan:").pack(padx=14, pady=(0, 4), anchor="w")
        self.bat_plan_menu = ctk.CTkOptionMenu(self.plan_frame, values=self.scheme_display_list, command=self.on_bat_plan_selected)
        self.bat_plan_menu.pack(padx=14, pady=(0, 10), fill="x")

        btn_row = ctk.CTkFrame(self.plan_frame, corner_radius=0, fg_color="transparent")
        btn_row.pack(fill="x", padx=14, pady=(0, 14))

        self.btn_restore_defaults = ctk.CTkButton(btn_row, text="Standart Planları Geri Yükle", command=self.ui_restore_default_schemes)
        self.btn_restore_defaults.pack(side="left", expand=True, fill="x", padx=(0, 8))

        self.btn_make_plans = ctk.CTkButton(btn_row, text="Balanced'tan 2 Plan Üret", command=self.ui_make_two_plans)
        self.btn_make_plans.pack(side="left", expand=True, fill="x", padx=(8, 0))

        self._sync_plan_menus_from_config()

        # CPU policy
        self.cpu_frame = ctk.CTkFrame(self.container, corner_radius=14)
        self.cpu_frame.pack(padx=16, pady=10, fill="x")

        ctk.CTkLabel(self.cpu_frame, text="CPU Policy (Gerçek Tasarruf/Performans)", font=("Roboto", 14, "bold")).pack(
            padx=14, pady=(14, 8), anchor="w"
        )

        self.tabs = ctk.CTkTabview(self.cpu_frame, height=320)
        self.tabs.pack(padx=14, pady=(0, 14), fill="x")

        self.tab_ac = self.tabs.add("AC (Priz)")
        self.tab_bat = self.tabs.add("Battery (Pil)")

        self._build_cpu_tab(self.tab_ac, mode="ac")
        self._build_cpu_tab(self.tab_bat, mode="bat")

        # Apply buttons
        self.btn_frame = ctk.CTkFrame(self.container, corner_radius=14)
        self.btn_frame.pack(padx=16, pady=10, fill="x")

        self.btn_apply_now = ctk.CTkButton(self.btn_frame, text="Şimdi Uygula (Mevcut Duruma Göre)", command=self.apply_for_current_power_state)
        self.btn_apply_now.pack(padx=14, pady=(14, 8), fill="x")

        self.btn_manual_ac = ctk.CTkButton(
            self.btn_frame,
            text="Manuel: AC Ayarlarını Uygula",
            command=lambda: self.manual_apply("ac"),
            fg_color="#1f538d",
            hover_color="#14375e",
        )
        self.btn_manual_ac.pack(padx=14, pady=6, fill="x")

        self.btn_manual_bat = ctk.CTkButton(
            self.btn_frame,
            text="Manuel: Battery Ayarlarını Uygula",
            command=lambda: self.manual_apply("bat"),
            fg_color="#333333",
            hover_color="#444444",
        )
        self.btn_manual_bat.pack(padx=14, pady=(6, 14), fill="x")

        self.hint = ctk.CTkLabel(self.container, text=f"Ayarlar: {CONFIG_PATH}", font=("Roboto", 11), text_color="#aaaaaa")
        self.hint.pack(pady=(6, 10))

        # Monitor threads
        self.check_thread = threading.Thread(target=self.battery_monitor, daemon=True)
        self.check_thread.start()
# start hidden
        self.withdraw()
        self.create_tray_icon()

    # ----------------- LIVE INFO -----------------
    def _live_text(self) -> str:
        if psutil is None:
            return "Canlı CPU bilgisi: psutil yok (istersen: pip install psutil)"
        try:
            p = psutil.cpu_percent(interval=None)
            f = psutil.cpu_freq()
            ghz = (f.current / 1000.0) if f and f.current else None
            ghz_txt = f"{ghz:.2f} GHz" if ghz is not None else "?"
            return f"CPU: {p:.0f}%  •  Frekans: {ghz_txt}"
        except Exception:
            return "CPU: ?"


    def _schedule_live_update(self):
        # Only update when window is visible (not withdrawn)
        try:
            if not self.winfo_viewable():
                self._live_after_id = None
                return
        except Exception:
            pass

        try:
            self.live_label.configure(text=self._live_text())
        except Exception:
            pass

        # schedule next
        self._live_after_id = self.after(1200, self._schedule_live_update)

    def start_live_updates(self):
        # start immediately
        if self._live_after_id is None:
            self._live_after_id = self.after(10, self._schedule_live_update)

    def stop_live_updates(self):
        if self._live_after_id is not None:
            try:
                self.after_cancel(self._live_after_id)
            except Exception:
                pass
            self._live_after_id = None

    # ----------------- CPU TAB UI -----------------
    def _build_cpu_tab(self, parent, mode: str):
        pol = self.cfg.cpu_ac if mode == "ac" else self.cfg.cpu_bat

        # EPP
        ctk.CTkLabel(parent, text="EPP (0=Performans, 100=Tasarruf)").pack(anchor="w", padx=10, pady=(10, 2))
        epp_value_label = ctk.CTkLabel(parent, text=f"{int(pol.epp)}")
        epp_value_label.pack(anchor="e", padx=10, pady=(0, 0))

        epp_slider = ctk.CTkSlider(
            parent, from_=0, to=100, number_of_steps=100,
            command=lambda v: self._on_cpu_slider(mode, "epp", int(v), epp_value_label)
        )
        epp_slider.set(int(pol.epp))
        epp_slider.pack(fill="x", padx=10, pady=(0, 8))

        # Boost
        ctk.CTkLabel(parent, text="Boost Mode").pack(anchor="w", padx=10, pady=(6, 2))
        boost_menu = ctk.CTkOptionMenu(
            parent,
            values=[BOOST_LABELS[0], BOOST_LABELS[1], BOOST_LABELS[2]],
            command=lambda label: self._on_cpu_option(mode, "boost_mode", boost_label_to_value(label)),
        )
        boost_menu.set(BOOST_LABELS.get(int(pol.boost_mode), BOOST_LABELS[1]))
        boost_menu.pack(fill="x", padx=10, pady=(0, 8))

        # CPU Max
        ctk.CTkLabel(parent, text="CPU Max (%)").pack(anchor="w", padx=10, pady=(6, 2))
        max_label = ctk.CTkLabel(parent, text=f"{int(pol.cpu_max)}")
        max_label.pack(anchor="e", padx=10, pady=(0, 0))
        max_slider = ctk.CTkSlider(
            parent, from_=1, to=100, number_of_steps=99,
            command=lambda v: self._on_cpu_slider(mode, "cpu_max", int(v), max_label)
        )
        max_slider.set(int(pol.cpu_max))
        max_slider.pack(fill="x", padx=10, pady=(0, 8))

        # CPU Min
        ctk.CTkLabel(parent, text="CPU Min (%)").pack(anchor="w", padx=10, pady=(6, 2))
        min_label = ctk.CTkLabel(parent, text=f"{int(pol.cpu_min)}")
        min_label.pack(anchor="e", padx=10, pady=(0, 0))
        min_slider = ctk.CTkSlider(
            parent, from_=0, to=100, number_of_steps=100,
            command=lambda v: self._on_cpu_slider(mode, "cpu_min", int(v), min_label)
        )
        min_slider.set(int(pol.cpu_min))
        min_slider.pack(fill="x", padx=10, pady=(0, 8))

        # Core parking
        ctk.CTkLabel(parent, text="Core Parking Min Cores (%)").pack(anchor="w", padx=10, pady=(6, 2))
        park_label = ctk.CTkLabel(parent, text=f"{int(pol.core_parking_min)}")
        park_label.pack(anchor="e", padx=10, pady=(0, 0))
        park_slider = ctk.CTkSlider(
            parent, from_=0, to=100, number_of_steps=100,
            command=lambda v: self._on_cpu_slider(mode, "core_parking_min", int(v), park_label)
        )
        park_slider.set(int(pol.core_parking_min))
        park_slider.pack(fill="x", padx=10, pady=(0, 8))

        # Cooling
        ctk.CTkLabel(parent, text="Cooling Policy").pack(anchor="w", padx=10, pady=(6, 2))
        cool_menu = ctk.CTkOptionMenu(
            parent,
            values=[COOLING_LABELS[1], COOLING_LABELS[0]],
            command=lambda label: self._on_cpu_option(mode, "cooling_policy", cooling_label_to_value(label)),
        )
        cool_menu.set(COOLING_LABELS.get(int(pol.cooling_policy), COOLING_LABELS[1]))
        cool_menu.pack(fill="x", padx=10, pady=(0, 8))

        apply_btn = ctk.CTkButton(
            parent,
            text="Bu Sekmenin CPU Policy'sini Seçili Planına Yaz",
            command=lambda: self.ui_apply_cpu_policy_tab(mode),
        )
        apply_btn.pack(fill="x", padx=10, pady=(10, 10))

    def _on_cpu_slider(self, mode: str, field_name: str, value: int, label_widget):
        label_widget.configure(text=str(int(value)))
        pol = self.cfg.cpu_ac if mode == "ac" else self.cfg.cpu_bat
        setattr(pol, field_name, int(value))

        if pol.cpu_min > pol.cpu_max:
            pol.cpu_min = pol.cpu_max

        save_config(self.cfg)

    def _on_cpu_option(self, mode: str, field_name: str, value: int):
        pol = self.cfg.cpu_ac if mode == "ac" else self.cfg.cpu_bat
        setattr(pol, field_name, int(value))
        save_config(self.cfg)

    # ----------------- STATUS -----------------
    def _status_text(self) -> str:
        cur = get_current_hz()
        cur_txt = f"{cur}Hz" if cur is not None else "Bilinmiyor"
        plug = is_plugged_in()
        if plug is True:
            ptxt = "Prizde"
        elif plug is False:
            ptxt = "Bataryada"
        else:
            ptxt = "Güç durumu bilinmiyor"

        _, plan_name = get_active_power_scheme()
        plan_txt = plan_name if plan_name else "?"
        return f"Şu anki: {cur_txt}  •  Durum: {ptxt}  •  Aktif Plan: {plan_txt}"

    def refresh_status(self):
        self.status_label.configure(text=self._status_text())

    # ----------------- TOGGLES -----------------
    def on_toggle(self):
        self.cfg.auto_mode = bool(self.switch_auto.get())
        self.cfg.set_power_plan = bool(self.switch_power.get())
        self.cfg.set_cpu_policy = bool(self.switch_cpu.get())
        save_config(self.cfg)

    # ----------------- Hz selection -----------------
    def on_ac_hz_selected(self, v: str):
        try:
            self.cfg.ac_hz = int(v)
            save_config(self.cfg)
        except Exception:
            pass

    def on_bat_hz_selected(self, v: str):
        try:
            self.cfg.battery_hz = int(v)
            save_config(self.cfg)
        except Exception:
            pass

    # ----------------- POWER PLANS -----------------
    def refresh_power_plans(self, initial: bool = False):
        self.schemes = list_power_schemes()

        self.scheme_display_list = []
        self.display_to_guid = {}
        self.guid_to_display = {}

        for guid, name, _ in self.schemes:
            disp = f"{name} — {guid}"
            self.scheme_display_list.append(disp)
            self.display_to_guid[disp] = guid
            self.guid_to_display[guid.lower()] = disp

        if not self.scheme_display_list:
            self.scheme_display_list = ["(Power plan listesi okunamadı) — "]
            self.display_to_guid[self.scheme_display_list[0]] = ""
            self.guid_to_display[""] = self.scheme_display_list[0]

        active_guid, _ = get_active_power_scheme()
        if initial:
            if not self.cfg.ac_plan_guid:
                self.cfg.ac_plan_guid = active_guid
            if not self.cfg.battery_plan_guid:
                self.cfg.battery_plan_guid = active_guid
            save_config(self.cfg)

    def ui_refresh_plans(self):
        self.refresh_power_plans(initial=False)
        self.ac_plan_menu.configure(values=self.scheme_display_list)
        self.bat_plan_menu.configure(values=self.scheme_display_list)
        self._sync_plan_menus_from_config()
        self.refresh_status()

    def _sync_plan_menus_from_config(self):
        ac_disp = None
        if self.cfg.ac_plan_guid:
            ac_disp = self.guid_to_display.get(self.cfg.ac_plan_guid.lower())
        bat_disp = None
        if self.cfg.battery_plan_guid:
            bat_disp = self.guid_to_display.get(self.cfg.battery_plan_guid.lower())

        self.ac_plan_menu.set(ac_disp or self.scheme_display_list[0])
        self.bat_plan_menu.set(bat_disp or self.scheme_display_list[0])

    def on_ac_plan_selected(self, display: str):
        guid = self.display_to_guid.get(display, "")
        self.cfg.ac_plan_guid = guid if guid else None
        save_config(self.cfg)
        self.after(0, self.refresh_status)

    def on_bat_plan_selected(self, display: str):
        guid = self.display_to_guid.get(display, "")
        self.cfg.battery_plan_guid = guid if guid else None
        save_config(self.cfg)
        self.after(0, self.refresh_status)

    def ui_restore_default_schemes(self):
        restore_default_power_schemes()
        self.ui_refresh_plans()

    def ui_make_two_plans(self):
        """
        Balanced/Dengeli planını baz alıp iki plan üretir.
        Not: Admin isteyebilir.
        """
        schemes = list_power_schemes()
        if not schemes:
            return

        # Balanced'ı bul
        base = None
        for guid, name, _ in schemes:
            n = name.lower()
            if "balanced" in n or "dengeli" in n:
                base = guid
                break
        if base is None:
            base = schemes[0][0]

        # High Perf
        new1, _ = duplicate_scheme(base)
        if new1:
            change_scheme_name(new1, "KEMAL - High Performance", "AC için")

        # Power Saver
        new2, _ = duplicate_scheme(base)
        if new2:
            change_scheme_name(new2, "KEMAL - Power Saver", "Battery için")

        self.ui_refresh_plans()

    # ----------------- APPLY LOGIC -----------------
    def manual_apply(self, mode: str):
        self.cfg.auto_mode = False
        self.switch_auto.deselect()
        save_config(self.cfg)

        if mode == "ac":
            self._apply_targets(plugged=True)
        else:
            self._apply_targets(plugged=False)

    def _apply_targets(self, plugged: bool):
        target_hz = self.cfg.ac_hz if plugged else self.cfg.battery_hz
        target_plan = self.cfg.ac_plan_guid if plugged else self.cfg.battery_plan_guid
        pol = self.cfg.cpu_ac if plugged else self.cfg.cpu_bat

        if self.cfg.set_power_plan and target_plan:
            set_power_scheme_by_guid(target_plan)

        if self.cfg.set_cpu_policy and target_plan:
            apply_cpu_policy_to_scheme(
                scheme_guid=target_plan,
                plugged=plugged,
                epp=pol.epp,
                boost_mode=pol.boost_mode,
                cpu_min=pol.cpu_min,
                cpu_max=pol.cpu_max,
                core_parking_min=pol.core_parking_min,
                cooling_policy=pol.cooling_policy,
            )

        if target_hz is not None:
            set_hz(int(target_hz))

        self.after(0, self.refresh_status)

    def apply_for_current_power_state(self):
        plug = is_plugged_in()
        if plug is None:
            return

        if not self._apply_lock.acquire(blocking=False):
            return
        try:
            self._apply_targets(plugged=bool(plug))
        finally:
            self._apply_lock.release()

    def ui_apply_cpu_policy_tab(self, mode: str):
        if mode == "ac":
            plan = self.cfg.ac_plan_guid
            pol = self.cfg.cpu_ac
            plugged = True
        else:
            plan = self.cfg.battery_plan_guid
            pol = self.cfg.cpu_bat
            plugged = False

        if not plan:
            return

        apply_cpu_policy_to_scheme(
            scheme_guid=plan,
            plugged=plugged,
            epp=pol.epp,
            boost_mode=pol.boost_mode,
            cpu_min=pol.cpu_min,
            cpu_max=pol.cpu_max,
            core_parking_min=pol.core_parking_min,
            cooling_policy=pol.cooling_policy,
        )
        self.after(0, self.refresh_status)

    # ----------------- MONITOR -----------------
    def battery_monitor(self):
        while True:
            try:
                plug = is_plugged_in()
                if self.cfg.auto_mode and plug is not None:
                    if plug != self._last_plug_state:
                        self._last_plug_state = plug
                        self.apply_for_current_power_state()
                else:
                    self.after(0, self.refresh_status)
            except Exception:
                pass
            time.sleep(MONITOR_INTERVAL_SEC)

    # ----------------- TRAY -----------------
    def hide_window(self):
        self.stop_live_updates()
        self.withdraw()
        self.create_tray_icon()

    def show_window(self):
        if self.icon:
            try:
                self.icon.stop()
            except Exception:
                pass
        self.icon = None
        self.deiconify()
        self.start_live_updates()
        self.after(0, self.refresh_status)

    def quit_app(self):
        self.stop_live_updates()
        if self.icon:
            try:
                self.icon.stop()
            except Exception:
                pass
        self.icon = None
        self.destroy()
        sys.exit()

    def create_tray_icon(self):
        if self.icon:
            return

        image = Image.new("RGB", (64, 64), color=(31, 83, 141))
        d = ImageDraw.Draw(image)
        d.rectangle([10, 10, 54, 54], outline="white", width=3)
        d.text((18, 22), "PRO", fill="white")

        menu = (
            item("Göster", lambda: self.after(0, self.show_window)),
            item("Şimdi Uygula", lambda: self.after(0, self.apply_for_current_power_state)),
            item("Çıkış", lambda: self.after(0, self.quit_app)),
        )
        self.icon = pystray.Icon("HzControlPro", image, APP_NAME, menu)
        threading.Thread(target=self.icon.run, daemon=True).start()

    def open_launcher_update(self):
        launcher_hint_path = r"C:\Saydut\launcher_path.txt"
        launcher_candidates = [
            r"C:\Saydut\SaydutLauncher\SaydutLauncher.exe",
            r"C:\Saydut\Saydut Launcher\SaydutLauncher.exe",
            r"C:\Saydut\SaydutLauncher\Saydut Launcher.exe",
        ]

        launcher_path = None
        if os.path.exists(launcher_hint_path):
            try:
                with open(launcher_hint_path, "r", encoding="utf-8") as handle:
                    candidate = handle.read().strip()
                if candidate and os.path.exists(candidate):
                    launcher_path = candidate
            except OSError:
                launcher_path = None

        if not launcher_path:
            launcher_path = next((p for p in launcher_candidates if os.path.exists(p)), None)

        if launcher_path:
            try:
                if launcher_path.lower().endswith(".py"):
                    subprocess.Popen([sys.executable, launcher_path])
                else:
                    subprocess.Popen([launcher_path])
                return
            except Exception as exc:
                messagebox.showerror("Hata", f"Launcher acilamadi:\n{exc}")
                return

        messagebox.showinfo(
            "Launcher Gerekli",
            "Guncelleme icin Saydut Launcher gerekli.\nIndirip kurduktan sonra tekrar deneyin.",
        )
        webbrowser.open("https://www.saydut.com")

    def show_update_prompt(self, latest_version: str) -> None:
        if self._update_prompted:
            return
        self._update_prompted = True
        if messagebox.askyesno(
            "Guncelleme mevcut",
            f"Mevcut surum: {APP_VERSION}\nYeni surum: {latest_version}\n\nLauncher acilsin mi?",
        ):
            self.open_launcher_update()

    def check_launcher_update_background(self) -> None:
        def worker():
            try:
                payload = _http_get_json(PROGRAMS_URL)
                latest = None
                for item in payload.get("programs", []):
                    if item.get("id") == PROGRAM_ID:
                        latest = item.get("version", "")
                        break
                if latest and _semver_tuple(latest) > _semver_tuple(APP_VERSION):
                    self.after(0, lambda: self.show_update_prompt(latest))
            except Exception:
                return

        threading.Thread(target=worker, daemon=True).start()


if __name__ == "__main__":
    app = HzApp()
    app.mainloop()
