import os
import time
import tkinter as tk
import re
from tkinter import ttk
from datetime import datetime, timedelta
import threading
import subprocess
from collections import deque
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import numpy as np
from matplotlib.patches import Wedge
import pythoncom
from pywinauto import Desktop
from pywinauto.timings import TimeoutError as PwaTimeoutError
import ctypes

current_test_state = None
good_popup = None


def safe_ui(func, *args, **kwargs):
    try:
        if root.winfo_exists():
            root.after(0, lambda: func(*args, **kwargs))
    except:
        pass

# =======================
#  BLOQUEO INPUT TEST PROGRAM (EnableWindow)
# =======================
_user32 = ctypes.windll.user32

def set_window_enabled(hwnd: int, enabled: bool) -> bool:
    """
    enabled=False  -> bloquea clicks/teclado en esa ventana
    enabled=True   -> permite clicks/teclado
    """
    try:
        if not hwnd:
            return False
        _user32.EnableWindow(int(hwnd), bool(enabled))
        return True
    except Exception as e:
        print("⚠️ set_window_enabled error:", e)
        return False


def is_good_state(state: str) -> bool:
    """
    True si el estado contiene Good o Bueno (case-insensitive).
    Ajusta aquí si quieres match exacto.
    """
    if not state:
        return False
    s = state.strip().lower()
    return ("good" in s) or ("bueno" in s)

EMBEDDED_QTY_EXCEPTIONS = [
    (2088701244, '2088701244', 9), (2088701390, '2088701390', 8), (2088701610, '2088701610', 9),
    (2088701746, '2088701746', 0), (2088702198, '2088702198', 7), (2088702199, '2088702199', 6),
    (2088702200, '2088702200', 8), (2088702201, '2088702201', 9), (2088702202, '2088702202', 5),
    (2088702203, '2088702203', 6), (2088702204, '2088702204', 6), (2088702205, '2088702205', 8),
    (2088702206, '2088702206', 8), (2088702207, '2088702207', 18), (2088702221, '2088702221', 9),
    (2088707042, '2088707042', 8), (2098700024, '2098700024', 8), (2098700026, '2098700026', 6),
    (2098700033, '2098700033', 8), (2098700046, '2098700046', 8), (2098700056, '2098700056', 7),
    (2098700058, '2098700058', 8), (2098700076, '2098700076', 7), (2098700083, '2098700083', 8),
    (2098700086, '2098700086', 8), (2098700108, '2098700108', 8), (2098700143, '2098700143', 7),
    (2098700154, '2098700154', 8), (2098700177, '2098700177', 7), (2098700189, '2098700189', 8),
    (2098700289, '2098700289', 7), (2098700290, '2098700290', 6), (2098700293, '2098700293', 7),
    (2098700296, '2098700296', 9), (2098700301, '2098700301', 7), (2098700302, '2098700302', 10),
    (2098700304, '2098700304', 7), (2098700305, '2098700305', 7), (2098700306, '2098700306', 10),
    (2098700307, '2098700307', 9), (2098700309, '2098700309', 9), (2098700315, '2098700315', 6),
    (2098700316, '2098700316', 9), (2098700320, '2098700320', 10), (2098700322, '2098700322', 9),
    (2098700353, '2098700353', 9), (2098700355, '2098700355', 9), (2098700356, '2098700356', 9),
    (2098700357, '2098700357', 8), (2098700358, '2098700358', 6), (2098700366, '2098700366', 8),
    (2098700371, '2098700371', 8), (2098700372, '2098700372', 7), (2098700373, '2098700373', 8),
    (2098700374, '2098700374', 7), (2098700378, '2098700378', 9), (2098706005, '2098706005', 8),
    (2098706008, '2098706008', 8), (2098706009, '2098706009', 8), (2098706010, '2098706010', 8),
    (2098706011, '2098706011', 8), (2098706013, '2098706013', 10), (2098706018, '2098706018', 8),
    (2098706021, '2098706021', 7), (2098706023, '2098706023', 7), (2098706024, '2098706024', 6),
    (2098706025, '2098706025', 7), (2098706026, '2098706026', 6), (2098706028, '2098706028', 8),
    (2098706031, '2098706031', 6), (2098706032, '2098706032', 6), (2098706040, '2098706040', 7),
    (2098706041, '2098706041', 8), (2098706044, '2098706044', 7), (2098706083, '2098706083', 8),
    (2098706084, '2098706084', 7), (2099700025, '2099700025', 9), (2099700030, '2099700030', 13),
    (2154140049, '2154140049', 2), (2154140222, '2154140222', 8), (2154140224, '2154140224', 8),
    (2154140225, '2154140225', 8), (2154140229, '2154140229', 8), (2154140243, '2154140243', 8),
    (2154140283, '2154140283', 8), (2154140284, '2154140284', 8), (2154140287, '2154140287', 8),
    (2154140289, '2154140289', 8), (2154140290, '2154140290', 8), (2154140293, '2154140293', 8),
    (2154140294, '2154140294', 9), (2154140295, '2154140295', 9), (2154140318, '2154140318', 8),
    (2154140319, '2154140319', 8), (2154150035, '2154150035', 7), (2154150089, '2154150089', 9),
    (2154150163, '2154150163', 9), (2154150214, '2154150214', 7), (2154150215, '2154150215', 9),
    (2154150247, '2154150247', 2), (2154150250, '2154150250', 2), (2154150337, '2154150337', 9),
    (2154150341, '2154150341', 8), (2154150342, '2154150342', 10), (2154150348, '2154150348', 9),
    (2154150380, '2154150380', 9), (2154150391, '2154150391', 9), (2154150447, '2154150447', 9),
    (2154150461, '2154150461', 9), (2154150554, '2154150554', 2), (2154150563, '2154150563', 4),
    (2154150582, '2154150582', 9), (2154150588, '2154150588', 7), (2154150603, '2154150603', 9),
    (2154150605, '2154150605', 8), (2154150633, '2154150633', 9), (2154150669, '2154150669', 9),
    (2154150672, '2154150672', 9), (2154150706, '2154150706', 9), (2154150707, '2154150707', 9),
    (2154150715, '2154150715', 8), (2154155136, '2154155136', 9), (2154160029, '2154160029', 9),
    (2154160030, 'SJ8T-18812-EB', 9), (2154160031, 'SJ8T-18812-REB', 9),
    (2154160032, 'SJ8T-14F662-KB', 8), (2154160033, 'SJ8T-14F662-SB', 8),
    (2154160034, 'SJ8T-14F662-JA', 9), (2154160035, 'SJ8T-18812-RCA', 9),
    (2154170049, 'SJ8T-19A397-REB', 17), (2154170050, 'SJ8T-19A397-EB', 19),
    (2154170052, 'SJ8T-19A397-LEA', 13)
]

class FileModifiedHandler(FileSystemEventHandler):
    def __init__(self, label, cycle_label, fail_label, yield_label, canvas, frame_color):
        self.label = label
        self.cycle_label = cycle_label
        self.fail_label = fail_label
        self.yield_label = yield_label
        self.canvas = canvas
        self.frame_color = frame_color
        self.modification_count = 0
        self.fail_count = 0
        self.yield_value = 100
        self.timestamps = deque(maxlen=50)
        self.qty_exceptions = {}
        self.np_equivalence_map = {}
        self.current_part_number = None
        self.forgiven_fails_remaining = 0
        self.forgiven_good_remaining = 0
        self.load_qty_exceptions()
        self.reset_times = [
            datetime.strptime("06:25", "%H:%M").time(),
            datetime.strptime("14:25", "%H:%M").time(),
            datetime.strptime("21:55", "%H:%M").time()
        ]
        self.folder_path = r"C:\\Users\\Public\\Documents\\Cirris\\printer"
        self.recently_incremented = False
        self.start_timer()

    def lerp_color(self, c1, c2, t):
        return tuple(c1[i] + (c2[i] - c1[i]) * t for i in range(3))

    def animate_gauge(self, target_value, duration=0.9, steps=40):
        start_value = self.yield_value
        delta = target_value - start_value

        def step(i):
            if i > steps:
                self.yield_value = target_value
                self.draw_gauge()

                if target_value >= 100:
                    self.animate_glow()

                return

            t = i / steps
            ease = 1 - (1 - t) ** 3  # Ease-out cúbico

            self.yield_value = start_value + delta * ease
            self.draw_gauge()

            root.after(int(duration * 1000 / steps), lambda: step(i + 1))

        step(0)
    def animate_glow(self):
        pulses = 6
        max_alpha = 0.25

        def pulse(i):
            if i > pulses:
                self.draw_gauge()
                return

            alpha = max_alpha * (1 - abs(i - pulses/2) / (pulses/2))
            self.draw_gauge(glow_alpha=alpha)

            root.after(60, lambda: pulse(i + 1))

        pulse(0)



    def start_timer(self):
        t = threading.Timer(60, self.check_reset)
        t.daemon = True
        t.start()

    def check_reset(self):
        current_time = datetime.now().time()
        if any(reset_time <= current_time < (datetime.combine(datetime.today(), reset_time) + timedelta(minutes=1)).time()
               for reset_time in self.reset_times):
            self.reset_counters()
        self.start_timer()

    def reset_counters(self):
        self.yield_value = 0
        self.modification_count = 0
        self.fail_count = 0
        self.timestamps.clear()
        self.current_part_number = None
        self.forgiven_fails_remaining = 0
        self.forgiven_good_remaining = 0
        self.load_qty_exceptions()

        safe_ui(self.yield_label.config, text="Yield: 100%")
        safe_ui(self.label.config, text="0")
        safe_ui(self.fail_label.config, text="0")
        safe_ui(self.cycle_label.config, text="N/A")

        safe_ui(self.animate_gauge, 100)

        turno = get_turno_actual()
        safe_ui(root.title, f"Corrida Actual - Turno {turno} --First Pass Yield--")

    def on_modified(self, event):
        if not root.winfo_exists():
            return

        self.update_yield()

        if event.is_directory or not event.src_path.lower().endswith(".prn"):
            return

        if not self.recently_incremented:
            self.timestamps.append(datetime.now())
            self.update_cycle_time()
            self.check_for_failures(event.src_path)
            self.recently_incremented = True
            threading.Timer(1, self.reset_increment_flag).start()

    def update_cycle_time(self):
        if len(self.timestamps) > 1:
            total_time = sum(
                (self.timestamps[i] - self.timestamps[i - 1]).total_seconds()
                for i in range(1, len(self.timestamps))
            )
            average_cycle_time = total_time / (len(self.timestamps) - 1)
            safe_ui(self.cycle_label.config, text=f"{average_cycle_time:.1f}s")

    def check_for_failures(self, file_path):
        for attempt in range(5):
            try:
                with open(file_path, "r", encoding="utf-8") as file:
                    first_line = file.readline().strip()
                    file.seek(0)

                    part_number = None
                    for line in file:
                        match = re.search(r"FD([^\s\^]+)", line)
                        if match:
                            part_number = match.group(1)
                            break

                    is_failure = "^XA" in first_line

                    key_np = self.np_equivalence_map.get(part_number, part_number)

                    normalized_current = self.np_equivalence_map.get(
                        self.current_part_number, self.current_part_number
                    )

                    if key_np != normalized_current:
                        self.current_part_number = key_np
                        self.forgiven_fails_remaining = self.qty_exceptions.get(key_np, 10)
                        self.forgiven_good_remaining = 2

                    if is_failure:
                        if self.forgiven_fails_remaining > 0:
                            self.forgiven_fails_remaining -= 1
                        else:
                            self.fail_count += 1
                            self.modification_count += 1
                            safe_ui(self.fail_label.config, text=f"{self.fail_count}")
                            safe_ui(self.label.config, text=f"{self.modification_count}")
                    else:
                        if self.forgiven_good_remaining > 0:
                            self.forgiven_good_remaining -= 1
                        else:
                            self.modification_count += 1
                            safe_ui(self.label.config, text=f"{self.modification_count}")
                break

            except PermissionError:
                time.sleep(0.5)
            except Exception as e:
                print("⚠️ Error inesperado:", e)
                break

    def reset_increment_flag(self):
        self.recently_incremented = False

    def update_yield(self):
        if self.modification_count == 0:
            self.yield_value = 100
        else:
            self.yield_value = max(
                0,
                ((self.modification_count - self.fail_count) / self.modification_count) * 100
            )

        safe_ui(self.yield_label.config, text=f"Yield: {self.yield_value:.2f}%")
        safe_ui(self.draw_gauge)

    def draw_gauge(self, glow_alpha=0):

        ax = self.ax
        ax.clear()

        self.fig.patch.set_facecolor("#E0E0E0")
        ax.set_facecolor("none")

        ax.text(0, 0.87, "Yield", ha="center", va="center", fontsize=12)

        angle = np.interp(self.yield_value, [0, 100], [180, 0])

        # 🎨 Color gradual inteligente
        red = (1, 0.39, 0.39)
        yellow = (1, 0.84, 0.25)
        green = (0.45, 0.77, 0.58)

        if self.yield_value <= 80:
            t = self.yield_value / 80
            bar_color = self.lerp_color(red, yellow, t)
        else:
            t = (self.yield_value - 80) / 20
            bar_color = self.lerp_color(yellow, green, t)

        # Fondo
        ax.add_patch(Wedge((0, 0), 1, 0, 180,
                        facecolor="#B0B0B0",
                        edgecolor="none",
                        linewidth=0,
                        antialiased=False))

        # Barra
        ax.add_patch(Wedge((0, 0), 1, angle, 180,
                        facecolor=bar_color,
                        edgecolor="none",
                        linewidth=0,
                        antialiased=False))

        # Glow opcional
        if glow_alpha > 0:
            ax.add_patch(Wedge((0, 0), 1.08, 0, 180,
                            facecolor=(0.4, 1, 0.4, glow_alpha),
                            edgecolor="none",
                            linewidth=0))

        # Interior
        ax.add_patch(Wedge((0, 0), 0.8, 0, 180,
                        facecolor="#E0E0E0",
                        edgecolor="none",
                        linewidth=0,
                        antialiased=False))

        ax.text(0, 0.2,
                f"{self.yield_value:.1f}%",
                ha='center',
                va='center',
                fontsize=18,
                fontweight='bold')

        ax.set_xticks([])
        ax.set_yticks([])
        ax.set_xlim(-1.2, 1.2)
        ax.set_ylim(-0.2, 1.2)

        self.canvas.draw_idle()

    def load_qty_exceptions(self):
        try:
            self.qty_exceptions.clear()
            self.np_equivalence_map.clear()

            for original, equivalent, qty in EMBEDDED_QTY_EXCEPTIONS:
                original = str(original).strip()
                equivalent = str(equivalent).strip()

                key = equivalent if equivalent else original

                self.qty_exceptions[key] = qty
                self.np_equivalence_map[original] = key
                self.np_equivalence_map[equivalent] = key

            print("📦 QTY excepciones cargadas correctamente.")

        except Exception as e:
            print("⚠️ Error cargando excepciones:", e)




def update_window_title():
    turno = get_turno_actual()

    if current_test_state:
        root.title(
            f"Corrida Actual - Turno {turno} --First Pass Yield--                  {current_test_state}"
        )
    else:
        root.title(
            f"Corrida Actual - Turno {turno} --First Pass Yield--"
        )

def show_good_popup():
    global good_popup

    if good_popup is not None and good_popup.winfo_exists():
        return

    good_popup = tk.Toplevel(root)
    good_popup.overrideredirect(True)
    good_popup.attributes("-topmost", True)
    good_popup.attributes("-alpha", 0.35)  # Transparencia suave

    # Fondo verde muy discreto
    good_popup.configure(bg="#2ecc71")

    width = 140
    height = 30

    screen_width = good_popup.winfo_screenwidth()
    x = screen_width - width - 10
    y = 5

    good_popup.geometry(f"{width}x{height}+{x}+{y}")
    good_popup.update_idletasks()

    label = tk.Label(
        good_popup,
        text="GOOD",
        font=("Segoe UI", 10, "bold"),
        fg="white",
        bg="#2ecc71"
    )
    label.pack(expand=True, fill="both")

def hide_good_popup():
    global good_popup

    if good_popup is not None and good_popup.winfo_exists():
        good_popup.destroy()
        good_popup = None

def show_np_not_found_toast():
    toast = tk.Toplevel()
    toast.overrideredirect(True)
    toast.attributes("-topmost", True)
    toast.attributes("-alpha", 0.95)  # Ligera transparencia
    toast.configure(bg="#ff4d4d")
    screen_width = toast.winfo_screenwidth()
    screen_height = toast.winfo_screenheight()
    width, height = 300, 60
    x = (screen_width // 2) - (width // 2)
    y = (screen_height // 2) - (height // 2) - 100  # Arriba del centro
    toast.geometry(f"{width}x{height}+{x}+{y}")
    frame = tk.Frame(toast, bg="#ff4d4d", bd=0, relief="flat", highlightthickness=0)
    frame.pack(expand=True, fill="both")
    label = tk.Label(
        frame,
        text="⚠️ Agregar NP a ContadorGeneralParaCirris.pyw",
        font=("Segoe UI", 12, "bold"),
        fg="white",
        bg="#ff4d4d"
    )
    label.pack(expand=True)
    toast.after(2500, toast.destroy)

def get_turno_actual():
    now = datetime.now().time()
    if now >= datetime.strptime("06:24", "%H:%M").time() and now < datetime.strptime("14:24", "%H:%M").time():
        return 1
    elif now >= datetime.strptime("14:24", "%H:%M").time() and now < datetime.strptime("21:54", "%H:%M").time():
        return 2
    else:
        return 3
# =======================
#  MONITOR CONSOLA - TEST MONITOR
# =======================

TITLE_CONTAINS = "Test Program"
POLL_MS = 100
RETRY_MS = 300
#STATE_NAMES = {"Good", "Bad", "Attach then Start", "Ready to Test"}


def find_top_state_pane(win):
    candidates = []

    for p in win.descendants(control_type="Pane"):
        try:
            name = (p.window_text() or "").strip()
            if not name:
                continue

            r = p.rectangle()
            if r.width() <= 0 or r.height() <= 0:
                continue

            candidates.append((r.top, r.width() * r.height(), p))
        except Exception:
            pass

    if not candidates:
        return None

    candidates.sort(key=lambda x: (x[0], -x[1]))
    return candidates[0][2]

class ConsoleUiMonitor(threading.Thread):
    def __init__(self):
        super().__init__(daemon=True)
        self._running = threading.Event()
        self._running.set()

    def stop(self):
        self._running.clear()
        self.join(timeout=2)

    def run(self):
        initialized_here = False

        try:
            try:
                pythoncom.CoInitialize()
                initialized_here = True
            except pythoncom.com_error as e:
                # Ignorar RPC_E_CHANGED_MODE específicamente
                if hasattr(e, "hresult") and e.hresult == -2147417850:
                    # RPC_E_CHANGED_MODE → ya estaba inicializado, no es fatal
                    pass
                else:
                    print("Error COM inesperado:", e)

            win = None
            pane = None
            last_state = None
            test_hwnd = None
            test_locked = False
            waiting_printed = False

            while self._running.is_set():

                try:
                    if win is None:
                        try:
                            d = Desktop(backend="uia")
                            win = d.window(title_re=f".*{TITLE_CONTAINS}.*")
                            win.wait("exists", timeout=2)
                            test_hwnd = getattr(win, "handle", None)

                            try:
                                if test_hwnd:
                                    set_window_enabled(test_hwnd, True)  # siempre arrancar habilitado al reconectar
                            except:
                                pass

                            test_locked = False
                            pane = None
                            last_state = None

                            print("✅ Conectado a Test monitor")
                            waiting_printed = False

                        except PwaTimeoutError:
                            if not waiting_printed:
                                print("⏳ Esperando ventana 'Test monitor'...")
                                waiting_printed = True

                            time.sleep(RETRY_MS / 1000)
                            continue

                    try:
                        win.wait("exists", timeout=0.2)
                    except Exception:
                        print("⚠️ Ventana cerrada. Reintentando conexión...")

                        # si estaba bloqueada, intentar liberar por seguridad
                        try:
                            if test_hwnd:
                                set_window_enabled(test_hwnd, True)
                        except:
                            pass
                        test_locked = False
                        test_hwnd = None

                        # ocultar popup good
                        try:
                            if root.winfo_exists():
                                root.after(0, hide_good_popup)
                        except:
                            pass

                        win = None
                        pane = None
                        last_state = None
                        time.sleep(RETRY_MS / 1000)
                        continue

                    if pane is None:
                        pane = find_top_state_pane(win)
                        if pane is None:
                            time.sleep(RETRY_MS / 1000)
                            continue

                    state = " ".join((pane.window_text() or "").strip().split())
                    if state != last_state:
                        if state:
                            print(f"📡 Estado detectado: {state}")
                        else:
                            print("📡 Estado vacío o transición")

                        last_state = state

                        global current_test_state
                        current_test_state = state if state else None
                        # ---- BLOQUEO / DESBLOQUEO de Test Program según estado ----
                        try:
                            should_lock = is_good_state(state)

                            # si se perdió el handle por reconexión, refrescar
                            if not test_hwnd and win is not None:
                                test_hwnd = getattr(win, "handle", None)

                            if should_lock and not test_locked:
                                if set_window_enabled(test_hwnd, False):
                                    test_locked = True
                                    print("🔒 Test Program BLOQUEADO (estado Good/Bueno)")

                            elif (not should_lock) and test_locked:
                                if set_window_enabled(test_hwnd, True):
                                    test_locked = False
                                    print("🔓 Test Program DESBLOQUEADO (estado != Good/Bueno)")

                        except Exception as e:
                            print("⚠️ Error en lock/unlock:", e)
                        # Mostrar u ocultar indicador GOOD
                        if is_good_state(state):
                            if root.winfo_exists():
                                root.after(0, show_good_popup)
                        else:
                            if root.winfo_exists():
                                root.after(0, hide_good_popup)
                        try:
                            if root.winfo_exists():
                                root.after(0, update_window_title)
                        except Exception:
                            pass



                    time.sleep(POLL_MS / 1000)

                except Exception as e:
                    print("⚠️ UIA stale / reiniciando conexión:", e)
                    win = None
                    pane = None
                    last_state = None
                    time.sleep(RETRY_MS / 1000)

        finally:
            # Seguridad: si el hilo muere, intentar desbloquear la ventana
            try:
                if test_hwnd:
                    set_window_enabled(test_hwnd, True)
            except:
                pass

            # Ocultar popup si quedó prendido
            try:
                if root.winfo_exists():
                    root.after(0, hide_good_popup)
            except:
                pass

            try:
                pythoncom.CoUninitialize()
            except:
                pass

def main():
    global root, vbs_process
    vbs_process = None
    root = tk.Tk()
    # Establecer el título dinámico desde el arranque
    turno = get_turno_actual()
    root.protocol("WM_DELETE_WINDOW", lambda: print("❌ Botón cerrar deshabilitado"))
    #root.title(f"Corrida Actual - Turno {turno}")
    root.title(f"Corrida Actual - Turno {turno} --First Pass Yield--")
    #root.title("Monitor de Piezas Probadas")
    root.geometry("580x150+435+520")  #  Ampliamos el ancho para mejor distribución
    root.resizable(False, False)
    #root.overrideredirect(True)
    bg_color = "#F5F5F5"       # Fondo principal (Blanco humo)
    text_color = "#333333"      # Texto principal (Negro suave)
    highlight_color = "#0078D7" # Azul brillante para resaltar
    frame_color = "#E0E0E0"     # Marco (Gris claro)
    fail_color = "#D32F2F"      # Rojo oscuro para fallas
    root.configure(bg=frame_color)
    # **Dividimos la interfaz en dos partes**
    main_frame = tk.Frame(root, bg=frame_color)
    main_frame.pack(expand=True, fill="both", padx=10, pady=10)
    frame_left = tk.Frame(main_frame, bg=frame_color)
    frame_left.grid(row=0, column=0, sticky="w", padx=10)
    frame_right = tk.Frame(main_frame, bg=frame_color, width=320, height=120)  
    frame_right.grid(row=0, column=1, sticky="n", padx=10)
    frame_right.grid_propagate(False)  # Evita que se colapse por el contenido
    label_text = tk.Label(frame_left, text="Total Pzs:", font=("Arial", 16, "bold"),
                        fg=text_color, bg=frame_color)
    label_text.grid(row=0, column=0, sticky="w", padx=10)
    label = tk.Label(frame_left, text="0", font=("Arial", 20, "bold"),
                    fg=text_color, bg=frame_color)
    label.grid(row=0, column=1, sticky="w", padx=5)  #  Espaciado uniforme
    fail_text = tk.Label(frame_left, text="Fallas:", font=("Arial", 16, "bold"),
                        fg=text_color, bg=frame_color)
    fail_text.grid(row=1, column=0, sticky="w", padx=10)
    fail_label = tk.Label(frame_left, text="0", font=("Arial", 16, "bold"),
                        fg=text_color, bg=frame_color)
    fail_label.grid(row=1, column=1, sticky="w", padx=5)
    cycle_text = tk.Label(frame_left, text="Tiempo ciclo:", font=("Arial", 16, "bold"),
                        fg=text_color, bg=frame_color)
    cycle_text.grid(row=2, column=0, sticky="w", padx=10)
    cycle_label = tk.Label(frame_left, text="N/A", font=("Arial", 16, "bold"),
                        fg=text_color, bg=frame_color)
    cycle_label.grid(row=2, column=1, sticky="w", padx=5)
    yield_text = tk.Label(frame_left, text="Yield:", font=("Arial", 16, "bold"),
                        fg=text_color, bg=frame_color)
    yield_text.grid(row=3, column=0, sticky="w", padx=10)
    yield_label = tk.Label(frame_left, text="100%", font=("Arial", 16, "bold"),
                        fg=text_color, bg=frame_color)
    yield_text.grid_remove()  # Esto lo oculta sin eliminarlo
    gauge_frame = tk.Frame(frame_right, bg="lightgray", width=300, height=200)
    gauge_frame.place(x=10, y=-40)  # 🔼 Lo colocamos visible
    gauge_frame.pack_propagate(False)  # Mantener el tamaño fijo
    # Crear una figura de Matplotlib
    fig, ax = plt.subplots(figsize=(3, 2))
    canvas = FigureCanvasTkAgg(fig, master=gauge_frame)
    canvas.get_tk_widget().pack()

    file_handler = FileModifiedHandler(
        label, cycle_label, fail_label, yield_label, canvas, frame_color
    )

    file_handler.fig = fig
    file_handler.ax = ax
    file_handler.yield_value = 0
    file_handler.animate_gauge(100)

    canvas.draw_idle()
    #file_handler = FileModifiedHandler(label, cycle_label, fail_label, yield_label, canvas, frame_color)
    observer = Observer()
    observer.schedule(file_handler, file_handler.folder_path, recursive=True)
    observer.start()
    force_focus()
    # ---- Monitor consola Test monitor ----
    console_monitor = ConsoleUiMonitor()
    console_monitor.start()
    root.bind("<Button-1>", on_title_bar_click)
    root.bind("<B1-Motion>", on_drag_motion)
    root.bind("6", toggle_vbs_script)
    
    
    try:
        root.mainloop()
    except KeyboardInterrupt:
        print("Aplicación cerrada.")

    # ---- Limpieza ordenada ----
    try:
        console_monitor.stop()
    except:
        pass

    observer.stop()
    observer.join()

def toggle_vbs_script(event=None):
    global vbs_process
    tester_file_path = r"C:\Users\Public\Documents\Cirris\tester.txt"
    if os.path.exists(tester_file_path):
        with open(tester_file_path, 'r') as file:
            tester_name = file.read().strip()
        script_path = os.path.join(r"\\mlxgumvwfile01\Departamentos\Fakra\Pruebas\CyclesCounter", tester_name, "CyclesCounter.vbs")
        if vbs_process is None:
            if os.path.exists(script_path):
                try:
                    vbs_process = subprocess.Popen(['cscript', script_path], creationflags=subprocess.CREATE_NO_WINDOW)
                    print("✅ VBScript iniciado.")
                except Exception as e:
                    print(f"⚠️ Error al ejecutar el script: {e}")
            else:
                print(f"⚠️ El script {script_path} no existe.")
        else:
            try:
                vbs_process.terminate()
                vbs_process.wait()
                vbs_process = None
                print("VBScript detenido.")
            except Exception as e:
                print(f"Error al detener el script: {e}")

def force_focus():
    try:
        root.attributes('-topmost', True)        # Siempre al frente
        root.focus_force()                       # Forzar foco
        root.update_idletasks()
    except Exception as e:
        print(f"⚠️ Error forzando foco: {e}")
    root.after(1500, force_focus)  # Cada 1.5s, balanceado

def on_title_bar_click(event):
    """Guarda la posición inicial del mouse al hacer clic en la ventana."""
    global root
    root.x_offset = event.x_root - root.winfo_x()
    root.y_offset = event.y_root - root.winfo_y()

def on_drag_motion(event):
    """Mueve la ventana cuando se arrastra con el mouse."""
    global root
    new_x = event.x_root - root.x_offset
    new_y = event.y_root - root.y_offset
    root.geometry(f"+{new_x}+{new_y}")

def close_app(event=None):
    print("Aplicación cerrada por el usuario.")
    root.destroy()

if __name__ == "__main__":
    main()
