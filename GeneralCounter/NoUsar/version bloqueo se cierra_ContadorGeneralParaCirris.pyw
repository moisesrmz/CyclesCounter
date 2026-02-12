import os
import time
import tkinter as tk
import re
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
import queue
import ctypes
import pythoncom
from pywinauto import Desktop
from pywinauto.timings import TimeoutError as PwaTimeoutError


# =======================
#  QTY Exceptions (tu lista)
# =======================
EMBEDDED_QTY_EXCEPTIONS = [
    (2088701244, '2088701244', 9), (2088701390, '2088701390', 8), (2088701610, '2088701610', 9),
    (2088701746, '2088701746', 0), (2088702198, '2088702198', 7), (2088702199, '2088702199', 6),
    (2088702200, '2088702200', 8), (2088702201, '2088702201', 9), (2088702202, '2088702202', 5),
    (2088702203, '2088702203', 6), (2088702204, '2088702204', 6), (2088702205, '2088702205', 8),
    (2088702206, '2088702206', 8), (2088702207, '2088702207', 8), (2088702221, '2088702221', 9),
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


# =======================
#  Turno + Toast
# =======================
def get_turno_actual():
    now = datetime.now().time()
    if datetime.strptime("06:24", "%H:%M").time() <= now < datetime.strptime("14:24", "%H:%M").time():
        return 1
    elif datetime.strptime("14:24", "%H:%M").time() <= now < datetime.strptime("21:54", "%H:%M").time():
        return 2
    else:
        return 3


def show_np_not_found_toast():
    toast = tk.Toplevel()
    toast.overrideredirect(True)
    toast.attributes("-topmost", True)
    toast.attributes("-alpha", 0.95)
    toast.configure(bg="#ff4d4d")
    screen_width = toast.winfo_screenwidth()
    screen_height = toast.winfo_screenheight()
    width, height = 300, 60
    x = (screen_width // 2) - (width // 2)
    y = (screen_height // 2) - (height // 2) - 100
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


# =======================
#  File watcher handler (tu lógica)
# =======================
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
        self.draw_gauge()

    def start_timer(self):
        t = threading.Timer(60, self.check_reset)
        t.daemon = True
        t.start()

    def check_reset(self):
        current_time = datetime.now().time()
        if any(
            reset_time <= current_time <
            (datetime.combine(datetime.today(), reset_time) + timedelta(minutes=1)).time()
            for reset_time in self.reset_times
        ):
            self.reset_counters()
        self.start_timer()

    def reset_counters(self):
        self.yield_value = 100
        self.yield_label.config(text="Yield: 100%")
        self.modification_count = 0
        self.fail_count = 0
        self.timestamps.clear()
        self.label.config(text="0")
        self.fail_label.config(text="0")
        self.cycle_label.config(text="N/A")
        turno = get_turno_actual()
        root.title(f"Corrida Actual - Turno {turno} --First Pass Yield--")

        print(f"🔄 Reinicio automático - Turno {turno} a las {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        root.after(0, self.draw_gauge)
        self.current_part_number = None
        self.forgiven_fails_remaining = 0
        self.forgiven_good_remaining = 0
        self.load_qty_exceptions()
        print("Reinicio - Reset de red rabbits")

    def on_modified(self, event):
        self.update_yield()
        if event.is_directory or not event.src_path.lower().endswith(".prn"):
            return
        if not self.recently_incremented:
            self.timestamps.append(datetime.now())
            self.update_cycle_time()
            self.check_for_failures(event.src_path)
            self.recently_incremented = True
            t = threading.Timer(1, self.reset_increment_flag)
            t.daemon = True
            t.start()

    def update_cycle_time(self):
        if len(self.timestamps) > 1:
            total_time = sum(
                (self.timestamps[i] - self.timestamps[i - 1]).total_seconds()
                for i in range(1, len(self.timestamps))
            )
            average_cycle_time = total_time / (len(self.timestamps) - 1)
            self.cycle_label.config(text=f"{average_cycle_time:.1f}s")

    def check_for_failures(self, file_path):
        if file_path.lower().endswith(".prn"):
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

                        if is_failure:
                            key_np = self.np_equivalence_map.get(part_number)
                            if not key_np:
                                show_np_not_found_toast()
                                key_np = part_number
                        else:
                            key_np = part_number

                        normalized_current = self.np_equivalence_map.get(
                            self.current_part_number, self.current_part_number
                        )
                        if key_np != normalized_current:
                            print(f"🔁 Número de parte cambiado a: {part_number} (usado como {key_np})")
                            self.current_part_number = key_np
                            self.forgiven_fails_remaining = self.qty_exceptions.get(key_np, 10)
                            self.forgiven_good_remaining = 2
                            print(f"📄 Carga inicial de no contar: {self.forgiven_fails_remaining} fallas, 1 pieza buena para NP {key_np}")

                        if is_failure:
                            if self.forgiven_fails_remaining > 0:
                                self.forgiven_fails_remaining -= 1
                                print(f"Falla NO CONTADA. Quedan {self.forgiven_fails_remaining}")
                            else:
                                self.fail_count += 1
                                self.fail_label.config(text=f"{self.fail_count}")
                                self.modification_count += 1
                                self.label.config(text=f"{self.modification_count}")
                                print(f"Falla CONTADA. Total: {self.fail_count}")
                        else:
                            if self.forgiven_good_remaining > 0:
                                self.forgiven_good_remaining -= 1
                                print("Pieza buena (no contada)")
                            else:
                                self.modification_count += 1
                                self.label.config(text=f"{self.modification_count}")
                                print(f"Pieza buena contada. Total: {self.modification_count}")
                    break
                except PermissionError:
                    time.sleep(0.5)
                except Exception as e:
                    print(f"⚠️ Error inesperado: {e}")
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
        self.yield_label.config(text=f"Yield: {self.yield_value:.2f}%")
        root.after(0, self.draw_gauge)

    def draw_gauge(self):
        if hasattr(self.canvas, 'figure'):
            plt.close(self.canvas.figure)
        fig, ax = plt.subplots(figsize=(3, 2))
        fig.patch.set_facecolor("#E0E0E0")
        ax.set_facecolor("none")
        ax.text(0, 0.87, "Yield", ha="center", va="center", fontsize=12, color="black")
        angle = np.interp(self.yield_value, [0, 100], [180, 0])

        if self.yield_value <= 80:
            bar_color = (255/255, 126/255, 99/255, 0.9)
        elif self.yield_value <= 90:
            bar_color = (255/255, 215/255, 64/255, 0.9)
        else:
            bar_color = (114/255, 196/255, 149/255, 0.9)

        ax.add_patch(Wedge((0, 0), 1, 0, 180, facecolor="#B0B0B0", edgecolor="none", lw=1))
        ax.add_patch(Wedge((0, 0), 1, angle, 180, facecolor=bar_color))
        ax.add_patch(Wedge((0, 0), 0.8, 0, 180, facecolor="#E0E0E0", edgecolor="none", zorder=10))
        ax.text(0, 0.2, f"{self.yield_value:.1f}%", ha='center', va='center',
                fontsize=18, fontweight='bold', color="black", zorder=11)
        ax.text(-1.1, 0.1, "0", ha='center', va='center', fontsize=10, color="black")
        ax.text(1.1, 0.1, "   100", ha='center', va='center', fontsize=10, color="black")

        ax.set_xticks([]), ax.set_yticks([])
        for spine in ax.spines.values():
            spine.set_visible(False)
        ax.set_xlim(-1.2, 1.2), ax.set_ylim(-0.2, 1.2)

        self.canvas.figure = fig
        self.canvas.draw()

    def load_qty_exceptions(self):
        try:
            self.qty_exceptions.clear()
            self.np_equivalence_map.clear()
            for original, equivalent, qty in EMBEDDED_QTY_EXCEPTIONS:
                original = str(original).strip()
                equivalent = str(equivalent).strip()
                self.qty_exceptions[equivalent or original] = qty
                self.np_equivalence_map[original] = equivalent or original
                self.np_equivalence_map[equivalent] = equivalent or original
            print("📦 QTY excepciones cargadas desde datos embebidos.")
        except Exception as e:
            print(f"⚠️ Error al cargar datos embebidos: {e}")
# =======================
#  Monitor UIA config
# =======================
TITLE_CONTAINS = "Test Program"
POLL_MS = 50
RETRY_MS = 200
STATE_NAMES = {"Good", "Bad", "Testing","Attach then Start","Ready to Test",}


# =======================
#  Win32 overlay helpers (1 monitor)
# =======================
_user32 = ctypes.windll.user32

SM_CXSCREEN = 0
SM_CYSCREEN = 1

GWL_EXSTYLE = -20
WS_EX_NOACTIVATE = 0x08000000
WS_EX_TOOLWINDOW = 0x00000080

SWP_NOSIZE = 0x0001
SWP_NOMOVE = 0x0002
SWP_SHOWWINDOW = 0x0040
SWP_NOACTIVATE = 0x0010
HWND_TOPMOST = -1


def get_screen_geom():
    width = _user32.GetSystemMetrics(SM_CXSCREEN)
    height = _user32.GetSystemMetrics(SM_CYSCREEN)
    return 0, 0, width, height


def set_window_ex_noactivate(hwnd):
    try:
        GetPtr = getattr(_user32, "GetWindowLongPtrW", _user32.GetWindowLongW)
        SetPtr = getattr(_user32, "SetWindowLongPtrW", _user32.SetWindowLongW)
        ex = GetPtr(hwnd, GWL_EXSTYLE)
        new_ex = ex | WS_EX_NOACTIVATE | WS_EX_TOOLWINDOW
        SetPtr(hwnd, GWL_EXSTYLE, new_ex)
        _user32.SetWindowPos(
            hwnd, HWND_TOPMOST, 0, 0, 0, 0,
            SWP_NOMOVE | SWP_NOSIZE | SWP_SHOWWINDOW | SWP_NOACTIVATE
        )
    except Exception as e:
        print("set_window_ex_noactivate falló:", e)


# =======================
#  Fullscreen overlay (bloquea mouse cuando visible)
# =======================
class FullscreenOverlay:
    def __init__(self, root):
        self.root = root
        self.top = None
        self._visible = False

        # Timers para el cursor
        self._after_watch_id = None
        self._after_restore_id = None

        # Config del efecto
        self.watch_delay_ms = 500   # esperar 500ms después de show()
        self.watch_duration_ms = 2000  # durar 800ms con relojito

    def _cancel_cursor_timers(self):
        if self._after_watch_id is not None:
            try:
                self.root.after_cancel(self._after_watch_id)
            except Exception:
                pass
            self._after_watch_id = None

        if self._after_restore_id is not None:
            try:
                self.root.after_cancel(self._after_restore_id)
            except Exception:
                pass
            self._after_restore_id = None

    def _set_cursor_default(self):
        if self.top is not None:
            try:
                self.top.configure(cursor="")  # default del sistema/Tk
            except Exception:
                pass

    def _set_cursor_watch(self):
        if self.top is not None:
            try:
                self.top.configure(cursor="watch")  # “relojito” en Tk
            except Exception:
                pass

    def _schedule_watch_pulse(self):
        # Cancela pulsos previos si show() se llama de nuevo
        self._cancel_cursor_timers()
        self._set_cursor_default()

        def start_watch():
            # Si ya no está visible, no hagas nada
            if not self._visible or self.top is None:
                return

            self._set_cursor_watch()

            def restore():
                if not self._visible or self.top is None:
                    return
                self._set_cursor_default()

            self._after_restore_id = self.root.after(self.watch_duration_ms, restore)

        self._after_watch_id = self.root.after(self.watch_delay_ms, start_watch)

    def show(self):
        if self._visible:
            return

        left, top, width, height = get_screen_geom()
        self.top = tk.Toplevel(self.root)
        self.top.overrideredirect(True)
        self.top.geometry(f"{width}x{height}+{left}+{top}")
        self.top.attributes("-topmost", True)
        self.top.attributes("-alpha", 0.3)
        self.top.configure(bg="#000000")
        self.top.update_idletasks()

        try:
            hwnd = int(self.top.winfo_id())
            set_window_ex_noactivate(hwnd)
        except Exception as e:
            print("No pude aplicar estilos nativos al overlay:", e)

        self._visible = True

        # Pulso de relojito: +500ms, dura 800ms
        self._schedule_watch_pulse()

    def hide(self):
        if not self._visible:
            return

        # Asegura que no se quede el cursor “watch”
        self._cancel_cursor_timers()

        try:
            if self.top is not None:
                try:
                    self.top.configure(cursor="")
                except Exception:
                    pass
                self.top.destroy()
        except Exception:
            pass

        self.top = None
        self._visible = False
# =======================
#  UIA helper: buscar pane de estado
# =======================
def find_top_state_pane(win):
    candidates = []
    for p in win.descendants(control_type="Pane"):
        try:
            name = p.window_text()
            if name not in STATE_NAMES:
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


# =======================
#  Monitor thread (pywinauto/UIA)
# =======================
class UiMonitorThread(threading.Thread):
    """
    Hilo daemon: detecta estado (Good/Bad/Testing/Ready...) y lo manda a cola.
    Versión optimizada: NO re-scanea descendants a cada texto “raro”.
    """
    def __init__(self, out_q: queue.Queue):
        super().__init__(daemon=True)
        self._running = threading.Event()
        self._running.set()
        self.q = out_q

    def stop(self):
        self._running.clear()
        self.join(timeout=2)

    def connect_window(self):
        d = Desktop(backend="uia")
        w = d.window(title_re=f".*{TITLE_CONTAINS}.*")
        w.wait("exists", timeout=2)  # polling corto
        pid = w.element_info.process_id
        return w, pid

    def _push_state(self, s):
        """Mete estado a la cola; si está llena, la vacía y reintenta 1 vez."""
        try:
            self.q.put_nowait((s, None))
            return
        except queue.Full:
            try:
                while True:
                    self.q.get_nowait()
            except queue.Empty:
                pass
            try:
                self.q.put_nowait((s, None))
            except Exception:
                pass

    @staticmethod
    def _norm_text(t: str) -> str:
        t = (t or "").strip()
        # colapsa tabs/saltos de línea/espacios múltiples
        return " ".join(t.split())

    def run(self):
        initialized_here = False
        try:
            try:
                pythoncom.CoInitializeEx(pythoncom.COINIT_APARTMENTTHREADED)
                initialized_here = True
            except pythoncom.com_error as e:
                if getattr(e, "hresult", None) == getattr(pythoncom, "RPC_E_CHANGED_MODE", None):
                    print("Aviso: CoInitializeEx -> RPC_E_CHANGED_MODE. Continuando.")
                else:
                    print("Aviso: CoInitializeEx falló:", e)

            win = None
            pane = None
            last_state = None
            last_pid = None

            waiting_msg_printed = False
            misses = 0  # cuenta lecturas “no reconocidas” sin re-scan inmediato

            while self._running.is_set():
                try:
                    # 1) Conectar ventana
                    if win is None:
                        try:
                            win, last_pid = self.connect_window()
                            pane = None
                            last_state = None
                            misses = 0
                            if waiting_msg_printed:
                                print("[Monitor] Ventana encontrada. Iniciando monitoreo.")
                            waiting_msg_printed = False
                        except PwaTimeoutError:
                            if not waiting_msg_printed:
                                print("[Monitor] Esperando ventana de prueba...")
                                waiting_msg_printed = True
                            self._push_state(None)
                            time.sleep(RETRY_MS / 1000)
                            continue

                    # 2) Validar ventana + PID (si reinicia app)
                    try:
                        win.wait("exists", timeout=0.2)
                        current_pid = win.element_info.process_id
                        if last_pid is not None and current_pid != last_pid:
                            raise RuntimeError("PID changed (app restarted)")
                    except Exception:
                        win = None
                        pane = None
                        last_state = None
                        misses = 0
                        self._push_state(None)
                        time.sleep(RETRY_MS / 1000)
                        continue

                    # 3) Encontrar pane SOLO si no lo tenemos
                    if pane is None:
                        pane = find_top_state_pane(win)
                        if pane is None:
                            self._push_state(None)
                            time.sleep(RETRY_MS / 1000)
                            continue

                    # 4) Leer estado (normalizado)
                    s = self._norm_text(pane.window_text())

                    # Si es un estado esperado, todo bien
                    if s in STATE_NAMES:
                        misses = 0
                        if s != last_state:
                            self._push_state(s)
                            last_state = s
                        time.sleep(POLL_MS / 1000)
                        continue

                    # Si NO es esperado:
                    # - NO hagas pane=None inmediato (eso dispara descendants() y mete lag)
                    # - manda None para apagar overlay mientras está en pantalla “no test”
                    misses += 1
                    self._push_state(None)

                    # Solo después de varios misses, re-busca pane (por si cambió UI)
                    if misses >= 12:  # 12 * POLL_MS(100ms) ~= 1.2s, ajusta si quieres
                        pane = None
                        misses = 0

                    time.sleep(POLL_MS / 1000)

                except PwaTimeoutError:
                    win = None
                    pane = None
                    last_state = None
                    misses = 0
                    self._push_state(None)
                    time.sleep(RETRY_MS / 1000)

                except Exception as e:
                    # Aquí sí: el pane probablemente quedó stale => reset y reconectar
                    print("[Monitor] UIA stale / ventana cambió. Reintentando. Detalle:", e)
                    win = None
                    pane = None
                    last_state = None
                    misses = 0
                    self._push_state(None)
                    time.sleep(RETRY_MS / 1000)

        finally:
            if initialized_here:
                try:
                    pythoncom.CoUninitialize()
                except Exception as e:
                    print("Aviso: pythoncom.CoUninitialize falló:", e)
# =======================
#  Opcional: VBS toggle
# =======================
def toggle_vbs_script(event=None):
    global vbs_process
    tester_file_path = r"C:\Users\Public\Documents\Cirris\tester.txt"
    if os.path.exists(tester_file_path):
        with open(tester_file_path, 'r') as file:
            tester_name = file.read().strip()

        script_path = os.path.join(
            r"\\mlxgumvwfile01\Departamentos\Fakra\Pruebas\CyclesCounter",
            tester_name,
            "CyclesCounter.vbs"
        )

        if vbs_process is None:
            if os.path.exists(script_path):
                try:
                    vbs_process = subprocess.Popen(
                        ['cscript', script_path],
                        creationflags=subprocess.CREATE_NO_WINDOW
                    )
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


# =======================
#  Drag de ventana
# =======================
def on_title_bar_click(event):
    global root
    root.x_offset = event.x_root - root.winfo_x()
    root.y_offset = event.y_root - root.winfo_y()

def on_drag_motion(event):
    global root
    new_x = event.x_root - root.x_offset
    new_y = event.y_root - root.y_offset
    root.geometry(f"+{new_x}+{new_y}")


# =======================
#  Mantener foco
# =======================
#  Opcional: VBS toggle
# =======================
# =======================
#  Drag de ventana
# =======================
# =======================
#  Mantener foco
# =======================
def force_focus():
    """Mantiene la ventana principal al frente (autofocus)."""
    global root
    try:
        root.attributes("-topmost", True)
        root.update_idletasks()
        root.lift()
        root.focus_force()
        root.after(50, lambda: root.attributes("-topmost", False))
    except Exception as e:
        print(f"⚠️ Error forzando foco: {e}")

    root.after(1500, force_focus)
# =======================
#  Main
# =======================
def main():
    global root, vbs_process
    vbs_process = None

    root = tk.Tk()

    # ---- Log de excepciones de callbacks Tk ----
    import traceback
    def _tk_report_callback_exception(exc, val, tb):
        msg = "".join(traceback.format_exception(exc, val, tb))
        with open("tk_crash_log.txt", "a", encoding="utf-8") as f:
            f.write("\n\n=== TK CALLBACK EXCEPTION ===\n")
            f.write(msg)
        print(msg)
    root.report_callback_exception = _tk_report_callback_exception

    # ---- UI básica ----
    turno = get_turno_actual()
    root.protocol("WM_DELETE_WINDOW", lambda: print("❌ Botón cerrar deshabilitado"))
    root.title(f"Corrida Actual - Turno {turno}                 --First Pass Yield--")
    root.geometry("580x150+435+520")
    root.resizable(False, False)

    frame_color = "#E0E0E0"
    text_color = "#333333"
    root.configure(bg=frame_color)

    main_frame = tk.Frame(root, bg=frame_color)
    main_frame.pack(expand=True, fill="both", padx=10, pady=10)

    frame_left = tk.Frame(main_frame, bg=frame_color)
    frame_left.grid(row=0, column=0, sticky="w", padx=10)

    frame_right = tk.Frame(main_frame, bg=frame_color, width=320, height=120)
    frame_right.grid(row=0, column=1, sticky="n", padx=10)
    frame_right.grid_propagate(False)

    # Labels
    tk.Label(frame_left, text="Total Pzs:", font=("Arial", 16, "bold"),
             fg=text_color, bg=frame_color).grid(row=0, column=0, sticky="w", padx=10)
    label = tk.Label(frame_left, text="0", font=("Arial", 20, "bold"),
                     fg=text_color, bg=frame_color)
    label.grid(row=0, column=1, sticky="w", padx=5)

    tk.Label(frame_left, text="Fallas:", font=("Arial", 16, "bold"),
             fg=text_color, bg=frame_color).grid(row=1, column=0, sticky="w", padx=10)
    fail_label = tk.Label(frame_left, text="0", font=("Arial", 16, "bold"),
                          fg=text_color, bg=frame_color)
    fail_label.grid(row=1, column=1, sticky="w", padx=5)

    tk.Label(frame_left, text="Tiempo ciclo:", font=("Arial", 16, "bold"),
             fg=text_color, bg=frame_color).grid(row=2, column=0, sticky="w", padx=10)
    cycle_label = tk.Label(frame_left, text="N/A", font=("Arial", 16, "bold"),
                           fg=text_color, bg=frame_color)
    cycle_label.grid(row=2, column=1, sticky="w", padx=5)

    yield_text = tk.Label(frame_left, text="Yield:", font=("Arial", 16, "bold"),
                          fg=text_color, bg=frame_color)
    yield_text.grid(row=3, column=0, sticky="w", padx=10)
    yield_label = tk.Label(frame_left, text="100%", font=("Arial", 16, "bold"),
                           fg=text_color, bg=frame_color)
    yield_text.grid_remove()

    # Gauge
    gauge_frame = tk.Frame(frame_right, bg="lightgray", width=300, height=200)
    gauge_frame.place(x=10, y=-40)
    gauge_frame.pack_propagate(False)

    fig, ax = plt.subplots(figsize=(3, 1.6))
    canvas = FigureCanvasTkAgg(fig, master=gauge_frame)
    canvas.get_tk_widget().pack()

    # ---- Watchdog (conteo) ----
    file_handler = FileModifiedHandler(label, cycle_label, fail_label, yield_label, canvas, frame_color)
    observer = Observer()
    observer.schedule(file_handler, file_handler.folder_path, recursive=True)
    observer.start()

    force_focus()

    # ---- Overlay + monitor Good/Bad ----
    q = queue.Queue(maxsize=50)
    monitor = UiMonitorThread(q)
    overlay = FullscreenOverlay(root)
    current_state = None

    # Mejora #1: apagado robusto (incluye None) y hide forzado
    def process_queue():
        nonlocal current_state
        try:
            while True:
                s, _ = q.get_nowait()

                # None = ventana perdida / reconectando / app cerrada
                if s is None:
                    overlay.hide()
                    current_state = None
                    continue

                if s == "Good":
                    if current_state != "Good":
                        overlay.show()
                        print("Overlay ON (Good) -> bloqueando clics")
                else:
                    overlay.hide()  # fuerza apagado aunque current_state no sea Good
                    if current_state == "Good":
                        print("Overlay OFF (no Good)")

                current_state = s

        except queue.Empty:
            pass

        root.after(20, process_queue)

    monitor.start()
    root.after(20, process_queue)

    # Mejora #3: failsafe si el hilo muere
    def overlay_failsafe():
        if not monitor.is_alive():
            overlay.hide()
        root.after(1000, overlay_failsafe)

    root.after(1000, overlay_failsafe)

    # ---- binds opcionales ----
    root.bind("<Button-1>", on_title_bar_click)
    root.bind("<B1-Motion>", on_drag_motion)
    root.bind("6", toggle_vbs_script)

    # ---- cierre limpio ----
    def on_quit(event=None):
        print("Cerrando aplicación (Ctrl+M). Limpiando recursos...")
        try:
            observer.stop()
            observer.join(timeout=2)
        except Exception:
            pass
        try:
            monitor.stop()
        except Exception:
            pass
        try:
            overlay.hide()
        except Exception:
            pass
        try:
            root.destroy()
        except Exception:
            pass

    root.bind_all("<Control-m>", on_quit)

    try:
        root.mainloop()
    except KeyboardInterrupt:
        on_quit()

    # Limpieza extra por si sale por otra vía
    try:
        observer.stop()
        observer.join(timeout=2)
    except Exception:
        pass
    try:
        monitor.stop()
    except Exception:
        pass
    try:
        overlay.hide()
    except Exception:
        pass


if __name__ == "__main__":
    main()