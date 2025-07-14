import os
import tkinter as tk
from datetime import datetime, timedelta
import threading
import subprocess
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler

class FileModifiedHandler(FileSystemEventHandler):
    def __init__(self, label, cycle_label):
        self.label = label
        self.cycle_label = cycle_label
        self.modification_count = 0
        self.timestamps = []
        self.max_cycles = 50  # Limitar el cálculo a los últimos 50 ciclos
        self.reset_times = [
            datetime.strptime("06:30", "%H:%M").time(),
            datetime.strptime("14:30", "%H:%M").time(),
            datetime.strptime("22:00", "%H:%M").time()
        ]
        
        self.folder_path = r"C:\Users\Public\Documents\Cirris\printer"
        self.recently_incremented = False

        # Iniciar el temporizador para verificar la hora cada minuto
        self.start_timer()

    def start_timer(self):
        threading.Timer(60, self.check_reset).start()

    def check_reset(self):
        current_time = datetime.now().time()

        if any(reset_time <= current_time < (datetime.combine(datetime.today(), reset_time) + timedelta(minutes=1)).time()
               for reset_time in self.reset_times):
            self.reset_counters()

        self.start_timer()

    def reset_counters(self):
        self.modification_count = 0
        self.timestamps.clear()
        self.label.config(text="Piezas probadas: 0", fg="white", bg="blue")
        self.cycle_label.config(text="Tiempo ciclo promedio: N/A")
        reset_time_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        print(f"Se reinició el conteo a las {reset_time_str}")

    def on_modified(self, event):
        if event.is_directory:
            return

        if not self.recently_incremented:
            self.modification_count += 1
            self.timestamps.append(datetime.now())
            
            # Limitar la lista a los últimos 50 ciclos
            if len(self.timestamps) > self.max_cycles:
                self.timestamps.pop(0)
            
            pieces_tested = self.modification_count
            self.label.config(text=f"Piezas probadas: {pieces_tested}", fg="white", bg="blue")
            self.update_cycle_time()
            self.recently_incremented = True
            threading.Timer(1, self.reset_increment_flag).start()

    def update_cycle_time(self):
        if len(self.timestamps) > 1:
            total_time = sum(
                (self.timestamps[i] - self.timestamps[i - 1]).total_seconds()
                for i in range(1, len(self.timestamps))
            )
            average_cycle_time = total_time / (len(self.timestamps) - 1)
            self.cycle_label.config(text=f"Tiempo ciclo promedio: {average_cycle_time:.2f} s")
            print(f"Tiempo ciclo promedio: {average_cycle_time:.2f} s")

    def reset_increment_flag(self):
        self.recently_incremented = False

def set_window_topmost(root):
    root.lift()
    root.attributes('-topmost', True)
    root.after(500, lambda: set_window_topmost(root))

def ensure_focus(root):
    root.focus_force()
    root.after(500, lambda: ensure_focus(root))

def on_title_bar_click(event):
    root._drag_start_x = event.x
    root._drag_start_y = event.y

def on_drag_motion(event):
    x_win = root.winfo_x() + (event.x - root._drag_start_x)
    y_win = root.winfo_y() + (event.y - root._drag_start_y)
    root.geometry(f"{root.winfo_width()}x{root.winfo_height()}+{x_win}+{y_win}")

def toggle_vbs_script():
    global vbs_process
    tester_file_path = r"C:\Users\Public\Documents\Cirris\tester.txt"

    if os.path.exists(tester_file_path):
        with open(tester_file_path, 'r') as file:
            tester_name = file.read().strip()

        script_path = fr"\\mlxgumvwfile01\Departamentos\Fakra\Pruebas\CyclesCounter\{tester_name}\CyclesCounter.vbs"

        if vbs_process is None:
            if os.path.exists(script_path):
                vbs_process = subprocess.Popen(['cscript', script_path])
                print(f"{script_path} iniciado.")
            else:
                print(f"El script {script_path} no existe.")
        else:
            vbs_process.terminate()
            vbs_process = None
            print(f"{script_path} terminado.")
    else:
        print(f"El archivo {tester_file_path} no existe.")

def toggle_camera_script():
    global camera_process
    camera_script_path = r"\\mlxgumvwfile01\Departamentos\Fakra\Pruebas\CyclesCounter\GeneralCounter\camara.pyw"

    if camera_process is None:
        if os.path.exists(camera_script_path):
            camera_process = subprocess.Popen(['pythonw', camera_script_path])
            print(f"Script de la cámara {camera_script_path} iniciado.")
        else:
            print(f"El script {camera_script_path} no existe.")
    else:
        camera_process.terminate()
        camera_process = None
        print(f"Script de la cámara {camera_script_path} terminado.")

def main():
    global root, vbs_process, camera_process
    vbs_process = None
    camera_process = None
    root = tk.Tk()
    root.title("Monitor de Piezas Probadas")

    # Configurar el fondo de `root` para que coincida con el color del frame
    root.configure(bg="white")
    root.geometry("350x100+630+30")  # Ajustar el tamaño de la ventana
    root.overrideredirect(True)  # Desactivar la barra de título y botones de la ventana

    # Crear un Frame para ajustar el tamaño del fondo gris y darle más espacio a los labels
    frame = tk.Frame(root, bg="white", padx=10, pady=10)
    frame.pack()

    label_font = ("Arial", 25)
    label = tk.Label(frame, text="Piezas probadas: 0", font=label_font, fg="white", bg="blue", padx=5, pady=2)
    label.pack(pady=3)

    cycle_label = tk.Label(frame, text="Tiempo ciclo promedio: N/A", font=("Arial", 16), fg="white", bg="blue", padx=5, pady=2)
    cycle_label.pack(pady=3)

    file_handler = FileModifiedHandler(label, cycle_label)
    observer = Observer()
    observer.schedule(file_handler, file_handler.folder_path, recursive=True)
    observer.start()

    root.bind("<Button-1>", on_title_bar_click)
    root.bind("<B1-Motion>", on_drag_motion)

    root.bind("6", lambda event: toggle_vbs_script())
    root.bind("5", lambda event: toggle_camera_script())

    set_window_topmost(root)
    ensure_focus(root)

    root.mainloop()

    observer.stop()
    observer.join()

if __name__ == "__main__":
    main()
