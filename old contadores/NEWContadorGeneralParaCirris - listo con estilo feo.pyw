import os
import time
import tkinter as tk
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
from matplotlib.patches import Arc, Circle

class FileModifiedHandler(FileSystemEventHandler):
    def __init__(self, label, cycle_label, fail_label, yield_label, canvas, frame_color):
        self.label = label
        self.cycle_label = cycle_label
        self.fail_label = fail_label  
        self.yield_label = yield_label  
        self.canvas = canvas  
        self.frame_color = frame_color  # ✅ Guardamos el color del frame en self.frame_color

        self.modification_count = 0
        self.fail_count = 0  
        self.yield_value = 100  
        self.timestamps = deque(maxlen=50)
        self.reset_times = [
            datetime.strptime("06:30", "%H:%M").time(),
            datetime.strptime("14:30", "%H:%M").time(),
            datetime.strptime("22:00", "%H:%M").time()
        ]
        
        self.folder_path = r"C:\Users\Public\Documents\Cirris\printer"
        self.recently_incremented = False
        self.start_timer()
        self.draw_gauge()


    def start_timer(self):
        threading.Timer(60, self.check_reset).start()

    def check_reset(self):
        current_time = datetime.now().time()
        if any(reset_time <= current_time < (datetime.combine(datetime.today(), reset_time) + timedelta(minutes=1)).time()
               for reset_time in self.reset_times):
            self.reset_counters()
        self.start_timer()

    def reset_counters(self):
        self.yield_value = 100  # Yield inicia en 100%
        self.yield_label.config(text="Yield: 100%")
        self.modification_count = 0
        self.fail_count = 0  
        self.timestamps.clear()
        self.label.config(text="Piezas probadas: 0")
        self.fail_label.config(text="Fallas: 0")  
        self.cycle_label.config(text="Tiempo ciclo promedio: N/A")
        print(f"🔄 Reinicio a las {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")

    def on_modified(self, event):
        """Se ejecuta cuando un archivo en la carpeta monitoreada es modificado."""
        self.update_yield()
        if event.is_directory:
            return

        # Solo procesar archivos .prn
        if not event.src_path.lower().endswith(".prn"):
            print(f"⚠️ Archivo ignorado: {event.src_path}")  # Debug: Ignorar archivos no PRN
            return

        print(f"📂 Archivo PRN modificado detectado: {event.src_path}")  # Debug: Solo archivos PRN

        if not self.recently_incremented:
            self.modification_count += 1
            self.timestamps.append(datetime.now())

            self.label.config(text=f"Piezas probadas: {self.modification_count}")
            self.update_cycle_time()

            # Confirmar que estamos revisando solo archivos .prn
            print(f"🔎 Revisando archivo PRN para fallas: {event.src_path}")  
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
            self.cycle_label.config(text=f"Tiempo ciclo promedio: {average_cycle_time:.2f} s")

    def check_for_failures(self, file_path):
        """Verifica si el archivo .prn contiene la cadena '^XA' en la primera línea para contar fallas."""
        if file_path.lower().endswith(".prn"):  
            attempts = 5  
            for attempt in range(attempts):
                try:
                    print(f"📂 Intentando abrir el archivo: {file_path} (Intento {attempt+1}/{attempts})")
                    with open(file_path, "r", encoding="utf-8") as file:
                        first_line = file.readline().strip()
                        print(f"📌 Primera línea leída: '{first_line}'")  # Debug: Verificar contenido

                        if "^XA" in first_line:
                            self.fail_count += 1
                            self.fail_label.config(text=f"Fallas: {self.fail_count}")
                            print(f"⚠️ Falla detectada en {file_path} (Total: {self.fail_count})")
                    return  
                except PermissionError:
                    print(f"🚨 Permiso denegado al acceder a {file_path}, reintentando ({attempt+1}/{attempts})...")
                    time.sleep(0.5)  
                except Exception as e:
                    print(f"⚠️ Error inesperado al leer {file_path}: {e}")
                    return  

    def reset_increment_flag(self):
        self.recently_incremented = False

    def update_yield(self):
        """Calcula y actualiza el Yield en el gauge."""
        if self.modification_count == 0:
            self.yield_value = 100  # Si no hay piezas probadas, el Yield es 100%
        else:
            self.yield_value = max(0, ((self.modification_count - self.fail_count) / self.modification_count) * 100)
        
        self.yield_label.config(text=f"Yield: {self.yield_value:.2f}%")
        
    # ✅ Usamos root.after() para evitar problemas de hilos con Matplotlib
        root.after(0, self.draw_gauge)


    def draw_gauge(self):
        """Dibuja un Gauge moderno y minimalista con la mitad superior del círculo."""

        # ✅ Cerrar la figura anterior antes de crear una nueva
        if hasattr(self.canvas, 'figure'):
            plt.close(self.canvas.figure)

        # 🔹 Crear una nueva figura de Matplotlib
        fig, ax = plt.subplots(figsize=(3, 2), facecolor=self.frame_color)

        # 🔹 Definir el ángulo del Yield (0 - 100% → 180° a 0° en la mitad superior)
        angle = np.interp(self.yield_value, [0, 100], [180, 0])

        # 🔹 Definir los colores por rango de Yield
        if self.yield_value > 90:
            gauge_color = "#00C853"  # Verde
        elif self.yield_value > 75:
            gauge_color = "#FFD600"  # Amarillo
        else:
            gauge_color = "#D32F2F"  # Rojo

        # 🔹 Dibujar arcos de colores usando `Arc`
        for start, end, color in [(180, 120, "#D32F2F"), (120, 60, "#FFD600"), (60, 0, "#00C853")]:
            arc = Arc((0, 0), 2, 2, theta1=start, theta2=end, color=color, lw=10, alpha=0.8)
            ax.add_patch(arc)

        # 🔹 Dibujar la aguja del Yield
        x = np.cos(np.radians(angle))
        y = np.sin(np.radians(angle))
        ax.plot([0, x], [0, y], color=gauge_color, lw=5, marker="o", markersize=8, alpha=0.9)

        # 🔹 Mostrar el porcentaje en el centro
        ax.text(0, -0.4, f"{self.yield_value:.1f}%", ha='center', va='center', fontsize=18, fontweight='bold', color=gauge_color)

        # 🔹 Agregar etiquetas de 0% y 100% en los extremos
        ax.text(-1.2, 0, "0%", ha='center', va='center', fontsize=12, fontweight='bold', color="#D32F2F")
        ax.text(1.2, 0, "100%", ha='center', va='center', fontsize=12, fontweight='bold', color="#00C853")

        # 🔹 Ocultar ejes
        ax.set_xticks([]), ax.set_yticks([])
        ax.spines['top'].set_visible(False), ax.spines['right'].set_visible(False)
        ax.spines['left'].set_visible(False), ax.spines['bottom'].set_visible(False)
        ax.set_xlim(-1.2, 1.2), ax.set_ylim(-0.2, 1.2)  # Ajustamos límites para mejorar la visualización

        # ✅ Actualizar la figura en Tkinter
        self.canvas.figure = fig
        self.canvas.draw()


def main():
    global root, vbs_process
    vbs_process = None

    root = tk.Tk()
    root.title("Monitor de Piezas Probadas")
    root.geometry("400x350+600+50")
    root.resizable(False, False)
    # root.overrideredirect(True)

    # 🌟 Estilos modernos
    bg_color = "#F5F5F5"       # Fondo principal (Blanco humo)
    text_color = "#333333"      # Texto principal (Negro suave)
    highlight_color = "#0078D7" # Azul brillante para resaltar
    frame_color = "#E0E0E0"     # Marco (Gris claro)
    fail_color = "#D32F2F"      # Rojo oscuro para fallas
    root.configure(bg=frame_color)

    frame = tk.Frame(root, bg=frame_color, bd=5, relief="ridge", highlightbackground="#CCCCCC", highlightthickness=2)
    frame.pack(expand=True, fill="both", padx=10, pady=10)

    label = tk.Label(frame, text="Piezas probadas: 0", font=("Arial", 18, "bold"),
                     fg=highlight_color, bg=frame_color)
    label.pack(pady=5)

    fail_label = tk.Label(frame, text="Fallas: 0", font=("Arial", 16, "bold"),
                          fg=fail_color, bg=frame_color)
    fail_label.pack(pady=5)

    cycle_label = tk.Label(frame, text="Tiempo ciclo promedio: N/A", font=("Arial", 14),
                           fg=text_color, bg=frame_color)
    cycle_label.pack(pady=5)

    yield_label = tk.Label(frame, text="Yield: 100%", font=("Arial", 16, "bold"),
                           fg="#0078D7", bg=frame_color)
    yield_label.pack(pady=5)

    # ✅ NUEVA SECCIÓN: Gauge con Matplotlib
    gauge_frame = tk.Frame(frame, bg=frame_color)
    gauge_frame.pack(pady=5)

    # Crear una figura de Matplotlib
    fig, ax = plt.subplots(figsize=(3, 1.6), subplot_kw={'projection': 'polar'})
    canvas = FigureCanvasTkAgg(fig, master=gauge_frame)
    canvas.get_tk_widget().pack()

    # ✅ Crear el objeto FileModifiedHandler con el nuevo canvas
    file_handler = FileModifiedHandler(label, cycle_label, fail_label, yield_label, canvas, frame_color)

    # ✅ Iniciar el observador ANTES de root.mainloop()
    observer = Observer()
    observer.schedule(file_handler, file_handler.folder_path, recursive=True)
    observer.start()

    # ✅ Mantener la ventana en primer plano
    force_focus()

    # ✅ Permitir mover la ventana con el mouse
    root.bind("<Button-1>", on_title_bar_click)
    root.bind("<B1-Motion>", on_drag_motion)

    # ✅ Asociar tecla para iniciar/detener VBScript
    root.bind("6", toggle_vbs_script)

    try:
        root.mainloop()
    except KeyboardInterrupt:
        print("🛑 Aplicación cerrada.")

    # ✅ Detener el observer cuando la ventana se cierra
    observer.stop()
    observer.join()

def toggle_vbs_script(event=None):
    """Inicia o detiene el VBScript al presionar la tecla '6'."""
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
                print("🛑 VBScript detenido.")
            except Exception as e:
                print(f"⚠️ Error al detener el script: {e}")

def force_focus():
    root.attributes('-topmost', True)
    root.after(500, force_focus)

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
    print("🛑 Aplicación cerrada por el usuario.")
    root.destroy()

if __name__ == "__main__":
    main()
