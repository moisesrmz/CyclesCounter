import os
import tkinter as tk
from datetime import datetime, timedelta
import threading
import subprocess
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler


class FileModifiedHandler(FileSystemEventHandler):
    def __init__(self, label):
        self.label = label
        self.modification_count = 0
        self.reset_times = [datetime.strptime("06:30", "%H:%M").time(),
                            datetime.strptime("14:30", "%H:%M").time(),
                            datetime.strptime("22:00", "%H:%M").time()]

        # Agrega aquí la ruta de la carpeta que deseas monitorear
        self.folder_path = r"\\mlxgumvwfile01\Departamentos\Fakra\Pruebas\-PRN\EOL1"#\\mlxgumvwfile01\Departamentos\Fakra\Pruebas\-PRN\EOL1

        # Iniciar el temporizador para verificar la hora cada minuto
        self.start_timer()
        self.recently_incremented = False

    def start_timer(self):
        # Programar la función de verificación cada minuto
        threading.Timer(60, self.check_reset).start()

    def check_reset(self):
        current_time = datetime.now().time()

        # Verificar si la hora actual es una hora de reinicio
        if any(reset_time <= current_time < (datetime.combine(datetime.today(), reset_time) + timedelta(minutes=1)).time()
               for reset_time in self.reset_times):
            self.reset_counters()

        # Reiniciar el temporizador para la próxima verificación
        self.start_timer()

    def reset_counters(self):
        self.modification_count = 0
        # Establecer el color azul para el número mostrado
        self.label.config(text="Piezas probadas: 0", fg="white", bg="blue")
        # Imprimir en consola cada vez que se reinicia
        reset_time_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        print(f"Se reinició el conteo a las {reset_time_str}")

    def on_modified(self, event):
        if event.is_directory:
            return

        # Incrementar el contador cada vez que se modifica un archivo, pero solo si no se ha incrementado recientemente
        if not self.recently_incremented:
            self.modification_count += 1
            pieces_tested = self.modification_count
            self.label.config(text=f"Piezas probadas: {pieces_tested}", fg="white", bg="blue")
            self.recently_incremented = True
            threading.Timer(1, self.reset_increment_flag).start()

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

#def kill_easywire():
    # Usar taskkill para terminar el proceso easywire.exe
    #subprocess.call(['taskkill', '/F', '/IM', 'easywire.exe'])
    #print("easywire.exe terminado.")


def main():
    global root, vbs_process
    vbs_process = None  # Variable global para rastrear el proceso VBS
    root = tk.Tk()
    root.title("Monitor de Piezas Probadas")

    # Ajustar el tamaño de la ventana
    root.geometry("410x100+650+350")  # Iniciar a la derecha y un poco abajo

    # Eliminar el botón de cerrar
    root.overrideredirect(True)

    # Cambiar la fuente de las etiquetas
    label_font = ("Arial", 25)
    label = tk.Label(root, text="Piezas probadas: 0", font=label_font, padx=10, pady=10, fg="white", bg="blue")
    label.pack(pady=20)

    # Inicializar el manejador de modificaciones
    file_handler = FileModifiedHandler(label)

    # Configurar el observador de cambios en la carpeta
    observer = Observer()
    observer.schedule(file_handler, file_handler.folder_path, recursive=True)
    observer.start()

    # Permitir que el usuario mueva la ventana
    root.bind("<Button-1>", on_title_bar_click)
    root.bind("<B1-Motion>", on_drag_motion)

    # Vincular la tecla "6" a la función de alternar el script VBS
    root.bind("6", lambda event: toggle_vbs_script())
    # Vincular la tecla "2" a la función de terminar el proceso easywire.exe
    #root.bind("2", lambda event: kill_easywire())

    set_window_topmost(root)
    ensure_focus(root)

    root.mainloop()

    # Detener el observador cuando la aplicación se cierre
    observer.stop()
    observer.join()


if __name__ == "__main__":
    main()
