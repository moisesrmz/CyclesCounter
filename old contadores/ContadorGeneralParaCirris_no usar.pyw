import os
import tkinter as tk
import threading
import cv2
from datetime import datetime, timedelta
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler


class FileModifiedHandler(FileSystemEventHandler):
    def __init__(self, label):
        self.label = label
        self.modification_count = 0
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
        self.label.config(text="Piezas probadas: 0", fg="white", bg="blue")
        reset_time_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        print(f"Se reinició el conteo a las {reset_time_str}")

    def on_modified(self, event):
        if event.is_directory:
            return

        if not self.recently_incremented:
            self.modification_count += 1
            pieces_tested = self.modification_count
            self.label.config(text=f"Piezas probadas: {pieces_tested}", fg="white", bg="blue")
            self.recently_incremented = True
            threading.Timer(1, self.reset_increment_flag).start()

    def reset_increment_flag(self):
        self.recently_incremented = False


class CameraHandler:
    def __init__(self):
        self.camera_process = None

    def toggle_camera(self):
        if self.camera_process is None:
            self.camera_process = threading.Thread(target=self.start_camera)
            self.camera_process.start()
            print("Cámara iniciada.")
        else:
            self.camera_process = None
            print("Cámara terminada.")

    def start_camera(self):
        cap = cv2.VideoCapture(0)
        while self.camera_process is not None:
            ret, frame = cap.read()
            cv2.imshow('Camera', frame)
            if cv2.waitKey(1) & 0xFF == ord('5'):
                break
        cap.release()
        cv2.destroyAllWindows()


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


def main():
    global root, vbs_process, camera_handler
    vbs_process = None
    root = tk.Tk()
    root.title("Monitor de Piezas Probadas")

    camera_handler = CameraHandler()

    label_font = ("Arial", 25)
    label = tk.Label(root, text="Piezas probadas: 0", font=label_font, padx=10, pady=10, fg="white", bg="blue")
    label.pack(pady=20)

    file_handler = FileModifiedHandler(label)
    observer = Observer()
    observer.schedule(file_handler, file_handler.folder_path, recursive=True)
    observer.start()

    button = tk.Button(root, text="Toggle Camera", command=camera_handler.toggle_camera)
    button.pack(pady=20)

    root.bind("<Button-1>", on_title_bar_click)
    root.bind("<B1-Motion>", on_drag_motion)

    root.bind("6", lambda event: toggle_vbs_script())

    set_window_topmost(root)
    ensure_focus(root)

    root.mainloop()

    observer.stop()
    observer.join()


if __name__ == "__main__":
    main()
