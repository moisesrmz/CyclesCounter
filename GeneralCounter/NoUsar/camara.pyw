import cv2
import time

def main():
    start_time = time.time()
    
    # Crear un objeto VideoCapture para acceder a la cámara (0 para la cámara predeterminada)
    print("Intentando abrir la cámara...")
    cap = cv2.VideoCapture(0)

    # Verificar si la cámara se abrió correctamente
    if not cap.isOpened():
        print("Error: No se pudo abrir la cámara.")
        return
    else:
        print(f"Cámara abierta en {time.time() - start_time:.2f} segundos.")

    # Bucle principal para capturar y mostrar el video
    while True:
        # Capturar un fotograma de la cámara
        ret, frame = cap.read()

        # Verificar si se capturó correctamente el fotograma
        if not ret:
            print("Error: No se pudo capturar el fotograma.")
            break

        # Mostrar el fotograma en una ventana
        cv2.imshow('Camera', frame)

        # Esperar 1 milisegundo y verificar si se presionó la tecla 'q' para salir del bucle
        if cv2.waitKey(1) & 0xFF == ord('q'):
            break

    # Liberar los recursos y cerrar la ventana
    cap.release()
    cv2.destroyAllWindows()

if __name__ == "__main__":
    main()
