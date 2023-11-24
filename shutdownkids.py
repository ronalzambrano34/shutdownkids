from win32com.client import Dispatch
import tkinter as tk
from datetime import datetime
import time
import pystray
import threading
import sys
import subprocess
from PIL import Image
from pystray import MenuItem as item
import config
# from plyer import notification
# from win10toast import ToastNotifier
from notifypy import Notify

# Iniciar el contador en segundo plano
running = True
hora_militar = config.hora_militar
img = "./_internal/icon.ico"


######################### addink ################################
def addink():
    import os

    def crear_acceso_directo(ruta_archivo, ruta_destino, nombre_acceso_directo):
        shell = Dispatch('WScript.Shell')
        acceso_directo = shell.CreateShortCut(os.path.join(
            ruta_destino, nombre_acceso_directo + '.lnk'))
        acceso_directo.Targetpath = ruta_archivo
        acceso_directo.WorkingDirectory = os.path.dirname(ruta_archivo)
        acceso_directo.save()

    # Ejemplo de uso
    archivo = 'shutdownkids.exe'
    ruta_carpeta_raiz = os.path.dirname(os.path.abspath(__file__))
    print("ruta run " + ruta_carpeta_raiz)
    ruta_carpeta_raiz = os.path.dirname(ruta_carpeta_raiz)
    ruta_archivo = os.path.join(ruta_carpeta_raiz, archivo)

    print("ruta ink " + ruta_archivo)
    ruta_destino = os.path.join(
        os.environ['userprofile'], 'AppData\\Roaming\\Microsoft\\Windows\\Start Menu\\Programs\\Startup')
    nombre_acceso_directo = 'shutdownkids'

    crear_acceso_directo(ruta_archivo, ruta_destino, nombre_acceso_directo)
#################################################################


######################### Ventana ###############################
def ventana():
    # Función que se ejecuta cuando se hace clic en el botón
    def on_button_click(event=None):
        contraseña = entry.get()
        if contraseña == config.psw:
            # Ejecutar un comando en CMD
            comando = "shutdown -a"  # Reemplaza "dir" con el comando que desees ejecutar
            resultado = subprocess.run(comando, capture_output=True, text=True)

            print(resultado.stdout)
            print("Contraseña correcta")
            root.destroy()  # Cierra la ventana
        else:
            sms("Contraseña incorrecta")
            print("Contraseña incorrecta")
            entry.delete(0, tk.END)  # Borra el contenido del campo de entrada

    # Crear la ventana
    root = tk.Tk(className="Contraseña.")

    # Ajustar el margen de la ventana
    root.configure(padx=40, pady=10)

    # Crea una etiqueta label
    label = tk.Label(root, text="Ingrese Contraseña")
    label.pack()

    # Crear el campo de entrada
    entry = tk.Entry(root, show="*", borderwidth=10)
    entry.pack()
    entry.focus_set()  # Seleccionar automáticamente el campo de entrada al ejecutar la ventana

    # Asociar el evento de presionar la tecla Enter directamente con la función on_button_click()
    entry.bind("<Return>", on_button_click)

    # Crear el botón
    button = tk.Button(root, text="Aceptar", command=on_button_click)
    button.pack()
    
    # root.lift()  # Elevar la ventana al primer plano
    # root.attributes("-topmost", True)
    root.mainloop()
#################################################################


########################### Noti ################################
def sms(text):
    print("notification")
    notification = Notify(
        default_notification_title="Apagado programado:",
        default_application_name="Great Application",
        default_notification_icon="./_internal/icon.ico",
        # default_notification_audio="path/to/sound.wav"
    )
    notification.message = text
    notification.send()

    # notification.notify(
    #     title="Apagado programado:",
    #     message=str(text),
    #     timeout=5,  # Duración de la notificación en segundos
    #     app_icon="./_internal/icon.ico",
    # )

    # toaster = ToastNotifier()
    # toaster.show_toast(str(text), " ")
    # toaster.show_toast(str(text), " ", "./_internal/icon.ico")
#################################################################

# Función para obtener la diferencia de tiempo
def get_diferencia():
    hora_objetivo = datetime.strptime(hora_militar, "%H:%M")  # Hora objetivo
    hora_actual_aux = datetime.strftime(datetime.now(), "%H:%M")  # Hora actual
    hora_actual = datetime.strptime(hora_actual_aux, "%H:%M")
    print("hora objetivo: ", hora_objetivo)

    if hora_actual < hora_objetivo:
        diferencia = hora_objetivo - hora_actual
        print("ok")
        return diferencia
    elif hora_actual == hora_objetivo:
        on_shutdown()
    else:
        print("no")
        # on_exit_clicked(icon, None)
        return None


# Función para manejar el clic en el botón de salida
def on_exit_clicked(icon, item):
    global running
    running = False
    if icon is not None:
        icon.stop()
    print("detenido")
    sys.exit()


def on_sms():
    # noti.sms(str(text)) # no puedo usarlo pq se cierra el icono :()
    sms(str(text))


def on_runf():
    print("running off")
    global running
    running = False


# Función para abrir la GUI ventana
def open_win():
    print("open_win")  # Mostrar la ventana
    ventana()
    # subprocess.run(["pythonw", "ventana.py"])

# Función para apagar PC


def on_shutdown():
    print("on_shutdown")
    # Ejecutar un comando en CMD
    comando = "shutdown -s"  # Reemplaza "dir" con el comando que desees ejecutar
    subprocess.run(comando, capture_output=True, text=True)
    open_win()

# Función para cancelar apagado


def on_shutdown_off():
    print("shutdown_off")
    # Ejecutar un comando en CMD
    comando = "shutdown -a"  # Reemplaza "dir" con el comando que desees ejecutar
    subprocess.run(comando, capture_output=True, text=True)


def on_start():
    print("on_start")
    # Ejecutar el programa en segundo plano utilizando pythonw
    # subprocess.Popen(["pythonw", "addink.py"])
    addink()
    # Cerrar la consola
    # subprocess.call("taskkill /IM cmd.exe /F", shell=True)

# Función para actualizar el menú de la bandeja del sistema


def update_menu(text="algo"):
    icon.menu = (
        item("Apagado: " + hora_militar, lambda: None),
        item(str(text), on_sms),
        # item("SMS", on_sms),
        item("-" * 20, lambda: None),  # Separador visual con guiones
        item(f"Cancelar Apagado", open_win),
        # item(f"Ejecutar al inicio", on_start),
        # item("-" * 20, lambda: None),  # Separador visual con guiones
        # item(f"RUNNING OFF", on_runf),
        item(f"EXIT", on_exit_clicked),
    )
    icon.title = str(text)


# Función para ejecutar el ícono en segundo plano
def run_icon():
    icon.run()


# Crear un ícono en la bandeja del sistema
menu = (
    item("algo", on_sms),
    item("Salir", on_exit_clicked)
)

# image = Image.open("./_internal/icon.ico")  # Ruta a la imagen del ícono
image = Image.open("./_internal/icon.ico")  # Ruta a la imagen del ícono
# Crear el ícono de la bandeja del sistema
icon = pystray.Icon(title="Apagado programado en:",
                    icon=image, name="name", menu=menu)


# Crear y ejecutar el hilo para el ícono de la bandeja del sistema
icon_thread = threading.Thread(target=run_icon)
icon_thread.start()

# para ejecutar el programa al inicio del sistema
on_start()

while running:
    diferencia = get_diferencia()
    if diferencia is not None:
        contador = diferencia.seconds
        while contador > 0 and running:
            horas = contador // 3600 % 24
            minutos = contador // 60 % 60
            text = "Faltan " + str(horas) + " horas y " + str(minutos) + " min"
            print(text)
            if minutos == 0:
                on_sms()
            update_menu(text)
            time.sleep(30)  # Pausa de 5 segundos
            # contador -= 30
            contador = get_diferencia().seconds                     
    else:
        running = False


print("Ejecución finalizada")
on_shutdown()


# Esperar a que el hilo del ícono termine antes de finalizar el programa
icon_thread.join()
