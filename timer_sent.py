import win32com.client           # Para controlar Outlook desde Python
import pythoncom                # Para manejar la inicialización de COM en hilos
import time                     # Para manejar temporizadores
import tkinter as tk            # Para construir la interfaz gráfica
from tkinter import ttk, messagebox  # Widgets mejorados y mensajes de alerta
import threading                # Para ejecutar tareas en segundo plano
import sys                      # Para recibir argumentos desde línea de comandos
import os

# Evento global que controla si el envío está activo o detenido
enviar_event = threading.Event()

# Obtener el nombre de la cuenta (perfil Outlook) desde el argumento del script
cuenta_seleccionada = sys.argv[1] if len(sys.argv) > 1 else ""
print(f"Cuenta seleccionada: {cuenta_seleccionada}")

# Función para buscar la carpeta "Borradores" dentro del perfil seleccionado
def obtener_carpeta_borradores(namespace, cuenta):
    for folder in namespace.Folders:
        if folder.Name == cuenta:
            for posible_nombre in ["Borradores", "Drafts"]:  # Español e inglés
                try:
                    return folder.Folders[posible_nombre]
                except:
                    pass
            # Buscar en subcarpetas si no se encontró directamente
            for subfolder in folder.Folders:
                for posible_nombre in ["Borradores", "Drafts"]:
                    try:
                        return subfolder.Folders[posible_nombre]
                    except:
                        continue
    return None  # Si no se encuentra la carpeta

# Función que cuenta los borradores en la carpeta correspondiente
def contar_borradores(cuenta):
    pythoncom.CoInitialize()
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        drafts_folder = obtener_carpeta_borradores(namespace, cuenta)
        if drafts_folder:
            return drafts_folder.Items.Count
        else:
            messagebox.showerror("Error", "No se encontró la carpeta de borradores.")
            return 0
    except Exception as e:
        messagebox.showerror("Error al contar borradores", str(e))
        return 0
    finally:
        pythoncom.CoUninitialize()

# Habilita o deshabilita el botón de "Iniciar Envío" dependiendo del valor del intervalo
def validar_intervalo():
    intervalo_str = combo_intervalo.get()
    if intervalo_str == "Seleccione intervalo..." or not intervalo_str.isdigit():
        start_button.config(state="disabled")
    else:
        start_button.config(state="normal")

# Actualiza el contador de borradores y tiempo estimado cuando se selecciona intervalo
def actualizar_contador(event=None):
    intervalo_str = combo_intervalo.get()
    if intervalo_str == "Seleccione intervalo..." or not intervalo_str.isdigit():
        intervalo = 15
    else:
        intervalo = int(intervalo_str)
    combo_intervalo.set(str(intervalo))

    validar_intervalo()

    total_borradores = contar_borradores(cuenta_seleccionada)
    status_label.config(text=f"Borradores restantes: {total_borradores} | Enviados: 0")

    if total_borradores > 0:
        tiempo_total = intervalo * total_borradores
        horas, resto = divmod(tiempo_total, 3600)
        minutos, segundos = divmod(resto, 60)
        estimado_label.config(text=f"Tiempo total estimado: {horas:02}:{minutos:02}:{segundos:02}")
    else:
        estimado_label.config(text="Tiempo restante: 00:00:00")

# Muestra una cuenta regresiva estimada durante el envío de correos
def iniciar_temporizador_dinamico(tiempo_total):
    def actualizar_reloj():
        nonlocal tiempo_total
        if tiempo_total <= 0 or not enviar_event.is_set():
            estimado_label.config(text="Tiempo restante: 00:00:00")
            return
        horas, resto = divmod(tiempo_total, 3600)
        minutos, segundos = divmod(resto, 60)
        estimado_label.config(text=f"Tiempo restante: {horas:02}:{minutos:02}:{segundos:02}")
        tiempo_total -= 1
        root.after(1000, actualizar_reloj)

    actualizar_reloj()

# Envía los borradores uno a uno con intervalo definido
def enviar_borradores(cuenta, status_label):
    enviar_event.set()
    intervalo = int(combo_intervalo.get())
    pythoncom.CoInitialize()
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        drafts_folder = obtener_carpeta_borradores(namespace, cuenta)

        if not drafts_folder:
            messagebox.showerror("Error", "No se encontró la carpeta de borradores.")
            return

        total_borradores = drafts_folder.Items.Count
        enviados = 0

        time.sleep(intervalo)  # Espera inicial antes de enviar el primer correo

        while enviados < total_borradores and enviar_event.is_set():
            item = drafts_folder.Items.GetFirst()
            if not item:
                break

            try:
                destinatarios_no_resueltos = [r.Name for r in item.Recipients if not r.Resolve()]
                if destinatarios_no_resueltos:
                    raise Exception(f"Destinatarios no resueltos: {', '.join(destinatarios_no_resueltos)}")

                item.Send()
                enviados += 1

            except Exception as e:
                messagebox.showerror("Error en borrador", f"Error en el borrador #{enviados + 1}: {e}")

            restantes = total_borradores - enviados
            status_label.config(text=f"Borradores restantes: {restantes} | Enviados: {enviados}")
            time.sleep(intervalo)

    except Exception as e:
        messagebox.showerror("Error general", f"{e}")

    finally:
        status_label.config(text="Proceso finalizado")
        estimado_label.config(text="Tiempo restante: 00:00:00")
        enviar_event.clear()

        if enviados > 0:
            print(f"{enviados}")

        pythoncom.CoUninitialize()

# Inicia el envío en segundo plano
def iniciar_envio():
    total_borradores = contar_borradores(cuenta_seleccionada)
    intervalo = int(combo_intervalo.get())
    if total_borradores > 0:
        enviar_event.set()
        tiempo_total = intervalo * total_borradores
        threading.Thread(target=lambda: iniciar_temporizador_dinamico(tiempo_total), daemon=True).start()
        threading.Thread(target=enviar_borradores, args=(cuenta_seleccionada, status_label), daemon=True).start()

# Detiene el envío si está activo
def detener_envio():
    enviar_event.clear()
    status_label.config(text="Envío detenido")
    estimado_label.config(text="Tiempo restante: --")

# =========================
# INICIO DE INTERFAZ GRÁFICA
# =========================

root = tk.Tk()
root.title("Enviar Borradores Outlook")
root.geometry("400x280")

# Establecer ícono de forma dinámica
try:
    base_path = getattr(sys, '_MEIPASS', os.path.abspath("."))
    icon_path = os.path.join(base_path, "config", "icono.ico")
    root.iconbitmap(icon_path)
except Exception as e:
    print(f"No se pudo cargar el icono: {e}")

label_cuenta = tk.Label(root, text=f"Correo seleccionado: {cuenta_seleccionada}", font=("Arial", 12), fg="green")
label_cuenta.pack(pady=5)

label_intervalo = tk.Label(root, text="Intervalo de envío (segundos):", font=("Arial", 12))
label_intervalo.pack(pady=5)

combo_intervalo = ttk.Combobox(root, values=["Seleccione intervalo...", "60",  "120", "180"], state="readonly", font=("Arial", 10), width=20)
combo_intervalo.bind("<<ComboboxSelected>>", actualizar_contador)
combo_intervalo.pack(pady=5)
combo_intervalo.current(0)

start_button = tk.Button(root, text="Iniciar Envío", command=iniciar_envio, font=("Arial", 12), bg="lightgreen", state="disabled")
start_button.pack(pady=5)

stop_button = tk.Button(root, text="Detener Envío", command=detener_envio, font=("Arial", 10), bg="red")
stop_button.pack(pady=5)

status_label = tk.Label(root, text="", font=("Arial", 14))
status_label.pack(pady=5)

estimado_label = tk.Label(root, text="Tiempo restante: --", font=("Arial", 12), fg="blue")
estimado_label.pack(pady=5)

validar_intervalo()

root.mainloop()