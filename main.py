import win32com.client
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from PIL import Image, ImageTk
import subprocess
import os
import sys
import time
import winreg
import tempfile
import shutil
from typing import List, Optional, Any

cuenta_seleccionada: Optional[str] = None
ruta_excel: Optional[str] = None
ruta_docx: Optional[str] = None

# GUI
root = tk.Tk()
root.title("Creador de Borradores")  # Título solo en la barra de la ventana
root.geometry("480x410")

# Establecer icono y logo
base_path = getattr(sys, '_MEIPASS', os.path.abspath("."))

# Establecer ícono de la aplicación
try:
    icon_path = os.path.join(base_path, "config", "icono.ico")
    root.iconbitmap(icon_path)
except (FileNotFoundError, OSError) as e:
    print(f"No se pudo cargar el icono: {e}")

# Cargar imagen de portada
try:
    cover_path = os.path.join(base_path, "config", "cover_borradores.jpg")
    cover_image = Image.open(cover_path)
    cover_image = cover_image.resize((500, 90))
    cover_img = ImageTk.PhotoImage(cover_image)
    etiqueta_cover = tk.Label(root, image=cover_img)
    etiqueta_cover.image = cover_img  # Para evitar que sea eliminada por el recolector de basura
    etiqueta_cover.pack(pady=10)
except Exception as e:
    print(f"No se pudo cargar la imagen de portada: {e}")

# Variables para mostrar nombres de archivos (deben ir después de crear root)
ruta_excel_var = tk.StringVar()
ruta_docx_var = tk.StringVar()

def obtener_perfiles_outlook() -> List[str]:
    perfiles = ["Seleccione perfil..."]
    try:
        office_versions = ["16.0", "15.0", "14.0"]
        for version in office_versions:
            path = fr"Software\\Microsoft\\Office\\{version}\\Outlook\\Profiles"
            try:
                with winreg.OpenKey(winreg.HKEY_CURRENT_USER, path) as key:
                    i = 0
                    while True:
                        try:
                            perfil = winreg.EnumKey(key, i)
                            perfiles.append(perfil)
                            i += 1
                        except OSError:
                            break
                break
            except FileNotFoundError:
                continue
        if len(perfiles) == 1:
            perfiles.append("No se encontraron perfiles.")
    except Exception as e:
        perfiles.append("Error al obtener perfiles")
    return perfiles

def cerrar_outlook() -> None:
    subprocess.run("taskkill /F /IM outlook.exe", shell=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)

def obtener_ruta_outlook() -> str:
    rutas = [
        r"C:\\Program Files\\Microsoft Office\\root\\Office16\\OUTLOOK.EXE",
        r"C:\\Program Files (x86)\\Microsoft Office\\Office16\\OUTLOOK.EXE",
        r"C:\\Program Files\\Microsoft Office\\Office16\\OUTLOOK.EXE"
    ]
    for ruta in rutas:
        if os.path.exists(ruta):
            return ruta
    raise FileNotFoundError("No se encontró Outlook.")

def iniciar_outlook_con_perfil(perfil: str) -> None:
    try:
        ruta_outlook = obtener_ruta_outlook()
        subprocess.Popen([ruta_outlook, "/profile", perfil])
        time.sleep(7)
    except Exception as e:
        print(f"No se pudo iniciar Outlook: {e}")

def obtener_cuentas_activas(max_intentos: int = 10, intervalo: int = 1) -> List[str]:
    for intento in range(max_intentos):
        try:
            outlook = win32com.client.Dispatch("Outlook.Application")
            namespace = outlook.GetNamespace("MAPI")
            cuentas = [account.SmtpAddress for account in namespace.Accounts]
            if cuentas:
                return cuentas
        except Exception:  # Considera capturar errores más específicos
            pass
        time.sleep(intervalo)
    return []

def cuenta_asociada_seleccionada(_event: Any) -> None:  # El parámetro 'event' no se usa, se renombra como '_event'
    global cuenta_seleccionada
    cuenta_seleccionada = combo_cuentas_asociadas.get()

def mostrar_cuenta_seleccionada(_event: Any) -> None:  # El parámetro 'event' no se usa, se renombra como '_event'
    global cuenta_seleccionada
    perfil = combo_cuentas.get()
    cuenta_seleccionada = None
    ruta_excel_var.set("")
    ruta_docx_var.set("")
    combo_cuentas_asociadas.pack_forget()
    combo_cuentas_asociadas.set("")

    if perfil == "Seleccione perfil...":
        label_cuenta_var.set("")
        label_cuenta.config(fg="black")
        return

    cerrar_outlook()
    iniciar_outlook_con_perfil(perfil)

    cuentas = obtener_cuentas_activas()
    if not cuentas:
        label_cuenta_var.set("No se encontraron cuentas.")
        label_cuenta.config(fg="red")
    elif len(cuentas) == 1:
        cuenta_seleccionada = cuentas[0]
        label_cuenta_var.set(f"{cuenta_seleccionada}... Listo   ✔️")
        label_cuenta.config(font=("Arial", 12, "bold"))
    else:
        cuenta_seleccionada = cuentas[0]
        label_cuenta_var.set("Selecciona una cuenta:")
        label_cuenta.config(fg="blue")
        combo_cuentas_asociadas['values'] = cuentas
        combo_cuentas_asociadas.current(0)
        combo_cuentas_asociadas.pack()

def ruta_script(nombre_script: str) -> str:
    base_path = getattr(sys, '_MEIPASS', os.path.dirname(__file__))
    return os.path.join(base_path, nombre_script)

def ejecutar_script(nombre_script_txt: str, perfil: str, mostrar_mensaje: bool = False) -> None:
    if perfil == "Seleccione perfil...":
        messagebox.showerror("Error", "Selecciona un perfil.")
        return

    global cuenta_seleccionada, ruta_excel, ruta_docx
    if not cuenta_seleccionada:
        messagebox.showerror("Error", "Selecciona una cuenta asociada.")
        return

    solo_envio = "timer_sent" in nombre_script_txt.lower()

    if not solo_envio and (not ruta_excel or not ruta_docx):
        messagebox.showerror("Error", "Carga Excel y Word antes de continuar.")
        return

    try:
        script_name = ""
        args = []

        if "txt_1" in nombre_script_txt.lower():
            script_name = "txt_1"
            args = [cuenta_seleccionada, perfil, ruta_excel, ruta_docx]
        elif "timer_sent" in nombre_script_txt.lower():
            script_name = "timer_sent"
            args = [cuenta_seleccionada]

        # Detectar entorno
        if getattr(sys, 'frozen', False):
            base_path = os.path.dirname(sys.executable)
            script_path = os.path.join(base_path, f"{script_name}.exe")
            cmd = [script_path] + args
        else:
            base_path = os.path.dirname(__file__)
            script_path = os.path.join(base_path, f"{script_name}.py")
            cmd = [sys.executable, script_path] + args

        if not os.path.exists(script_path):
            raise FileNotFoundError(f"No se encontró el archivo: {script_path}")

        result = subprocess.run(cmd, capture_output=True, text=True)
        print(result.stdout)

        if result.returncode != 0:
            messagebox.showerror("Error", f"Error:\n{result.stderr}")
        else:
            if mostrar_mensaje:
                if "timer_sent" in script_name:
                    cantidad = result.stdout.strip()
                    if cantidad.isdigit() and int(cantidad) > 0:
                        messagebox.showinfo("Éxito",
                            f"Fueron enviados {cantidad} borradores del perfil '{perfil}' con éxito.")
                else:
                    messagebox.showinfo("Éxito", f"{script_name}.py/.exe ejecutado correctamente.")

    except Exception as e:
        messagebox.showerror("Error", f"No se pudo ejecutar {script_name}:\n{e}")

def cargar_excel() -> None:
    global ruta_excel
    archivo = filedialog.askopenfilename(filetypes=[("Archivos Excel", "*.xlsx *.xls *xlsm")])
    if archivo:
        ruta_excel = archivo
        nombre_archivo = os.path.basename(archivo)
        ruta_excel_var.set(f"... {nombre_archivo}   ✔️")

def cargar_docx() -> None:
    global ruta_docx
    archivo = filedialog.askopenfilename(filetypes=[("Documentos de Word", "*.docx")])
    if archivo:
        ruta_docx = archivo
        nombre_archivo = os.path.basename(archivo)
        ruta_docx_var.set(f"... {nombre_archivo}   ✔️")

def ejecutar_timer_send() -> None:
    perfil = combo_cuentas.get()
    if perfil == "Seleccione perfil...":
        messagebox.showerror("Error", "Selecciona un perfil.")
        return
    if not cuenta_seleccionada:
        messagebox.showerror("Error", "Selecciona una cuenta asociada.")
        return

    try:
        # Detectar entorno: empaquetado vs desarrollo
        if getattr(sys, 'frozen', False):
            exe_dir = os.path.dirname(sys.executable)
        else:
            exe_dir = os.path.dirname(__file__)

        exe_path = os.path.join(exe_dir, "timer_sent.exe")

        if not os.path.exists(exe_path):
            raise FileNotFoundError(f"No se encontró el archivo: {exe_path}")

        # Agrega log temporal
        print(">>> Ejecutando:", [exe_path, cuenta_seleccionada])

        result = subprocess.run([exe_path, cuenta_seleccionada], capture_output=True, text=True)

        if result.returncode != 0:
            messagebox.showerror("Error", f"Error ejecutando timer_sent.exe:\n{result.stderr}")
        else:
            enviados = result.stdout.strip()
            if enviados.isdigit():
                messagebox.showinfo("Éxito", f"Se enviaron {enviados} borradores correctamente.")
            else:
                messagebox.showinfo("Ejecutado", f"Salida:\n{result.stdout}")

    except Exception as e:
        messagebox.showerror("Error", f"No se pudo ejecutar timer_sent:\n{e}")

def crear_boton(nombre_script: str, texto_boton: str) -> tk.Button:
    return tk.Button(root, text=texto_boton,
        command=lambda: ejecutar_script(nombre_script, combo_cuentas.get(), False),
        font=("Arial", 10), padx=10, pady=5)

perfiles = obtener_perfiles_outlook()
max_length = max([len(perfil) for perfil in perfiles], default=20)
width_combobox = min(max_length + 2, 50)

tk.Label(root, text="Selecciona un perfil de Outlook:", font=("Arial", 10, "bold"), anchor="w").pack(anchor="w", padx=10)

frame_combo_perfiles = tk.Frame(root)
frame_combo_perfiles.pack(anchor="w", pady=5, padx=10)

combo_cuentas = ttk.Combobox(frame_combo_perfiles, values=perfiles, state="readonly", font=("Arial", 10), width=width_combobox)
combo_cuentas.bind("<<ComboboxSelected>>", mostrar_cuenta_seleccionada)
combo_cuentas.pack(side="left")
combo_cuentas.current(0)

label_cuenta_var = tk.StringVar()
label_cuenta = tk.Label(root, textvariable=label_cuenta_var, font=("Arial", 10))
label_cuenta.pack(pady=5)

combo_cuentas_asociadas = ttk.Combobox(root, state="readonly", font=("Arial", 10))
combo_cuentas_asociadas.bind("<<ComboboxSelected>>", cuenta_asociada_seleccionada)
combo_cuentas_asociadas.pack(pady=5)
combo_cuentas_asociadas.pack_forget()

frame_excel = tk.Frame(root)
frame_excel.pack(anchor="w", pady=5, padx=10)
btn_excel = tk.Button(frame_excel, text="Cargar Excel", command=cargar_excel,
                      font=("Arial", 10), padx=10, pady=5)

btn_excel.pack(side="left")
label_excel = tk.Label(frame_excel, textvariable=ruta_excel_var, font=("Arial", 9), wraplength=300, justify="left", fg="green")
label_excel.pack(side="left", padx=10)

frame_docx = tk.Frame(root)
frame_docx.pack(anchor="w", pady=5, padx=10)
btn_docx = tk.Button(frame_docx, text="Cargar Texto Mail", command=cargar_docx,
                     font=("Arial", 10), padx=10, pady=5)
btn_docx.pack(side="left")
label_docx = tk.Label(frame_docx, textvariable=ruta_docx_var, font=("Arial", 9), wraplength=300, justify="left", fg="green")
label_docx.pack(side="left", padx=10)

frame_crear = tk.Frame(root)
frame_crear.pack(anchor="w", pady=5, padx=10)
tk.Button(frame_crear, text="Crear Borradores",
    command=lambda: ejecutar_script("txt_1", combo_cuentas.get(), False),
    font=("Arial", 10), padx=10, pady=5).pack(side="left")

tk.Button(root, text="Enviar Borradores", command=ejecutar_timer_send,
          font=("Arial", 12), padx=10, pady=5, bg="purple", fg="white").pack(pady=10)

root.mainloop()