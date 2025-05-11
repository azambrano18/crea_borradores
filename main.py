# Librerías estándar
import os
import sys
import time
import json
import logging
import subprocess
import urllib.request
import winreg
#verifica que el .docx tenga texto
from docx import Document
# Tipado y otros
from typing import List, Optional, Any
# Interfaz gráfica
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
# Imágenes
from PIL import Image, ImageTk
# Automatización con Outlook
import win32com.client

# Versión actual del programa
__version__ = "1.0.2"

# Configuración del log de errores
logging.basicConfig(
    filename='errores_main.log',
    level=logging.ERROR,
    format='%(asctime)s - %(levelname)s - %(message)s'
)
# Verificación de actualización desde GitHub
def verificar_actualizacion(forzar: bool = False):
    try:
        url_api = "https://api.github.com/repos/azambrano18/crea_borradores/releases/latest"
        with urllib.request.urlopen(url_api) as response:
            data = json.loads(response.read())
            ultima_version = data["tag_name"].lstrip("v")
            assets = data["assets"]

            # Mostrar barra de progreso
            barra_progreso["value"] = 0
            porcentaje_var.set("0%")
            barra_progreso.pack(side="left", padx=(0, 10))
            etiqueta_porcentaje.pack(side="left")
            frame_progreso.pack(side="bottom", fill="x", padx=10, pady=5)
            root.update_idletasks()

        if ultima_version != __version__ or forzar:
            if messagebox.askyesno("Actualización disponible",
                f"Hay una nueva versión ({ultima_version}) disponible.\n¿Deseas descargarla ahora?"):

                exe_dir = os.path.dirname(sys.executable)
                archivos = {
                    "main.exe": "CreadorBorradores_Nuevo.exe",
                    "txt_1.exe": "txt_1_nuevo.exe",
                    "timer_sent.exe": "timer_sent_nuevo.exe"
                }

                descargas = [asset for asset in assets if asset["name"] in archivos]
                avance_por_archivo = 100 // len(descargas)
                base = 0

                for asset in descargas:
                    nombre = asset["name"]
                    url = asset["browser_download_url"]
                    destino = os.path.join(exe_dir, archivos[nombre])

                    hook = crear_hook_barra_inferior(base, avance_por_archivo)
                    urllib.request.urlretrieve(url, destino, reporthook=hook)

                    base += avance_por_archivo

                barra_progreso["value"] = 100
                porcentaje_var.set("100%")
                root.update_idletasks()  # Esta sí se mantiene

                # Ocultar la barra de progreso
                frame_progreso.pack_forget()  # No se vuelve a llamar update_idletasks aquí

                messagebox.showinfo("Actualización descargada", "Se lanzará la nueva versión ahora.")
                subprocess.Popen([os.path.join(exe_dir, "CreadorBorradores_Nuevo.exe")])
                sys.exit()

    except Exception as e:
        print(f"No se pudo verificar actualización: {e}")
        logging.error(f"No se pudo verificar actualización desde {url_api}", exc_info=True)

cuenta_seleccionada: Optional[str] = None
ruta_excel: Optional[str] = None
ruta_docx: Optional[str] = None

#Crea un hook para actualizar la barra de progreso de descarga.
def crear_hook_barra_inferior(base: int, avance_por_archivo: int):
    def hook(count, block_size, total_size):
        if total_size > 0:
            porcentaje = int((count * block_size * 100) / total_size)
            total = min(100, base + int(porcentaje * avance_por_archivo / 100))
            barra_progreso["value"] = total
            porcentaje_var.set(f"{total}%")
            root.update_idletasks()
    return hook

# GUI
root = tk.Tk()
# ---------------------- BARRA DE MENÚ ----------------------
def salir_aplicacion():
    """Cierra la aplicación."""
    root.quit()

def forzar_actualizacion_manual():
    """Permite forzar la descarga de la última versión, útil si falló la auto-actualización."""
    if messagebox.askyesno("Forzar actualización", "¿Deseas forzar la descarga e instalación de la última versión?"):
        verificar_actualizacion(forzar=True)

def mostrar_info_ayuda():
    """Muestra un mensaje de ayuda básico."""
    messagebox.showinfo("Ayuda", f"Versión: {__version__}\n\nEsta aplicación automatiza la creación y envío de borradores en Outlook.")

def mostrar_info_ver():
    """Muestra información de ejemplo para el menú Ver."""
    messagebox.showinfo("Ver", "Aquí se mostrarán opciones visuales en el futuro.")

menu_bar = tk.Menu(root)

# Menú Archivo
menu_archivo = tk.Menu(menu_bar, tearoff=0)
menu_archivo.add_command(label="Actualizar app", command=forzar_actualizacion_manual)
menu_archivo.add_separator()
menu_archivo.add_command(label="Salir", command=salir_aplicacion)
menu_bar.add_cascade(label="Archivo", menu=menu_archivo)

# Menú Ver
menu_ver = tk.Menu(menu_bar, tearoff=0)
menu_ver.add_command(label="Mostrar info", command=mostrar_info_ver)
menu_bar.add_cascade(label="Ver", menu=menu_ver)

# Menú Ayuda
menu_ayuda = tk.Menu(menu_bar, tearoff=0)
menu_ayuda.add_command(label="Acerca de", command=mostrar_info_ayuda)
menu_bar.add_cascade(label="Ayuda", menu=menu_ayuda)

# Asociar el menú a la ventana
root.config(menu=menu_bar)

root.title("Automatización de Borradores y envios en Outlook")  # Título solo en la barra de la ventana
root.geometry("480x430")

#barra de estado
frame_progreso = tk.Frame(root)
frame_progreso.pack(side="bottom", fill="x", padx=10, pady=5)

barra_progreso = ttk.Progressbar(frame_progreso, length=400, mode='determinate', maximum=100)
barra_progreso.pack(side="left", padx=(0, 10))

porcentaje_var = tk.StringVar(value="0%")
etiqueta_porcentaje = tk.Label(frame_progreso, textvariable=porcentaje_var)
etiqueta_porcentaje.pack(side="left")

# Establecer icono y logo
base_path = getattr(sys, '_MEIPASS', os.path.abspath("."))

# Establecer ícono de la aplicación
try:
    icon_path = os.path.join(base_path, "config", "icono.ico")
    icon_image = Image.open(icon_path)
    icon_tk = ImageTk.PhotoImage(icon_image)
    root.iconphoto(False, icon_tk)
except Exception as e:
    logging.error("No se pudo cargar el icono", exc_info=True)
    print(f"No se pudo cargar el icono: {e}")

# Cargar imagen de portada
try:
    cover_path = os.path.join(base_path, "config", "cover_borradores.jpg")
    cover_image = Image.open(cover_path)
    cover_image = cover_image.resize((500, 90))
    cover_img = ImageTk.PhotoImage(cover_image)
    etiqueta_cover = tk.Label(root, image=cover_img)
    etiqueta_cover.image = cover_img  # Para evitar que sea eliminada por el recolector
    etiqueta_cover.pack(pady=10)
except Exception as e:
    logging.error("No se pudo cargar la imagen de portada", exc_info=True)
    print(f"No se pudo cargar la imagen de portada: {e}")

# Variables para mostrar nombres de archivos (deben ir después de crear root)
ruta_excel_var = tk.StringVar()
ruta_docx_var = tk.StringVar()

#Busca los perfiles configurados en Outlook a través del registro de Windows. Devuelve una lista con los nombres de los perfiles encontrados.
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
            logging.error("No se encontraron perfiles.", exc_info=True)
            perfiles.append("No se encontraron perfiles.")
    except Exception as e:
        logging.error("Error al obtener perfiles", exc_info=True)
        perfiles.append("Error al obtener perfiles")
    return perfiles

def cerrar_outlook() -> None:
    #Fuerza el cierre de Outlook utilizando el comando taskkill.
    subprocess.run("taskkill /F /IM outlook.exe", shell=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)

#Busca la ruta del ejecutable de Outlook en ubicaciones comunes. Lanza una excepción si no se encuentra.
def obtener_ruta_outlook() -> str:
    rutas = [
        r"C:\\Program Files\\Microsoft Office\\root\\Office16\\OUTLOOK.EXE",
        r"C:\\Program Files (x86)\\Microsoft Office\\Office16\\OUTLOOK.EXE",
        r"C:\\Program Files\\Microsoft Office\\Office16\\OUTLOOK.EXE"
    ]
    for ruta in rutas:
        if os.path.exists(ruta):
            return ruta
        logging.error("No se encontró Outlook.", exc_info=True)
    raise FileNotFoundError("No se encontró Outlook.")

#Inicia Outlook utilizando el perfil especificado. Espera unos segundos tras lanzarlo para asegurar que esté listo.
def iniciar_outlook_con_perfil(perfil: str) -> None:
    try:
        ruta_outlook = obtener_ruta_outlook()
        subprocess.Popen([ruta_outlook, "/profile", perfil])
        time.sleep(7)
    except Exception as e:
        logging.error("No se pudo iniciar Outlook", exc_info=True)
        print(f"No se pudo iniciar Outlook: {e}")

#Intenta obtener las cuentas activas configuradas en Outlook. Reintenta varias veces en caso de fallo.
def obtener_cuentas_activas(max_intentos: int = 10, intervalo: int = 1) -> List[str]:
    for intento in range(max_intentos):
        try:
            outlook = win32com.client.Dispatch("Outlook.Application")
            namespace = outlook.GetNamespace("MAPI")
            cuentas = [account.SmtpAddress for account in namespace.Accounts]
            if cuentas:
                return cuentas
        except Exception as e:
            messagebox.showerror("Error al obtener cuentas", f"No se pudo acceder a las cuentas de Outlook:\n{e}")
            logging.error("No se pudo acceder a las cuentas de Outlook", exc_info=True)
            break  # detiene el bucle si ya falló definitivamente

        time.sleep(intervalo)
    return []

#Al seleccionar una cuenta en el combobox, se guarda como cuenta seleccionada.
def cuenta_asociada_seleccionada(_event: Any) -> None:  # El parámetro 'event' no se usa, se renombra como '_event'
    global cuenta_seleccionada
    cuenta_seleccionada = combo_cuentas_asociadas.get()

#Cuando se selecciona un perfil, este metodo cierra Outlook, lo vuelve a abrir con ese perfil y actualiza la interfaz para seleccionar la cuenta asociada a ese perfil.
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

def validar_datos_para_ejecucion(perfil: str, requiere_archivos: bool = True) -> bool:
    if perfil == "Seleccione perfil...":
        messagebox.showerror("Error", "Selecciona un perfil de Outlook.")
        logging.error(f"No se seleccionó un perfil válido: '{perfil}'")
        return False

    if not cuenta_seleccionada:
        messagebox.showerror("Error", "Selecciona una cuenta asociada.")
        logging.error(f"No se seleccionó cuenta asociada para el perfil: '{perfil}'")
        return False

    if requiere_archivos and (not ruta_excel or not ruta_docx):
        messagebox.showerror("Error", "Debes cargar tanto el archivo Excel como el Word antes de continuar.")
        logging.error(f"Archivos faltantes. Excel: {bool(ruta_excel)}, Word: {bool(ruta_docx)}")
        return False

    return True

#Devuelve la ruta completa del script según el entorno. Si está empaquetado con PyInstaller, utiliza la carpeta temporal (_MEIPASS).
def ruta_script(nombre_script: str) -> str:
    base_path = getattr(sys, '_MEIPASS', os.path.dirname(__file__))
    return os.path.join(base_path, nombre_script)

#Ejecuta un script externo (txt_1 o timer_sent), dependiendo del nombre proporcionado.
def ejecutar_script(nombre_script_txt: str, perfil: str, mostrar_mensaje: bool = False) -> None:
    solo_envio = "timer_sent" in nombre_script_txt.lower()

    if not validar_datos_para_ejecucion(perfil, requiere_archivos=not solo_envio):
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

        # Detectar si estamos en entorno empaquetado (PyInstaller)
        if getattr(sys, 'frozen', False):
            base_path = os.path.dirname(sys.executable)
            script_path = os.path.join(base_path, f"{script_name}.exe")
            cmd = [script_path] + args
        else:
            base_path = os.path.dirname(__file__)
            script_path = os.path.join(base_path, f"{script_name}.py")
            cmd = [sys.executable, script_path] + args

        if not os.path.exists(script_path):
            logging.error(f"No se encontró el archivo del script: {script_path}", exc_info=True)
            raise FileNotFoundError(f"No se encontró el archivo: {script_path}")

        result = subprocess.run(cmd, capture_output=True, text=True)
        print(result.stdout)

        if result.returncode != 0:
            logging.error(f"Error ejecutando '{script_name}' con args: {args}\nSTDERR: {result.stderr.strip()}", exc_info=True)
            error_msg = result.stderr.strip() or "Error desconocido."
            messagebox.showerror("Error en ejecución", f"Ocurrió un error:\n{error_msg}")
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
        logging.error("No se pudo ejecutar la script", exc_info=True)
        messagebox.showerror("Error", f"No se pudo ejecutar {script_name}:\n{e}")

#Abre un diálogo para seleccionar un archivo Excel y guarda la ruta seleccionada.
def cargar_excel() -> None:
    global ruta_excel
    archivo = filedialog.askopenfilename(filetypes=[("Archivos Excel", "*.xlsx *.xls *xlsm")])
    if archivo:
        ruta_excel = archivo
        nombre_archivo = os.path.basename(archivo)
        ruta_excel_var.set(f"... {nombre_archivo}   ✔️")

#Abre un diálogo para seleccionar un archivo Word (.docx), valida que no esté vacío y guarda la ruta.
def cargar_docx() -> None:
    global ruta_docx
    archivo = filedialog.askopenfilename(filetypes=[("Documentos de Word", "*.docx")])
    if archivo:
        try:
            if os.path.getsize(archivo) == 0:
                logging.error("El archivo .docx seleccionado está vacío", exc_info=True)
                messagebox.showerror("Archivo vacío", "El archivo .docx seleccionado está vacío.")
                return

            # Validación adicional: revisar si contiene texto
            doc = Document(archivo)
            contenido = "\n".join([p.text for p in doc.paragraphs]).strip()
            if not contenido:
                logging.error("El archivo .docx no contiene texto legible", exc_info=True)
                messagebox.showerror("Sin contenido", "El archivo .docx no contiene texto legible.")
                return

        except Exception as e:
            logging.error("No se pudo verificar el archivo", exc_info=True)
            messagebox.showerror("Error al validar archivo", f"No se pudo verificar el archivo:\n{e}")
            return

        ruta_docx = archivo
        nombre_archivo = os.path.basename(archivo)
        ruta_docx_var.set(f"... {nombre_archivo}   ✔️")

#Ejecuta el archivo timer_sent.exe con la cuenta seleccionada. Muestra mensaje de error si falta algún dato.
def ejecutar_timer_send() -> None:
    perfil = combo_cuentas.get()

    if not validar_datos_para_ejecucion(perfil, requiere_archivos=False):
        return

    try:
        exe_dir = os.path.dirname(sys.executable) if getattr(sys, 'frozen', False) else os.path.dirname(__file__)
        exe_path = os.path.join(exe_dir, "timer_sent.exe")

        if not os.path.exists(exe_path):
            logging.error(f"No se encontró el archivo 'timer_sent.exe' en {exe_path}", exc_info=True)

            raise FileNotFoundError(f"No se encontró el archivo: {exe_path}")

        print(">>> Ejecutando:", [exe_path, cuenta_seleccionada])

        result = subprocess.run([exe_path, cuenta_seleccionada], capture_output=True, text=True)

        if result.returncode != 0:
            logging.error(f"Error ejecutando 'timer_sent.exe' para la cuenta {cuenta_seleccionada}. STDERR: {result.stderr}", exc_info=True)
            messagebox.showerror("Error", f"Error ejecutando timer_sent.exe:\n{result.stderr}")
        else:
            enviados = result.stdout.strip()
            if enviados.isdigit():
                messagebox.showinfo("Éxito", f"Se enviaron {enviados} borradores correctamente.")
            else:
                messagebox.showinfo("Ejecutado", f"Salida:\n{result.stdout}")

    except Exception as e:
        logging.error("No se pudo ejecutar timer_sent", exc_info=True)
        messagebox.showerror("Error", f"No se pudo ejecutar timer_sent:\n{e}")

# Obtener perfiles disponibles en Outlook
perfiles = obtener_perfiles_outlook()
max_length = max([len(perfil) for perfil in perfiles], default=20)
width_combobox = min(max_length + 2, 50)
# Etiqueta para selección de perfil
tk.Label(root, text="Selecciona un perfil de Outlook:", font=("Arial", 10, "bold"), anchor="w").pack(anchor="w", padx=10)
# Frame contenedor del combobox de perfiles
frame_combo_perfiles = tk.Frame(root)
frame_combo_perfiles.pack(anchor="w", pady=5, padx=10)
# Combobox con los perfiles de Outlook
combo_cuentas = ttk.Combobox(frame_combo_perfiles, values=perfiles, state="readonly", font=("Arial", 10), width=width_combobox)
combo_cuentas.bind("<<ComboboxSelected>>", mostrar_cuenta_seleccionada)
combo_cuentas.pack(side="left")
combo_cuentas.current(0)
# Variable de etiqueta para mostrar cuenta seleccionada
label_cuenta_var = tk.StringVar()
label_cuenta = tk.Label(root, textvariable=label_cuenta_var, font=("Arial", 10))
label_cuenta.pack(pady=5)
# Segundo combobox (para cuentas asociadas al perfil)
combo_cuentas_asociadas = ttk.Combobox(root, state="readonly", font=("Arial", 10))
combo_cuentas_asociadas.bind("<<ComboboxSelected>>", cuenta_asociada_seleccionada)
combo_cuentas_asociadas.pack(pady=5)
combo_cuentas_asociadas.pack_forget()
# Frame contenedor para botón de Excel
frame_excel = tk.Frame(root)
frame_excel.pack(anchor="w", pady=5, padx=10)
# Botón para cargar Excel
btn_excel = tk.Button(frame_excel, text="Cargar Excel", command=cargar_excel,
                      font=("Arial", 10), padx=10, pady=5)
btn_excel.pack(side="left")
# Etiqueta para mostrar nombre del archivo Excel
label_excel = tk.Label(frame_excel, textvariable=ruta_excel_var, font=("Arial", 9), wraplength=300, justify="left", fg="green")
label_excel.pack(side="left", padx=10)

# Frame contenedor para botón de Word
frame_docx = tk.Frame(root)
frame_docx.pack(anchor="w", pady=5, padx=10)
# Botón para cargar Word
btn_docx = tk.Button(frame_docx, text="Cargar Texto Mail", command=cargar_docx,
                     font=("Arial", 10), padx=10, pady=5)
btn_docx.pack(side="left")
# Etiqueta para mostrar nombre del archivo Word
label_docx = tk.Label(frame_docx, textvariable=ruta_docx_var, font=("Arial", 9), wraplength=300, justify="left", fg="green")
label_docx.pack(side="left", padx=10)

# Frame contenedor para botón de Borradores
frame_crear = tk.Frame(root)
frame_crear.pack(anchor="w", pady=5, padx=10)
# Botón para cargar Borradores
tk.Button(frame_crear, text="Crear Borradores",
    command=lambda: ejecutar_script("txt_1", combo_cuentas.get(), False),
    font=("Arial", 10), padx=10, pady=5).pack(side="left")

# Botón para Enviar Borradores
tk.Button(root, text="Enviar Borradores", command=ejecutar_timer_send,
          font=("Arial", 12), padx=10, pady=5, bg="purple", fg="white").pack(pady=10)

# Inicia el loop principal de la interfaz gráfica
root.mainloop()