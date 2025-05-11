# Librerías estándar
import sys  # Permite trabajar con argumentos desde la línea de comandos y manejar salidas y errores
import os  # Proporciona funciones para interactuar con el sistema de archivos (verificar rutas, etc.)
import traceback  # Permite imprimir rastros de error detallados para depuración
# Librerías externas
import pandas as pd  # Utilizada para leer y manipular datos desde archivos Excel
import win32com.client  # Permite automatizar Outlook a través de COM (crear correos, acceder a cuentas, etc.)
import mammoth  # Convierte archivos .docx (Word) en HTML, útil para correos con formato

#-------------cargar_cuerpo_desde_docx: --------------
"""
Carga el contenido de un archivo DOCX, lo convierte a HTML y reemplaza la variable [Nombre].
Args:
    archivo_docx (str): Ruta al archivo .docx que contiene el cuerpo del mensaje.
    nombre (str): Nombre que se reemplazará en el contenido donde aparezca [Nombre].
Returns:
    str: Cuerpo del mensaje en formato HTML con el nombre integrado.
Raises:
    SystemExit: Si el archivo no existe o ocurre un error durante la conversión.
"""
def cargar_cuerpo_desde_docx(archivo_docx, nombre):
    try:
        if not os.path.exists(archivo_docx):
            raise FileNotFoundError(f"El archivo '{archivo_docx}' no existe.")
        with open(archivo_docx, "rb") as docx_file:
            resultado = mammoth.convert_to_html(docx_file)
            cuerpo = resultado.value
        cuerpo = cuerpo.replace("[Nombre]", nombre)
        cuerpo = f'<div style="font-family: Calibri, sans-serif; font-size: 11pt;">{cuerpo}</div>'
        return cuerpo
    except Exception as e:
        print(f"Error al cargar el archivo .docx con formato: {e}", file=sys.stderr)
        traceback.print_exc()
        sys.exit(1)
#----------------crear_borrador----------------
"""
Crea un borrador en Outlook con la cuenta, destinatario, asunto y cuerpo especificados.
Args:
    cuenta (str): Dirección de correo del remitente.
    destinatario (str): Dirección de correo del destinatario.
    asunto (str): Asunto del mensaje.
    cuerpo (str): Cuerpo del mensaje en formato HTML.
    perfil_outlook (str, opcional): Nombre del perfil de Outlook a usar. Default es "".
Returns:
    bool: True si se creó el borrador correctamente.
Raises:
    SystemExit: Si Outlook no puede inicializarse o no se encuentra la cuenta.
"""
def crear_borrador(cuenta, destinatario, asunto, cuerpo, perfil_outlook=""):
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        if perfil_outlook:
            namespace.Logon(Profile=perfil_outlook, ShowDialog=False, NewSession=True)
        cuenta_encontrada = None
        for account in namespace.Accounts:
            if account.SmtpAddress.lower() == cuenta.lower():
                cuenta_encontrada = account
                break
        if not cuenta_encontrada:
            print(f"No se encontró la cuenta: {cuenta}", file=sys.stderr)
            sys.exit(1)
        mensaje = outlook.CreateItem(0)
        mensaje._oleobj_.Invoke(*(64209, 0, 8, 0, cuenta_encontrada))
        mensaje.Display()
        signature = mensaje.HTMLBody
        mensaje.Subject = asunto
        mensaje.To = destinatario
        mensaje.BodyFormat = 2
        mensaje.HTMLBody = cuerpo + signature
        mensaje.Save()
        mensaje.Close(1)
        print(f"Borrador creado exitosamente para: {destinatario}")
        return True
    except Exception as e:
        print(f"Error al crear el borrador para {destinatario}: {e}", file=sys.stderr)
        traceback.print_exc()
        sys.exit(1)

#---------------procesar_excel---------------
"""
Procesa el archivo Excel con las columnas 'Correo', 'Asunto' y 'Nombre', genera un cuerpo HTML para cada fila y crea un borrador en Outlook.
Args:
        correo_cuenta (str): Dirección de correo usada como remitente.
        perfil_outlook (str): Perfil de Outlook a utilizar.
        ruta_excel (str): Ruta al archivo Excel.
        ruta_docx (str): Ruta al archivo DOCX con la plantilla del mensaje.
 Raises:
        SystemExit: Si faltan archivos, columnas requeridas o ocurre un error de ejecución.
"""
def procesar_excel(correo_cuenta, perfil_outlook="", ruta_excel=None, ruta_docx=None):
    if not ruta_excel or not os.path.exists(ruta_excel):
        sys.exit("ERROR: Ruta del archivo Excel inválida o no encontrada.")
    if not ruta_docx or not os.path.exists(ruta_docx):
        sys.exit("ERROR: Ruta del archivo DOCX inválida o no encontrada.")
    try:
        df = pd.read_excel(ruta_excel, sheet_name='Hoja1')
        if {'Correo', 'Asunto', 'Nombre'}.issubset(df.columns):
            for index, row in df.iterrows():
                destinatario = str(row['Correo']).strip()
                asunto = str(row['Asunto']).strip()
                nombre = str(row['Nombre']).strip()
                cuerpo = cargar_cuerpo_desde_docx(ruta_docx, nombre)
                crear_borrador(correo_cuenta, destinatario, asunto, cuerpo, perfil_outlook)
        else:
            sys.exit("ERROR: El archivo Excel no contiene las columnas necesarias: 'Correo', 'Asunto', 'Nombre'.")
    except Exception as e:
        print(f"Error al procesar el archivo Excel: {e}", file=sys.stderr)
        traceback.print_exc()
        sys.exit(1)

if __name__ == "__main__":
    if len(sys.argv) != 5:
        print("Uso: python txt_1.py <cuenta> <perfil_outlook> <ruta_excel> <ruta_docx>")
        sys.exit(1)

    cuenta = sys.argv[1]
    perfil = sys.argv[2]
    ruta_excel = sys.argv[3]
    ruta_docx = sys.argv[4]

    print(">> Argumentos recibidos correctamente")
    print(f"Cuenta: {cuenta}")
    print(f"Perfil: {perfil}")
    print(f"Excel:  {ruta_excel}")
    print(f"Word:   {ruta_docx}")

    procesar_excel(cuenta, perfil_outlook=perfil, ruta_excel=ruta_excel, ruta_docx=ruta_docx)