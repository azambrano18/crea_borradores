import sys
import win32com.client
import mammoth
import pandas as pd
import os
import traceback

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