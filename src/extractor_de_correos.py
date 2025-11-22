import os
import re
import win32com.client
import pandas as pd
import unicodedata

from datetime import datetime
from datetime import timedelta
from datetime import date

from pathlib import Path

from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo

from colorama import init, Fore

# === Funciones auxiliares ===
def limpiar_texto(nombre):
    """Limpia nombres de archivo sin eliminar la extensión."""
    if not nombre:
        return ""
    nombre = str(nombre).strip()
    
    # Detectar extensión válida (opcional)
    m = re.match(r"(.+)\.([A-Za-z0-9]{1,5})$", nombre)
    if m:
        base, ext = m.group(1), "." + m.group(2).lower()
    else:
        base, ext = nombre, ""

    # Separar nombre y extensión
    # base, ext = os.path.splitext(nombre)

    # Reemplazar caracteres inválidos solo en la parte base
    # base = re.sub(r'[\\/:\*\?"<>\|\r\n\t]', "_", base)
    
    # Reemplazar caracteres inválidos + comillas Unicode
    base = re.sub(
        r'[\\/:*?"<>|\r\n\t“”‘’´`]',
        "_",
        base
    )
    
    # Normalizar espacios
    base = re.sub(r'\s+', " ", base)
    base = base.strip()

    # Limitar longitud sin afectar extensión
    if len(base) > 70:
        base = base[:70]

    # Volver a unir con extensión (en minúsculas)
    return f"{base}{ext.lower()}"

def quitar_acentos(texto):
    reemplazos = {
        "á": "a", "é": "e", "í": "i", "ó": "o", "ú": "u",
        "Á": "A", "É": "E", "Í": "I", "Ó": "O", "Ú": "U",
        "ñ": "n", "Ñ": "N"
    }
    for a, b in reemplazos.items():
        texto = texto.replace(a, b)
    return texto

def filtrar_cc(cc):
    """Filtra nombres autorizados del campo CC."""
    if not cc:
        return "-----"
    autorizados = [
        "Arq. Belty Espinoza Santos", "Arq. Ricardo Valverde Paredes", "Ing. Jhon Araujo Borjas", "Arq. Ader Intriago Arteaga",
        "Ing. Jorge Maldonado Granizo", "Arq. Roberto Vivanco Calderon", "Ing. Lissette Moreno Balladares",
        "Ing. Luis Vargas Orozco", "Ing. Patricia Fuentes Moran", "Lic. Narcisa A. Munoz Feraud", "Lic. Leonardo Rodriguez Molina",
        "Ing. Jaime Franco Baquerizo", "Arq. Jhonther Cardenas", "Ing. Miguel Flores Poveda", "Ing. Jorge Tohaquiza Jacho",
        "Ing. Humberto Rodriguez Gonzalez", "Ing. Eddy Alfonso Garcia", "Gilda Suarez Crespin", "Arq. Pilar Zalamea Garcia",
        "Ing. Isaac Munoz Mindiola", "Arq. Franklin Medina Gonzalez", "Ing. Victor Velasco Matute", "Sra. Bethzaida Villamil"
    ]
    cc_norm = quitar_acentos(cc.lower())
    resultado = []
    for nombre in autorizados:
        n_norm = quitar_acentos(nombre.lower())
        partes = n_norm.split()
        if len(partes) >= 2 and partes[1] in cc_norm and partes[-1] in cc_norm:
            resultado.append(nombre)
    if len(resultado) == 0:
        return "-----"
    else:
        return "\n".join(resultado)

def Corregir_rem(rem):
    """Pone el título a los nombres que aparecen en la lista, en el campo remitente"""
    if not rem:
        return ""
    autorizados = [
        "Arq. Belty Espinoza Santos", "Arq. Ricardo Valverde Paredes", "Ing. Jhon Araujo Borjas", "Arq. Ader Intriago Arteaga",
        "Ing. Jorge Maldonado Granizo", "Arq. Roberto Vivanco Calderon", "Ing. Lissette Moreno Balladares",
        "Ing. Luis Vargas Orozco", "Ing. Patricia Fuentes Moran", "Lic. Narcisa A. Munoz Feraud", "Lic. Leonardo Rodriguez Molina",
        "Ing. Jaime Franco Baquerizo", "Arq. Jhonther Cardenas", "Ing. Miguel Flores Poveda", "Ing. Jorge Tohaquiza Jacho",
        "Ing. Humberto Rodriguez Gonzalez", "Ing. Eddy Alfonso Garcia", "Gilda Suarez Crespin", "Arq. Pilar Zalamea Garcia",
        "Ing. Isaac Munoz Mindiola", "Arq. Franklin Medina Gonzalez", "Ing. Victor Velasco Matute", "Sra. Bethzaida Villamil"
    ]
    rem_norm = quitar_acentos(rem.lower())
    resultado = []
    for nombre in autorizados:
        n_norm = quitar_acentos(nombre.lower())
        partes = n_norm.split()
        if len(partes) >= 2 and partes[1] in rem_norm and partes[-1] in rem_norm:
            resultado.append(nombre)
    
    if len(resultado) == 0:
        return rem
    else:
        return "\n".join(resultado)

def obtener_info_persona(remitente):
    """
    Devuelve (cargo, dependencia) aunque el remitente llegue sin título,
    abreviado, en otro orden o en minúsculas.
    """
    nombre_limpio = limpiar_nombre(remitente)
    palabras = nombre_limpio.split()

    for nombre_base, info in PERSONAL_LIMPIO.items():
        coincidencias = 0
        for p in palabras:
            if p in nombre_base:
                coincidencias += 1

        # Se considera coincidencia válida si al menos 2 palabras coinciden
        if coincidencias >= 2:
            return info["cargo"], info["dependencia"]

    return None, None

def limpiar_nombre(nombre):
    if not nombre:
        return ""

    # minúsculas
    nombre = nombre.lower()

    # quitar tildes
    nombre = ''.join(
        c for c in unicodedata.normalize('NFD', nombre)
        if unicodedata.category(c) != 'Mn'
    )

    # quitar títulos comunes
    nombre = re.sub(r'\b(arq|ing|lic|sr|sra|dra|dr)\.?', '', nombre)

    # quitar puntos y comas
    nombre = nombre.replace('.', '').replace(',', '')

    # quitar espacios dobles
    nombre = re.sub(r'\s+', ' ', nombre).strip()

    return nombre

def nompropio_python(asunto):
    return asunto.title()

def limpiar_destinatarios(cadena_to: str) -> str:
    if not cadena_to:
        return ""
    
    # Outlook suele devolver algo como: "Juan <juan@mail.com>; Maria <maria@mail.com>"
    partes = [p.strip() for p in cadena_to.split(";") if p.strip()]
    
    # Unir con salto de línea
    return "\n".join(partes)

PERSONAL_INFO = {
    # Dirección DIOU
    "Arq. Belty Espinoza Santos": {
        "cargo": "Directora",
        "dependencia": "Dirección DIOU"
    },

    # Jefatura de Planificación
    "Arq. Ricardo Valverde Paredes": {
        "cargo": "Jefe de Planificación",
        "dependencia": "Jefatura de Planificación"
    },
    "Ing. Jhon Araujo Borjas": {
        "cargo": "Analista de Planificación de Obras 3",
        "dependencia": "Jefatura de Planificación"
    },
    "Arq. Ader Intriago Arteaga": {
        "cargo": "Analista de Planificación de Obras 3",
        "dependencia": "Jefatura de Planificación"
    },
    "Ing. Jorge Maldonado Granizo": {
        "cargo": "Analista de Planificación de Obras 3",
        "dependencia": "Jefatura de Planificación"
    },
    "Arq. Roberto Vivanco Calderón": {
        "cargo": "Analista de Planificación de Obras 3",
        "dependencia": "Jefatura de Planificación"
    },
    "Ing. Lissette Moreno Balladares": {
        "cargo": "Analista de Planificación de Obras 3",
        "dependencia": "Jefatura de Planificación"
    },
    "Ing. Luis Vargas Orozco": {
        "cargo": "Analista de Planificación de Obras 3",
        "dependencia": "Jefatura de Planificación"
    },
    "Ing. Patricia Fuentes Morán": {
        "cargo": "Analista de Planificación 1",
        "dependencia": "Jefatura de Planificación"
    },
    "Lic. Narcisa A. Muñoz Feraud": {
        "cargo": "Asistente de Planificación de Obras 2",
        "dependencia": "Jefatura de Planificación"
    },
    "Lic. Leonardo Rodríguez Molina": {
        "cargo": "Asistente de Planificación de Obras 2",
        "dependencia": "Jefatura de Planificación"
    },
    "Ing. Jaime Franco Baquerizo": {
        "cargo": "Asistente 2 de Infraestructura y Obras Universitarias",
        "dependencia": "Jefatura de Planificación"
    },
    "Arq. Jhonther Cárdenas": {
        "cargo": "Supervisor de Construcciones Civiles",
        "dependencia": "Jefatura de Planificación"
    },

    # Jefatura de Infraestructura
    "Ing. Miguel Flores Poveda": {
        "cargo": "Jefe de Infraestructura",
        "dependencia": "Jefatura de Infraestructura"
    },
    "Ing. Jorge Tohaquiza Jacho": {
        "cargo": "Analista de Infraestructura de Obras Universitarias 3",
        "dependencia": "Jefatura de Infraestructura"
    },
    "Ing. Humberto Rodríguez González": {
        "cargo": "Analista de Mantenimiento de Infraestructura 3",
        "dependencia": "Jefatura de Infraestructura"
    },
    "Ing. Eddy Alfonso García": {
        "cargo": "Asistente de Infraestructura de Obras Universitarias 2",
        "dependencia": "Jefatura de Infraestructura"
    },
    "Gilda Suárez Crespín": {
        "cargo": "Asistente de Infraestructura de Obras Universitarias 2",
        "dependencia": "Jefatura de Infraestructura"
    },

    # Jefatura de Diseño y Fiscalización
    "Arq. Pilar Zalamea García": {
        "cargo": "Jefe de Diseño y Fiscalización",
        "dependencia": "Jefatura de Diseño y Fiscalización"
    },
    "Ing. Isaac Muñoz Mindiola": {
        "cargo": "Analista de Diseño y Fiscalización 1",
        "dependencia": "Jefatura de Diseño y Fiscalización"
    },
    "Arq. Franklin Medina González": {
        "cargo": "Analista de Diseño y Fiscalización 3",
        "dependencia": "Jefatura de Diseño y Fiscalización"
    },
    "Ing. Victor Velasco Matute": {
        "cargo": "Asistente de Infraestructura 2",
        "dependencia": "Jefatura de Diseño y Fiscalización"
    },
    "Sr. Bethzaida Villamil": {
        "cargo": "Asistente 2 Diseño y Planificación (DIOU)",
        "dependencia": "Jefatura de Diseño y Fiscalización"
    }
}

PERSONAL_LIMPIO = {
    limpiar_nombre(nombre_original): info
    for nombre_original, info in PERSONAL_INFO.items()
}

# === Script principal ===
def exportar_correos():
    # Inicializa colorama
    init(autoreset=True)  # autoreset=True hace que después de cada print el color vuelva al normal
    
    while True:
        print("=== Exportador de Correos con Anexos ===", end="\n\n")

        # Pedir fechas al usuario
        dia = input("Día (1-31): ")
        mes = input("Mes (1-12): ")
        anio = input("Año (4 dígitos): ")
        
        # try:
        # fecha_inicio_raw = date(int(anio), int(mes), int(dia))
        # fecha_fin_raw = fecha_inicio_raw + timedelta(days=1)
        fecha_inicio_raw = datetime(int(anio), int(mes), int(dia), 0, 0, 0)
        fecha_fin_raw = fecha_inicio_raw + timedelta(days=1)
        
        # print(f"{fecha_inicio_raw}")
        # print(f"{fecha_fin_raw}")

        fecha_inicio_str = fecha_inicio_raw.strftime("%Y-%m-%d")
        # fecha_fin_str = fecha_fin_raw.strftime("%Y-%m-%d")

        # fecha_inicio = datetime.strptime(fecha_inicio_str, "%Y-%m-%d")
        # fecha_fin = datetime.strptime(fecha_fin_str, "%Y-%m-%d")

        mes_int= int(mes)

        meses = {1:"Enero", 2:"Febrero", 3:"Marzo", 4:"Abril",
                 5:"Mayo", 6:"Junio", 7:"Julio", 8:"Agosto",
                 9:"Septiembre", 10:"Octubre", 11:"Noviembre", 12:"Diciembre"}

        mes_str = meses[mes_int]
        
        # Crear carpeta del día
        carpeta_base = Path.home()/"Downloads"/f"Correos {mes_str} - {fecha_inicio_raw.strftime("%Y")}"/fecha_inicio_str
        
        # Conectar con Outlook
        outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        carpeta = outlook.GetDefaultFolder(6)  # 6 = Bandeja de entrada

        registros = []

        # filtro = f"[ReceivedTime] >= '{fecha_inicio_raw.strftime('%d/%m/%Y %I:%M %p')}' AND [ReceivedTime] < '{fecha_fin_raw.strftime('%d/%m/%Y %I:%M %p')}'"
        filtro = f"[ReceivedTime] >= '{fecha_inicio_raw.strftime('%d/%m/%Y %H:%M')}' AND [ReceivedTime] < '{fecha_fin_raw.strftime('%d/%m/%Y %H:%M')}'"
        
        # print(fecha_inicio_raw.strftime('%d/%m/%Y %H:%M'))
        # print(fecha_fin_raw.strftime('%d/%m/%Y %H:%M'))
        
        items = carpeta.Items
        items.Sort("[ReceivedTime]", True)   # True = descendente, False = ascendente
        items_filtrados = items.Restrict(filtro)
        
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False
        
        if len(items_filtrados) > 0:
            print("\nProcesando correos", end="")

        for mail in items_filtrados:
            try:
                if mail.Class != 43:
                    continue

                remitente = mail.SenderName
                destinatario = mail.To
                asunto = mail.Subject or ""
                cc = mail.CC or ""

                recibido = mail.ReceivedTime
                recibido_py = datetime(recibido.year, recibido.month, recibido.day,
                                    recibido.hour, recibido.minute, recibido.second)

                if not (fecha_inicio_raw <= recibido_py < fecha_fin_raw):
                    continue
                
                if remitente in ["QUIPUX", "UNIVERSIDAD DE GUAYAQUIL", "INFO UG", "Comunicados DVSBE"]:
                    continue
                
                print(".", end="", flush=True)
                
                id_correo = recibido_py.strftime("%H%M%S") + "_" + limpiar_texto(remitente)
                asunto_limpio = limpiar_texto(asunto)
                carpeta_nombre = f"{id_correo}"
                if len(carpeta_nombre) > 100:
                    carpeta_nombre = carpeta_nombre[:100]

                carpeta_base.mkdir(parents=True, exist_ok=True) # crear carpeta base

                carpeta_correo = carpeta_base / carpeta_nombre
                os.makedirs(carpeta_correo, exist_ok=True) # crear carpeta de correo

                lista_anexos = []

                if mail.Attachments.Count > 0:
                    i = 0
                    for anexo in mail.Attachments:
                        nombre_archivo = limpiar_texto(anexo.FileName)
                        if "image" not in nombre_archivo.lower() and "outlook" not in nombre_archivo.lower():
                            carpeta_anexos = carpeta_correo / "Anexos"
                            os.makedirs(carpeta_anexos, exist_ok=True) # crear carpeta Anexos, si existen
                            
                            ruta_anexo = carpeta_anexos / nombre_archivo
                            anexo.SaveAsFile(str(ruta_anexo))
                            i+=1
                            # lista_anexos.append(nombre_archivo)
                            lista_anexos.append(i)

                mht_path = carpeta_correo / f"{asunto_limpio}.mht"
                pdf_path = carpeta_correo / f"{asunto_limpio}.pdf"
                msg_path = carpeta_correo / f"{asunto_limpio}.msg"

                try:
                    mail.SaveAs(str(mht_path), 10)
                    doc = word.Documents.Open(str(mht_path))
                    doc.ExportAsFixedFormat(OutputFileName=str(pdf_path), ExportFormat=17)
                    doc.Close(False)
                    os.remove(mht_path)
                except Exception:
                    mail.SaveAs(str(msg_path), 3)
                
                cargo, dependencia = obtener_info_persona(remitente)
                
                cant_anexos = len(lista_anexos)
                observaciones = "No contiene anexos" if cant_anexos == 0 else f"Anexa {cant_anexos} documento(s)"

                cc_filtrado = filtrar_cc(str(cc))
                remitente_filtrado = Corregir_rem(remitente)
                destinatario_filtrado = limpiar_destinatarios(destinatario)

                registros.append({
                    "Fecha del Documento": recibido_py.strftime("%Y-%m-%d %H:%M:%S"),
                    "Remitente": nompropio_python(remitente_filtrado),
                    "Cargo": cargo,
                    "Facultad/Dependencia": dependencia,
                    # "Destinatario": nompropio_python(destinatario),
                    "Destinatario": nompropio_python(destinatario_filtrado),
                    "Empresa/Cargo": "",
                    "Asunto": nompropio_python(asunto),
                    "Con Copia": cc_filtrado,
                    "Observaciones": observaciones,
                    # "Anexo(s)": "; ".join(lista_anexos)
                    # "Anexo(s)": "; ".join([str(a) for a in lista_anexos])
                    # "Anexo(s)": len(lista_anexos)
                })
            except Exception as e:
                print(f"Error procesando correo: {e}")
        word.Quit()

        # === Exportar resultados a Excel con formato ===
        if registros:
            df = pd.DataFrame(registros)
            ruta_excel = carpeta_base / f"{fecha_inicio_str}_CorreosExportados.xlsx"
            df.to_excel(ruta_excel, index=False)

            # === Aplicar formato con openpyxl ===
            wb = load_workbook(ruta_excel)
            ws = wb.active

            # Cuerpo
            for row in ws.iter_rows():
                for celda in row:
                    celda.font = Font(name="Arial", size=12)
                    if celda.value is not None:   # opcional: solo celdas con contenido
                        celda.alignment = Alignment(wrap_text=True, horizontal="left", vertical="top")

            for celda in ws["I"]:
                celda.font = Font(name="Arial", size=12, color="FF0000", bold=True)  # rojo
                celda.alignment = Alignment(horizontal="center", vertical="center")
            
            # Ocultar la columna J
            # ws.column_dimensions['J'].hidden = True

            # Encabezados
            header_font = Font(name="Arial", size=12, bold=True, color="FFFFFF")
            header_fill = PatternFill(start_color="4F81BD", end_color="4F81BD", fill_type="solid")

            for col in range(1, ws.max_column + 1):
                cell = ws.cell(row=1, column=col)
                cell.font = header_font
                cell.fill = header_fill
                cell.alignment = Alignment(wrap_text=True, horizontal="center", vertical="center")

            # Bordes
            thin_border = Border(
                left=Side(style="thin"),
                right=Side(style="thin"),
                top=Side(style="thin"),
                bottom=Side(style="thin")
            )

            # Aplicar bordes a las celdas
            for row in ws.iter_rows(min_row=1, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
                for cell in row:
                    cell.border = thin_border

            # Ajustar ancho de columnas
            for col in ws.columns:
                column = get_column_letter(col[0].column)
                ws.column_dimensions[column].width = 20.71

            ws.column_dimensions["G"].width = 30.71
            ws.column_dimensions["I"].width = 30.71
            
            # Ajustar alto de filas
            for row in ws.iter_rows():
                celda = row[0]
                ws.row_dimensions[celda.row+1].height = 150.04

            # Convertir a tabla con estilo
            ultima_fila = ws.max_row
            ultima_col = ws.max_column
            rango_tabla = f"A1:{get_column_letter(ultima_col)}{ultima_fila}"

            # Llenar con color amarillo las celdas vacías
            body_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid") #amarillo
            
            for fila in ws[rango_tabla]:     # fila = tuple de celdas
                for celda in fila:           # celda = objeto real de openpyxl
                    if celda.value is None or celda.value == "":
                        celda.fill = body_fill

            # Tabla
            tabla = Table(displayName="CorreosExportados", ref=rango_tabla)
            estilo = TableStyleInfo(
                name="TableStyleMedium2",
                showRowStripes=True,
                showColumnStripes=False
            )
            tabla.tableStyleInfo = estilo
            ws.add_table(tabla)
            ws.sheet_view.zoomScale = 80 # Ajustar zoom
            wb.save(ruta_excel)

            print(Fore.GREEN + f"\n\nExportación completada correctamente: {len(registros)} correos.")
            input("\nPresione ENTER para ver los correos.")
            os.system("cls")
            os.startfile(carpeta_base)
        else:
            print(Fore.RED + "\nNo se encontraron correos en el rango especificado.")
            input("\nPresione ENTER para continuar.")
            os.system("cls")    

if __name__ == "__main__":
    exportar_correos()

