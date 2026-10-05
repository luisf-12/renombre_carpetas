import sys
import subprocess
import os
import glob
import shutil
import zipfile
import pandas as pd
import win32com.client as win32
import tempfile

# CONFIGURACIÓN DE RUTAS
USER_PROFILE = os.environ.get('USERPROFILE', '') # Detecta "C:\Users\NombreUsuario" automáticamente
ESCRITORIO_PATH = os.path.join(USER_PROFILE, "OneDrive - emprestur.com", "Escritorio")
EXCEL_MAESTRO_PATH = os.path.join(ESCRITORIO_PATH, "ORDENES OK PARA FACTURAR.xlsx")
PLANTILLA_PATH = os.path.join(ESCRITORIO_PATH, "PLANTILLA-FORMATO DETALLE DE FACTURACION.xlsx")
MES_ACTUAL = ""  # Carpeta en Escritorio
NOMBRE_HOJA = ""  # hoja del Excel

def generar_e_integrar_carpeta(excel_app, ruta_carpeta_fact, nombre_carpeta, fila_maestro, ruta_pendientes):
    id_paciente = str(fila_maestro['ID PACIENTE']).strip()
    nombre_paciente = str(fila_maestro['NOMBRE']).strip()
    numero_orden = str(fila_maestro['ORDEN']).strip()
    numero_factura = str(fila_maestro['FACTURA']).strip()
    
    nuevo_nombre_factura = str(fila_maestro["NOMBRE FACTURA"]).strip() + ".pdf"
    nuevo_nombre_detalle = str(fila_maestro["NOMBRE DETALLE"]).strip() + ".pdf"
    nuevo_nombre_soportes = str(fila_maestro["NOMBRE IMÁGENES"]).strip() + ".pdf"
    nombre_zip_final = str(fila_maestro["TEXTO ZIP"]).strip() + ".zip"

    # 1. ENCONTRAR Y CONVERTIR EL EXCEL DEL SISTEMA USANDO EXCEL NATIVO
    archivos_excel = glob.glob(os.path.join(ruta_carpeta_fact, "*.xlsx")) + glob.glob(os.path.join(ruta_carpeta_fact, "*.xls"))
    archivos_excel = [f for f in archivos_excel if "PLANTILLA" not in f.upper() and "TEMP" not in f.upper()]
    
    if not archivos_excel:
        print(f"No se encontró EXCEL del TM en: {nombre_carpeta}.")
        return False

    excel_sistema_path = archivos_excel[0]
    
    # --- CORRECCIÓN 1: Enviar temporal del sistema a carpeta Temp de Windows ---
    temp_dir = tempfile.gettempdir()
    archivo_reparado_path = os.path.join(temp_dir, f"temp_sistema_convertido_{nombre_carpeta}.xlsx")
    
    # Forzar a Excel a abrir el archivo (repara formatos HTML/viejos sin depender de xlrd)
    try:
        wb_sis = excel_app.Workbooks.Open(excel_sistema_path)
        if os.path.exists(archivo_reparado_path): os.remove(archivo_reparado_path)
        wb_sis.SaveAs(archivo_reparado_path, FileFormat=51) # 51 = Formato XLSX Estándar
        wb_sis.Close(False)
        df_sistema = pd.read_excel(archivo_reparado_path)
        if os.path.exists(archivo_reparado_path): os.remove(archivo_reparado_path)
    except Exception as e:
        print(f"Falló la conversión del archivo excel del sistema en {nombre_carpeta}: {e}")
        return False

    # 2. LIMPIEZA DE FILAS EN EL DATA FRAME
    if 'Nro Remesa' not in df_sistema.columns:
        for i in range(min(15, len(df_sistema))):
            valores_fila = df_sistema.iloc[i].astype(str).str.strip().tolist()
            if 'Nro Remesa' in valores_fila:
                df_sistema.columns = valores_fila
                df_sistema = df_sistema.iloc[i+1:]
                break

    df_sistema.columns = df_sistema.columns.str.strip()

    if 'Nro Remesa' in df_sistema.columns:
        df_sistema = df_sistema[df_sistema['Nro Remesa'].notna()]
        df_sistema['Nro Remesa'] = df_sistema['Nro Remesa'].astype(str).str.strip()
        df_sistema = df_sistema[df_sistema['Nro Remesa'].str.startswith('OS', na=False)]
        df_sistema = df_sistema.reset_index(drop=True)

    if df_sistema.empty:
        print(f"Reporte vacío tras filtrado en {nombre_carpeta}.")
        return False

    # 3. EDICIÓN DE LA PLANTILLA ORIGINAL
    try:
        wb = excel_app.Workbooks.Open(PLANTILLA_PATH)
        ws = wb.ActiveSheet
    except Exception as e:
        print(f"Error abriendo la plantilla de detalle base: {e}")
        return False

    # Escaneo dinámico para encontrar la fila del TOTAL real en la plantilla
    fila_total = None
    for r in range(3, 200):
        for c in range(1, 20):
            celda_val = str(ws.Cells(r, c).Value or "").strip().upper()
            if "TOTAL" in celda_val:
                fila_total = r
                break
        if fila_total:
            break

    # Escribir número de factura en J1 (o donde corresponda)
    ws.Cells(1, 10).Value = numero_factura

    # Inyectar datos alineados estrictamente desde la FILA 3 según la plantilla original
    fila_inicio_plantilla = 3
    for idx, fila_sis in df_sistema.iterrows():
        fila_actual = fila_inicio_plantilla + idx
        
        ws.Cells(fila_actual, 1).Value = fila_sis.get('Nro Remesa', '')          # A
        ws.Cells(fila_actual, 2).Value = ""                                         # B: Obligatoriamente vacío por solicitud
        ws.Cells(fila_actual, 3).Value = fila_sis.get('Origen', '')           # C
        ws.Cells(fila_actual, 4).Value = fila_sis.get('Destino', '')          # D
        ws.Cells(fila_actual, 5).Value = id_paciente                                # E
        ws.Cells(fila_actual, 6).Value = nombre_paciente                            # F
        ws.Cells(fila_actual, 7).Value = numero_orden                               # G
        ws.Cells(fila_actual, 8).Value = fila_sis.get('Recepcion Mercancia', '')     # H
        ws.Cells(fila_actual, 9).Value = ""                                         # I: Obligatoriamente vacío por solicitud
        
        # Columnas de control intermedias en limpio
        ws.Cells(fila_actual, 10).Value = fila_sis.get('Total (1+2+3)', 0) # J: Und (Mapeado con el Valor Total del sistema)
        ws.Cells(fila_actual, 11).Value = 0   # K: Peso Declarado
        ws.Cells(fila_actual, 12).Value = 0   # L: Peso Cobrado
        ws.Cells(fila_actual, 13).Value = 0   # M: Valor Declarado
        ws.Cells(fila_actual, 14).Value = 0   # N: Flete(1)
        ws.Cells(fila_actual, 15).Value = 0   # O: Costo Manejo(2)
        ws.Cells(fila_actual, 16).Value = 0   # P: Otros(3)
        
        # Q: Valor Total (Columna 17)
        ws.Cells(fila_actual, 17).Value = fila_sis.get('Total (1+2+3)', 0)

    # 4. ELIMINACIÓN DE FILAS VACÍAS (CIERRE DE ESPACIOS)
    ultima_fila_escrita = fila_inicio_plantilla + len(df_sistema) - 1
    
    if fila_total and ultima_fila_escrita < fila_total - 1:
        fila_vacia_inicio = ultima_fila_escrita + 1
        fila_vacia_fin = fila_total - 1
        try:
            ws.Range(f"{fila_vacia_inicio}:{fila_vacia_fin}").EntireRow.Delete()
            print(f"Espacio limpio Filas vacías eliminadas desde la {fila_vacia_inicio} hasta la {fila_vacia_fin}.")
        except Exception as e:
            print(f"No se pudieron borrar las filas vacías: {e}")

    # Guardar cambios y exportar a PDF de forma nativa
    
    # --- CORRECCIÓN 2: Enviar temporal de la plantilla a carpeta Temp de Windows ---
    excel_temporal = os.path.join(temp_dir, f"temp_formato_{nombre_carpeta}.xlsx")
    pdf_detalle_path = os.path.join(ruta_carpeta_fact, nuevo_nombre_detalle)
    
    try:
        if os.path.exists(excel_temporal): os.remove(excel_temporal)
        wb.SaveAs(excel_temporal)
        wb.ExportAsFixedFormat(0, pdf_detalle_path)
        wb.Close(False)
        if os.path.exists(excel_temporal): os.remove(excel_temporal)
    except Exception as e:
        print(f"Error al guardar PDF en {nombre_carpeta}: {e}")
        wb.Close(False)
        return False

    # 5. RENOMBRAR LOS ARCHIVOS RESTANTES EN LA CARPETA
    archivos_en_carpeta = os.listdir(ruta_carpeta_fact)
    pdfs_listos_para_zip = [pdf_detalle_path]

    for archivo in archivos_en_carpeta:
        if not archivo.lower().endswith(".pdf") or archivo == nuevo_nombre_detalle:
            continue
            
        ruta_archivo_original = os.path.join(ruta_carpeta_fact, archivo)
        archivo_lower = archivo.lower()
        nombre_final_archivo = None
        
        if "facturas" in archivo_lower:
            nombre_final_archivo = nuevo_nombre_factura
        elif archivo_lower.startswith("932"):
            nombre_final_archivo = nuevo_nombre_soportes
            
        if nombre_final_archivo:
            ruta_archivo_final = os.path.join(ruta_carpeta_fact, nombre_final_archivo)
            if archivo != nombre_final_archivo:
                try:
                    if os.path.exists(ruta_archivo_final): os.remove(ruta_archivo_final)
                    os.rename(ruta_archivo_original, ruta_archivo_final)
                except Exception as e:
                    print(f"No se pudo renombrar '{archivo}': {e}")
                    continue
            
            if ruta_archivo_final not in pdfs_listos_para_zip:
                pdfs_listos_para_zip.append(ruta_archivo_final)

    # 6. ENVASE EN ZIP Y ENVÍO A PENDIENTES
    if pdfs_listos_para_zip:
        ruta_zip_temporal = os.path.join(ruta_carpeta_fact, nombre_zip_final)
        try:
            if os.path.exists(ruta_zip_temporal): os.remove(ruta_zip_temporal)
            with zipfile.ZipFile(ruta_zip_temporal, 'w', zipfile.ZIP_DEFLATED) as zipf:
                for pdf_completo in pdfs_listos_para_zip:
                    if os.path.exists(pdf_completo):
                        zipf.write(pdf_completo, os.path.basename(pdf_completo))
            
            ruta_zip_destino_final = os.path.join(ruta_pendientes, nombre_zip_final)
            if os.path.exists(ruta_zip_destino_final): os.remove(ruta_zip_destino_final)
            shutil.move(ruta_zip_temporal, ruta_zip_destino_final)
        except Exception as e:
            print(f"Error procesando el ZIP de {nombre_carpeta}: {e}")

    # 7. ELIMINACIÓN DEL EXCEL ORIGINAL DEL SISTEMA
    try:
        if os.path.exists(excel_sistema_path):
            os.remove(excel_sistema_path)

    except Exception as e:
        print(f"No se pudo eliminar el archivo original excel sobrante en {nombre_carpeta}: {e}")

    return True

def procesar_facturacion_completa():
    
    if not os.path.exists(EXCEL_MAESTRO_PATH):
        print(f"no se encontro archivo maestro en: {EXCEL_MAESTRO_PATH}")
        return
    
    try:
        df_maestro = pd.read_excel(EXCEL_MAESTRO_PATH, sheet_name=NOMBRE_HOJA)
    except Exception as e:
        print(f"Error no existe la hoja '{NOMBRE_HOJA}' en el Excel maestro {e}")
        return

    df_maestro.columns = df_maestro.columns.str.strip()
    df_maestro['N CARPETA'] = df_maestro['N CARPETA'].astype(str).str.strip()

    ruta_mes = os.path.join(ESCRITORIO_PATH, MES_ACTUAL)
    ruta_pendientes = os.path.join(ruta_mes, "pendientes por radicar")
    os.makedirs(ruta_pendientes, exist_ok=True)

    if not os.path.exists(ruta_mes):
        print(f"La carpeta '{MES_ACTUAL}' no está en el Escritorio.")
        return

    carpetas_fisicas = [d for d in os.listdir(ruta_mes) if os.path.isdir(os.path.join(ruta_mes, d)) and d.upper().startswith("FACT")]

    if not carpetas_fisicas:
        return

    # Instanciar Excel una sola vez para optimizar velocidad
    excel_app = win32.Dispatch("Excel.Application")
    excel_app.Visible = False
    excel_app.DisplayAlerts = False

    try:
        for carpeta in carpetas_fisicas:
            ruta_carpeta_fact = os.path.join(ruta_mes, carpeta)
            
            fila_maestro = df_maestro[df_maestro['N CARPETA'] == carpeta]
            if fila_maestro.empty:
                print(f"'{carpeta}' no encontrada en la columna 'N CARPETA'")
                continue
                
            generar_e_integrar_carpeta(excel_app, ruta_carpeta_fact, carpeta, fila_maestro.iloc[0], ruta_pendientes)
    finally:
        excel_app.Quit()

    print("CARPETAS RENOMBRADAS")

if __name__ == "__main__":
    procesar_facturacion_completa()
