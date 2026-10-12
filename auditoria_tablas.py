"""Auditoría estricta de tablas MER contra diccionario fijo del formato oficial."""
from io import BytesIO
from pathlib import Path
import json
import re
from openpyxl import load_workbook
import pandas as pd

CAMPOS = ["Archivo", "Ruta", "Hoja", "Tabla esperada", "Tabla encontrada", "Estado", "Detalle"]
PATRON_COLUMNAS = ["Hoja", "Tabla esperada", "Archivos que la contienen", "Origen"]
DICCIONARIO = json.loads((Path(__file__).parent / 'diccionario_tablas_mer.json').read_text(encoding='utf-8'))['tablas_por_hoja']


def auditar_tablas_excel(entradas, referencia=None):
    """Referencia intencionalmente fija. Compara exclusivamente tablas de indicadores.
    Ignora Tabla1, tablas auxiliares y cualquier hoja que no figure en el diccionario.
    Una variación del nombre de un indicador se informa aunque lleve sufijo numérico.
    """
    filas = []
    for archivo, contenido, ruta in entradas:
        try:
            wb = load_workbook(BytesIO(contenido), read_only=False, data_only=True)
            inventario = {ws.title: set(ws.tables.keys()) for ws in wb.worksheets}
            wb.close()
        except Exception as exc:
            filas.append([archivo, ruta, "", "", "", "ERROR LECTURA", str(exc)[:350]])
            continue
        for hoja, esperados in DICCIONARIO.items():
            # Ausencia completa de la hoja no cuenta como error: algunos centros usan formatos parciales.
            if hoja not in inventario:
                continue
            nombres = inventario[hoja]
            for esperado in esperados:
                if esperado in nombres:
                    filas.append([archivo, ruta, hoja, esperado, esperado, "CORRECTO", "Nombre exacto del indicador."])
                    continue
                # Buscar el nombre cambiado solo entre objetos parecidos al esperado.
                parecidos = sorted(n for n in nombres if n.casefold().startswith(esperado.casefold()) and n != esperado)
                if not parecidos:
                    # Variantes triviales de separación / sufijos, sin confundir indicadores distintos
                    norm = lambda s: re.sub(r'[^a-z0-9]', '',s.casefold())
                    parecidos = sorted(n for n in nombres if norm(n).startswith(norm(esperado)) and n != esperado)
                if parecidos:
                    filas.append([archivo,ruta,hoja,esperado,", ".join(parecidos),"NOMBRE DIFERENTE","La tabla del indicador debe llamarse exactamente como en el formato base."])
                else:
                    filas.append([archivo,ruta,hoja,esperado,"","TABLA FALTANTE","No se encontró la tabla estructurada oficial en esta hoja."])
    patron = [{"Hoja":h,"Tabla esperada":", ".join(v),"Archivos que la contienen":sum(1 for f in filas if f[2]==h),"Origen":"Diccionario oficial integrado V6"} for h,v in DICCIONARIO.items()]
    return pd.DataFrame(filas,columns=CAMPOS),pd.DataFrame(patron,columns=PATRON_COLUMNAS)
