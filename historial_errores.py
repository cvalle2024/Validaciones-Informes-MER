"""Registro histórico local MER (SQLite). No almacena identificadores personales en claro."""
import hashlib
import io
import json
import os
import re
import sqlite3
import unicodedata
import uuid
from datetime import datetime, timezone
from pathlib import Path
import pandas as pd

DB_PATH = Path(os.environ.get('MER_HISTORY_DB', str(Path(__file__).parent / 'data' / 'historial_mer.sqlite3')))
MESES = {'ENE':1,'FEB':2,'MAR':3,'ABR':4,'MAY':5,'JUN':6,'JUL':7,'AGO':8,'SEP':9,'OCT':10,'NOV':11,'DIC':12,
         'JAN':1,'APR':4,'AUG':8,'DEC':12}

def _norm(value):
    s = unicodedata.normalize('NFKD', str(value or '')).encode('ascii', 'ignore').decode().casefold()
    return re.sub(r'\s+', ' ', s.replace('_', ' ')).strip()

def _periodo(value):
    s = _norm(value).upper()
    match = re.search(r'(?<!\d)(20\d{2})[-/](0?[1-9]|1[0-2])', s)
    if match: return f'{match.group(1)}-{int(match.group(2)):02d}'
    match = re.search(r'(ENE|FEB|MAR|ABR|MAY|JUN|JUL|AGO|SEP|OCT|NOV|DIC|JAN|APR|AUG|DEC)[A-Z]*\s*[-/ ]?\s*(20\d{2})', s)
    if match: return f'{match.group(2)}-{MESES[match.group(1)]:02d}'
    match = re.search(r'(?<!\d)(0?[1-9]|1[0-2])[-/](20\d{2})', s)
    if match: return f'{match.group(2)}-{int(match.group(1)):02d}'
    return ''

def _fiscal(periodo):
    if not periodo: return ('SIN PERIODO', 'SIN PERIODO')
    y, m = [int(x) for x in periodo.split('-')]
    fy = y + (m >= 10)
    q = (m-10)%12//3+1
    return f'FY{fy%100:02d}', f'FY{fy%100:02d}-Q{q}'

def conectar():
    DB_PATH.parent.mkdir(parents=True, exist_ok=True)
    con = sqlite3.connect(str(DB_PATH), timeout=20)
    con.execute('PRAGMA journal_mode=WAL')
    con.execute('''CREATE TABLE IF NOT EXISTS ejecuciones (
        run_id TEXT PRIMARY KEY, fecha TEXT NOT NULL, origen TEXT NOT NULL, total INTEGER NOT NULL)''')
    con.execute('''CREATE TABLE IF NOT EXISTS hallazgos (
        run_id TEXT NOT NULL, clave TEXT NOT NULL, fecha TEXT NOT NULL,
        periodo TEXT, fy TEXT, trimestre TEXT, pais TEXT, sitio TEXT,
        archivo TEXT, hoja TEXT, tipo TEXT, estado TEXT, fila TEXT,
        detalle TEXT, ocurrencia INTEGER NOT NULL, primera_vez TEXT NOT NULL,
        PRIMARY KEY (run_id, clave), FOREIGN KEY(run_id) REFERENCES ejecuciones(run_id))''')
    con.execute('CREATE INDEX IF NOT EXISTS idx_hallazgos_clave ON hallazgos(clave)')
    con.execute('CREATE INDEX IF NOT EXISTS idx_hallazgos_periodo ON hallazgos(periodo, trimestre)')
    return con

# No conservar IDs/expedientes, resultados clínicos u otros valores de pacientes.
def _safe_detail(row):
    for k in ('Estado','Error identificado','Detalle','Motivo'):
        v = row.get(k)
        if v is not None and str(v).strip() and str(v).lower() != 'nan':
            return str(v)[:180] if k in ('Estado','Error identificado') else ''
    return ''

def _value(row, *keys):
    for k in keys:
        v=row.get(k)
        if v is not None and pd.notna(v) and str(v).strip() and str(v).lower()!='nan': return str(v).strip()
    return ''

def registrar(diccionario, mes_default='', origen='Validación MER'):
    """Registra una ejecución real una sola vez; calcula recurrencia por firma de hallazgo."""
    fecha=datetime.now(timezone.utc).isoformat(timespec='seconds')
    run_id=uuid.uuid4().hex
    filas=[]; seen=set()
    con=conectar()
    try:
        con.execute('BEGIN IMMEDIATE')
        for tipo,df in diccionario.items():
            if df is None or df.empty or tipo == 'Auditoria_Sitio TX_CURR': continue
            for posicion,series in df.iterrows():
                row=series.to_dict()
                archivo=_value(row,'Archivo','Archivo(s) analizados','Ruta')
                sitio=_value(row,'Sitio','Sitio encontrado','Servicio de Salud','Establecimiento de Salud')
                pais=_value(row,'País','País reportado')
                hoja=_value(row,'Hoja')
                fila=_value(row,'Fila Excel','Fila','Número de fila')
                if not fila: fila=f'REG-{posicion}'
                detalle=_safe_detail(row)
                mes=_value(row,'Mes de reporte')
                periodo=_periodo(mes) or _periodo(mes_default)
                fy,q=_fiscal(periodo)
                # Firma no incluye valores identificadores de pacientes en claro.
                origen_id=[_norm(archivo),_norm(pais),_norm(sitio),_norm(hoja),_norm(tipo),_norm(fila),_norm(detalle),periodo]
                clave=hashlib.sha256(json.dumps(origen_id,ensure_ascii=False).encode()).hexdigest()
                if clave in seen: continue
                seen.add(clave)
                prev=con.execute('SELECT ocurrencia, primera_vez FROM hallazgos WHERE clave=? ORDER BY ocurrencia DESC LIMIT 1',(clave,)).fetchone()
                count=prev[0]+1 if prev else 1
                primera=prev[1] if prev else fecha
                filas.append((run_id,clave,fecha,periodo,fy,q,pais,sitio,archivo,hoja,tipo,detalle,fila,'',count,primera))
        con.execute('INSERT INTO ejecuciones VALUES (?,?,?,?)',(run_id,fecha,origen,len(filas)))
        con.executemany('''INSERT INTO hallazgos (run_id,clave,fecha,periodo,fy,trimestre,pais,sitio,archivo,hoja,tipo,estado,fila,detalle,ocurrencia,primera_vez)
                           VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)''',filas)
        con.commit()
    except Exception:
        con.rollback();raise
    finally: con.close()
    return run_id,len(filas)

def tablas():
    con=conectar()
    try:
        df=pd.read_sql_query('SELECT fecha AS "Fecha de validación", periodo AS "Mes", fy AS "Año fiscal", trimestre AS "Trimestre", pais AS "País", sitio AS "Sitio", archivo AS "Archivo", hoja AS "Hoja", tipo AS "Tipo de error", estado AS "Descripción", fila AS "Fila Excel", ocurrencia AS "N° detección", primera_vez AS "Primera detección", clave AS "Clave", run_id AS "Ejecución" FROM hallazgos ORDER BY fecha DESC',con)
        ejec=pd.read_sql_query('SELECT fecha AS "Fecha", origen AS "Origen", total AS "Hallazgos" FROM ejecuciones ORDER BY fecha DESC',con)
    finally: con.close()
    if df.empty: return df,pd.DataFrame(),pd.DataFrame(),ejec
    mensual=(df.groupby(['Mes','País','Tipo de error'],dropna=False).agg(
       Detecciones=('Clave','size'),Errores_distintos=('Clave','nunique'),Reincidencias=('N° detección',lambda x:(x>1).sum()),
       Archivos=('Archivo','nunique')).reset_index())
    trimestral=(df.groupby(['Año fiscal','Trimestre','País','Tipo de error'],dropna=False).agg(
       Detecciones=('Clave','size'),Errores_distintos=('Clave','nunique'),Reincidencias=('N° detección',lambda x:(x>1).sum()),
       Archivos=('Archivo','nunique')).reset_index())
    return df,mensual,trimestral,ejec

def exportar(df,mensual,trimestral,ejec):
    b=io.BytesIO()
    with pd.ExcelWriter(b,engine='openpyxl') as writer:
        for name,data in [('Historial',df),('Resumen mensual',mensual),('Resumen trimestral',trimestral),('Ejecuciones',ejec)]:
            data.drop(columns=['Clave','Ejecución'],errors='ignore').to_excel(writer,index=False,sheet_name=name)
            ws=writer.sheets[name]
            ws.freeze_panes='A2'; ws.auto_filter.ref=ws.dimensions
            from openpyxl.styles import Font,PatternFill
            for cell in ws[1]:
                cell.font=Font(bold=True,color='FFFFFF')
                cell.fill=PatternFill('solid',fgColor='176B83')
            from openpyxl.utils import get_column_letter
            for col in ws.columns:
                letter=get_column_letter(col[0].column)
                ws.column_dimensions[letter].width=min(36,max(14,max(len(str(c.value or '')) for c in list(col)[:100])+2))
    return b.getvalue()
