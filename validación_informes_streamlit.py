# -*- coding: utf-8 -*-
import io
import os
import re
import zipfile
import unicodedata
from datetime import datetime
from collections import defaultdict
from typing import Optional, Tuple, List, Dict
from PIL import Image
import pandas as pd
import streamlit as st
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
from openpyxl.utils import get_column_letter
from pathlib import Path

#st.set_page_config(page_title="Script validar indicadores MER", layout="wide", page_icon="✅")

LOGO_PATH = Path(__file__).parent / "logo.png"
logo_img = Image.open(LOGO_PATH)

st.set_page_config(page_title="Validaciones Maestro VIH", page_icon=logo_img, layout="wide")

col_logo, col_title = st.columns([1, 9])
with col_logo:
    st.image(logo_img, width=100)
with col_title:
    st.title("Validaciones Maestro VIH")
    st.caption("TX_PVLS / TX_CURR / HTS_TST • Revisión y métricas con filtros instantáneos")



# ============================
# --------- UI HEADER --------
# ============================
st.title("✅ Script de validación de indicadores MER (VIHCA)")
st.caption(
    "TX_PVLS / TX_CURR / HTS_TST "
    "• Comparación TX_CURR vs Dispensación_TARV por Sexo y Rango de edad"
)

with st.expander("ℹ️ Cómo usar", expanded=False):
    st.markdown(
        """
1) **Puedes cargar archivos excel uno a uno o si quiere leer varios archivos subelos como ZIP**.  
2) Pulsa **Procesar** (se ejecuta una sola vez).  
3) Usa los **segmentadores** (País / Depto / Sitio) — el filtrado es **instantáneo**.  
4) **Descarga** el Excel (completo o filtrado).  
        """
    )

# ====== DOCUMENTACIÓN (como expander) ======
DOC_MD = r"""
# 📖 Documentación de validaciones

## Indicadores y reglas
- **Numerador > Denominador (TX_PVLS):** Por sexo y edad, `Numerador ≤ Denominador`.Se verifica el valor de cada cela del Denominador y se compara con cada celda del Numerador para verificar que no sea mayor los valores del Numerador, Se detectan secciones “Numerador” y “Denominador”.
- **Denominador > TX_CURR (PVLS vs TX_CURR):** Por **sexo + tipo de población + edad**, Se verifica las celdas del Denominador TX_PVLS y se compara con las celdas del TX_CURR para validar que las celdas del Denominadpr no sean mayor al TX_CURR `Denominador (PVLS) ≤ TX_CURR`.
- **TX_CURR ≠ Dispensación_TARV (en hoja TX_CURR):** Dos cuadros en la misma hoja; se comparan por **sexo + edad** (no por población) y se reporta la **Diferencia (TX_CURR − Disp_TARV)** y si **Disp_TARV > TX_CURR**.
- **CD4 vacío positivo (HTS_TST):** Si `Resultado = Positivo`, **CD4 Basal** no puede estar vacío.
- **Fecha TARV < Diagnóstico (HTS_TST):** **Fecha inicio TARV** no puede ser anterior a la **Fecha del diagnóstico**.
- **Formato fecha diagnóstico (HTS_TST):** Si la fecha viene con `/`, debe cumplir **dd/mm/yyyy** válido (mes ≤ 12).

## Fuentes de “Mes de reporte”
- **HTS_TST:** desde **Fecha del diagnóstico** (por fila) y se normaliza a `MMM YYYY` (p.ej., `JUL 2025`).
- **TX_PVLS / TX_CURR:** prioridad **Fecha de reporte** > **Mes de reporte**; si no existen, se usa el fallback de la UI.

## Cómo se leen las tablas
- Se localizan encabezados con **“Sexo”** (y en su caso **“Tipo de población”**) y **columnas de edad** (contienen “año/ano/65/+”).
- En **TX_PVLS** se separan dinámicamente las secciones **Numerador** y **Denominador**.
- En **TX_CURR** se detectan los dos cuadros por rótulos: `TX CURR` y `Dispensación/Entrega TAR(V)`; se suman totales por **(Sexo, Edad)**.

## Métricas
- Por indicador se acumulan **checks** (comprobaciones) y **errors** (fallos).
- Se calcula `% Error = errors / checks` global y por combinación **(País, Depto, Sitio, Mes, Indicador)**.

## Exportación a Excel
- Hoja **Resumen** con total de errores por indicador.
- Una hoja por **indicador** con todas las filas detectadas.
- Hojas de **Métricas** (globales y por mes).
- En cada hoja de errores se **resalta en rojo** la columna crítica (p.ej., “CD4 Basal”, “Fecha inicio TARV”, “Diferencia (TX_CURR - Disp_TARV)”).

## Diagnóstico adicional
- Vista por archivo para **TX_CURR vs Dispensación_TARV** con orden por `|Diferencia|` y descarga **CSV**.
"""

with st.expander("📖 Documentación (clic para ver)", expanded=False):
    st.markdown(DOC_MD)
    st.download_button(
        "⬇️ Descargar documentación (Markdown)",
        DOC_MD.encode("utf-8"),
        file_name="documentacion_validaciones.md",
        mime="text/markdown",
        use_container_width=True,
    )
# ====== FIN DOCUMENTACIÓN ======


# ============================
# ------ CARGA DE INPUTS -----
# ============================
col_u1, col_u2 = st.columns([3, 2])
with col_u1:
    subida_multiple = st.file_uploader(
        "📂 Cargar .xlsx (varios) o 1 .zip con subcarpetas",
        type=["xlsx", "zip"],
        accept_multiple_files=True
    )
with col_u2:
    default_pais = st.text_input("País por defecto", "Desconocido")
    default_mes = st.text_input("Mes por defecto ", "Desconocido")

procesar = st.button("▶️ Procesar", use_container_width=True)

# ============================
# ---- ESTADO (CACHE/STORE) --
# ============================
for key, val in {
    "processed": False,
    "df_num": pd.DataFrame(),
    "df_txpv": pd.DataFrame(),
    "df_cd4": pd.DataFrame(),
    "df_tarv": pd.DataFrame(),
    "df_fdiag": pd.DataFrame(),
    "df_currq": pd.DataFrame(),  # diferencias TX_CURR vs Dispensación_TARV
    "metrics_global": defaultdict(lambda: {"errors": 0, "checks": 0}),
    "metrics_by_pds": defaultdict(lambda: {"errors": 0, "checks": 0}),
    "currq_debug": dict(),  # por archivo: (Sexo, Edad, TX_CURR, Dispensación_TARV, Dif)
}.items():
    if key not in st.session_state:
        st.session_state[key] = val

# ============================
# ----- CONSTANTES / HELPERS -
# ============================
IND_NUM_GT_DEN  = "num_gt_den"
IND_DEN_GT_CURR = "den_gt_curr"
IND_CD4_MISSING = "cd4_missing"
IND_TARV_LT_DIAG= "tarv_lt_diag"
IND_DIAG_BAD_FMT= "diag_bad_format"
IND_CURR_Q1Q2_DIFF = "curr_q1q2_diff"  # TX_CURR (cuadro) ≠ Dispensación_TARV

DISPLAY_NAMES = {
    IND_NUM_GT_DEN:      "Numerador > Denominador",
    IND_DEN_GT_CURR:     "Denominador > TX_CURR",
    IND_CD4_MISSING:     "CD4 vacío positivo",
    IND_TARV_LT_DIAG:    "Fecha TARV < Diagnóstico",
    IND_DIAG_BAD_FMT:    "Formato fecha diagnóstico",
    IND_CURR_Q1Q2_DIFF:  "TX_CURR ≠ Dispensación_TARV",
}

MESES = {
    "enero","febrero","marzo","abril","mayo","junio",
    "julio","agosto","septiembre","setiembre","octubre","noviembre","diciembre",
    "ene","feb","mar","abr","may","jun","jul","ago","sep","oct","nov","dic",
    "january","february","march","april","may","june",
    "july","august","september","october","november","december",
}
RUIDO_DIRS = {"", ".", "..", "__MACOSX", ".ds_store", ".git"}
INVALID_SHEET_CHARS = r'[:\\/?*\[\]]'

SPAN_ABBR = ["ENE","FEB","MAR","ABR","MAY","JUN","JUL","AGO","SEP","OCT","NOV","DIC"]

def _pct(e, c): return round((e / c * 100.0), 2) if c else 0.0

def _safe_sheet_name(name: str, used: set) -> str:
    base = unicodedata.normalize("NFKD", name).encode("ascii","ignore").decode("ascii")
    base = re.sub(INVALID_SHEET_CHARS, "-", base).strip() or "Hoja"
    base = base[:31]
    candidate = base; i = 1
    while candidate in used:
        suf = f"_{i}"; candidate = base[:31-len(suf)] + suf; i += 1
    used.add(candidate)
    return candidate

def _norm(s) -> str:
    if s is None: return ""
    s = str(s)
    s = unicodedata.normalize("NFKD", s).encode("ascii","ignore").decode("ascii")
    return re.sub(r"\s+", " ", s.strip().lower())

def _normalize_sexo(x) -> str:
    sx = _norm(x)
    if sx.startswith("masc"): return "Masculino"
    if sx.startswith("fem"):  return "Femenino"
    return str(x).strip()

def buscar_columna_multi(columnas, *patrones) -> Optional[str]:
    cols_norm = [_norm(c) for c in columnas]
    for i, cn in enumerate(cols_norm):
        if all(p in cn for p in patrones):
            return columnas[i]
    return None

def month_label_from_value(v: object) -> str:
    """
    Devuelve 'MMM YYYY' si v es fecha; si no es fecha, devuelve str(v).
    """
    if v is None or (isinstance(v, float) and pd.isna(v)):  # NaN
        return ""
    # intentar parsear como fecha
    try:
        dt = pd.to_datetime(v, dayfirst=True, errors="coerce")
        if pd.notna(dt):
            return f"{SPAN_ABBR[dt.month-1]} {dt.year}"
    except Exception:
        pass
    s = str(v).strip()
    return s

def inferir_pais_mes(path_rel: str, default_pais: str, default_mes: str):
    ruta = path_rel.replace("\\", "/")
    partes = [p for p in ruta.split("/") if p.strip().lower() not in RUIDO_DIRS]
    if partes and partes[-1].lower().endswith(".xlsx"): partes = partes[:-1]
    pais = partes[-2].strip() if len(partes) >= 2 else default_pais
    pais = pais or default_pais
    # Mes por ruta sólo si parece mes; si no, fallback
    mes = None
    for seg in reversed(partes):
        toks = re.split(r"[_\-\s/\.]+", _norm(seg))
        if any(t in MESES for t in toks):
            mes = seg.strip(); break
    if not mes:
        base = os.path.basename(path_rel)
        base = re.sub(r"\.xlsx$", "", base, flags=re.IGNORECASE)
        toks = re.split(r"[_\-\s/\.]+", _norm(base))
        if any(t in MESES for t in toks):
            mes = base
    mes = mes or default_mes
    return pais, mes

def leer_excel_desde_bytes(nombre, data_bytes): return pd.ExcelFile(io.BytesIO(data_bytes))

def encontrar_fila_encabezado(df_raw: pd.DataFrame, needles) -> Optional[int]:
    try:
        mask = df_raw.astype(str).apply(
            lambda r: all(any(needle.lower() in str(x).lower() for x in r.values) for needle in needles), axis=1
        ); idxs = df_raw[mask].index.tolist()
        return idxs[0] if idxs else None
    except Exception:
        return None

def normalizar_tabla_por_encabezado(df_raw: pd.DataFrame, idx_header: int):
    columnas = df_raw.iloc[idx_header].fillna("").astype(str).tolist()
    df_body = df_raw.iloc[idx_header + 1:].copy(); df_body.columns = columnas
    return df_body.reset_index(drop=True), columnas

def numeros_seguro(v): return pd.to_numeric(v, errors="coerce")

def _add_metric(ind_key: str, pais: str, mes_rep: str, depto: str = "", sitio: str = "",
                checks_add: int = 0, errors_add: int = 0):
    if checks_add:
        st.session_state.metrics_global[ind_key]["checks"] += checks_add
        st.session_state.metrics_by_pds[(pais or "", depto or "", sitio or "", mes_rep or "", ind_key)]["checks"] += checks_add
    if errors_add:
        st.session_state.metrics_global[ind_key]["errors"] += errors_add
        st.session_state.metrics_by_pds[(pais or "", depto or "", sitio or "", mes_rep or "", ind_key)]["errors"] += errors_add

# --- Normalizador de encabezados (evita KeyError por variantes) ---
def _rename_standard_columns(df: pd.DataFrame) -> pd.DataFrame:
    mapping: Dict[str, str] = {}
    for c in df.columns:
        cn = _norm(c)
        if not cn:
            continue
        if "sexo" in cn or "genero" in cn or "género" in cn:
            mapping[c] = "Sexo"
        elif ("tipo" in cn and "pobl" in cn) or "poblacion clave" in cn or "población clave" in cn:
            mapping[c] = "Tipo de población"
        elif "pais" in cn:
            mapping[c] = "País"
        elif "departamento" in cn or "depto" in cn or "provincia" in cn:
            mapping[c] = "Departamento"
        elif ("servicio" in cn and "salud" in cn) or "sitio" in cn or "clinica" in cn or "clínica" in cn:
            mapping[c] = "Sitio"
        elif "mes" in cn and "rep" in cn:
            mapping[c] = "Mes de reporte"
        elif "fecha" in cn and "reporte" in cn:
            mapping[c] = "Fecha de reporte"
    return df.rename(columns=mapping)

# --- Helpers específicos para HTS_TST ---
def _dedupe_columns(cols: List[str]) -> List[str]:
    seen: Dict[str, int] = {}
    out: List[str] = []
    for c in cols:
        if c in seen:
            seen[c] += 1
            out.append(f"{c}__{seen[c]}")
        else:
            seen[c] = 0
            out.append(c)
    return out

def _coerce_scalar(v):
    if isinstance(v, pd.Series):
        for x in v.tolist():
            if pd.notna(x) and str(x).strip():
                return x
        return ""
    return v

def _first_col(df: pd.DataFrame, *tokens) -> Optional[str]:
    for c in df.columns:
        cn = _norm(c)
        if all(t in cn for t in tokens):
            return c
    return None

# --- UI helper para evitar cajas "empty" ---
def show_df_or_note(df, note="— Sin filas para mostrar —", height=300):
    if df is None or (isinstance(df, pd.DataFrame) and df.empty):
        st.caption(note)
        return False
    st.dataframe(df, use_container_width=True, height=height)
    return True

# ============================
# ------- VALIDACIONES -------
# ============================
def procesar_tx_pvls_y_curr(
    xl: pd.ExcelFile, pais_inferido: str, mes_inferido: str, nombre_archivo: str,
    errores_numerador, errores_txpvls
):
    if "TX_PVLS" not in xl.sheet_names: return
    pvls_raw = xl.parse("TX_PVLS", header=None)
    idx_header = encontrar_fila_encabezado(pvls_raw, ["Sexo", "Tipo"])
    if idx_header is None: return
    df_data, columnas = normalizar_tabla_por_encabezado(pvls_raw, idx_header)
    df_data = _rename_standard_columns(df_data)

    # Localizar filas de numerador/denominador
    try:
        idx_num = df_data[df_data.iloc[:,0].astype(str).str.contains("Numerador", case=False, na=False)].index[0]
        idx_den = df_data[df_data.iloc[:,0].astype(str).str.contains("Denominador", case=False, na=False)].index[0]
    except IndexError:
        try:
            idx_num = df_data[df_data.get("TX_PVLS Numerador", "").astype(str).str.contains("Numerador", case=False, na=False)].index[0]
            idx_den = df_data[df_data.get("TX_PVLS Numerador", "").astype(str).str.contains("Denominador", case=False, na=False)].index[0]
        except Exception:
            return

    df_num = _rename_standard_columns(df_data.iloc[idx_num + 1:idx_den].copy())
    df_den = _rename_standard_columns(df_data.iloc[idx_den + 1:].copy())

    # Columnas de edad (tolerante)
    columnas_edad = [c for c in df_data.columns
                     if ("año" in c.lower()) or ("ano" in _norm(c)) or ("65" in c) or ("+" in c) or ("más" in c.lower() and "65" in c.lower())]

    # País/Depto/Sitio
    col_pais   = buscar_columna_multi(df_data.columns, "pais")
    col_depto  = buscar_columna_multi(df_data.columns, "departamento") or buscar_columna_multi(df_data.columns, "depto") or buscar_columna_multi(df_data.columns, "provincia")
    col_sitio  = buscar_columna_multi(df_data.columns, "servicio", "salud") or buscar_columna_multi(df_data.columns, "sitio") or buscar_columna_multi(df_data.columns, "clinica")

    # >>>>> MES / FECHA DE REPORTE (PRIORIDAD EN ESTAS HOJAS)
    col_fecha_rep = buscar_columna_multi(df_data.columns, "fecha", "reporte")
    col_mesrep    = buscar_columna_multi(df_data.columns, "mes", "reporte")
    def _ctx(row):
        p = str(row.get(col_pais)) if col_pais else pais_inferido
        d = str(row.get(col_depto)) if col_depto else ""
        s = str(row.get(col_sitio)) if col_sitio else ""
        # Prioridad: Fecha de reporte > Mes de reporte > fallback
        raw_mes = row.get(col_fecha_rep) if col_fecha_rep else (row.get(col_mesrep) if col_mesrep else None)
        m = month_label_from_value(raw_mes) or month_label_from_value(mes_inferido)
        return (p if str(p).strip() else pais_inferido,
                d if str(d).strip() else "",
                s if str(s).strip() else "",
                m if str(m).strip() else month_label_from_value(mes_inferido))
    # <<<<<

    fila_base_num = idx_header + 3 + idx_num + 1

    # Numerador > Denominador
    for i, row_num in df_num.iterrows():
        sexo = str(row_num.get("Sexo", "")).strip()
        pob  = str(row_num.get("Tipo de población", "")).strip()
        if _normalize_sexo(sexo) not in ["Masculino", "Femenino"]: continue
        row_den = df_den[(df_den["Sexo"].astype(str).str.strip()==sexo) &
                         (df_den["Tipo de población"].astype(str).str.strip()==pob)]
        if row_den.empty: continue
        row_den = row_den.iloc[0]
        pais_row, depto_row, sitio_row, mes_rep = _ctx(row_num)

        for col in columnas_edad:
            val_num = numeros_seguro(row_num.get(col))
            val_den = numeros_seguro(row_den.get(col))
            if pd.notna(val_num) and pd.notna(val_den):
                _add_metric(IND_NUM_GT_DEN, pais_row, mes_rep, depto_row, sitio_row, checks_add=1)
                if val_num > val_den:
                    _add_metric(IND_NUM_GT_DEN, pais_row, mes_rep, depto_row, sitio_row, errors_add=1)
                    col_idx = df_data.columns.tolist().index(col)
                    errores_numerador.append({
                        "País": pais_row, "Departamento": depto_row, "Sitio": sitio_row, "Mes de reporte": mes_rep,
                        "Archivo": nombre_archivo, "Sexo": sexo, "Tipo de población": pob, "Rango de edad": col,
                        "Numerador": float(val_num), "Denominador": float(val_den),
                        "Fila Excel": int(fila_base_num + i), "Columna Excel": get_column_letter(col_idx + 1)
                    })

    # Denominador (PVLS) > TX_CURR
    if "TX_CURR" in xl.sheet_names:
        curr_raw = xl.parse("TX_CURR", header=None)
        idx_curr = encontrar_fila_encabezado(curr_raw, ["Sexo", "Tipo"])
        if idx_curr is None: return
        df_curr, _ = normalizar_tabla_por_encabezado(curr_raw, idx_curr)
        df_curr = _rename_standard_columns(df_curr)
        fila_base_excel_den = idx_header + 3 + idx_den + 1

        for i, row_den in df_den.iterrows():
            sexo = str(row_den.get("Sexo", "")).strip()
            pob  = str(row_den.get("Tipo de población", "")).strip()
            if _normalize_sexo(sexo) not in ["Masculino", "Femenino"]: continue
            row_curr = df_curr[(df_curr["Sexo"].astype(str).str.strip()==sexo) &
                               (df_curr["Tipo de población"].astype(str).str.strip()==pob)]
            if row_curr.empty: continue
            row_curr = row_curr.iloc[0]

            pais_row, depto_row, sitio_row, mes_rep = _ctx(row_den)
            for col in columnas_edad:
                val_den = numeros_seguro(row_den.get(col))
                val_curr = numeros_seguro(row_curr.get(col))
                if pd.notna(val_den) and pd.notna(val_curr):
                    _add_metric(IND_DEN_GT_CURR, pais_row, mes_rep, depto_row, sitio_row, checks_add=1)
                    if val_den > val_curr:
                        _add_metric(IND_DEN_GT_CURR, pais_row, mes_rep, depto_row, sitio_row, errors_add=1)
                        col_idx = df_data.columns.tolist().index(col)
                        errores_txpvls.append({
                            "País": pais_row, "Departamento": depto_row, "Sitio": sitio_row, "Mes de reporte": mes_rep,
                            "Archivo": nombre_archivo, "Sexo": sexo, "Tipo de población": pob, "Rango de edad": col,
                            "Denominador (PVLS)": float(val_den), "TX_CURR": float(val_curr),
                            "Fila Excel": int(fila_base_excel_den + i), "Columna Excel": get_column_letter(col_idx + 1)
                        })

def procesar_hts_tst(
    xl: pd.ExcelFile, pais_inferido: str, mes_inferido: str, nombre_archivo: str,
    errores_cd4, errores_fecha_tarv, errores_formato_fecha_diag
):
    """
    HTS_TST: 'Mes de reporte' se toma de la columna 'Fecha del diagnóstico ...' (por fila).
    """
    if "HTS_TST" not in xl.sheet_names:
        return

    df_raw = xl.parse("HTS_TST", header=None)
    idx_hts = encontrar_fila_encabezado(df_raw, ["Resultado", "CD4"])
    if idx_hts is None:
        return

    df_data, _ = normalizar_tabla_por_encabezado(df_raw, idx_hts)
    df_data.columns = _dedupe_columns(df_data.columns)
    df_data = _rename_standard_columns(df_data)

    # localizar columnas
    col_resultado = _first_col(df_data, "resultado")
    col_cd4       = _first_col(df_data, "cd4")
    col_tarv      = _first_col(df_data, "inicio", "tar")
    col_diag      = _first_col(df_data, "fecha", "diagn")  # <- usar como Mes de reporte
    col_sitio     = _first_col(df_data, "servicio", "salud") or _first_col(df_data, "sitio") or _first_col(df_data, "clinica")
    col_pais      = _first_col(df_data, "pais")
    col_depto     = _first_col(df_data, "departamento") or _first_col(df_data, "depto") or _first_col(df_data, "provincia")

    if not all([col_resultado, col_cd4, col_diag]):
        return

    fila_base_hts = idx_hts + 3
    for i, row in df_data.iterrows():
        resultado  = str(_coerce_scalar(row.get(col_resultado))).strip().lower()
        cd4        = _coerce_scalar(row.get(col_cd4))
        fecha_diag = _coerce_scalar(row.get(col_diag))
        fecha_tarv = _coerce_scalar(row.get(col_tarv)) if col_tarv else None
        sitio      = _coerce_scalar(row.get(col_sitio)) if col_sitio else ""
        pais_row   = _coerce_scalar(row.get(col_pais))  if col_pais else pais_inferido
        depto_row  = _coerce_scalar(row.get(col_depto)) if col_depto else ""

        # >>> Mes de reporte desde Fecha del diagnóstico
        mes_rep    = month_label_from_value(fecha_diag) or month_label_from_value(mes_inferido)
        # <<<

        pais_row   = str(pais_row).strip() or pais_inferido
        depto_row  = str(depto_row).strip()
        sitio_row  = str(sitio).strip()
        mes_rep    = str(mes_rep).strip() or month_label_from_value(mes_inferido)

        # CD4 vacío cuando Resultado = Positivo
        if resultado == "positivo":
            _add_metric(IND_CD4_MISSING, pais_row, mes_rep, depto_row, sitio_row, checks_add=1)
            if pd.isna(cd4) or str(cd4).strip() == "":
                _add_metric(IND_CD4_MISSING, pais_row, mes_rep, depto_row, sitio_row, errors_add=1)
                errores_cd4.append({
                    "País": pais_row, "Departamento": depto_row, "Sitio": sitio_row, "Mes de reporte": mes_rep,
                    "Archivo": nombre_archivo, "Resultado": "Positivo", "CD4 Basal": "",
                    "Fila Excel": int(fila_base_hts + i), "Columna Excel": col_cd4
                })

        # Fecha TARV < Diagnóstico
        if pd.notna(fecha_diag) and pd.notna(fecha_tarv) and str(fecha_tarv).strip():
            try:
                fd = pd.to_datetime(fecha_diag, dayfirst=True, errors="coerce")
                ft = pd.to_datetime(fecha_tarv, dayfirst=True, errors="coerce")
                if pd.notna(fd) and pd.notna(ft):
                    _add_metric(IND_TARV_LT_DIAG, pais_row, mes_rep, depto_row, sitio_row, checks_add=1)
                    if ft < fd:
                        _add_metric(IND_TARV_LT_DIAG, pais_row, mes_rep, depto_row, sitio_row, errors_add=1)
                        errores_fecha_tarv.append({
                            "País": pais_row, "Departamento": depto_row, "Sitio": sitio_row, "Mes de reporte": mes_rep,
                            "Archivo": nombre_archivo, "Resultado": "Positivo" if resultado == "positivo" else "",
                            "Fecha diagnóstico": fd.date(), "Fecha inicio TARV": ft.date(),
                            "Fila Excel": int(fila_base_hts + i), "Columna Excel": col_tarv
                        })
            except Exception:
                pass

        # Validación de formato de Fecha Diagnóstico (dd/mm/yyyy)
        try:
            fecha_texto = str(fecha_diag).strip()
            if fecha_texto and "/" in fecha_texto:
                _add_metric(IND_DIAG_BAD_FMT, pais_row, mes_rep, depto_row, sitio_row, checks_add=1)
                partes = fecha_texto.split("/")
                if len(partes) == 3:
                    dia, mes_, anio = partes
                    if int(mes_) > 12: raise ValueError
                    datetime.strptime(fecha_texto, "%d/%m/%Y")
        except Exception:
            _add_metric(IND_DIAG_BAD_FMT, pais_row, mes_rep, depto_row, sitio_row, errors_add=1)
            errores_formato_fecha_diag.append({
                "País": pais_row, "Departamento": depto_row, "Sitio": sitio_row, "Mes de reporte": mes_rep,
                "Archivo": nombre_archivo,
                "Fecha del diagnóstico de la prueba": fecha_diag,
                "Fila Excel": int(fila_base_hts + i), "Columna Excel": col_diag
            })

# ====== TX_CURR: comparar "TX_CURR" vs "Dispensación_TARV" (o Entrega TARV) solo por SEXO y EDAD ======
def procesar_tx_curr_cuadros(
    xl: pd.ExcelFile, pais_inferido: str, mes_inferido: str,
    nombre_archivo: str, errores_currq, debug_store: Optional[dict] = None
):
    sheet_name = "TX_CURR"
    if sheet_name not in xl.sheet_names: return
    df_raw = xl.parse(sheet_name, header=None)
    nrows, ncols = df_raw.shape

    def _find_label_positions(alternatives: List[List[str]]) -> List[Tuple[int,int]]:
        pos = []
        for r in range(nrows):
            for c in range(ncols):
                s = _norm(df_raw.iat[r, c])
                if not s:
                    continue
                for toks in alternatives:
                    if all(tok in s for tok in toks):
                        pos.append((r, c)); break
        return pos

    # Rótulos tolerantes
    pos_tx = _find_label_positions([["tx", "curr"]])
    pos_et = _find_label_positions([
        ["dispens", "tar"],  # Dispensación TARV
        ["dispensacion", "tar"],
        ["dispensación", "tar"],
        ["entrega", "tar"],  # Entrega TARV/TAVR
        ["entrega", "tavr"],
    ])
    if not pos_tx or not pos_et:
        return

    fila_tx = min(pos_tx)[0]
    fila_et = min(pos_et)[0]

    def _find_header_after(start_row: int) -> Optional[int]:
        for r in range(start_row, min(start_row + 50, nrows)):
            row_vals = df_raw.iloc[r].tolist()
            if any("sexo" in _norm(x) for x in row_vals):
                cols = [str(x) for x in row_vals]
                cols_norm = [_norm(x) for x in cols]
                try:
                    col_sexo = next(i for i, cn in enumerate(cols_norm) if "sexo" in cn)
                except StopIteration:
                    continue
                edad_ok = any(
                    ("ano" in _norm(cols[j])) or ("año" in _norm(cols[j])) or re.search(r"\b65\b", _norm(cols[j])) or
                    ("+" in _norm(cols[j])) or ("mas" in _norm(cols[j]) and "65" in _norm(cols[j]))
                    for j in range(col_sexo + 1, ncols)
                )
                if edad_ok:
                    return r
        return None

    hdr_tx = _find_header_after(fila_tx)
    hdr_et = _find_header_after(fila_et)
    if hdr_tx is None or hdr_et is None:
        return

    def _extract_table_totals(header_row: int, stop_row: Optional[int]):
        cols = df_raw.iloc[header_row].fillna("").astype(str).tolist()
        cols_norm = [_norm(x) for x in cols]
        try:
            col_sexo = next(i for i, cn in enumerate(cols_norm) if "sexo" in cn)
        except StopIteration:
            return {}, {}, None

        edad_idx, edad_key, edad_map = [], [], {}
        for j in range(col_sexo + 1, ncols):
            lab  = cols[j]
            labn = _norm(lab)
            if ("ano" in labn) or ("año" in labn) or re.search(r"\b65\b", labn) or ("+" in labn) or ("mas" in labn and "65" in labn):
                k = labn
                edad_idx.append(j); edad_key.append(k); edad_map[k] = str(lab)

        if not edad_idx:
            return {}, {}, col_sexo

        totals: Dict[Tuple[str,str], int] = {}
        r = header_row + 1
        while r < nrows:
            if stop_row is not None and r >= stop_row: break
            rown = [_norm(x) for x in df_raw.iloc[r].tolist()]
            if any("sexo" in x for x in rown): break  # nuevo header

            sx_raw = df_raw.iat[r, col_sexo]
            sx = _normalize_sexo(sx_raw)
            if sx in ["Masculino", "Femenino"]:
                for c_idx, ekey in zip(edad_idx, edad_key):
                    v = pd.to_numeric(df_raw.iat[r, c_idx], errors="coerce")
                    if pd.notna(v):
                        v = int(round(float(v)))
                        key = (sx, ekey)
                        totals[key] = totals.get(key, 0) + v
            r += 1

        return totals, edad_map, col_sexo

    # Delimitar rangos entre cuadros
    if hdr_tx < hdr_et:
        totals_tx, edades_tx, col_sexo_tx = _extract_table_totals(hdr_tx, hdr_et)
        totals_et, edades_et, col_sexo_et = _extract_table_totals(hdr_et, None)
    else:
        totals_et, edades_et, col_sexo_et = _extract_table_totals(hdr_et, hdr_tx)
        totals_tx, edades_tx, col_sexo_tx = _extract_table_totals(hdr_tx, None)

    # Contexto desde fila debajo del header de TX_CURR
    cols_hdr = df_raw.iloc[hdr_tx].fillna("").astype(str).tolist()
    cols_hdrn = [_norm(x) for x in cols_hdr]

    def _find_col_idx(cols_norm, *patrones):
        for i, cn in enumerate(cols_norm):
            if all(p in cn for p in patrones):
                return i
        return None

    col_pais_i   = _find_col_idx(cols_hdrn, "pais")
    col_depto_i  = (_find_col_idx(cols_hdrn, "departamento") or
                    _find_col_idx(cols_hdrn, "depto") or
                    _find_col_idx(cols_hdrn, "provincia"))
    col_sitio_i  = (_find_col_idx(cols_hdrn, "servicio") or
                    _find_col_idx(cols_hdrn, "salud") or
                    _find_col_idx(cols_hdrn, "sitio") or
                    _find_col_idx(cols_hdrn, "clinica"))
    # >>> Permite Fecha de reporte o Mes de reporte
    col_mesrep_i = (_find_col_idx(cols_hdrn, "fecha") if _find_col_idx(cols_hdrn, "reporte") is not None else None)
    if col_mesrep_i is None:
        col_mesrep_i = _find_col_idx(cols_hdrn, "mes")
    # <<<

    def _ctx_from_rowvals(row_vals):
        def _val(idx, fallback):
            try:
                v = row_vals[idx] if idx is not None else fallback
            except Exception:
                v = fallback
            v = str(v).strip()
            return v if v else fallback
        p = _val(col_pais_i, pais_inferido)
        d = _val(col_depto_i, "")
        s = _val(col_sitio_i, "")
        # Prioridad aquí: Fecha de reporte > Mes de reporte > fallback
        raw_mes = row_vals[col_mesrep_i] if col_mesrep_i is not None else None
        m = month_label_from_value(raw_mes) or month_label_from_value(mes_inferido)
        return p, d, s, m

    fila_ctx_vals = df_raw.iloc[hdr_tx + 1].fillna("").astype(str).tolist() if (hdr_tx + 1) < nrows else []
    pais_row, depto_row, sitio_row, mes_rep = _ctx_from_rowvals(fila_ctx_vals)

    # Comparación por (Sexo, Edad)
    rows_debug = []
    all_keys = set(totals_tx.keys()) | set(totals_et.keys())
    for (sexo, edad_key) in sorted(all_keys):
        v_tx = int(totals_tx.get((sexo, edad_key), 0))
        v_et = int(totals_et.get((sexo, edad_key), 0))
        etiqueta_edad = edades_tx.get(edad_key) or edades_et.get(edad_key) or edad_key

        _add_metric(IND_CURR_Q1Q2_DIFF, pais_row, mes_rep, depto_row, sitio_row, checks_add=1)
        if v_tx != v_et:
            _add_metric(IND_CURR_Q1Q2_DIFF, pais_row, mes_rep, depto_row, sitio_row, errors_add=1)
            errores_currq.append({
                "País": pais_row,
                "Departamento": depto_row,
                "Sitio": sitio_row,
                "Mes de reporte": mes_rep,
                "Archivo": nombre_archivo,
                "Sexo": sexo,
                "Rango de edad": etiqueta_edad,
                "TX_CURR": v_tx,
                "Dispensación_TARV": v_et,
                "Diferencia (TX_CURR - Disp_TARV)": v_tx - v_et,
                "Disp_TARV > TX_CURR": "Sí" if v_et > v_tx else "No",
            })

        rows_debug.append({
            "Sexo": sexo,
            "Rango de edad": etiqueta_edad,
            "TX_CURR": v_tx,
            "Dispensación_TARV": v_et,
            "Diferencia": v_tx - v_et
        })

    if debug_store is not None:
        df_dbg = pd.DataFrame(rows_debug).sort_values(["Sexo", "Rango de edad"])
        df_dbg.attrs["modo"] = "por_rotulos"
        debug_store[nombre_archivo] = df_dbg

# ============================
# ------- EXPORTACIÓN --------
# ============================
def exportar_excel_resultados(errores_dict, df_metricas_global: pd.DataFrame, df_metricas_group: pd.DataFrame) -> bytes:
    config_resaltado = {
        "Numerador > Denominador": "Numerador",
        "Denominador > TX_CURR": "Denominador (PVLS)",
        "CD4 vacío positivo": "CD4 Basal",
        "Fecha TARV < Diagnóstico": "Fecha inicio TARV",
        "Formato fecha diagnóstico": "Fecha del diagnóstico de la prueba",
        "TX_CURR ≠ Dispensación_TARV": "Diferencia (TX_CURR - Disp_TARV)",
    }
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        used = set()
        resumen = pd.DataFrame({
            "Tipo de Error": list(errores_dict.keys()),
            "Total": [len(df) for df in errores_dict.values()]
        })
        resumen.to_excel(writer, index=False, sheet_name=_safe_sheet_name("Resumen", used))
        for nombre, df in errores_dict.items():
            if df is not None and not df.empty:
                df.to_excel(writer, index=False, sheet_name=_safe_sheet_name(nombre, used))
        if df_metricas_global is not None and not df_metricas_global.empty:
            df_metricas_global.to_excel(writer, index=False, sheet_name=_safe_sheet_name("Metricas Globales Seleccion", used))
        if df_metricas_group is not None and not df_metricas_group.empty:
            df_metricas_group.to_excel(writer, index=False, sheet_name=_safe_sheet_name("Metricas por Mes", used))

    buffer.seek(0); wb = load_workbook(buffer)
    rojo = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")

    for nombre, df in errores_dict.items():
        if df is None or df.empty: continue
        target = None
        for ws in wb.worksheets:
            if ws.title.lower().startswith(_safe_sheet_name(nombre, set()).lower()[:5]):
                target = ws.title; break
        if not target: continue
        ws = wb[target]
        campo_rojo = config_resaltado.get(nombre)
        if campo_rojo and campo_rojo in df.columns:
            col_idx = list(df.columns).index(campo_rojo) + 1
            for row in range(2, ws.max_row + 1):
                ws.cell(row=row, column=col_idx).fill = rojo
        ws.auto_filter.ref = ws.dimensions

    out = io.BytesIO(); wb.save(out); out.seek(0)
    return out.getvalue()

# ============================
# --------- PROCESO ----------
# ============================
if procesar:
    if not subida_multiple:
        st.warning("Primero carga archivos .xlsx o un .zip.")
        st.stop()

    entradas = []
    for up in subida_multiple:
        nombre = up.name; data = up.read()
        if nombre.lower().endswith(".zip"):
            with zipfile.ZipFile(io.BytesIO(data)) as zf:
                for info in zf.infolist():
                    if info.is_dir(): continue
                    if info.filename.lower().endswith(".xlsx") and not os.path.basename(info.filename).startswith("~$"):
                        entradas.append((os.path.basename(info.filename), zf.read(info.filename), info.filename))
        else:
            if nombre.lower().endswith(".xlsx") and not os.path.basename(nombre).startswith("~$"):
                entradas.append((os.path.basename(nombre), data, nombre))

    if not entradas:
        st.error("No se encontraron archivos .xlsx válidos.")
        st.stop()

    errores_numerador = []
    errores_txpvls = []
    errores_cd4 = []
    errores_fecha_tarv = []
    errores_formato_fecha_diag = []
    errores_currq = []

    # Reiniciar métricas y debug
    st.session_state.metrics_global = defaultdict(lambda: {"errors": 0, "checks": 0})
    st.session_state.metrics_by_pds = defaultdict(lambda: {"errors": 0, "checks": 0})
    st.session_state.currq_debug = {}

    progreso = st.progress(0.0, text="Procesando archivos…"); total = len(entradas)
    for idx, (nombre_archivo, data_bytes, ruta_rel) in enumerate(entradas, start=1):
        try:
            pais_inf, mes_inf = inferir_pais_mes(ruta_rel.replace("\\", "/"), default_pais, default_mes)
            xl = leer_excel_desde_bytes(nombre_archivo, data_bytes)
            procesar_tx_pvls_y_curr(xl, pais_inf, mes_inf, nombre_archivo, errores_numerador, errores_txpvls)
            procesar_hts_tst(xl, pais_inf, mes_inf, nombre_archivo, errores_cd4, errores_fecha_tarv, errores_formato_fecha_diag)
            procesar_tx_curr_cuadros(xl, pais_inf, mes_inf, nombre_archivo, errores_currq, debug_store=st.session_state.currq_debug)
        except Exception as e:
            st.warning(f"⚠️ Este error es por el nombre del campo que no es igual[ {nombre_archivo}: {e}")
        progreso.progress(idx/total, text=f"Procesando {idx} de {total}…")

    # Guardar DataFrames en sesión
    st.session_state.df_num   = pd.DataFrame(errores_numerador)
    st.session_state.df_txpv  = pd.DataFrame(errores_txpvls)
    st.session_state.df_cd4   = pd.DataFrame(errores_cd4)
    st.session_state.df_tarv  = pd.DataFrame(errores_fecha_tarv)
    st.session_state.df_fdiag = pd.DataFrame(errores_formato_fecha_diag)
    st.session_state.df_currq = pd.DataFrame(errores_currq)
    st.session_state.processed = True
    st.success("Procesamiento completado. Ahora puedes filtrar al instante ✅")

# ============================
# ------- INTERFAZ (LIVE) ----
# ============================
if not st.session_state.processed:
    st.info("Carga tus archivos y pulsa **Procesar**.")
    st.stop()

# Asegurar columnas para filtros
for dfname in ["df_num","df_txpv","df_cd4","df_tarv","df_fdiag","df_currq"]:
    df = st.session_state[dfname]
    if not df.empty:
        for col in ["País","Departamento","Sitio","Mes de reporte"]:
            if col not in df.columns:
                st.session_state[dfname][col] = ""

# Resumen de conteos
st.subheader("📌 Resumen (conteo de filas de error)")
c1, c2, c3, c4, c5, c6 = st.columns(6)
c1.metric("Numerador > Denominador", len(st.session_state.df_num))
c2.metric("Denominador > TX_CURR", len(st.session_state.df_txpv))
c3.metric("CD4 vacío positivo", len(st.session_state.df_cd4))
c4.metric("TARV < Diagnóstico", len(st.session_state.df_tarv))
c5.metric("Fecha diag. mal formateada", len(st.session_state.df_fdiag))
c6.metric("TX_CURR ≠ Dispensación_TARV", len(st.session_state.df_currq))

# Segmentadores
st.subheader("🎛️ Segmentadores")
df_all = pd.concat(
    [df for df in [
        st.session_state.df_num, st.session_state.df_txpv, st.session_state.df_cd4,
        st.session_state.df_tarv, st.session_state.df_fdiag, st.session_state.df_currq
    ] if not df.empty],
    ignore_index=True
) if any([not st.session_state[k].empty for k in ["df_num","df_txpv","df_cd4","df_tarv","df_fdiag","df_currq"]]) else pd.DataFrame(columns=["País","Departamento","Sitio","Mes de reporte"])

paises  = ["Todos"] + sorted([p for p in df_all["País"].dropna().unique().tolist() if str(p).strip()]) if not df_all.empty else ["Todos"]
departs = ["Todos"] + sorted([d for d in df_all["Departamento"].dropna().unique().tolist() if str(d).strip()]) if not df_all.empty else ["Todos"]
sitios  = ["Todos"] + sorted([s for s in df_all["Sitio"].dropna().unique().tolist() if str(s).strip()]) if not df_all.empty else ["Todos"]

fc1, fc2, fc3 = st.columns(3)
with fc1: sel_pais = st.selectbox("País", paises, index=0)
with fc2: sel_depto = st.selectbox("Departamento", departs, index=0)
with fc3: sel_sitio = st.selectbox("Sitio", sitios, index=0)

def _aplicar_filtro(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty: return df
    m = pd.Series([True] * len(df))
    if sel_pais != "Todos": m &= (df["País"] == sel_pais)
    if sel_depto != "Todos": m &= (df["Departamento"] == sel_depto)
    if sel_sitio != "Todos": m &= (df["Sitio"] == sel_sitio)
    return df[m].copy()

df_num_f   = _aplicar_filtro(st.session_state.df_num)
df_txpv_f  = _aplicar_filtro(st.session_state.df_txpv)
df_cd4_f   = _aplicar_filtro(st.session_state.df_cd4)
df_tarv_f  = _aplicar_filtro(st.session_state.df_tarv)
df_fdiag_f = _aplicar_filtro(st.session_state.df_fdiag)
df_currq_f = _aplicar_filtro(st.session_state.df_currq)

# Métricas (adaptadas a la selección)
def _build_metrics_df_from_selection(sel_pais, sel_depto, sel_sitio):
    agg = defaultdict(lambda: {"errors": 0, "checks": 0})
    for (pais, depto, sitio, mes_rep, ind), v in st.session_state.metrics_by_pds.items():
        if (sel_pais == "Todos" or pais == sel_pais) and \
           (sel_depto == "Todos" or depto == sel_depto) and \
           (sel_sitio == "Todos" or sitio == sel_sitio):
            agg[ind]["errors"] += v["errors"]; agg[ind]["checks"] += v["checks"]

    df_global = (pd.DataFrame([
        {"Indicador": DISPLAY_NAMES.get(k, k), "Errores": v["errors"], "Chequeos": v["checks"], "% Error": _pct(v["errors"], v["checks"])}
        for k, v in agg.items()
    ]).sort_values("% Error", ascending=False)
    if agg else pd.DataFrame(columns=["Indicador","Errores","Chequeos","% Error"]))

    rows = []
    for (pais, depto, sitio, mes_rep, ind), v in st.session_state.metrics_by_pds.items():
        if (sel_pais == "Todos" or pais == sel_pais) and \
           (sel_depto == "Todos" or depto == sel_depto) and \
           (sel_sitio == "Todos" or sitio == sel_sitio):
            rows.append({
                "País": pais, "Departamento": depto, "Sitio": sitio, "Mes de reporte": mes_rep,
                "Indicador": DISPLAY_NAMES.get(ind, ind), "Errores": v["errors"], "Chequeos": v["checks"],
                "% Error": _pct(v["errors"], v["checks"])
            })
    df_group = pd.DataFrame(rows)
    if not df_group.empty:
        df_group = df_group[["País","Departamento","Sitio","Mes de reporte","Indicador","Errores","Chequeos","% Error"]]
        df_group = df_group.sort_values(["País","Departamento","Sitio","Indicador"])
    else:
        df_group = pd.DataFrame(columns=["País","Departamento","Sitio","Mes de reporte","Indicador","Errores","Chequeos","% Error"])
    return df_global, df_group

st.subheader("📈 Métricas de calidad (adaptadas al filtro)")
df_metricas_global_sel, df_metricas_por_mes_sel = _build_metrics_df_from_selection(sel_pais, sel_depto, sel_sitio)

gc1, gc2 = st.columns([1.2, 2])
with gc1:
    st.markdown("**Métricas – Selección actual**")
    show_df_or_note(df_metricas_global_sel, "— Sin métricas para la selección —", height=260)
with gc2:
    st.markdown("**Desglose por Mes – Selección**")
    show_df_or_note(df_metricas_por_mes_sel, "— Sin desglose para la selección —", height=260)

st.markdown("**Indicadores – % de error (selección)**")
cards = [IND_NUM_GT_DEN, IND_DEN_GT_CURR, IND_CD4_MISSING, IND_TARV_LT_DIAG, IND_DIAG_BAD_FMT, IND_CURR_Q1Q2_DIFF]
cc1, cc2, cc3, cc4, cc5, cc6 = st.columns(6)
cols = [cc1, cc2, cc3, cc4, cc5, cc6]
sel_map = {row["Indicador"]: row for _, row in df_metricas_global_sel.iterrows()} if not df_metricas_global_sel.empty else {}
for col, key in zip(cols, cards):
    name = DISPLAY_NAMES[key]
    v = sel_map.get(name, {"Errores":0, "Chequeos":0, "% Error":0})
    col.metric(label=name, value=f"{v.get('% Error',0)}%", delta=f"{v.get('Errores',0)} / {v.get('Chequeos',0)} err/cheq")

# Pestañas (sin cajas "empty")
tabs = st.tabs([
    "Numerador > Denominador",
    "Denominador > TX_CURR",
    "CD4 vacío positivo",
    "Fecha TARV < Diagnóstico",
    "Formato fecha diagnóstico",
    "TX_CURR ≠ Dispensación_TARV",
])
with tabs[0]: show_df_or_note(df_num_f,   "— Sin diferencias de Numerador > Denominador —")
with tabs[1]: show_df_or_note(df_txpv_f,  "— Sin casos Denominador > TX_CURR —")
with tabs[2]: show_df_or_note(df_cd4_f,   "— Sin positivos con CD4 vacío —")
with tabs[3]: show_df_or_note(df_tarv_f,  "— Sin casos TARV < Diagnóstico —")
with tabs[4]: show_df_or_note(df_fdiag_f, "— Sin problemas de formato de fecha —")
with tabs[5]: show_df_or_note(df_currq_f, "— TX_CURR = Dispensación_TARV en la selección —")

# Diagnóstico por archivo
st.subheader("🔍 Diagnóstico TX_CURR vs Dispensación_TARV por archivo")
if st.session_state.currq_debug:
    archs = sorted(st.session_state.currq_debug.keys())
    col_d1, col_d2 = st.columns([2, 1])
    with col_d1:
        sel_arch = st.selectbox("Archivo", archs, index=0)
    with col_d2:
        ordenar_por_abs = st.checkbox("Ordenar por |Diferencia| desc.", value=True)

    df_dbg = st.session_state.currq_debug.get(sel_arch, pd.DataFrame()).copy()
    if not df_dbg.empty:
        if ordenar_por_abs:
            df_dbg["abs_diff"] = df_dbg["Diferencia"].abs()
            df_dbg = df_dbg.sort_values(["abs_diff","Sexo","Rango de edad"], ascending=[False, True, True]).drop(columns=["abs_diff"])
        sexo_opts = ["Todos"] + sorted(df_dbg["Sexo"].dropna().unique().tolist())
        sel_sexo = st.selectbox("Filtrar por Sexo", sexo_opts, index=0)
        if sel_sexo != "Todos":
            df_dbg = df_dbg[df_dbg["Sexo"] == sel_sexo]

        show_df_or_note(df_dbg, "— Sin datos de diagnóstico para este archivo —", height=320)
        if not df_dbg.empty:
            csv = df_dbg.to_csv(index=False).encode("utf-8-sig")
            st.download_button("⬇️ Descargar diagnóstico (CSV)", data=csv,
                               file_name=f"DIAGNOSTICO_TX_CURR_{sel_arch}.csv",
                               mime="text/csv", use_container_width=True)
            modo = getattr(df_dbg, "attrs", {}).get("modo", "")
            if modo:
                st.caption(f"Modo de parsing para **{sel_arch}**: **{modo}**")
    else:
        st.caption("— Sin datos de diagnóstico —")
else:
    st.info("Procesa archivos para habilitar el diagnóstico TX_CURR.")

# ============================
# ---------- DESCARGA --------
# ============================
full_dict = {
    "Numerador > Denominador": st.session_state.df_num,
    "Denominador > TX_CURR": st.session_state.df_txpv,
    "CD4 vacío positivo": st.session_state.df_cd4,
    "Fecha TARV < Diagnóstico": st.session_state.df_tarv,
    "Formato fecha diagnóstico": st.session_state.df_fdiag,
    "TX_CURR ≠ Dispensación_TARV": st.session_state.df_currq,
}

rows_metrics_global = [
    {"Indicador": DISPLAY_NAMES[k], "Errores": v["errors"], "Chequeos": v["checks"], "% Error": _pct(v["errors"], v["checks"])}
    for k, v in st.session_state.metrics_global.items()
]
if rows_metrics_global:
    df_metricas_global_all = pd.DataFrame(rows_metrics_global).sort_values("% Error", ascending=False)
else:
    df_metricas_global_all = pd.DataFrame(columns=["Indicador","Errores","Chequeos","% Error"])

rows_all = []
for (pais, depto, sitio, mes_rep, ind), v in st.session_state.metrics_by_pds.items():
    rows_all.append({
        "País": pais, "Departamento": depto, "Sitio": sitio, "Mes de reporte": mes_rep,
        "Indicador": DISPLAY_NAMES[ind], "Errores": v["errors"], "Chequeos": v["checks"],
        "% Error": _pct(v["errors"], v["checks"])
    })
df_metricas_por_mes_all = pd.DataFrame(rows_all)
if not df_metricas_por_mes_all.empty:
    df_metricas_por_mes_all = df_metricas_por_mes_all[["País","Departamento","Sitio","Mes de reporte","Indicador","Errores","Chequeos","% Error"]]
    df_metricas_por_mes_all = df_metricas_por_mes_all.sort_values(["País","Departamento","Sitio","Indicador"])

bytes_excel_full = exportar_excel_resultados(full_dict, df_metricas_global_all, df_metricas_por_mes_all)

filt_dict = {
    "Numerador > Denominador": df_num_f,
    "Denominador > TX_CURR": df_txpv_f,
    "CD4 vacío positivo": df_cd4_f,
    "Fecha TARV < Diagnóstico": df_tarv_f,
    "Formato fecha diagnóstico": df_fdiag_f,
    "TX_CURR ≠ Dispensación_TARV": df_currq_f,
}
bytes_excel_filt = exportar_excel_resultados(filt_dict, df_metricas_global_sel, df_metricas_por_mes_sel)

fecha_str = datetime.now().strftime("%Y%m%d_%H%M")
cdl1, cdl2 = st.columns(2)
with cdl1:
    st.download_button("⬇️ Descargar Excel (COMPLETO)", data=bytes_excel_full,
        file_name=f"VALIDACIONES_MAESTRO_VIH_COMPLETO_{fecha_str}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
with cdl2:
    st.download_button("⬇️ Descargar Excel (FILTRADO)", data=bytes_excel_filt,
        file_name=f"VALIDACIONES_MAESTRO_VIH_FILTRADO_{fecha_str}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)





