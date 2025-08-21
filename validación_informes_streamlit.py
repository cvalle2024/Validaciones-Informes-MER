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

# ============================
# --------- CONFIG UI --------
# ============================
LOGO_PATH = Path(__file__).parent / "logo.png"
logo_img = Image.open(LOGO_PATH)
st.set_page_config(page_title="Validaciones Maestro VIH", page_icon=logo_img, layout="wide")

c_logo, c_title = st.columns([1, 9])
with c_logo:
    st.image(logo_img, width=90)
with c_title:
    st.title("✅ Script de validación de indicadores MER (VIHCA)")
    st.caption("TX_PVLS / TX_CURR / HTS_TST • Reglas por Sexo, Población, Rango de edad y campos clave")

with st.expander("ℹ️ Cómo usar", expanded=False):
    st.markdown(
        """
1) **Carga** uno o varios `.xlsx` o un `.zip` con subcarpetas.  
2) Pulsa **Procesar**.  
3) Usa los **segmentadores** (País / Depto / Sitio).  
4) **Descarga** el Excel (completo o filtrado).  
        """
    )

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
def _ss_default(key, val):
    if key not in st.session_state: st.session_state[key] = val

for key, val in {
    "processed": False,
    "df_num": pd.DataFrame(),      # Numerador > Denominador
    "df_txpv": pd.DataFrame(),     # Denominador > TX_CURR
    "df_cd4": pd.DataFrame(),      # CD4 vacío positivo
    "df_tarv": pd.DataFrame(),     # TARV < Diagnóstico
    "df_fdiag": pd.DataFrame(),    # Formato fecha diagnóstico
    "df_currq": pd.DataFrame(),    # TX_CURR ≠ Dispensación_TARV
    "df_iddup": pd.DataFrame(),    # ID duplicado (HTS_TST)
    "metrics_global": defaultdict(lambda: {"errors": 0, "checks": 0}),
    "metrics_by_pds": defaultdict(lambda: {"errors": 0, "checks": 0}),
}.items():
    _ss_default(key, val)

_ss_default("sel_pais", "Todos")
_ss_default("sel_depto", "Todos")
_ss_default("sel_sitio", "Todos")

# ============================
# ----- CONSTANTES / HELPERS -
# ============================
IND_NUM_GT_DEN      = "num_gt_den"
IND_DEN_GT_CURR     = "den_gt_curr"
IND_CD4_MISSING     = "cd4_missing"
IND_TARV_LT_DIAG    = "tarv_lt_diag"
IND_DIAG_BAD_FMT    = "diag_bad_format"
IND_CURR_Q1Q2_DIFF  = "curr_q1q2_diff"    # TX_CURR ≠ Dispensación_TARV
IND_ID_DUPLICADO    = "id_duplicado"      # ID (expediente) duplicado

DISPLAY_NAMES = {
    IND_NUM_GT_DEN:      "Numerador > Denominador",
    IND_DEN_GT_CURR:     "Denominador > TX_CURR",
    IND_CD4_MISSING:     "CD4 vacío positivo",
    IND_TARV_LT_DIAG:    "Fecha TARV < Diagnóstico",
    IND_DIAG_BAD_FMT:    "Formato fecha diagnóstico",
    IND_CURR_Q1Q2_DIFF:  "TX_CURR ≠ Dispensación_TARV",
    IND_ID_DUPLICADO:    "ID (expediente) duplicado",
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
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return ""
    try:
        dt = pd.to_datetime(v, dayfirst=True, errors="coerce")
        if pd.notna(dt):
            return f"{SPAN_ABBR[dt.month-1]} {dt.year}"
    except Exception:
        pass
    return str(v).strip()

def inferir_pais_mes(path_rel: str, default_pais: str, default_mes: str):
    ruta = path_rel.replace("\\", "/")
    partes = [p for p in ruta.split("/") if p.strip().lower() not in RUIDO_DIRS]
    if partes and partes[-1].lower().endswith(".xlsx"): partes = partes[:-1]
    pais = partes[-2].strip() if len(partes) >= 2 else default_pais
    # Evita meses como país
    if any(tok in MESES for tok in re.split(r"[_\-\s/\.]+", _norm(pais))):
        pais = default_pais
    pais = pais or default_pais
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

def show_df_or_note(df, note="— Sin filas para mostrar —", height=300):
    if df is None or (isinstance(df, pd.DataFrame) and df.empty):
        st.caption(note); return False
    st.dataframe(df, use_container_width=True, height=height); return True

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

    # Fecha/Mes de reporte (prioridad)
    col_fecha_rep = buscar_columna_multi(df_data.columns, "fecha", "reporte")
    col_mesrep    = buscar_columna_multi(df_data.columns, "mes", "reporte")
    def _ctx(row):
        p = str(row.get(col_pais)) if col_pais else pais_inferido
        d = str(row.get(col_depto)) if col_depto else ""
        s = str(row.get(col_sitio)) if col_sitio else ""
        raw_mes = row.get(col_fecha_rep) if col_fecha_rep else (row.get(col_mesrep) if col_mesrep else None)
        m = month_label_from_value(raw_mes) or month_label_from_value(mes_inferido)
        p = p if str(p).strip() else pais_inferido
        d = d if str(d).strip() else ""
        s = s if str(s).strip() else ""
        m = m if str(m).strip() else month_label_from_value(mes_inferido)
        return p, d, s, m

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

# ===== TX_CURR vs Dispensación_TARV (por Sexo/Edad) =====
def procesar_tx_curr_cuadros(
    xl: pd.ExcelFile, pais_inferido: str, mes_inferido: str,
    nombre_archivo: str, errores_currq
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
                if not s: continue
                for toks in alternatives:
                    if all(tok in s for tok in toks):
                        pos.append((r, c)); break
        return pos

    pos_tx = _find_label_positions([["tx", "curr"]])
    pos_et = _find_label_positions([
        ["dispens","tar"], ["dispensacion","tar"], ["dispensación","tar"],
        ["entrega","tar"], ["entrega","tavr"]
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
    if hdr_tx is None or hdr_et is None: return

    def _extract_table_totals(header_row: int, stop_row: Optional[int]):
        cols = df_raw.iloc[header_row].fillna("").astype(str).tolist()
        cols_norm = [_norm(x) for x in cols]
        try:
            col_sexo = next(i for i, cn in enumerate(cols_norm) if "sexo" in cn)
        except StopIteration:
            return {}, {}, None

        edad_idx, edad_key, edad_map = [], [], {}
        for j in range(col_sexo + 1, ncols):
            lab  = cols[j]; labn = _norm(lab)
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
            if any("sexo" in x for x in rown): break
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

    if hdr_tx < hdr_et:
        totals_tx, edades_tx, _ = _extract_table_totals(hdr_tx, hdr_et)
        totals_et, edades_et, _ = _extract_table_totals(hdr_et, None)
    else:
        totals_et, edades_et, _ = _extract_table_totals(hdr_et, hdr_tx)
        totals_tx, edades_tx, _ = _extract_table_totals(hdr_tx, None)

    # Contexto
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
    col_mesrep_i = (_find_col_idx(cols_hdrn, "fecha") if _find_col_idx(cols_hdrn, "reporte") is not None else None)
    if col_mesrep_i is None:
        col_mesrep_i = _find_col_idx(cols_hdrn, "mes")

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
        raw_mes = row_vals[col_mesrep_i] if col_mesrep_i is not None else None
        m = month_label_from_value(raw_mes) or month_label_from_value(mes_inferido)
        return p, d, s, m

    fila_ctx_vals = df_raw.iloc[hdr_tx + 1].fillna("").astype(str).tolist() if (hdr_tx + 1) < nrows else []
    pais_row, depto_row, sitio_row, mes_rep = _ctx_from_rowvals(fila_ctx_vals)

    all_keys = set(totals_tx.keys()) | set(totals_et.keys())
    for (sexo, edad_key) in sorted(all_keys):
        v_tx = int(totals_tx.get((sexo, edad_key), 0))
        v_et = int(totals_et.get((sexo, edad_key), 0))
        etiqueta_edad = (edades_tx.get(edad_key) or edades_et.get(edad_key) or edad_key)
        _add_metric(IND_CURR_Q1Q2_DIFF, pais_row, mes_rep, depto_row, sitio_row, checks_add=1)
        if v_tx != v_et:
            _add_metric(IND_CURR_Q1Q2_DIFF, pais_row, mes_rep, depto_row, sitio_row, errors_add=1)
            errores_currq.append({
                "País": pais_row, "Departamento": depto_row, "Sitio": sitio_row, "Mes de reporte": mes_rep,
                "Archivo": nombre_archivo, "Sexo": sexo, "Rango de edad": etiqueta_edad,
                "TX_CURR": v_tx, "Dispensación_TARV": v_et,
                "Diferencia (TX_CURR - Disp_TARV)": v_tx - v_et,
                "Disp_TARV > TX_CURR": "Sí" if v_et > v_tx else "No",
            })

# ===== HTS_TST + ID Duplicado =====
def procesar_hts_tst(
    xl: pd.ExcelFile, pais_inferido: str, mes_inferido: str, nombre_archivo: str,
    errores_cd4, errores_fecha_tarv, errores_formato_fecha_diag, errores_iddup
):
    if "HTS_TST" not in xl.sheet_names:
        return

    df_raw = xl.parse("HTS_TST", header=None)
    idx_hts = encontrar_fila_encabezado(df_raw, ["Resultado", "CD4"])
    if idx_hts is None:
        return

    df_data, _ = normalizar_tabla_por_encabezado(df_raw, idx_hts)
    df_data.columns = _dedupe_columns(df_data.columns)
    df_data = _rename_standard_columns(df_data)

    col_resultado = _first_col(df_data, "resultado")
    col_cd4       = _first_col(df_data, "cd4")
    col_tarv      = _first_col(df_data, "inicio", "tar")
    col_diag      = _first_col(df_data, "fecha", "diagn")
    col_sitio     = _first_col(df_data, "servicio", "salud") or _first_col(df_data, "sitio") or _first_col(df_data, "clinica")
    col_pais      = _first_col(df_data, "pais")
    col_depto     = _first_col(df_data, "departamento") or _first_col(df_data, "depto") or _first_col(df_data, "provincia")
    col_id = (_first_col(df_data, "id", "expediente") or
              _first_col(df_data, "numero", "expediente") or
              _first_col(df_data, "número", "expediente") or
              _first_col(df_data, "id"))

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

        mes_rep    = month_label_from_value(fecha_diag) or month_label_from_value(mes_inferido)

        pais_row   = str(pais_row).strip() or pais_inferido
        depto_row  = str(depto_row).strip()
        sitio_row  = str(sitio).strip()
        mes_rep    = str(mes_rep).strip() or month_label_from_value(mes_inferido)

        # CHECK para duplicados (para %)
        if col_id is not None:
            id_val = _coerce_scalar(row.get(col_id))
            if pd.notna(id_val) and str(id_val).strip():
                _add_metric(IND_ID_DUPLICADO, pais_row, mes_rep, depto_row, sitio_row, checks_add=1)

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

        # Formato de Fecha Diagnóstico (si trae '/')
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

    # Duplicados ID
    if col_id:
        try:
            col_id_idx = list(df_data.columns).index(col_id)
        except ValueError:
            col_id_idx = None

        ids_raw = df_data[col_id].astype(str).str.strip()
        mask_non_empty = ids_raw.replace({"nan": "", "NaN": ""}).astype(bool)
        vc = ids_raw[mask_non_empty].value_counts()
        duplicados = vc[vc > 1]

        if not duplicados.empty:
            for id_val, count in duplicados.items():
                idxs = df_data.index[ids_raw == id_val].tolist()
                r0 = df_data.loc[idxs[0]]

                pais_row  = str(_coerce_scalar(r0.get(col_pais)))  if col_pais  else pais_inferido
                depto_row = str(_coerce_scalar(r0.get(col_depto))) if col_depto else ""
                sitio_row = str(_coerce_scalar(r0.get(col_sitio)))  if col_sitio  else ""
                mes_rep   = month_label_from_value(_coerce_scalar(r0.get(col_diag))) or month_label_from_value(mes_inferido)

                pais_row  = pais_row.strip() or pais_inferido
                depto_row = depto_row.strip()
                sitio_row = sitio_row.strip()
                mes_rep   = mes_rep.strip() or month_label_from_value(mes_inferido)

                filas_excel = [int(idx_hts + 3 + i) for i in idxs]
                col_letter = get_column_letter(col_id_idx + 1) if col_id_idx is not None else col_id

                _add_metric(IND_ID_DUPLICADO, pais_row, mes_rep, depto_row, sitio_row, errors_add=(int(count) - 1))

                errores_iddup.append({
                    "País": pais_row, "Departamento": depto_row, "Sitio": sitio_row, "Mes de reporte": mes_rep,
                    "Archivo": nombre_archivo, "ID expediente": str(id_val), "Ocurrencias": int(count),
                    "Filas Excel": ", ".join(map(str, filas_excel)), "Columna Excel": col_letter
                })

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
    errores_iddup = []

    st.session_state.metrics_global = defaultdict(lambda: {"errors": 0, "checks": 0})
    st.session_state.metrics_by_pds = defaultdict(lambda: {"errors": 0, "checks": 0})

    progreso = st.progress(0.0, text="Procesando archivos…"); total = len(entradas)
    for idx, (nombre_archivo, data_bytes, ruta_rel) in enumerate(entradas, start=1):
        try:
            pais_inf, mes_inf = inferir_pais_mes(ruta_rel.replace("\\", "/"), default_pais, default_mes)
            xl = leer_excel_desde_bytes(nombre_archivo, data_bytes)
            procesar_tx_pvls_y_curr(xl, pais_inf, mes_inf, nombre_archivo, errores_numerador, errores_txpvls)
            procesar_hts_tst(xl, pais_inf, mes_inf, nombre_archivo, errores_cd4, errores_fecha_tarv, errores_formato_fecha_diag, errores_iddup)
            procesar_tx_curr_cuadros(xl, pais_inf, mes_inf, nombre_archivo, errores_currq)
        except Exception as e:
            st.warning(f"⚠️ Error procesando {nombre_archivo}: {e}")
        progreso.progress(idx/total, text=f"Procesando {idx} de {total}…")

    st.session_state.df_num   = pd.DataFrame(errores_numerador)
    st.session_state.df_txpv  = pd.DataFrame(errores_txpvls)
    st.session_state.df_cd4   = pd.DataFrame(errores_cd4)
    st.session_state.df_tarv  = pd.DataFrame(errores_fecha_tarv)
    st.session_state.df_fdiag = pd.DataFrame(errores_formato_fecha_diag)
    st.session_state.df_currq = pd.DataFrame(errores_currq)
    st.session_state.df_iddup = pd.DataFrame(errores_iddup)
    st.session_state.processed = True
    st.success("Procesamiento completado. Ahora puedes filtrar al instante ✅")

# ============================
# ------- INTERFAZ (LIVE) ----
# ============================
if not st.session_state.processed:
    st.info("Carga tus archivos y pulsa **Procesar**.")
    st.stop()

# Asegurar columnas base
for dfname in ["df_num","df_txpv","df_cd4","df_tarv","df_fdiag","df_currq","df_iddup"]:
    df = st.session_state[dfname]
    if not isinstance(df, pd.DataFrame):
        st.session_state[dfname] = pd.DataFrame()
        continue
    if not df.empty:
        for col in ["País","Departamento","Sitio","Mes de reporte"]:
            if col not in df.columns:
                st.session_state[dfname][col] = ""

# Universo para segmentadores a partir de TODOS los errores
df_all = pd.concat(
    [df for df in [
        st.session_state.df_num, st.session_state.df_txpv, st.session_state.df_cd4,
        st.session_state.df_tarv, st.session_state.df_fdiag, st.session_state.df_currq,
        st.session_state.df_iddup
    ] if isinstance(df, pd.DataFrame) and not df.empty],
    ignore_index=True
) if any([isinstance(st.session_state[k], pd.DataFrame) and not st.session_state[k].empty
          for k in ["df_num","df_txpv","df_cd4","df_tarv","df_fdiag","df_currq","df_iddup"]]) \
  else pd.DataFrame(columns=["País","Departamento","Sitio","Mes de reporte"])

# Normaliza valores para evitar mezcla (espacios, mayúsculas, etc.)
for c in ["País","Departamento","Sitio","Mes de reporte"]:
    if c in df_all.columns:
        df_all[c] = df_all[c].astype(str).str.strip()

def _limpia_opts(vals):
    arr = []
    for v in vals:
        s = str(v).strip()
        if not s: 
            continue
        if s.lower() == "desconocido":
            continue
        if s.lower() == "nan":
            continue
        arr.append(s)
    return ["Todos"] + sorted(set(arr))

# === Callbacks de cascada ===
def _on_change_pais():
    st.session_state.sel_depto = "Todos"
    st.session_state.sel_sitio = "Todos"

def _on_change_depto():
    st.session_state.sel_sitio = "Todos"

# ─────────────────────────────────────────────────────────────
# 1) Segmentadores (caja)
# ─────────────────────────────────────────────────────────────
seg = st.container(border=True)
with seg:
    st.subheader("🧊 Segmentadores")

    # País
    pais_opts = _limpia_opts(df_all["País"].dropna().tolist()) if "País" in df_all.columns else ["Todos"]
    if st.session_state.sel_pais not in pais_opts:
        st.session_state.sel_pais = "Todos"
    st.selectbox("País", pais_opts, key="sel_pais", on_change=_on_change_pais)

    # Depto condicionado por país
    df_p = df_all if st.session_state.sel_pais == "Todos" else df_all[df_all["País"] == st.session_state.sel_pais]
    depto_opts = _limpia_opts(df_p["Departamento"].dropna().tolist()) if "Departamento" in df_p.columns else ["Todos"]
    if st.session_state.sel_depto not in depto_opts:
        st.session_state.sel_depto = "Todos"
    st.selectbox("Departamento", depto_opts, key="sel_depto", on_change=_on_change_depto)

    # Sitio condicionado por país y depto
    if st.session_state.sel_depto == "Todos":
        df_pd = df_p
    else:
        df_pd = df_p[df_p["Departamento"] == st.session_state.sel_depto]
    sitio_opts = _limpia_opts(df_pd["Sitio"].dropna().tolist()) if "Sitio" in df_pd.columns else ["Todos"]
    if st.session_state.sel_sitio not in sitio_opts:
        st.session_state.sel_sitio = "Todos"
    st.selectbox("Sitio", sitio_opts, key="sel_sitio")

def _aplicar_filtro(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty: return df
    m = pd.Series([True] * len(df))
    if st.session_state.sel_pais != "Todos":  m &= (df["País"].astype(str).str.strip() == st.session_state.sel_pais)
    if st.session_state.sel_depto != "Todos": m &= (df["Departamento"].astype(str).str.strip() == st.session_state.sel_depto)
    if st.session_state.sel_sitio != "Todos": m &= (df["Sitio"].astype(str).str.strip() == st.session_state.sel_sitio)
    return df[m].copy()

df_num_f   = _aplicar_filtro(st.session_state.df_num)
df_txpv_f  = _aplicar_filtro(st.session_state.df_txpv)
df_cd4_f   = _aplicar_filtro(st.session_state.df_cd4)
df_tarv_f  = _aplicar_filtro(st.session_state.df_tarv)
df_fdiag_f = _aplicar_filtro(st.session_state.df_fdiag)
df_currq_f = _aplicar_filtro(st.session_state.df_currq)
df_iddup_f = _aplicar_filtro(st.session_state.df_iddup)

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

df_metricas_global_sel, df_metricas_por_mes_sel = _build_metrics_df_from_selection(
    st.session_state.sel_pais, st.session_state.sel_depto, st.session_state.sel_sitio
)

# ─────────────────────────────────────────────────────────────
# 2) Resumen
# ─────────────────────────────────────────────────────────────
res = st.container(border=True)
with res:
    st.subheader("⚫ *Resumen de errores por indicador*")
    c1, c2, c3, c4, c5, c6, c7 = st.columns(7)
    c1.metric("Numerador > Denominador", len(df_num_f))
    c2.metric("Denominador > TX_CURR", len(df_txpv_f))
    c3.metric("CD4 vacío positivo", len(df_cd4_f))
    c4.metric("TARV < Diagnóstico", len(df_tarv_f))
    c5.metric("Fecha diag. mal formateada", len(df_fdiag_f))
    c6.metric("TX_CURR ≠ Dispensación_TARV", len(df_currq_f))
    c7.metric("ID (expediente) duplicado", len(df_iddup_f))

# ─────────────────────────────────────────────────────────────
# 3) Indicadores – % de error (selección)
# ─────────────────────────────────────────────────────────────
sel = st.container(border=True)
with sel:
    st.subheader("📊 *Porcentaje de errores por indicador*")
    cards = [IND_NUM_GT_DEN, IND_DEN_GT_CURR, IND_CD4_MISSING, IND_TARV_LT_DIAG, IND_DIAG_BAD_FMT, IND_CURR_Q1Q2_DIFF, IND_ID_DUPLICADO]
    cols = st.columns(len(cards))
    sel_map = {row["Indicador"]: row for _, row in df_metricas_global_sel.iterrows()} if not df_metricas_global_sel.empty else {}
    for col, key in zip(cols, cards):
        name = DISPLAY_NAMES[key]
        v = sel_map.get(name, {"Errores":0, "Chequeos":0, "% Error":0})
        col.metric(label=name, value=f"{v.get('% Error',0)}%", delta=f"{v.get('Errores',0)} / {v.get('Chequeos',0)} err/cheq")

# ─────────────────────────────────────────────────────────────
# 4) Detalle por indicador
# ─────────────────────────────────────────────────────────────
det = st.container(border=True)
with det:
    st.subheader("🔎 *Detalle por indicador*")
    tabs = st.tabs([
        "Numerador > Denominador",
        "Denominador > TX_CURR",
        "CD4 vacío positivo",
        "Fecha TARV < Diagnóstico",
        "Formato fecha diagnóstico",
        "TX_CURR ≠ Dispensación_TARV",
        "ID (expediente) duplicado",
    ])
    with tabs[0]: show_df_or_note(df_num_f,   "— Sin diferencias de Numerador > Denominador —", height=340)
    with tabs[1]: show_df_or_note(df_txpv_f,  "— Sin casos Denominador > TX_CURR —", height=340)
    with tabs[2]: show_df_or_note(df_cd4_f,   "— Sin positivos con CD4 vacío —", height=340)
    with tabs[3]: show_df_or_note(df_tarv_f,  "— Sin casos TARV < Diagnóstico —", height=340)
    with tabs[4]: show_df_or_note(df_fdiag_f, "— Sin problemas de formato de fecha —", height=340)
    with tabs[5]: show_df_or_note(df_currq_f, "— TX_CURR = Dispensación_TARV en la selección —", height=340)
    with tabs[6]: show_df_or_note(df_iddup_f, "— Sin IDs (expediente) duplicados —", height=340)

# ─────────────────────────────────────────────────────────────
# 5) Métricas de calidad (adaptadas al filtro)
# ─────────────────────────────────────────────────────────────
met = st.container(border=True)
with met:
    st.subheader("📈 *Resumen de porcentajes de error por indicador y desglose por país*")
    gc1, gc2 = st.columns([1.2, 2])
    with gc1:
        st.markdown("**Métricas – Selección actual**")
        show_df_or_note(df_metricas_global_sel, "— Sin métricas para la selección —", height=260)
    with gc2:
        st.markdown("**Desglose por Mes – Selección**")
        show_df_or_note(df_metricas_por_mes_sel, "— Sin desglose para la selección —", height=260)

# ============================
# ---------- DESCARGA --------
# ============================
def exportar_excel_resultados(errores_dict, df_metricas_global: pd.DataFrame, df_metricas_group: pd.DataFrame) -> bytes:
    config_resaltado = {
        "Numerador > Denominador": "Numerador",
        "Denominador > TX_CURR": "Denominador (PVLS)",
        "CD4 vacío positivo": "CD4 Basal",
        "Fecha TARV < Diagnóstico": "Fecha inicio TARV",
        "Formato fecha diagnóstico": "Fecha del diagnóstico de la prueba",
        "TX_CURR ≠ Dispensación_TARV": "Diferencia (TX_CURR - Disp_TARV)",
        "ID (expediente) duplicado": "ID expediente",
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

full_dict = {
    "Numerador > Denominador": st.session_state.df_num,
    "Denominador > TX_CURR": st.session_state.df_txpv,
    "CD4 vacío positivo": st.session_state.df_cd4,
    "Fecha TARV < Diagnóstico": st.session_state.df_tarv,
    "Formato fecha diagnóstico": st.session_state.df_fdiag,
    "TX_CURR ≠ Dispensación_TARV": st.session_state.df_currq,
    "ID (expediente) duplicado": st.session_state.df_iddup,
}

rows_metrics_global = [
    {"Indicador": DISPLAY_NAMES[k], "Errores": v["errors"], "Chequeos": v["checks"], "% Error": _pct(v["errors"], v["checks"])}
    for k, v in st.session_state.metrics_global.items()
]
df_metricas_global_all = (
    pd.DataFrame(rows_metrics_global).sort_values("% Error", ascending=False)
    if rows_metrics_global else pd.DataFrame(columns=["Indicador","Errores","Chequeos","% Error"])
)

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
    "ID (expediente) duplicado": df_iddup_f,
}
bytes_excel_filt = exportar_excel_resultados(filt_dict, df_metricas_global_sel, df_metricas_por_mes_sel)

fecha_str = datetime.now().strftime("%Y%m%d_%H%M")
dl = st.container(border=True)
with dl:
    cdl1, cdl2 = st.columns(2)
    with cdl1:
        st.download_button("⬇️ Descargar Excel (COMPLETO)", data=bytes_excel_full,
            file_name=f"VALIDACIONES_DE_INDICADORES_MER_{pais} _ {fecha_str}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
    with cdl2:
        st.download_button("⬇️ Descargar Excel (FILTRADO)", data=bytes_excel_filt,
            file_name=f"VALIDACIONES_DE_INDICADORES_MER_{pais} _ {fecha_str}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
