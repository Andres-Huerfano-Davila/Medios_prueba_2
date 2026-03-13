import io
import math
import re
import traceback
import unicodedata
import zipfile
from pathlib import Path
from typing import Dict, List, Optional, Sequence, Tuple

import numpy as np
import pandas as pd
import streamlit as st

st.set_page_config(page_title="Medios Magnéticos Nómina JMC", layout="wide")

st.markdown(
    """
    <style>
    .block-container {padding-top: 1rem; padding-bottom: 1.5rem;}
    .hero {background: linear-gradient(135deg,#fff4eb 0%,#fffaf6 100%); border:1px solid #f3c9a5; border-radius:16px; padding:18px 20px; margin-bottom:16px;}
    .hero h1 {margin:0; color:#a85400; font-size:1.55rem;}
    .hero p {margin:8px 0 0 0; color:#714221;}
    .credit {margin-top:8px; display:inline-block; background:#fff0e1; color:#a85400; border:1px solid #f3c9a5; border-radius:999px; padding:4px 10px; font-size:.9rem; font-weight:600;}
    .footer-credit {text-align:center; color:#a06a3d; font-weight:600; border:1px solid #f0d6bf; border-radius:12px; padding:10px; background:#fffaf6; margin-top:18px;}
    </style>
    """,
    unsafe_allow_html=True,
)

st.markdown(
    """
    <div class="hero">
        <h1>Validación de medios magnéticos</h1>
        <p>CC-nóminas = pagado | PCP0 = contabilizado</p>
        <div class="credit">Creado por Andrés Huérfano Dávila - Nómina JMC</div>
    </div>
    """,
    unsafe_allow_html=True,
)

MAX_EXCEL_ROWS = 1_048_576
SUPPORTED = {".txt", ".csv", ".xls", ".xlsx", ".xlsb", ".ods"}

EXACT_ALIASES = {
    "Número de personal": ["número de personal", "numero de personal", "nº pers", "n° pers", "n pers", "n pers."],
    "Nombre del empleado o candidato": ["nombre del empleado o candidato"],
    "CC-nómina": ["cc-nómina", "cc-nomina", "cc nómina", "cc nomina", "cc-n.", "cc n"],
    "Texto expl.CC-nómina": ["texto expl.cc-nómina", "texto expl.cc-nomina", "texto expl.cc nómina", "texto expl.cc nomina", "texto expl.cc-nomina"],
    "Importe": ["importe"],
    "Fecha de pago": ["fecha de pago"],
    "Período para nómina": ["período para nómina", "periodo para nomina"],
    "Período En cálc.nóm.": ["período en cálc.nóm.", "período en cálc nóm", "periodo en calc nom", "periodo en cálculo nómina"],
    "Tipo": ["tipo"],
}

CONTAINS_ALIASES = {
    "Número de personal": ["numero de personal", "n pers", "nº pers", "sap", "pernr"],
    "Nombre del empleado o candidato": ["nombre del empleado", "nombre empleado", "candidato"],
    "CC-nómina": ["cc-n", "cc nomina", "cc nómina", "wagetype", "tipo de salario"],
    "Texto expl.CC-nómina": ["texto expl", "texto explicativo", "descripcion cc"],
    "Importe": ["importe", "valor", "monto"],
}

CC_KEEP = [
    "Número de personal", "Nombre del empleado o candidato", "CC-nómina", "Texto expl.CC-nómina",
    "Importe", "Fecha de pago", "Fecha de pago_display", "Fecha de pago_año", "Período para nómina",
    "__archivo_origen__", "__hoja_origen__"
]
PCP0_KEEP = [
    "Número de personal", "CC-nómina", "Texto expl.CC-nómina", "Importe",
    "Período para nómina", "Período En cálc.nóm.", "__archivo_origen__", "__hoja_origen__"
]


def normalize_text(value: object) -> str:
    if value is None:
        return ""
    text = str(value).strip().lower()
    text = unicodedata.normalize("NFKD", text)
    text = "".join(ch for ch in text if not unicodedata.combining(ch))
    text = re.sub(r"[\n\r\t]+", " ", text)
    text = re.sub(r"[^a-z0-9]+", " ", text)
    return re.sub(r"\s+", " ", text).strip()


def normalize_id(value: object) -> Optional[str]:
    if pd.isna(value):
        return None
    if isinstance(value, (int, np.integer)):
        return str(int(value))
    if isinstance(value, (float, np.floating)):
        if np.isnan(value):
            return None
        return str(int(value)) if float(value).is_integer() else str(value).strip()
    text = str(value).strip()
    text = re.sub(r"\.0+$", "", text)
    return text or None


def parse_number(value: object) -> float:
    if pd.isna(value):
        return 0.0
    if isinstance(value, (int, float, np.integer, np.floating)):
        return 0.0 if pd.isna(value) else float(value)
    text = str(value).strip()
    if not text:
        return 0.0
    text = text.replace("\xa0", "").replace(" ", "")
    text = re.sub(r"[^0-9,\.\-()]", "", text)
    if text.startswith("(") and text.endswith(")"):
        text = "-" + text[1:-1]
    if "," in text and "." in text:
        text = text.replace(".", "").replace(",", ".") if text.rfind(",") > text.rfind(".") else text.replace(",", "")
    elif text.count(",") == 1 and "." not in text:
        text = text.replace(",", ".")
    elif text.count(".") > 1 and "," not in text:
        text = text.replace(".", "")
    try:
        return float(text)
    except Exception:
        return 0.0


def normalize_tipo(value: object) -> Optional[str]:
    text = normalize_text(value)
    if not text:
        return None
    if text.startswith("salaria"):
        return "Salarial"
    if text.startswith("benef") or "beneficio" in text:
        return "Beneficio"
    if text in {"no aplica", "noaplica", "na", "n a", "no"}:
        return "No aplica"
    return None


def robust_parse_dates(series: pd.Series) -> pd.Series:
    if pd.api.types.is_datetime64_any_dtype(series):
        return pd.to_datetime(series, errors="coerce")
    non_null = series.dropna()
    if non_null.empty:
        return pd.to_datetime(series, errors="coerce")
    if pd.api.types.is_numeric_dtype(non_null):
        numeric = pd.to_numeric(series, errors="coerce")
        valid = numeric.dropna()
        if not valid.empty and valid.between(20000, 70000).mean() > 0.9:
            return pd.to_datetime(numeric, unit="D", origin="1899-12-30", errors="coerce")
    text = series.astype(str).str.strip().replace({"": np.nan, "nan": np.nan, "NaT": np.nan, "None": np.nan})
    result = pd.Series(pd.NaT, index=series.index, dtype="datetime64[ns]")
    iso_mask = text.str.match(r"^\d{4}-\d{2}-\d{2}(?:[ T]\d{2}:\d{2}:\d{2})?$", na=False)
    if iso_mask.any():
        iso_text = text[iso_mask].str.replace("T", " ", regex=False)
        parsed = pd.to_datetime(iso_text, format="%Y-%m-%d %H:%M:%S", errors="coerce")
        rem = parsed.isna()
        if rem.any():
            parsed.loc[rem] = pd.to_datetime(iso_text.loc[rem], format="%Y-%m-%d", errors="coerce")
        result.loc[iso_mask] = parsed
    dmy_mask = text.str.match(r"^\d{1,2}/\d{1,2}/\d{4}$", na=False)
    if dmy_mask.any():
        result.loc[dmy_mask] = pd.to_datetime(text[dmy_mask], format="%d/%m/%Y", errors="coerce")
    dot_mask = text.str.match(r"^\d{1,2}\.\d{1,2}\.\d{4}$", na=False)
    if dot_mask.any():
        result.loc[dot_mask] = pd.to_datetime(text[dot_mask], format="%d.%m.%Y", errors="coerce")
    rem = result.isna() & text.notna()
    if rem.any():
        result.loc[rem] = pd.to_datetime(text[rem], errors="coerce")
    return result


def preferred_column_match(df: pd.DataFrame, canonical_name: str) -> Optional[str]:
    normalized_cols = {col: normalize_text(col) for col in df.columns}
    exact_aliases = [normalize_text(canonical_name)] + [normalize_text(x) for x in EXACT_ALIASES.get(canonical_name, [])]
    for col, ncol in normalized_cols.items():
        if ncol in exact_aliases:
            return col
    if canonical_name == "Período para nómina":
        for col, ncol in normalized_cols.items():
            if ncol == "periodo para nomina":
                return col
        return None
    if canonical_name == "Período En cálc.nóm.":
        for col, ncol in normalized_cols.items():
            if ncol in {"periodo en calc nom", "periodo en calculo nomina", "periodo en calculo nom"}:
                return col
        return None
    contains_aliases = [normalize_text(x) for x in CONTAINS_ALIASES.get(canonical_name, [])]
    for col, ncol in normalized_cols.items():
        if any(alias in ncol for alias in contains_aliases):
            return col
    return None


def add_canonical_columns(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    for canonical in set(list(EXACT_ALIASES.keys()) + list(CONTAINS_ALIASES.keys())):
        source = preferred_column_match(df, canonical)
        if source and canonical not in df.columns:
            df[canonical] = df[source]
    if "Número de personal" in df.columns:
        df["Número de personal"] = df["Número de personal"].map(normalize_id)
    if "Nombre del empleado o candidato" in df.columns:
        df["Nombre del empleado o candidato"] = df["Nombre del empleado o candidato"].astype(str).replace("nan", "").str.strip()
    if "CC-nómina" in df.columns:
        df["CC-nómina"] = df["CC-nómina"].astype(str).replace("nan", "").str.strip()
    if "Texto expl.CC-nómina" in df.columns:
        df["Texto expl.CC-nómina"] = df["Texto expl.CC-nómina"].astype(str).replace("nan", "").str.strip()
    if "Importe" in df.columns:
        df["Importe"] = df["Importe"].map(parse_number)
    if "Período para nómina" in df.columns:
        df["Período para nómina"] = df["Período para nómina"].map(normalize_id)
    if "Período En cálc.nóm." in df.columns:
        df["Período En cálc.nóm."] = df["Período En cálc.nóm."].map(normalize_id)
    if "Fecha de pago" in df.columns:
        df["Fecha de pago"] = robust_parse_dates(df["Fecha de pago"])
        df["Fecha de pago_display"] = df["Fecha de pago"].dt.strftime("%d/%m/%Y")
        df["Fecha de pago_año"] = df["Fecha de pago"].dt.year.astype("Int64")
    return df


def read_text_file(data: bytes, filename: str) -> pd.DataFrame:
    encodings = ["utf-8", "utf-8-sig", "cp1252", "latin-1"]
    last_error = None
    for enc in encodings:
        try:
            text = data.decode(enc, errors="ignore")
            lines = [ln.rstrip() for ln in text.splitlines() if ln.strip()]
            pipe_lines = [ln for ln in lines if "|" in ln]
            if len(pipe_lines) >= 2:
                header = None
                rows = []
                for ln in pipe_lines:
                    parts = [p.strip() for p in ln.strip().strip("|").split("|")]
                    if len(parts) < 2:
                        continue
                    joined = " ".join(parts).lower()
                    if any(k in joined for k in ["importe", "cc", "per", "texto expl", "pers"]):
                        header = parts
                        continue
                    if header and len(parts) == len(header):
                        rows.append(parts)
                if header and rows:
                    return pd.DataFrame(rows, columns=header)
        except Exception as exc:
            last_error = exc
    for enc in encodings:
        for sep in [None, ";", "\t", "|", ","]:
            try:
                if sep is None:
                    df = pd.read_csv(io.BytesIO(data), sep=None, engine="python", encoding=enc, dtype=object)
                else:
                    df = pd.read_csv(io.BytesIO(data), sep=sep, encoding=enc, dtype=object)
                if df.shape[1] > 1:
                    return df
            except Exception as exc:
                last_error = exc
    for enc in encodings:
        try:
            df = pd.read_fwf(io.BytesIO(data), encoding=enc, dtype=object)
            if not df.empty:
                return df
        except Exception as exc:
            last_error = exc
    raise ValueError(f"No fue posible leer {filename}: {last_error}")


def read_excel_file(data: bytes, filename: str) -> pd.DataFrame:
    ext = Path(filename).suffix.lower()
    engine = None
    if ext == ".xlsb":
        engine = "pyxlsb"
    elif ext == ".xls":
        engine = "xlrd"
    elif ext == ".ods":
        engine = "odf"
    xls = pd.ExcelFile(io.BytesIO(data), engine=engine)
    frames = []
    for sheet in xls.sheet_names:
        try:
            temp = pd.read_excel(io.BytesIO(data), sheet_name=sheet, engine=engine, dtype=object)
            if temp is not None and not temp.empty:
                temp = temp.copy()
                temp["__archivo_origen__"] = filename
                temp["__hoja_origen__"] = sheet
                frames.append(temp)
        except Exception:
            continue
    if not frames:
        raise ValueError(f"No se pudieron leer hojas válidas en {filename}")
    return pd.concat(frames, ignore_index=True, sort=False)


def read_uploaded_table(uploaded_file, keep_columns: Sequence[str]) -> pd.DataFrame:
    filename = uploaded_file.name
    data = uploaded_file.getvalue()
    ext = Path(filename).suffix.lower()
    if ext not in SUPPORTED:
        raise ValueError(f"Formato no soportado: {filename}")
    if ext in {".xlsx", ".xls", ".xlsb", ".ods"}:
        df = read_excel_file(data, filename)
    else:
        df = read_text_file(data, filename)
        df["__archivo_origen__"] = filename
        df["__hoja_origen__"] = "TXT/CSV"
    df = add_canonical_columns(df)
    existing = [c for c in keep_columns if c in df.columns]
    return df[existing].copy()


def combine_uploaded_files(uploaded_files, label: str, keep_columns: Sequence[str]) -> Tuple[pd.DataFrame, pd.DataFrame]:
    frames, logs = [], []
    progress = st.progress(0, text=f"Leyendo archivos de {label}...")
    total = max(len(uploaded_files), 1)
    for idx, file in enumerate(uploaded_files, start=1):
        try:
            df = read_uploaded_table(file, keep_columns)
            frames.append(df)
            logs.append({"archivo": file.name, "estado": "OK", "registros": int(len(df)), "detalle": "Archivo leído correctamente"})
        except Exception as exc:
            logs.append({"archivo": file.name, "estado": "ERROR", "registros": 0, "detalle": str(exc)})
        progress.progress(idx / total, text=f"Leyendo archivos de {label}... {idx}/{total}")
    progress.empty()
    if not frames:
        raise ValueError(f"No se pudo leer ningún archivo de {label}")
    return pd.concat(frames, ignore_index=True, sort=False), pd.DataFrame(logs)


def require_columns(df: pd.DataFrame, required: Sequence[str], label: str) -> None:
    missing = [c for c in required if c not in df.columns]
    if missing:
        raise ValueError(f"En {label} faltan columnas obligatorias: {', '.join(missing)}")


def prepare_concepts_df(uploaded_file) -> pd.DataFrame:
    df = read_uploaded_table(uploaded_file, ["CC-nómina", "Texto expl.CC-nómina", "Tipo"])
    require_columns(df, ["CC-nómina", "Texto expl.CC-nómina", "Tipo"], "conceptos")
    concepts = df[["CC-nómina", "Texto expl.CC-nómina", "Tipo"]].copy()
    concepts["CC-nómina"] = concepts["CC-nómina"].astype(str).replace("nan", "").str.strip()
    concepts["Texto expl.CC-nómina"] = concepts["Texto expl.CC-nómina"].astype(str).replace("nan", "").str.strip()
    concepts["Tipo"] = concepts["Tipo"].map(normalize_tipo)
    concepts = concepts[concepts["CC-nómina"].ne("")].drop_duplicates(subset=["CC-nómina"], keep="first").reset_index(drop=True)
    invalid = concepts[concepts["Tipo"].isna()]
    if not invalid.empty:
        raise ValueError("Hay conceptos con Tipo inválido. Usa solamente: Salarial, Beneficio o No aplica.")
    return concepts


def process_cc(df: pd.DataFrame, concepts: pd.DataFrame) -> Tuple[pd.DataFrame, pd.DataFrame]:
    require_columns(df, ["Número de personal", "Nombre del empleado o candidato", "CC-nómina", "Importe"], "CC-nóminas")
    work = df.copy()
    work["Importe"] = work["Importe"].map(parse_number)
    work["Nombre del empleado o candidato"] = work["Nombre del empleado o candidato"].fillna("").astype(str).str.strip()
    merged = work.merge(concepts, on="CC-nómina", how="left", suffixes=("", "_map"))
    if "Texto expl.CC-nómina_map" in merged.columns:
        merged["Texto expl.CC-nómina"] = merged["Texto expl.CC-nómina_map"].fillna(merged.get("Texto expl.CC-nómina", ""))
    merged["Tipo_final"] = merged["Tipo_map"] if "Tipo_map" in merged.columns else merged["Tipo"]
    merged = merged[merged["Tipo_final"].isin(["Salarial", "Beneficio"])].copy()
    merged["Salariales"] = np.where(merged["Tipo_final"].eq("Salarial"), merged["Importe"], 0.0)
    merged["Beneficios adicionales"] = np.where(merged["Tipo_final"].eq("Beneficio"), merged["Importe"], 0.0)
    resumen = merged.groupby(["Número de personal", "Nombre del empleado o candidato"], dropna=False, as_index=False)[["Salariales", "Beneficios adicionales"]].sum().sort_values(["Número de personal", "Nombre del empleado o candidato"]).reset_index(drop=True)
    resumen["Importe total"] = resumen["Salariales"] + resumen["Beneficios adicionales"]
    detalle = merged.groupby(["Número de personal", "Nombre del empleado o candidato", "CC-nómina", "Texto expl.CC-nómina"], dropna=False, as_index=False)[["Salariales", "Beneficios adicionales"]].sum().sort_values(["Número de personal", "CC-nómina"]).reset_index(drop=True)
    detalle["Importe total"] = detalle["Salariales"] + detalle["Beneficios adicionales"]
    return resumen, detalle


def process_pcp0(df: pd.DataFrame, concepts: pd.DataFrame) -> Tuple[pd.DataFrame, pd.DataFrame]:
    require_columns(df, ["Número de personal", "CC-nómina", "Importe"], "PCP0")
    work = df.copy()
    work["Importe"] = work["Importe"].map(parse_number)
    merged = work.merge(concepts, on="CC-nómina", how="left", suffixes=("", "_map"))
    if "Texto expl.CC-nómina_map" in merged.columns:
        merged["Texto expl.CC-nómina"] = merged["Texto expl.CC-nómina_map"].fillna(merged.get("Texto expl.CC-nómina", ""))
    merged["Tipo_final"] = merged["Tipo_map"] if "Tipo_map" in merged.columns else merged["Tipo"]
    merged = merged[merged["Tipo_final"].isin(["Salarial", "Beneficio"])].copy()
    merged["Salariales"] = np.where(merged["Tipo_final"].eq("Salarial"), merged["Importe"], 0.0)
    merged["Beneficios adicionales"] = np.where(merged["Tipo_final"].eq("Beneficio"), merged["Importe"], 0.0)
    resumen = merged.groupby(["Número de personal"], dropna=False, as_index=False)[["Salariales", "Beneficios adicionales"]].sum().sort_values(["Número de personal"]).reset_index(drop=True)
    resumen["Importe total"] = resumen["Salariales"] + resumen["Beneficios adicionales"]
    detalle = merged.groupby(["Número de personal", "CC-nómina", "Texto expl.CC-nómina"], dropna=False, as_index=False)[["Salariales", "Beneficios adicionales"]].sum().sort_values(["Número de personal", "CC-nómina"]).reset_index(drop=True)
    detalle["Importe total"] = detalle["Salariales"] + detalle["Beneficios adicionales"]
    return resumen, detalle


def compare_cc_vs_pcp0(cc_summary: pd.DataFrame, cc_detail: pd.DataFrame, pcp0_summary: pd.DataFrame, pcp0_detail: pd.DataFrame) -> Tuple[pd.DataFrame, pd.DataFrame]:
    cc_sum = cc_summary[["Número de personal", "Nombre del empleado o candidato", "Importe total"]].copy().rename(columns={"Importe total": "Importe total CC-nóminas"})
    pcp0_sum = pcp0_summary[["Número de personal", "Importe total"]].copy().rename(columns={"Importe total": "Importe total Contabilización"})
    resumen = cc_sum.merge(pcp0_sum, on="Número de personal", how="outer")
    resumen["Nombre del empleado o candidato"] = resumen["Nombre del empleado o candidato"].fillna("")
    resumen["Importe total CC-nóminas"] = resumen["Importe total CC-nóminas"].fillna(0.0)
    resumen["Importe total Contabilización"] = resumen["Importe total Contabilización"].fillna(0.0)
    resumen["Diferencia"] = resumen["Importe total CC-nóminas"] - resumen["Importe total Contabilización"]
    resumen["Estado"] = np.select(
        [
            resumen["Diferencia"].abs().lt(0.005),
            (resumen["Importe total CC-nóminas"] > 0) & (resumen["Importe total Contabilización"] == 0),
            (resumen["Importe total CC-nóminas"] == 0) & (resumen["Importe total Contabilización"] > 0),
        ],
        ["OK", "Pagado no contabilizado", "Contabilizado sin pago"],
        default="Diferencia por revisar",
    )
    cc_det = cc_detail[["Número de personal", "CC-nómina", "Texto expl.CC-nómina", "Importe total"]].copy().rename(columns={"Importe total": "Importe total CC-nóminas"})
    pcp0_det = pcp0_detail[["Número de personal", "CC-nómina", "Texto expl.CC-nómina", "Importe total"]].copy().rename(columns={"Importe total": "Importe total Contabilización"})
    detalle = cc_det.merge(pcp0_det, on=["Número de personal", "CC-nómina"], how="outer", suffixes=("_cc", "_pcp0"))
    detalle["Texto expl.CC-nómina"] = detalle["Texto expl.CC-nómina_cc"].fillna(detalle["Texto expl.CC-nómina_pcp0"])
    detalle["Importe total CC-nóminas"] = detalle["Importe total CC-nóminas"].fillna(0.0)
    detalle["Importe total Contabilización"] = detalle["Importe total Contabilización"].fillna(0.0)
    detalle["Diferencia"] = detalle["Importe total CC-nóminas"] - detalle["Importe total Contabilización"]
    saps_diff = resumen.loc[resumen["Estado"] != "OK", "Número de personal"].tolist()
    detalle = detalle[detalle["Número de personal"].isin(saps_diff)].copy()
    detalle = detalle[detalle["Diferencia"].abs() >= 0.005].copy()
    detalle = detalle[["Número de personal", "CC-nómina", "Texto expl.CC-nómina", "Importe total CC-nóminas", "Importe total Contabilización", "Diferencia"]].sort_values(["Número de personal", "CC-nómina"]).reset_index(drop=True)
    resumen = resumen.sort_values(["Estado", "Número de personal"]).reset_index(drop=True)
    return resumen, detalle


def build_period_summary(cc_src: pd.DataFrame, pcp0_src: pd.DataFrame, concepts: pd.DataFrame) -> pd.DataFrame:
    cc_period = cc_src.merge(concepts[["CC-nómina", "Tipo"]], on="CC-nómina", how="left", suffixes=("", "_map"))
    cc_period["Tipo_final"] = cc_period["Tipo_map"] if "Tipo_map" in cc_period.columns else cc_period["Tipo"]
    cc_period = cc_period[cc_period["Tipo_final"].isin(["Salarial", "Beneficio"])].copy()
    cc_period["Importe"] = cc_period["Importe"].map(parse_number)
    cc_period["Periodo"] = cc_period["Período para nómina"].astype(str)
    cc_by_period = cc_period.groupby("Periodo", as_index=False)["Importe"].sum().rename(columns={"Importe": "Pagado CC"})

    pcp0_period = pcp0_src.merge(concepts[["CC-nómina", "Tipo"]], on="CC-nómina", how="left", suffixes=("", "_map"))
    pcp0_period["Tipo_final"] = pcp0_period["Tipo_map"] if "Tipo_map" in pcp0_period.columns else pcp0_period["Tipo"]
    pcp0_period = pcp0_period[pcp0_period["Tipo_final"].isin(["Salarial", "Beneficio"])].copy()
    pcp0_period["Importe"] = pcp0_period["Importe"].map(parse_number)
    pcp0_period["Periodo"] = np.where(
        pcp0_period.get("Período para nómina", pd.Series(index=pcp0_period.index, dtype=object)).notna(),
        pcp0_period.get("Período para nómina", pd.Series(index=pcp0_period.index, dtype=object)).astype(str),
        pcp0_period.get("Período En cálc.nóm.", pd.Series(index=pcp0_period.index, dtype=object)).astype(str),
    )
    pcp0_by_period = pcp0_period.groupby("Periodo", as_index=False)["Importe"].sum().rename(columns={"Importe": "Contabilizado PCP0"})

    resumen = cc_by_period.merge(pcp0_by_period, on="Periodo", how="outer").fillna(0.0)
    resumen["Diferencia"] = resumen["Pagado CC"] - resumen["Contabilizado PCP0"]
    return resumen.sort_values("Periodo").reset_index(drop=True)


def chunk_dataframe(df: pd.DataFrame, max_rows: int = MAX_EXCEL_ROWS - 1) -> List[pd.DataFrame]:
    if len(df) <= max_rows:
        return [df]
    parts = math.ceil(len(df) / max_rows)
    return [df.iloc[i * max_rows:(i + 1) * max_rows].copy() for i in range(parts)]


def write_sheet(writer: pd.ExcelWriter, sheet_name: str, df: pd.DataFrame) -> None:
    df.to_excel(writer, sheet_name=sheet_name, index=False)
    workbook = writer.book
    ws = writer.sheets[sheet_name]
    header_fmt = workbook.add_format({"bold": True, "bg_color": "#D9EAF7", "border": 1, "text_wrap": True, "valign": "top"})
    num_fmt = workbook.add_format({"num_format": "#,##0.00"})
    for col_idx, col in enumerate(df.columns):
        ws.write(0, col_idx, col, header_fmt)
        max_len = len(col) if df.empty else max(len(col), int(df[col].astype(str).str.len().max()))
        width = min(max_len + 2, 40)
        if col in {"Salariales", "Beneficios adicionales", "Importe total", "Importe", "Pagado CC", "Contabilizado PCP0", "Diferencia", "Importe total CC-nóminas", "Importe total Contabilización"}:
            ws.set_column(col_idx, col_idx, max(width, 16), num_fmt)
        else:
            ws.set_column(col_idx, col_idx, width)
    ws.freeze_panes(1, 0)
    ws.autofilter(0, 0, max(len(df), 1), len(df.columns) - 1)


def to_excel_bytes(named_sheets: Dict[str, pd.DataFrame]) -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        for base_name, df in named_sheets.items():
            parts = chunk_dataframe(df)
            if len(parts) == 1:
                write_sheet(writer, base_name[:31], parts[0])
            else:
                for i, part in enumerate(parts, start=1):
                    write_sheet(writer, f"{base_name}_{i}"[:31], part)
    output.seek(0)
    return output.getvalue()


def to_csv_bytes(df: pd.DataFrame) -> bytes:
    return df.to_csv(index=False, encoding="utf-8-sig").encode("utf-8-sig")


def build_zip(files: Dict[str, bytes]) -> bytes:
    output = io.BytesIO()
    with zipfile.ZipFile(output, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        for name, data in files.items():
            zf.writestr(name, data)
    output.seek(0)
    return output.getvalue()


def init_state() -> None:
    defaults = {
        "cc_df": None,
        "cc_log": pd.DataFrame(columns=["archivo", "estado", "registros", "detalle"]),
        "pcp0_df": None,
        "pcp0_log": pd.DataFrame(columns=["archivo", "estado", "registros", "detalle"]),
        "concepts_df": None,
        "results": None,
        "last_error": None,
    }
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v


init_state()

st.info("Usa este flujo: 1) Cargar CC-nóminas  2) Elegir año y períodos  3) Cargar conceptos  4) Cargar PCP0  5) Procesar validación.")

with st.expander("1) Cargar CC-nóminas", expanded=True):
    cc_files = st.file_uploader("Archivos de CC-nóminas", type=[x.lstrip('.') for x in SUPPORTED], accept_multiple_files=True, key="cc_files")
    if st.button("Leer CC-nóminas", type="primary"):
        try:
            if not cc_files:
                raise ValueError("Carga uno o varios archivos de CC-nóminas.")
            cc_df, cc_log = combine_uploaded_files(cc_files, "CC-nóminas", CC_KEEP)
            require_columns(cc_df, ["Número de personal", "Nombre del empleado o candidato", "CC-nómina", "Importe", "Fecha de pago", "Período para nómina"], "CC-nóminas")
            st.session_state["cc_df"] = cc_df
            st.session_state["cc_log"] = cc_log
            st.session_state["results"] = None
            st.success("CC-nóminas cargado correctamente.")
        except Exception as exc:
            st.session_state["last_error"] = traceback.format_exc()
            st.error(str(exc))
    if isinstance(st.session_state["cc_df"], pd.DataFrame):
        df = st.session_state["cc_df"]
        c1, c2, c3 = st.columns(3)
        c1.metric("Registros CC", f"{len(df):,}".replace(",", "."))
        c2.metric("SAP únicos", f"{df['Número de personal'].nunique():,}".replace(",", "."))
        c3.metric("Períodos", f"{df['Período para nómina'].dropna().nunique():,}".replace(",", "."))
        st.dataframe(st.session_state["cc_log"], width="stretch", height=180)
        prev = df[[c for c in ["__archivo_origen__", "__hoja_origen__", "Número de personal", "CC-nómina", "Importe", "Fecha de pago_display", "Período para nómina"] if c in df.columns]].head(100).copy()
        if "Fecha de pago_display" in prev.columns:
            prev = prev.rename(columns={"Fecha de pago_display": "Fecha de pago"})
        st.dataframe(prev, width="stretch", height=260)

cc_df = st.session_state["cc_df"]
selected_year = None
selected_periods: List[str] = []
if isinstance(cc_df, pd.DataFrame):
    with st.expander("2) Seleccionar año y períodos", expanded=True):
        years = sorted([int(x) for x in cc_df.get("Fecha de pago_año", pd.Series(dtype="Int64")).dropna().unique().tolist()])
        if years:
            selected_year = st.selectbox("Año de Fecha de pago", options=years, index=len(years)-1)
            cc_year = cc_df[cc_df["Fecha de pago_año"].eq(selected_year)].copy()
            periods = sorted([str(x) for x in cc_year["Período para nómina"].dropna().astype(str).unique().tolist() if str(x).strip()])
            selected_periods = st.multiselect("Período para nómina", options=periods, default=periods)
            st.caption(f"Registros CC en el año seleccionado: {len(cc_year):,}".replace(",", "."))
        else:
            st.warning("No se encontraron años válidos en Fecha de pago.")

with st.expander("3) Cargar conceptos", expanded=True):
    concepts_file = st.file_uploader("Archivo de conceptos", type=[x.lstrip('.') for x in SUPPORTED], accept_multiple_files=False, key="concepts_file")
    if st.button("Leer conceptos"):
        try:
            if concepts_file is None:
                raise ValueError("Carga el archivo de conceptos.")
            concepts_df = prepare_concepts_df(concepts_file)
            st.session_state["concepts_df"] = concepts_df
            st.session_state["results"] = None
            st.success("Conceptos cargados correctamente.")
        except Exception as exc:
            st.session_state["last_error"] = traceback.format_exc()
            st.error(str(exc))
    if isinstance(st.session_state["concepts_df"], pd.DataFrame):
        st.dataframe(st.session_state["concepts_df"].head(100), width="stretch", height=220)

with st.expander("4) Cargar PCP0", expanded=True):
    pcp0_files = st.file_uploader("Archivos de PCP0 / contabilización", type=[x.lstrip('.') for x in SUPPORTED], accept_multiple_files=True, key="pcp0_files")
    if st.button("Leer PCP0"):
        try:
            if not pcp0_files:
                raise ValueError("Carga uno o varios archivos de PCP0.")
            pcp0_df, pcp0_log = combine_uploaded_files(pcp0_files, "PCP0", PCP0_KEEP)
            require_columns(pcp0_df, ["Número de personal", "CC-nómina", "Importe"], "PCP0")
            st.session_state["pcp0_df"] = pcp0_df
            st.session_state["pcp0_log"] = pcp0_log
            st.session_state["results"] = None
            st.success("PCP0 cargado correctamente.")
        except Exception as exc:
            st.session_state["last_error"] = traceback.format_exc()
            st.error(str(exc))
    if isinstance(st.session_state["pcp0_df"], pd.DataFrame):
        p = st.session_state["pcp0_df"]
        c1, c2, c3 = st.columns(3)
        c1.metric("Registros PCP0", f"{len(p):,}".replace(",", "."))
        c2.metric("SAP únicos", f"{p['Número de personal'].nunique():,}".replace(",", "."))
        c3.metric("Conceptos", f"{p['CC-nómina'].dropna().nunique():,}".replace(",", "."))
        st.dataframe(st.session_state["pcp0_log"], width="stretch", height=180)

st.markdown("### 5) Procesar validación")
if st.button("Procesar validación", type="primary"):
    try:
        if not isinstance(st.session_state["cc_df"], pd.DataFrame):
            raise ValueError("Primero carga CC-nóminas.")
        if not isinstance(st.session_state["concepts_df"], pd.DataFrame):
            raise ValueError("Primero carga conceptos.")
        if not isinstance(st.session_state["pcp0_df"], pd.DataFrame):
            raise ValueError("Primero carga PCP0.")
        if selected_year is None:
            raise ValueError("Selecciona un año válido.")
        if not selected_periods:
            raise ValueError("Selecciona al menos un período para nómina.")

        with st.spinner("Procesando validación..."):
            cc_filtered = st.session_state["cc_df"].copy()
            cc_filtered = cc_filtered[cc_filtered["Fecha de pago_año"].eq(selected_year)].copy()
            cc_filtered = cc_filtered[cc_filtered["Período para nómina"].astype(str).isin([str(x) for x in selected_periods])].copy()

            pcp0_filtered = st.session_state["pcp0_df"].copy()
            sel = {str(x) for x in selected_periods}
            mask = pd.Series(False, index=pcp0_filtered.index)
            if "Período para nómina" in pcp0_filtered.columns:
                mask = mask | pcp0_filtered["Período para nómina"].astype(str).isin(sel)
            if "Período En cálc.nóm." in pcp0_filtered.columns:
                mask = mask | pcp0_filtered["Período En cálc.nóm."].astype(str).isin(sel)
            pcp0_filtered = pcp0_filtered[mask].copy()

            cc_resumen, cc_detalle = process_cc(cc_filtered, st.session_state["concepts_df"])
            pcp0_resumen, pcp0_detalle = process_pcp0(pcp0_filtered, st.session_state["concepts_df"])
            comp_resumen, comp_detalle = compare_cc_vs_pcp0(cc_resumen, cc_detalle, pcp0_resumen, pcp0_detalle)
            periodos = build_period_summary(cc_filtered, pcp0_filtered, st.session_state["concepts_df"])

            parametros = pd.DataFrame({"Parametro": ["Año seleccionado", "Períodos seleccionados"], "Valor": [str(selected_year), ", ".join(selected_periods)]})
            excel_bytes = to_excel_bytes({
                "Parametros": parametros,
                "Pagado_CC": cc_resumen,
                "Detalle_CC": cc_detalle,
                "Contabilizado_PCP0": pcp0_resumen,
                "Detalle_PCP0": pcp0_detalle,
                "Comparativo_Empleado": comp_resumen,
                "Detalle_Diferencias": comp_detalle,
                "Comparativo_Periodo": periodos,
                "Conceptos": st.session_state["concepts_df"],
                "Log_CC": st.session_state["cc_log"],
                "Log_PCP0": st.session_state["pcp0_log"],
            })
            zip_bytes = build_zip({
                "resultado_medios_magneticos.xlsx": excel_bytes,
                "parametros.csv": to_csv_bytes(parametros),
                "pagado_cc.csv": to_csv_bytes(cc_resumen),
                "detalle_cc.csv": to_csv_bytes(cc_detalle),
                "contabilizado_pcp0.csv": to_csv_bytes(pcp0_resumen),
                "detalle_pcp0.csv": to_csv_bytes(pcp0_detalle),
                "comparativo_empleado.csv": to_csv_bytes(comp_resumen),
                "detalle_diferencias.csv": to_csv_bytes(comp_detalle),
                "comparativo_periodo.csv": to_csv_bytes(periodos),
                "conceptos.csv": to_csv_bytes(st.session_state["concepts_df"]),
                "log_cc.csv": to_csv_bytes(st.session_state["cc_log"]),
                "log_pcp0.csv": to_csv_bytes(st.session_state["pcp0_log"]),
            })
            st.session_state["results"] = {
                "parametros": parametros,
                "cc_resumen": cc_resumen,
                "cc_detalle": cc_detalle,
                "pcp0_resumen": pcp0_resumen,
                "pcp0_detalle": pcp0_detalle,
                "comp_resumen": comp_resumen,
                "comp_detalle": comp_detalle,
                "periodos": periodos,
                "excel_bytes": excel_bytes,
                "zip_bytes": zip_bytes,
            }
            st.session_state["last_error"] = None
        st.success("Validación procesada correctamente.")
    except Exception as exc:
        st.session_state["last_error"] = traceback.format_exc()
        st.error(f"No fue posible procesar la validación: {exc}")

results = st.session_state.get("results")
if isinstance(results, dict):
    st.markdown("### Resultados")
    c1, c2, c3 = st.columns(3)
    c1.metric("SAP comparados", f"{results['comp_resumen']['Número de personal'].nunique():,}".replace(",", "."))
    c2.metric("SAP con diferencia", f"{(results['comp_resumen']['Estado'] != 'OK').sum():,}".replace(",", "."))
    c3.metric("Diferencia total", f"{results['comp_resumen']['Diferencia'].sum():,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
    st.dataframe(results["comp_resumen"].head(300), width="stretch", height=320)
    st.dataframe(results["comp_detalle"].head(300), width="stretch", height=320)
    col1, col2 = st.columns(2)
    col1.download_button("Descargar Excel", data=results["excel_bytes"], file_name="resultado_medios_magneticos.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    col2.download_button("Descargar ZIP", data=results["zip_bytes"], file_name="resultado_medios_magneticos.zip", mime="application/zip")

if st.session_state.get("last_error"):
    with st.expander("Ver detalle técnico del error"):
        st.code(st.session_state["last_error"])

st.markdown("<div class='footer-credit'>Creado por Andrés Huérfano Dávila - Nómina JMC</div>", unsafe_allow_html=True)
