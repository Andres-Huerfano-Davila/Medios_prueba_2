import io
import re
import zipfile
import unicodedata
import hashlib
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import numpy as np
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="Unificador, comparativo y generador de medios magnéticos",
    layout="wide",
)

st.markdown(
    """
    <style>
    .block-container {padding-top: 1rem; padding-bottom: 1.5rem;}
    .hero {
        background: linear-gradient(135deg,#fff4eb 0%,#fffaf6 100%);
        border:1px solid #f3c9a5;
        border-radius:16px;
        padding:18px 20px;
        margin-bottom:16px;
    }
    .hero h1 {margin:0; color:#a85400; font-size:1.5rem;}
    .hero p {margin:8px 0 0 0; color:#714221;}
    .credit {
        margin-top:8px;
        display:inline-block;
        background:#fff0e1;
        color:#a85400;
        border:1px solid #f3c9a5;
        border-radius:999px;
        padding:4px 10px;
        font-size:.9rem;
        font-weight:600;
    }
    .footer-credit {
        text-align:center;
        color:#a06a3d;
        font-weight:600;
        border:1px solid #f0d6bf;
        border-radius:12px;
        padding:10px;
        background:#fffaf6;
        margin-top:18px;
    }
    .soft-note {
        color:#8c5a2b;
        font-size:.92rem;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

st.markdown(
    """
    <div class="hero">
        <h1>Unificador, comparativo y generador de medios magnéticos</h1>
        <p>Carga archivos de CC-nóminas y contabilizaciones, clasifica conceptos y genera salidas en Excel, CSV y ZIP.</p>
        <div class="credit">Creado por Andrés Huérfano Dávila - Nómina JMC</div>
    </div>
    """,
    unsafe_allow_html=True,
)

MAX_EXCEL_ROWS = 1_048_576

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

SUMMARY_COLUMNS_CC = [
    "Número de personal",
    "Nombre del empleado o candidato",
    "Salariales",
    "Beneficios adicionales",
    "Importe total",
]
DETAIL_COLUMNS_CC = [
    "Número de personal",
    "Nombre del empleado o candidato",
    "CC-nómina",
    "Texto expl.CC-nómina",
    "Salariales",
    "Beneficios adicionales",
    "Importe total",
]
SUMMARY_COLUMNS_CONTAB = [
    "Número de personal",
    "Salariales",
    "Beneficios adicionales",
    "Importe total",
]
DETAIL_COLUMNS_CONTAB = [
    "Número de personal",
    "CC-nómina",
    "Texto expl.CC-nómina",
    "Salariales",
    "Beneficios adicionales",
    "Importe total",
]

CC_KEEP_COLUMNS = [
    "Número de personal",
    "Nombre del empleado o candidato",
    "CC-nómina",
    "Texto expl.CC-nómina",
    "Importe",
    "Fecha de pago",
    "Fecha de pago_display",
    "Fecha de pago_año",
    "Período para nómina",
    "__archivo_origen__",
    "__hoja_origen__",
]
CONTAB_KEEP_COLUMNS = [
    "Número de personal",
    "CC-nómina",
    "Texto expl.CC-nómina",
    "Importe",
    "Período para nómina",
    "Período En cálc.nóm.",
    "__archivo_origen__",
    "__hoja_origen__",
]


def normalize_text(value) -> str:
    if value is None:
        return ""
    text = str(value).strip().lower()
    text = unicodedata.normalize("NFKD", text)
    text = "".join(ch for ch in text if not unicodedata.combining(ch))
    text = re.sub(r"[\n\r\t]+", " ", text)
    text = re.sub(r"[^a-z0-9]+", " ", text)
    return re.sub(r"\s+", " ", text).strip()


def normalize_id(value) -> Optional[str]:
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


def parse_number(value) -> float:
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
        if text.rfind(",") > text.rfind("."):
            text = text.replace(".", "").replace(",", ".")
        else:
            text = text.replace(",", "")
    elif text.count(",") == 1 and "." not in text:
        text = text.replace(",", ".")
    elif text.count(".") > 1 and "," not in text:
        text = text.replace(".", "")
    try:
        return float(text)
    except Exception:
        return 0.0


def normalize_tipo(value) -> Optional[str]:
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

    # 1) exact canonical or explicit exact aliases
    exact_aliases = [normalize_text(canonical_name)] + [normalize_text(x) for x in EXACT_ALIASES.get(canonical_name, [])]
    for col, ncol in normalized_cols.items():
        if ncol in exact_aliases:
            return col

    # 2) strong rules for ambiguous columns
    if canonical_name == "Período para nómina":
        for col, ncol in normalized_cols.items():
            if ncol == "periodo para nomina":
                return col
        return None
    if canonical_name == "Período En cálc.nóm.":
        strong = {
            "periodo en calc nom",
            "periodo en calculo nomina",
            "periodo en calculo de nomina",
            "periodo en cal nom",
            "periodo en calculo n m",
        }
        for col, ncol in normalized_cols.items():
            if ncol in strong or ("periodo en" in ncol and "nom" in ncol):
                return col
        return None
    if canonical_name == "Fecha de pago":
        for col, ncol in normalized_cols.items():
            if ncol == "fecha de pago":
                return col
        return None

    # 3) contains aliases for non-ambiguous columns
    for col, ncol in normalized_cols.items():
        for alias in CONTAINS_ALIASES.get(canonical_name, []):
            if normalize_text(alias) in ncol:
                return col
    return None


def add_canonical_columns(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]

    for canonical in EXACT_ALIASES:
        original = preferred_column_match(df, canonical)
        if original and canonical not in df.columns:
            df[canonical] = df[original]

    if "Número de personal" in df.columns:
        df["Número de personal"] = df["Número de personal"].map(normalize_id)
    if "Nombre del empleado o candidato" in df.columns:
        df["Nombre del empleado o candidato"] = df["Nombre del empleado o candidato"].astype("string").fillna("").str.strip()
    if "CC-nómina" in df.columns:
        df["CC-nómina"] = df["CC-nómina"].astype("string").fillna("").str.strip()
    if "Texto expl.CC-nómina" in df.columns:
        df["Texto expl.CC-nómina"] = df["Texto expl.CC-nómina"].astype("string").fillna("").str.strip()
    if "Importe" in df.columns:
        df["Importe"] = df["Importe"].map(parse_number).astype("float64")
    if "Período para nómina" in df.columns:
        df["Período para nómina"] = df["Período para nómina"].map(normalize_id)
    if "Período En cálc.nóm." in df.columns:
        df["Período En cálc.nóm."] = df["Período En cálc.nóm."].map(normalize_id)
    if "Fecha de pago" in df.columns:
        df["Fecha de pago"] = robust_parse_dates(df["Fecha de pago"])
        df["Fecha de pago_display"] = df["Fecha de pago"].dt.strftime("%d/%m/%Y")
        df["Fecha de pago_año"] = df["Fecha de pago"].dt.year.astype("Int64").astype("string").fillna("")

    return df


def project_dataset(df: pd.DataFrame, dataset_type: str) -> pd.DataFrame:
    df = add_canonical_columns(df)

    if dataset_type == "cc":
        required = ["Número de personal", "Nombre del empleado o candidato", "CC-nómina", "Importe"]
        missing = [col for col in required if col not in df.columns]
        if missing:
            raise ValueError("Faltan columnas obligatorias: " + ", ".join(missing))
        if "Texto expl.CC-nómina" not in df.columns:
            df["Texto expl.CC-nómina"] = ""
        if "Período para nómina" not in df.columns:
            df["Período para nómina"] = ""
        if "Fecha de pago" not in df.columns:
            df["Fecha de pago"] = pd.NaT
            df["Fecha de pago_display"] = ""
            df["Fecha de pago_año"] = ""
        for col in ["__archivo_origen__", "__hoja_origen__"]:
            if col not in df.columns:
                df[col] = ""
        out = df[[col for col in CC_KEEP_COLUMNS if col in df.columns]].copy()
        out["Nombre del empleado o candidato"] = out["Nombre del empleado o candidato"].astype("string")
        out["CC-nómina"] = out["CC-nómina"].astype("string")
        out["Texto expl.CC-nómina"] = out["Texto expl.CC-nómina"].astype("string")
        if "Período para nómina" in out.columns:
            out["Período para nómina"] = out["Período para nómina"].astype("string")
        if "Fecha de pago_display" in out.columns:
            out["Fecha de pago_display"] = out["Fecha de pago_display"].astype("string")
        if "Fecha de pago_año" in out.columns:
            out["Fecha de pago_año"] = out["Fecha de pago_año"].astype("string")
        return out

    if dataset_type == "contab":
        required = ["Número de personal", "CC-nómina", "Importe"]
        missing = [col for col in required if col not in df.columns]
        if missing:
            raise ValueError("Faltan columnas obligatorias: " + ", ".join(missing))
        if "Texto expl.CC-nómina" not in df.columns:
            df["Texto expl.CC-nómina"] = ""
        if "Período para nómina" not in df.columns:
            df["Período para nómina"] = ""
        if "Período En cálc.nóm." not in df.columns:
            df["Período En cálc.nóm."] = ""
        for col in ["__archivo_origen__", "__hoja_origen__"]:
            if col not in df.columns:
                df[col] = ""
        out = df[[col for col in CONTAB_KEEP_COLUMNS if col in df.columns]].copy()
        out["CC-nómina"] = out["CC-nómina"].astype("string")
        out["Texto expl.CC-nómina"] = out["Texto expl.CC-nómina"].astype("string")
        out["Período para nómina"] = out["Período para nómina"].astype("string")
        out["Período En cálc.nóm."] = out["Período En cálc.nóm."].astype("string")
        return out

    raise ValueError("Tipo de dataset no soportado")


def hash_uploaded_file(uploaded_file) -> str:
    data = uploaded_file.getvalue()
    h = hashlib.md5()
    h.update(uploaded_file.name.encode("utf-8", errors="ignore"))
    h.update(str(len(data)).encode())
    h.update(data[:4096])
    return h.hexdigest()


def hash_uploaded_files(files) -> str:
    h = hashlib.md5()
    for f in files:
        data = f.getvalue()
        h.update(f.name.encode("utf-8", errors="ignore"))
        h.update(str(len(data)).encode())
        h.update(data[:4096])
    return h.hexdigest()


def read_text_file(data: bytes, filename: str) -> pd.DataFrame:
    encodings = ["utf-8", "utf-8-sig", "cp1252", "latin-1"]
    last_error = None

    # 1) SAP-style report with pipes
    for enc in encodings:
        try:
            text = data.decode(enc, errors="ignore")
            lines = [ln.rstrip("\n\r") for ln in text.splitlines() if ln.strip()]
            pipe_lines = [ln for ln in lines if "|" in ln]
            if len(pipe_lines) >= 2:
                header = None
                rows = []
                for ln in pipe_lines:
                    parts = [p.strip() for p in ln.strip().strip("|").split("|")]
                    if len(parts) < 2:
                        continue
                    compact = " ".join(parts).lower()
                    if any(key in compact for key in ["importe", "lib.mayor", "fe.contab", "fecha doc", "nº ejec"]):
                        header = parts
                        continue
                    if header and len(parts) == len(header):
                        rows.append(parts)
                if header and rows:
                    return pd.DataFrame(rows, columns=header)
        except Exception as exc:
            last_error = exc

    # 2) delimited txt/csv
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

    # 3) fixed width fallback
    for enc in encodings:
        try:
            df = pd.read_fwf(io.BytesIO(data), encoding=enc, dtype=object)
            if not df.empty:
                return df
        except Exception as exc:
            last_error = exc

    raise ValueError(f"No fue posible leer el archivo de texto: {filename}. {last_error}")


def read_excel_file(data: bytes, filename: str) -> pd.DataFrame:
    ext = Path(filename).suffix.lower()
    engine = None
    if ext == ".xlsb":
        engine = "pyxlsb"
    elif ext == ".xls":
        engine = "xlrd"
    elif ext == ".ods":
        engine = "odf"

    workbook = pd.ExcelFile(io.BytesIO(data), engine=engine)
    frames = []
    for sheet in workbook.sheet_names:
        try:
            part = pd.read_excel(io.BytesIO(data), sheet_name=sheet, engine=engine, dtype=object)
            if part is not None and not part.empty:
                part = part.copy()
                part["__archivo_origen__"] = filename
                part["__hoja_origen__"] = sheet
                frames.append(part)
        except Exception:
            continue
    if not frames:
        raise ValueError(f"No se pudieron leer hojas válidas en {filename}")
    return pd.concat(frames, ignore_index=True, sort=False)


def read_uploaded_table(uploaded_file) -> pd.DataFrame:
    filename = uploaded_file.name
    data = uploaded_file.getvalue()
    ext = Path(filename).suffix.lower()
    if ext in {".xlsx", ".xlsm", ".xls", ".xlsb", ".ods"}:
        return read_excel_file(data, filename)
    if ext in {".csv", ".txt"}:
        df = read_text_file(data, filename)
        df["__archivo_origen__"] = filename
        df["__hoja_origen__"] = "TXT/CSV"
        return df
    raise ValueError(f"Formato no soportado: {filename}")


def combine_uploaded_files(uploaded_files, dataset_type: str, label: str) -> Tuple[pd.DataFrame, pd.DataFrame]:
    frames: List[pd.DataFrame] = []
    logs: List[Dict[str, object]] = []

    progress = st.progress(0, text=f"Leyendo archivos de {label}... 0/{len(uploaded_files)}")
    status = st.empty()

    for idx, file in enumerate(uploaded_files, start=1):
        status.info(f"Leyendo: {file.name}")
        try:
            raw = read_uploaded_table(file)
            reduced = project_dataset(raw, dataset_type)
            frames.append(reduced)
            logs.append(
                {
                    "archivo": file.name,
                    "estado": "OK",
                    "registros": int(len(reduced)),
                    "columnas_detectadas": ", ".join(reduced.columns.astype(str).tolist()),
                    "detalle": "Archivo leído correctamente",
                }
            )
        except Exception as exc:
            logs.append(
                {
                    "archivo": file.name,
                    "estado": "ERROR",
                    "registros": 0,
                    "columnas_detectadas": "",
                    "detalle": str(exc),
                }
            )
        progress.progress(idx / len(uploaded_files), text=f"Leyendo archivos de {label}... {idx}/{len(uploaded_files)}")

    status.empty()
    progress.empty()

    if not frames:
        raise ValueError("No se pudo leer ninguno de los archivos cargados.")

    combined = pd.concat(frames, ignore_index=True, sort=False)
    log_df = pd.DataFrame(logs)
    return combined, log_df


def create_concepts_template(source_df: pd.DataFrame) -> pd.DataFrame:
    template = source_df.copy()
    if "CC-nómina" not in template.columns:
        template["CC-nómina"] = ""
    if "Texto expl.CC-nómina" not in template.columns:
        template["Texto expl.CC-nómina"] = ""
    template = template[["CC-nómina", "Texto expl.CC-nómina"]].copy()
    template["CC-nómina"] = template["CC-nómina"].astype("string").fillna("").str.strip()
    template["Texto expl.CC-nómina"] = template["Texto expl.CC-nómina"].astype("string").fillna("").str.strip()
    template = template[template["CC-nómina"].ne("")].drop_duplicates().sort_values(["CC-nómina", "Texto expl.CC-nómina"])
    template["Tipo"] = ""
    return template.reset_index(drop=True)


def prepare_concepts_df(df: pd.DataFrame) -> pd.DataFrame:
    df = add_canonical_columns(df.copy())
    for col in ["CC-nómina", "Texto expl.CC-nómina", "Tipo"]:
        if col not in df.columns:
            df[col] = ""
    out = df[["CC-nómina", "Texto expl.CC-nómina", "Tipo"]].copy()
    out["CC-nómina"] = out["CC-nómina"].astype("string").fillna("").str.strip()
    out["Texto expl.CC-nómina"] = out["Texto expl.CC-nómina"].astype("string").fillna("").str.strip()
    out["Tipo"] = out["Tipo"].map(normalize_tipo)
    out = out[out["CC-nómina"].ne("")].drop_duplicates(subset=["CC-nómina"], keep="first")
    return out.reset_index(drop=True)


def extract_concepts_from_upload(uploaded_file) -> pd.DataFrame:
    raw = read_uploaded_table(uploaded_file)
    return prepare_concepts_df(raw)


def process_cc_nominas(data_df: pd.DataFrame, concepts_df: pd.DataFrame):
    progress = st.progress(0, text="Procesando CC-nóminas...")
    work = data_df.copy()
    work["Nombre del empleado o candidato"] = work["Nombre del empleado o candidato"].astype("string").fillna("").str.strip()
    work["Importe"] = work["Importe"].map(parse_number)

    progress.progress(0.25, text="Preparando conceptos...")
    concepts = prepare_concepts_df(concepts_df)
    merged = work.merge(concepts, on="CC-nómina", how="left", suffixes=("", "_map"))

    if "Texto expl.CC-nómina_map" in merged.columns:
        merged["Texto expl.CC-nómina"] = merged["Texto expl.CC-nómina_map"].fillna(merged.get("Texto expl.CC-nómina", ""))
    if "Tipo_map" in merged.columns:
        merged["Tipo_concepto"] = merged["Tipo_map"]
    else:
        merged["Tipo_concepto"] = merged.get("Tipo", None)

    progress.progress(0.55, text="Filtrando conceptos aplicables...")
    merged = merged[merged["Tipo_concepto"].isin(["Salarial", "Beneficio"])].copy()
    merged["Salariales"] = np.where(merged["Tipo_concepto"].eq("Salarial"), merged["Importe"], 0.0)
    merged["Beneficios adicionales"] = np.where(merged["Tipo_concepto"].eq("Beneficio"), merged["Importe"], 0.0)

    progress.progress(0.8, text="Generando resumen y detalle...")
    resumen = (
        merged.groupby(["Número de personal", "Nombre del empleado o candidato"], dropna=False, as_index=False)[["Salariales", "Beneficios adicionales"]]
        .sum()
        .sort_values(["Número de personal", "Nombre del empleado o candidato"])
    )
    resumen["Importe total"] = resumen["Salariales"] + resumen["Beneficios adicionales"]
    resumen = resumen[SUMMARY_COLUMNS_CC].reset_index(drop=True)

    detalle = (
        merged.groupby(["Número de personal", "Nombre del empleado o candidato", "CC-nómina", "Texto expl.CC-nómina"], dropna=False, as_index=False)[["Salariales", "Beneficios adicionales"]]
        .sum()
        .sort_values(["Número de personal", "CC-nómina"])
    )
    detalle["Importe total"] = detalle["Salariales"] + detalle["Beneficios adicionales"]
    detalle = detalle[DETAIL_COLUMNS_CC].reset_index(drop=True)

    progress.progress(1.0, text="CC-nóminas procesadas")
    progress.empty()
    return resumen, detalle, concepts


def process_contabilizaciones(data_df: pd.DataFrame, concepts_df: pd.DataFrame):
    progress = st.progress(0, text="Procesando contabilizaciones...")
    work = data_df.copy()
    work["Importe"] = work["Importe"].map(parse_number)

    progress.progress(0.25, text="Preparando conceptos...")
    concepts = prepare_concepts_df(concepts_df)
    merged = work.merge(concepts, on="CC-nómina", how="left", suffixes=("", "_map"))

    if "Texto expl.CC-nómina_map" in merged.columns:
        merged["Texto expl.CC-nómina"] = merged["Texto expl.CC-nómina_map"].fillna(merged.get("Texto expl.CC-nómina", ""))
    if "Tipo_map" in merged.columns:
        merged["Tipo_concepto"] = merged["Tipo_map"]
    else:
        merged["Tipo_concepto"] = merged.get("Tipo", None)

    progress.progress(0.55, text="Filtrando conceptos aplicables...")
    merged = merged[merged["Tipo_concepto"].isin(["Salarial", "Beneficio"])].copy()
    merged["Salariales"] = np.where(merged["Tipo_concepto"].eq("Salarial"), merged["Importe"], 0.0)
    merged["Beneficios adicionales"] = np.where(merged["Tipo_concepto"].eq("Beneficio"), merged["Importe"], 0.0)

    progress.progress(0.8, text="Generando resumen y detalle...")
    resumen = (
        merged.groupby(["Número de personal"], dropna=False, as_index=False)[["Salariales", "Beneficios adicionales"]]
        .sum()
        .sort_values(["Número de personal"])
    )
    resumen["Importe total"] = resumen["Salariales"] + resumen["Beneficios adicionales"]
    resumen = resumen[SUMMARY_COLUMNS_CONTAB].reset_index(drop=True)

    detalle = (
        merged.groupby(["Número de personal", "CC-nómina", "Texto expl.CC-nómina"], dropna=False, as_index=False)[["Salariales", "Beneficios adicionales"]]
        .sum()
        .sort_values(["Número de personal", "CC-nómina"])
    )
    detalle["Importe total"] = detalle["Salariales"] + detalle["Beneficios adicionales"]
    detalle = detalle[DETAIL_COLUMNS_CONTAB].reset_index(drop=True)

    progress.progress(1.0, text="Contabilizaciones procesadas")
    progress.empty()
    return resumen, detalle, concepts


def chunk_dataframe(df: pd.DataFrame, chunk_size: int = MAX_EXCEL_ROWS) -> List[pd.DataFrame]:
    if len(df) <= chunk_size:
        return [df]
    parts = []
    for start in range(0, len(df), chunk_size):
        parts.append(df.iloc[start : start + chunk_size].copy())
    return parts


def format_excel_sheet(writer, sheet_name: str, df: pd.DataFrame):
    workbook = writer.book
    worksheet = writer.sheets[sheet_name]
    header_fmt = workbook.add_format({"bold": True, "bg_color": "#D9EAF7", "border": 1, "text_wrap": True, "valign": "top"})
    num_fmt = workbook.add_format({"num_format": "#,##0.00"})
    int_fmt = workbook.add_format({"num_format": "#,##0"})

    for col_num, value in enumerate(df.columns.values):
        worksheet.write(0, col_num, value, header_fmt)

    for idx, col in enumerate(df.columns):
        max_len = len(col) if df.empty else max(len(col), int(df[col].astype(str).head(500).map(len).max()))
        width = min(max_len + 2, 40)
        if col in {"Salariales", "Beneficios adicionales", "Importe total", "Importe", "Diferencia", "Importe total CC-nóminas", "Importe total Contabilización"}:
            worksheet.set_column(idx, idx, max(width, 16), num_fmt)
        elif col in {"Cantidad"}:
            worksheet.set_column(idx, idx, max(width, 12), int_fmt)
        else:
            worksheet.set_column(idx, idx, width)
    worksheet.freeze_panes(1, 0)
    worksheet.autofilter(0, 0, max(len(df), 1), len(df.columns) - 1)


def to_excel_bytes(sheets: Dict[str, pd.DataFrame]) -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        for base_name, df in sheets.items():
            parts = chunk_dataframe(df)
            if len(parts) == 1:
                sheet_name = base_name[:31]
                parts[0].to_excel(writer, sheet_name=sheet_name, index=False)
                format_excel_sheet(writer, sheet_name, parts[0])
            else:
                for i, part in enumerate(parts, start=1):
                    sheet_name = f"{base_name}_{i}"[:31]
                    part.to_excel(writer, sheet_name=sheet_name, index=False)
                    format_excel_sheet(writer, sheet_name, part)
    output.seek(0)
    return output.getvalue()


def to_csv_bytes(df: pd.DataFrame) -> bytes:
    return df.to_csv(index=False, encoding="utf-8-sig").encode("utf-8-sig")


def build_zip_bytes(files: Dict[str, bytes]) -> bytes:
    output = io.BytesIO()
    with zipfile.ZipFile(output, mode="w", compression=zipfile.ZIP_DEFLATED) as zf:
        for name, data in files.items():
            zf.writestr(name, data)
    output.seek(0)
    return output.getvalue()


def clean_generated_summary(df: pd.DataFrame) -> pd.DataFrame:
    work = add_canonical_columns(df.copy())
    work["Número de personal"] = work["Número de personal"].map(normalize_id)
    if "Importe total" not in work.columns:
        if {"Salariales", "Beneficios adicionales"}.issubset(work.columns):
            work["Importe total"] = work["Salariales"].map(parse_number) + work["Beneficios adicionales"].map(parse_number)
        elif "Importe" in work.columns:
            work["Importe total"] = work["Importe"].map(parse_number)
        else:
            work["Importe total"] = 0.0
    work["Importe total"] = work["Importe total"].map(parse_number)
    if "Nombre del empleado o candidato" not in work.columns:
        work["Nombre del empleado o candidato"] = ""
    work = work[["Número de personal", "Nombre del empleado o candidato", "Importe total"]].copy()
    work = work[work["Número de personal"].notna()].copy()
    return work.groupby(["Número de personal", "Nombre del empleado o candidato"], dropna=False, as_index=False)["Importe total"].sum()


def clean_generated_detail(df: pd.DataFrame) -> pd.DataFrame:
    work = add_canonical_columns(df.copy())
    for col in ["Número de personal", "CC-nómina", "Texto expl.CC-nómina"]:
        if col not in work.columns:
            work[col] = ""
    work["Número de personal"] = work["Número de personal"].map(normalize_id)
    work["CC-nómina"] = work["CC-nómina"].astype("string").fillna("").str.strip()
    work["Texto expl.CC-nómina"] = work["Texto expl.CC-nómina"].astype("string").fillna("").str.strip()
    if "Importe total" not in work.columns:
        if {"Salariales", "Beneficios adicionales"}.issubset(work.columns):
            work["Importe total"] = work["Salariales"].map(parse_number) + work["Beneficios adicionales"].map(parse_number)
        elif "Importe" in work.columns:
            work["Importe total"] = work["Importe"].map(parse_number)
        else:
            work["Importe total"] = 0.0
    work["Importe total"] = work["Importe total"].map(parse_number)
    work = work[work["Número de personal"].notna()].copy()
    return work[["Número de personal", "CC-nómina", "Texto expl.CC-nómina", "Importe total"]]


def summarize_from_detail(detail_df: pd.DataFrame) -> pd.DataFrame:
    work = clean_generated_detail(detail_df)
    summary = work.groupby(["Número de personal"], as_index=False)["Importe total"].sum()
    summary["Nombre del empleado o candidato"] = ""
    return summary[["Número de personal", "Nombre del empleado o candidato", "Importe total"]]


def parse_generated_result(uploaded_file):
    filename = uploaded_file.name
    ext = Path(filename).suffix.lower()
    data = uploaded_file.getvalue()

    if ext in {".xlsx", ".xlsm", ".xls", ".xlsb"}:
        engine = None
        if ext == ".xlsb":
            engine = "pyxlsb"
        elif ext == ".xls":
            engine = "xlrd"
        xls = pd.ExcelFile(io.BytesIO(data), engine=engine)
        resumen = None
        detalle = None
        for sheet in xls.sheet_names:
            temp = pd.read_excel(io.BytesIO(data), sheet_name=sheet, engine=engine, dtype=object)
            temp = add_canonical_columns(temp)
            cols = set(temp.columns)
            nsheet = normalize_text(sheet)
            if nsheet == "resumen":
                resumen = temp.copy()
            elif nsheet == "detalle":
                detalle = temp.copy()
            elif {"Número de personal", "Importe total", "CC-nómina"}.issubset(cols) and detalle is None:
                detalle = temp.copy()
            elif {"Número de personal", "Importe total"}.issubset(cols) and resumen is None:
                resumen = temp.copy()
        if detalle is not None and resumen is None:
            resumen = summarize_from_detail(detalle)
        if resumen is None:
            raise ValueError(f"No se pudo identificar una hoja resumen válida en {filename}")
        if detalle is None:
            detalle = pd.DataFrame(columns=["Número de personal", "CC-nómina", "Texto expl.CC-nómina", "Importe total"])
        return clean_generated_summary(resumen), clean_generated_detail(detalle)

    if ext in {".csv", ".txt"}:
        df = add_canonical_columns(read_text_file(data, filename))
        cols = set(df.columns)
        if {"Número de personal", "Importe total", "CC-nómina"}.issubset(cols):
            detalle = clean_generated_detail(df)
            resumen = summarize_from_detail(detalle)
            return resumen, detalle
        if {"Número de personal", "Importe total"}.issubset(cols):
            resumen = clean_generated_summary(df)
            detalle = pd.DataFrame(columns=["Número de personal", "CC-nómina", "Texto expl.CC-nómina", "Importe total"])
            return resumen, detalle
        raise ValueError(f"El archivo {filename} no tiene las columnas necesarias para el comparativo.")

    raise ValueError(f"Formato no soportado para comparativo: {filename}")


def generate_comparative(cc_summary, cc_detail, contab_summary, contab_detail, tolerance=0.0):
    cc_sum = clean_generated_summary(cc_summary).rename(columns={"Importe total": "Importe total CC-nóminas"})
    contab_sum = clean_generated_summary(contab_summary).rename(columns={"Importe total": "Importe total Contabilización"})
    resumen = cc_sum.merge(contab_sum[["Número de personal", "Importe total Contabilización"]], on="Número de personal", how="outer")
    resumen["Nombre del empleado o candidato"] = resumen.get("Nombre del empleado o candidato", "").fillna("")
    resumen["Importe total CC-nóminas"] = resumen["Importe total CC-nóminas"].fillna(0.0)
    resumen["Importe total Contabilización"] = resumen["Importe total Contabilización"].fillna(0.0)
    resumen["Diferencia"] = resumen["Importe total CC-nóminas"] - resumen["Importe total Contabilización"]
    resumen = resumen.sort_values(["Número de personal", "Nombre del empleado o candidato"]).reset_index(drop=True)

    cc_det = clean_generated_detail(cc_detail).rename(columns={"Importe total": "Importe total CC-nóminas"})
    contab_det = clean_generated_detail(contab_detail).rename(columns={"Importe total": "Importe total Contabilización"})
    detalle = cc_det.merge(contab_det, on=["Número de personal", "CC-nómina"], how="outer", suffixes=("_cc", "_contab"))
    detalle["Texto expl.CC-nómina"] = detalle.get("Texto expl.CC-nómina_cc", "").fillna(detalle.get("Texto expl.CC-nómina_contab", ""))
    detalle["Importe total CC-nóminas"] = detalle["Importe total CC-nóminas"].fillna(0.0)
    detalle["Importe total Contabilización"] = detalle["Importe total Contabilización"].fillna(0.0)
    detalle["Diferencia"] = detalle["Importe total CC-nóminas"] - detalle["Importe total Contabilización"]

    saps = resumen.loc[resumen["Diferencia"].abs() > tolerance, "Número de personal"].tolist()
    detalle = detalle[detalle["Número de personal"].isin(saps)].copy()
    detalle = detalle[detalle["Diferencia"].abs() > tolerance].copy()
    detalle = detalle[["Número de personal", "CC-nómina", "Texto expl.CC-nómina", "Importe total CC-nóminas", "Importe total Contabilización", "Diferencia"]].sort_values(["Número de personal", "CC-nómina"]).reset_index(drop=True)
    return resumen, detalle


def show_metrics(summary_df: pd.DataFrame, total_column: str, label_prefix: str):
    c1, c2, c3 = st.columns(3)
    c1.metric(f"{label_prefix} registros resumen", f"{len(summary_df):,}".replace(",", "."))
    c2.metric(f"{label_prefix} SAP únicos", f"{summary_df['Número de personal'].nunique():,}".replace(",", "."))
    c3.metric(f"{label_prefix} total importe", f"{summary_df[total_column].sum():,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))


def render_log(df_log: pd.DataFrame, title: str, key_prefix: str):
    st.markdown(f"**{title}**")
    if df_log is None or df_log.empty:
        st.info("Aún no hay registros de lectura.")
        return
    ok_count = int((df_log["estado"] == "OK").sum()) if "estado" in df_log.columns else 0
    err_count = int((df_log["estado"] == "ERROR").sum()) if "estado" in df_log.columns else 0
    c1, c2 = st.columns(2)
    c1.success(f"Archivos leídos correctamente: {ok_count}")
    c2.warning(f"Archivos con error: {err_count}")
    st.dataframe(df_log, width="stretch", height=220)
    st.download_button(
        f"Descargar log {title}",
        data=to_csv_bytes(df_log),
        file_name=f"log_{key_prefix}.csv",
        mime="text/csv",
        key=f"dl_log_{key_prefix}",
    )


def get_year_options(df: pd.DataFrame) -> List[str]:
    if "Fecha de pago_año" not in df.columns:
        return []
    vals = df["Fecha de pago_año"].astype("string").fillna("")
    return sorted([x for x in vals.unique().tolist() if x])


def get_filter_options(df: pd.DataFrame, column: str) -> List[str]:
    if column not in df.columns:
        return []
    vals = df[column].astype("string").fillna("").str.strip()
    return sorted([x for x in vals.unique().tolist() if x])


def init_state():
    defaults = {
        "cc_raw": None,
        "cc_log": pd.DataFrame(columns=["archivo", "estado", "registros", "columnas_detectadas", "detalle"]),
        "cc_file_hash": None,
        "cc_resumen": None,
        "cc_detalle": None,
        "contab_raw": None,
        "contab_log": pd.DataFrame(columns=["archivo", "estado", "registros", "columnas_detectadas", "detalle"]),
        "contab_file_hash": None,
        "contab_resumen": None,
        "contab_detalle": None,
        "conceptos_df": None,
        "conceptos_hash": None,
        "comparativo_resumen": None,
        "comparativo_detalle": None,
        "cc_filter_years": [],
        "cc_filter_dates": [],
        "cc_filter_use_dates": False,
        "cc_filter_periods": [],
        "cc_filter_use_periods": False,
        "contab_filter_periods_nomina": [],
        "contab_filter_use_periods_nomina": False,
        "contab_filter_periods_calc": [],
        "contab_filter_use_periods_calc": False,
    }
    for key, value in defaults.items():
        if key not in st.session_state:
            st.session_state[key] = value


def apply_cc_filters(df: pd.DataFrame) -> pd.DataFrame:
    work = df.copy()
    years = st.session_state.get("cc_filter_years", [])
    if years and "Fecha de pago_año" in work.columns:
        work = work[work["Fecha de pago_año"].astype("string").isin(years)].copy()
    if st.session_state.get("cc_filter_use_dates", False):
        dates = st.session_state.get("cc_filter_dates", [])
        if dates:
            work = work[work["Fecha de pago_display"].astype("string").isin(dates)].copy()
        else:
            work = work.iloc[0:0].copy()
    if st.session_state.get("cc_filter_use_periods", False):
        periods = st.session_state.get("cc_filter_periods", [])
        if periods:
            work = work[work["Período para nómina"].astype("string").isin(periods)].copy()
        else:
            work = work.iloc[0:0].copy()
    return work


def apply_contab_filters(df: pd.DataFrame) -> pd.DataFrame:
    work = df.copy()
    if st.session_state.get("contab_filter_use_periods_nomina", False):
        periods = st.session_state.get("contab_filter_periods_nomina", [])
        if periods:
            work = work[work["Período para nómina"].astype("string").isin(periods)].copy()
        else:
            work = work.iloc[0:0].copy()
    if st.session_state.get("contab_filter_use_periods_calc", False):
        periods = st.session_state.get("contab_filter_periods_calc", [])
        if periods:
            work = work[work["Período En cálc.nóm."].astype("string").isin(periods)].copy()
        else:
            work = work.iloc[0:0].copy()
    return work


init_state()

with st.expander("Notas importantes", expanded=False):
    st.markdown(
        """
        - La app intenta leer **TXT, CSV, XLS, XLSX, XLSB y ODS**.
        - Si un archivo falla, queda registrado en el log y la app sigue con los demás.
        - La app guarda en memoria solo las columnas necesarias para que sea más estable y rápida.
        - Los filtros no recalculan el consolidado hasta que oprimas **Generar**.
        - En **CC-nóminas**, el filtro principal es **Año de Fecha de pago** y luego **Fecha de pago**; después aparece **Período para nómina**.
        - En **Contabilizaciones**, solo se filtra por **Período para nómina** y **Período En cálc.nóm.**.
        - Si una hoja supera **1.048.576 filas**, el Excel se divide automáticamente en varias hojas.
        """
    )


tab1, tab2, tab3 = st.tabs(["cc-nóminas", "Contabilizaciones", "comparativo"])

with tab1:
    st.subheader("1) Consolidación de CC-nóminas")
    uploaded_cc_files = st.file_uploader(
        "Carga uno o varios archivos de CC-nóminas / acumulados",
        type=["txt", "csv", "xls", "xlsx", "xlsb", "ods"],
        accept_multiple_files=True,
        key="cc_files",
    )

    if st.button("Leer archivos CC-nóminas", type="primary", key="btn_load_cc"):
        if not uploaded_cc_files:
            st.warning("Primero carga uno o varios archivos.")
        else:
            try:
                file_hash = hash_uploaded_files(uploaded_cc_files)
                cc_raw, cc_log = combine_uploaded_files(uploaded_cc_files, "cc", "CC-nóminas")
                st.session_state["cc_raw"] = cc_raw
                st.session_state["cc_log"] = cc_log
                st.session_state["cc_file_hash"] = file_hash
                years = get_year_options(cc_raw)
                st.session_state["cc_filter_years"] = years
                st.session_state["cc_filter_dates"] = get_filter_options(cc_raw[cc_raw["Fecha de pago_año"].astype("string").isin(years)] if years else cc_raw, "Fecha de pago_display") if "Fecha de pago_display" in cc_raw.columns else []
                st.session_state["cc_filter_periods"] = get_filter_options(cc_raw, "Período para nómina")
                st.session_state["cc_filter_use_dates"] = False
                st.session_state["cc_filter_use_periods"] = False
                st.success("Archivos de CC-nóminas cargados correctamente.")
            except Exception as exc:
                st.error(f"No fue posible procesar los archivos de CC-nóminas: {exc}")

    cc_raw = st.session_state.get("cc_raw")
    cc_log = st.session_state.get("cc_log")
    render_log(cc_log, "Log de lectura CC-nóminas", "cc")

    if isinstance(cc_raw, pd.DataFrame) and not cc_raw.empty:
        c1, c2, c3 = st.columns(3)
        c1.metric("Registros cargados", f"{len(cc_raw):,}".replace(",", "."))
        c2.metric("SAP detectados", f"{cc_raw['Número de personal'].nunique():,}".replace(",", "."))
        c3.metric("Conceptos detectados", f"{cc_raw['CC-nómina'].nunique():,}".replace(",", "."))

        preview = cc_raw[[c for c in ["__archivo_origen__", "__hoja_origen__", "Número de personal", "CC-nómina", "Importe", "Nombre del empleado o candidato", "Fecha de pago_display", "Período para nómina"] if c in cc_raw.columns]].head(100).copy()
        if "Fecha de pago_display" in preview.columns:
            preview = preview.rename(columns={"Fecha de pago_display": "Fecha de pago"})
        st.markdown("**Vista previa de archivos cargados**")
        st.dataframe(preview, width="stretch", height=280)

        st.markdown("### Filtros de consolidación")
        st.markdown("<div class='soft-note'>Los filtros quedan guardados en memoria. El consolidado solo se recalcula cuando oprimes Generar.</div>", unsafe_allow_html=True)

        year_options = get_year_options(cc_raw)
        current_years = [x for x in st.session_state.get("cc_filter_years", year_options) if x in year_options]
        temp_after_year = cc_raw[cc_raw["Fecha de pago_año"].astype("string").isin(current_years)].copy() if current_years else cc_raw.iloc[0:0].copy()
        date_options = get_filter_options(temp_after_year, "Fecha de pago_display")
        current_dates = [x for x in st.session_state.get("cc_filter_dates", date_options) if x in date_options]
        temp_after_date = temp_after_year.copy()
        if st.session_state.get("cc_filter_use_dates", False) and current_dates:
            temp_after_date = temp_after_date[temp_after_date["Fecha de pago_display"].astype("string").isin(current_dates)].copy()
        period_options = get_filter_options(temp_after_date, "Período para nómina")
        current_periods = [x for x in st.session_state.get("cc_filter_periods", period_options) if x in period_options]

        with st.form("form_cc_filters"):
            st.markdown("**Fecha de pago**")
            selected_years = st.multiselect("Año", options=year_options, default=current_years)
            use_dates = st.checkbox("Usar filtro para Fecha de pago", value=st.session_state.get("cc_filter_use_dates", False))
            selected_dates = st.multiselect("Fechas disponibles", options=date_options, default=current_dates)

            st.markdown("**Período para nómina**")
            use_periods = st.checkbox("Usar filtro para Período para nómina", value=st.session_state.get("cc_filter_use_periods", False))
            selected_periods = st.multiselect("Períodos disponibles", options=period_options, default=current_periods)

            apply_cc_filters_btn = st.form_submit_button("Aplicar filtros CC-nóminas", type="primary")

            if apply_cc_filters_btn:
                st.session_state["cc_filter_years"] = selected_years
                st.session_state["cc_filter_use_dates"] = use_dates
                st.session_state["cc_filter_dates"] = selected_dates if selected_dates else date_options
                st.session_state["cc_filter_use_periods"] = use_periods
                st.session_state["cc_filter_periods"] = selected_periods if selected_periods else period_options
                st.success("Filtros actualizados")

        cc_filtered = apply_cc_filters(cc_raw)
        st.info(f"Registros después de filtros: {len(cc_filtered):,}".replace(",", "."))

        employees_cc = cc_filtered[["Número de personal", "Nombre del empleado o candidato"]].dropna(subset=["Número de personal"]).drop_duplicates().sort_values(["Número de personal", "Nombre del empleado o candidato"]).head(200)
        st.markdown("**SAP únicos detectados**")
        st.dataframe(employees_cc, width="stretch", height=220)

        concepts_template = create_concepts_template(cc_filtered if not cc_filtered.empty else cc_raw)
        st.download_button(
            "Descargar plantilla de conceptos",
            data=to_excel_bytes({"Conceptos": concepts_template}),
            file_name="plantilla_conceptos_cc_nominas.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="dl_template_cc",
        )

        uploaded_concepts = st.file_uploader(
            "Carga la base de conceptos diligenciada (CC-nómina, Texto expl.CC-nómina, Tipo)",
            type=["txt", "csv", "xls", "xlsx", "xlsb", "ods"],
            accept_multiple_files=False,
            key="concepts_file_cc",
        )
        if uploaded_concepts is not None:
            current_hash = hash_uploaded_file(uploaded_concepts)
            if st.session_state.get("conceptos_hash") != current_hash:
                try:
                    st.session_state["conceptos_df"] = extract_concepts_from_upload(uploaded_concepts)
                    st.session_state["conceptos_hash"] = current_hash
                    st.success("Conceptos cargados en memoria.")
                except Exception as exc:
                    st.error(f"No fue posible leer el archivo de conceptos: {exc}")

        concepts_df = st.session_state.get("conceptos_df")
        if isinstance(concepts_df, pd.DataFrame) and not concepts_df.empty:
            invalid_types = concepts_df[concepts_df["Tipo"].isna()].copy()
            if not invalid_types.empty:
                st.warning("Hay conceptos sin Tipo válido. Usa solamente: Salarial, Beneficio o No aplica.")
                st.dataframe(invalid_types.head(100), width="stretch", height=180)

            if st.button("Generar consolidado CC-nóminas", type="primary", key="btn_generate_cc"):
                try:
                    resumen_cc, detalle_cc, conceptos_cc = process_cc_nominas(cc_filtered, concepts_df)
                    st.session_state["cc_resumen"] = resumen_cc
                    st.session_state["cc_detalle"] = detalle_cc
                    st.session_state["conceptos_df"] = conceptos_cc
                    st.success("Consolidado de CC-nóminas generado.")
                except Exception as exc:
                    st.error(f"No fue posible generar el consolidado de CC-nóminas: {exc}")

        if isinstance(st.session_state.get("cc_resumen"), pd.DataFrame):
            resumen_cc = st.session_state["cc_resumen"]
            detalle_cc = st.session_state["cc_detalle"]
            conceptos_cc = st.session_state.get("conceptos_df", pd.DataFrame())
            show_metrics(resumen_cc, "Importe total", "CC-nóminas")
            st.markdown("**Resumen totalizado por empleado**")
            st.dataframe(resumen_cc.head(300), width="stretch", height=300)
            st.markdown("**Detalle por concepto**")
            st.dataframe(detalle_cc.head(300), width="stretch", height=300)
            excel_cc = to_excel_bytes({"Resumen": resumen_cc, "Detalle": detalle_cc, "Conceptos": conceptos_cc})
            zip_cc = build_zip_bytes({
                "cc_nominas_resultado.xlsx": excel_cc,
                "cc_nominas_resumen.csv": to_csv_bytes(resumen_cc),
                "cc_nominas_detalle.csv": to_csv_bytes(detalle_cc),
                "conceptos_utilizados.csv": to_csv_bytes(conceptos_cc),
                "log_cc_nominas.csv": to_csv_bytes(cc_log),
            })
            c1, c2 = st.columns(2)
            c1.download_button("Descargar Excel CC-nóminas", data=excel_cc, file_name="cc_nominas_resultado.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            c2.download_button("Descargar ZIP CC-nóminas", data=zip_cc, file_name="cc_nominas_resultados.zip", mime="application/zip")

with tab2:
    st.subheader("2) Consolidación de Contabilizaciones / PCP0")
    uploaded_contab_files = st.file_uploader(
        "Carga uno o varios archivos de contabilización / PCP0",
        type=["txt", "csv", "xls", "xlsx", "xlsb", "ods"],
        accept_multiple_files=True,
        key="contab_files",
    )

    if st.button("Leer archivos Contabilizaciones", type="primary", key="btn_load_contab"):
        if not uploaded_contab_files:
            st.warning("Primero carga uno o varios archivos.")
        else:
            try:
                file_hash = hash_uploaded_files(uploaded_contab_files)
                contab_raw, contab_log = combine_uploaded_files(uploaded_contab_files, "contab", "Contabilizaciones")
                st.session_state["contab_raw"] = contab_raw
                st.session_state["contab_log"] = contab_log
                st.session_state["contab_file_hash"] = file_hash
                st.session_state["contab_filter_periods_nomina"] = get_filter_options(contab_raw, "Período para nómina")
                st.session_state["contab_filter_periods_calc"] = get_filter_options(contab_raw, "Período En cálc.nóm.")
                st.session_state["contab_filter_use_periods_nomina"] = False
                st.session_state["contab_filter_use_periods_calc"] = False
                st.success("Archivos de contabilización cargados correctamente.")
            except Exception as exc:
                st.error(f"No fue posible procesar los archivos de contabilización: {exc}")

    contab_raw = st.session_state.get("contab_raw")
    contab_log = st.session_state.get("contab_log")
    render_log(contab_log, "Log de lectura Contabilizaciones", "contab")

    if isinstance(contab_raw, pd.DataFrame) and not contab_raw.empty:
        c1, c2, c3 = st.columns(3)
        c1.metric("Registros cargados", f"{len(contab_raw):,}".replace(",", "."))
        c2.metric("SAP detectados", f"{contab_raw['Número de personal'].nunique():,}".replace(",", "."))
        c3.metric("Conceptos detectados", f"{contab_raw['CC-nómina'].nunique():,}".replace(",", "."))

        preview = contab_raw[[c for c in ["__archivo_origen__", "__hoja_origen__", "Número de personal", "CC-nómina", "Importe", "Período para nómina", "Período En cálc.nóm."] if c in contab_raw.columns]].head(100).copy()
        st.markdown("**Vista previa de archivos cargados**")
        st.dataframe(preview, width="stretch", height=280)

        st.markdown("### Filtros de consolidación")
        st.markdown("<div class='soft-note'>En contabilizaciones solo se muestran los filtros que realmente aplican: Período para nómina y Período En cálc.nóm.</div>", unsafe_allow_html=True)

        period_nom_options = get_filter_options(contab_raw, "Período para nómina")
        current_period_nom = [x for x in st.session_state.get("contab_filter_periods_nomina", period_nom_options) if x in period_nom_options]
        temp_after_nom = contab_raw.copy()
        if st.session_state.get("contab_filter_use_periods_nomina", False) and current_period_nom:
            temp_after_nom = temp_after_nom[temp_after_nom["Período para nómina"].astype("string").isin(current_period_nom)].copy()
        calc_options = get_filter_options(temp_after_nom, "Período En cálc.nóm.")
        current_calc = [x for x in st.session_state.get("contab_filter_periods_calc", calc_options) if x in calc_options]

        with st.form("form_contab_filters"):
            use_nom = st.checkbox("Usar filtro para Período para nómina", value=st.session_state.get("contab_filter_use_periods_nomina", False))
            selected_nom = st.multiselect("Período para nómina", options=period_nom_options, default=current_period_nom)
            use_calc = st.checkbox("Usar filtro para Período En cálc.nóm.", value=st.session_state.get("contab_filter_use_periods_calc", False))
            selected_calc = st.multiselect("Período En cálc.nóm.", options=calc_options, default=current_calc)
            apply_contab_filters_btn = st.form_submit_button("Aplicar filtros Contabilizaciones", type="primary")
            if apply_contab_filters_btn:
                st.session_state["contab_filter_use_periods_nomina"] = use_nom
                st.session_state["contab_filter_periods_nomina"] = selected_nom if selected_nom else period_nom_options
                st.session_state["contab_filter_use_periods_calc"] = use_calc
                st.session_state["contab_filter_periods_calc"] = selected_calc if selected_calc else calc_options
                st.success("Filtros actualizados")

        contab_filtered = apply_contab_filters(contab_raw)
        st.info(f"Registros después de filtros: {len(contab_filtered):,}".replace(",", "."))
        employees_contab = contab_filtered[["Número de personal"]].dropna(subset=["Número de personal"]).drop_duplicates().sort_values(["Número de personal"]).head(200)
        st.markdown("**SAP únicos detectados**")
        st.dataframe(employees_contab, width="stretch", height=220)

        st.info("Se usarán los conceptos cargados previamente en CC-nóminas. Si quieres, puedes reemplazarlos abajo.")
        uploaded_concepts_contab = st.file_uploader(
            "Opcional: carga nuevamente el archivo de conceptos para reemplazar el anterior",
            type=["txt", "csv", "xls", "xlsx", "xlsb", "ods"],
            accept_multiple_files=False,
            key="concepts_file_contab",
        )
        if uploaded_concepts_contab is not None:
            current_hash = hash_uploaded_file(uploaded_concepts_contab)
            if st.session_state.get("conceptos_hash") != current_hash:
                try:
                    st.session_state["conceptos_df"] = extract_concepts_from_upload(uploaded_concepts_contab)
                    st.session_state["conceptos_hash"] = current_hash
                    st.success("Conceptos actualizados en memoria.")
                except Exception as exc:
                    st.error(f"No fue posible leer el archivo de conceptos: {exc}")

        concepts_df = st.session_state.get("conceptos_df")
        if not isinstance(concepts_df, pd.DataFrame) or concepts_df.empty:
            st.warning("Primero carga y clasifica los conceptos en la pestaña CC-nóminas o súbelos aquí.")
        else:
            invalid_types = concepts_df[concepts_df["Tipo"].isna()].copy()
            if not invalid_types.empty:
                st.warning("Hay conceptos sin Tipo válido. Usa solamente: Salarial, Beneficio o No aplica.")
                st.dataframe(invalid_types.head(100), width="stretch", height=180)

            if st.button("Generar consolidado Contabilizaciones", type="primary", key="btn_generate_contab"):
                try:
                    resumen_contab, detalle_contab, conceptos_contab = process_contabilizaciones(contab_filtered, concepts_df)
                    st.session_state["contab_resumen"] = resumen_contab
                    st.session_state["contab_detalle"] = detalle_contab
                    st.session_state["conceptos_df"] = conceptos_contab
                    st.success("Consolidado de contabilizaciones generado.")
                except Exception as exc:
                    st.error(f"No fue posible generar el consolidado de contabilizaciones: {exc}")

        if isinstance(st.session_state.get("contab_resumen"), pd.DataFrame):
            resumen_contab = st.session_state["contab_resumen"]
            detalle_contab = st.session_state["contab_detalle"]
            conceptos_contab = st.session_state.get("conceptos_df", pd.DataFrame())
            show_metrics(resumen_contab, "Importe total", "Contabilización")
            st.markdown("**Resumen totalizado por SAP**")
            st.dataframe(resumen_contab.head(300), width="stretch", height=300)
            st.markdown("**Detalle por concepto**")
            st.dataframe(detalle_contab.head(300), width="stretch", height=300)
            excel_contab = to_excel_bytes({"Resumen": resumen_contab, "Detalle": detalle_contab, "Conceptos": conceptos_contab})
            zip_contab = build_zip_bytes({
                "contabilizaciones_resultado.xlsx": excel_contab,
                "contabilizaciones_resumen.csv": to_csv_bytes(resumen_contab),
                "contabilizaciones_detalle.csv": to_csv_bytes(detalle_contab),
                "conceptos_utilizados.csv": to_csv_bytes(conceptos_contab),
                "log_contabilizaciones.csv": to_csv_bytes(contab_log),
            })
            c1, c2 = st.columns(2)
            c1.download_button("Descargar Excel Contabilizaciones", data=excel_contab, file_name="contabilizaciones_resultado.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            c2.download_button("Descargar ZIP Contabilizaciones", data=zip_contab, file_name="contabilizaciones_resultados.zip", mime="application/zip")

with tab3:
    st.subheader("3) Comparativo")
    st.markdown("Compara **Número de personal + Importe total** entre CC-nóminas y Contabilización, y genera el detalle de diferencias por concepto.")
    tolerance = st.number_input("Tolerancia para considerar diferencia (0 = exacto)", min_value=0.0, value=0.0, step=0.01, key="tolerance_compare")

    can_compare_session = all(isinstance(st.session_state.get(k), pd.DataFrame) for k in ["cc_resumen", "cc_detalle", "contab_resumen", "contab_detalle"])
    if can_compare_session:
        st.success("Hay resultados en memoria de las pestañas anteriores.")
        if st.button("Generar comparativo con los datos ya procesados", type="primary", key="btn_compare_session"):
            try:
                resumen_comp, detalle_comp = generate_comparative(
                    st.session_state["cc_resumen"],
                    st.session_state["cc_detalle"],
                    st.session_state["contab_resumen"],
                    st.session_state["contab_detalle"],
                    tolerance=tolerance,
                )
                st.session_state["comparativo_resumen"] = resumen_comp
                st.session_state["comparativo_detalle"] = detalle_comp
                st.success("Comparativo generado con resultados en memoria.")
            except Exception as exc:
                st.error(f"No fue posible generar el comparativo: {exc}")

    st.markdown("**O, si prefieres, vuelve a cargar los archivos generados**")
    c1, c2 = st.columns(2)
    with c1:
        compare_cc_file = st.file_uploader("Archivo resultado de CC-nóminas (Excel, CSV o TXT)", type=["txt", "csv", "xls", "xlsx", "xlsb"], accept_multiple_files=False, key="compare_cc_file")
    with c2:
        compare_contab_file = st.file_uploader("Archivo resultado de Contabilizaciones (Excel, CSV o TXT)", type=["txt", "csv", "xls", "xlsx", "xlsb"], accept_multiple_files=False, key="compare_contab_file")

    if st.button("Generar comparativo con archivos cargados", key="btn_compare_upload"):
        if compare_cc_file is None or compare_contab_file is None:
            st.warning("Carga ambos archivos resultado para generar el comparativo por esta vía.")
        else:
            try:
                cc_summary_file, cc_detail_file = parse_generated_result(compare_cc_file)
                contab_summary_file, contab_detail_file = parse_generated_result(compare_contab_file)
                resumen_comp, detalle_comp = generate_comparative(cc_summary_file, cc_detail_file, contab_summary_file, contab_detail_file, tolerance=tolerance)
                st.session_state["comparativo_resumen"] = resumen_comp
                st.session_state["comparativo_detalle"] = detalle_comp
                st.success("Comparativo generado con archivos cargados.")
            except Exception as exc:
                st.error(f"No fue posible generar el comparativo con archivos cargados: {exc}")

    if isinstance(st.session_state.get("comparativo_resumen"), pd.DataFrame):
        resumen_comp = st.session_state["comparativo_resumen"]
        detalle_comp = st.session_state["comparativo_detalle"]
        c1, c2, c3 = st.columns(3)
        c1.metric("SAP comparados", f"{resumen_comp['Número de personal'].nunique():,}".replace(",", "."))
        c2.metric("SAP con diferencia", f"{(resumen_comp['Diferencia'].abs() > tolerance).sum():,}".replace(",", "."))
        c3.metric("Diferencia total", f"{resumen_comp['Diferencia'].sum():,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
        st.markdown("**Resumen comparativo**")
        st.dataframe(resumen_comp.head(300), width="stretch", height=320)
        st.markdown("**Detalle de diferencias por CC-nómina**")
        st.dataframe(detalle_comp.head(300), width="stretch", height=320)
        excel_compare = to_excel_bytes({"Resumen_Comparativo": resumen_comp, "Detalle_Diferencias": detalle_comp})
        zip_compare = build_zip_bytes({
            "comparativo_resultado.xlsx": excel_compare,
            "comparativo_resumen.csv": to_csv_bytes(resumen_comp),
            "comparativo_detalle.csv": to_csv_bytes(detalle_comp),
        })
        c1, c2 = st.columns(2)
        c1.download_button("Descargar Excel Comparativo", data=excel_compare, file_name="comparativo_resultado.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        c2.download_button("Descargar ZIP Comparativo", data=zip_compare, file_name="comparativo_resultados.zip", mime="application/zip")

st.markdown("<div class='footer-credit'>Creado por Andrés Huérfano Dávila - Nómina JMC</div>", unsafe_allow_html=True)
