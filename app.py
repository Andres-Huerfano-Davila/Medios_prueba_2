
import io
import math
import re
import unicodedata
import zipfile
import hashlib
from pathlib import Path
from typing import Dict, List, Optional, Sequence, Tuple

import numpy as np
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="Medios Magnéticos Nómina JMC",
    page_icon="🧾",
    layout="wide",
)

st.markdown(
    """
    <style>
    .block-container {padding-top: 1rem; padding-bottom: 1.5rem;}
    .hero {
        background: linear-gradient(135deg, #fff4eb 0%, #fffaf6 100%);
        border: 1px solid #f3c9a5;
        border-radius: 16px;
        padding: 18px 20px;
        margin-bottom: 16px;
    }
    .hero h1 {margin: 0; color: #a85400; font-size: 1.55rem;}
    .hero p {margin: 8px 0 0 0; color: #714221;}
    .credit {
        margin-top: 8px;
        display: inline-block;
        background: #fff0e1;
        color: #a85400;
        border: 1px solid #f3c9a5;
        border-radius: 999px;
        padding: 4px 10px;
        font-size: .9rem;
        font-weight: 600;
    }
    .card {
        background: #fffaf6;
        border: 1px solid #f0d6bf;
        border-radius: 14px;
        padding: 14px 16px;
        margin-bottom: 12px;
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
    </style>
    """,
    unsafe_allow_html=True,
)

st.markdown(
    """
    <div class="hero">
        <h1>Validador pagado vs contabilizado</h1>
        <p>CC-nóminas = pagado | PCP0 = contabilizado | Selección por año y Período para nómina</p>
        <div class="credit">Creado por Andrés Huérfano Dávila - Nómina JMC</div>
    </div>
    """,
    unsafe_allow_html=True,
)

MAX_EXCEL_ROWS = 1_048_576
SUPPORTED_EXTENSIONS = {".txt", ".csv", ".xls", ".xlsx", ".xlsb", ".ods"}

ALIASES = {
    "Número de personal": ["número de personal", "numero de personal", "pernr", "sap", "n pers", "nº pers", "n° pers", "numero personal"],
    "Nombre del empleado o candidato": ["nombre del empleado o candidato", "nombre del empleado", "nombre empleado", "nombre"],
    "CC-nómina": ["cc-nómina", "cc-nomina", "cc nómina", "cc nomina", "wagetype", "tipo salario", "tipo de salario", "cc-n.", "cc n"],
    "Texto expl.CC-nómina": ["texto expl.cc-nómina", "texto expl.cc-nomina", "texto expl cc nómina", "texto expl cc nomina", "texto explicativo cc nómina", "texto explicativo cc nomina", "texto expl", "descripcion cc nómina", "descripcion cc nomina"],
    "Importe": ["importe", "valor", "monto", "importe total"],
    "Fecha de pago": ["fecha de pago", "fecha pago"],
    "Período para nómina": ["período para nómina", "periodo para nomina", "periodo para nómina", "período para nomina", "per.para"],
    "Período En cálc.nóm.": ["período en cálc.nóm.", "período en cálc nom", "periodo en calc nom", "periodo en cálculo nomina", "periodo en calculo nomina", "periodo en"],
    "Tipo": ["tipo"],
}

EXACT_PREFERRED = {
    "Número de personal": ["Número de personal"],
    "Nombre del empleado o candidato": ["Nombre del empleado o candidato"],
    "CC-nómina": ["CC-nómina"],
    "Texto expl.CC-nómina": ["Texto expl.CC-nómina"],
    "Importe": ["Importe"],
    "Fecha de pago": ["Fecha de pago"],
    "Período para nómina": ["Período para nómina"],
    "Período En cálc.nóm.": ["Período En cálc.nóm."],
    "Tipo": ["Tipo"],
}


def init_state() -> None:
    defaults = {
        "cc_raw": None,
        "cc_log": pd.DataFrame(columns=["archivo", "estado", "registros", "detalle"]),
        "pcp0_raw": None,
        "pcp0_log": pd.DataFrame(columns=["archivo", "estado", "registros", "detalle"]),
        "conceptos_df": None,
        "cc_files_hash": None,
        "pcp0_files_hash": None,
        "conceptos_hash": None,
        "selected_year": None,
        "selected_periods": [],
        "cc_resumen": None,
        "cc_detalle": None,
        "pcp0_resumen": None,
        "pcp0_detalle": None,
        "comparativo_resumen": None,
        "comparativo_detalle": None,
        "comparativo_periodo": None,
        "parametros_df": None,
    }
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v


init_state()


def normalize_text(value: object) -> str:
    if value is None:
        return ""
    text = str(value).strip().lower()
    text = unicodedata.normalize("NFKD", text)
    text = "".join(ch for ch in text if not unicodedata.combining(ch))
    text = re.sub(r"[\n\r\t]+", " ", text)
    text = re.sub(r"[^a-z0-9]+", " ", text)
    return re.sub(r"\s+", " ", text).strip()


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

    rem = result.isna() & text.notna()
    if rem.any():
        result.loc[rem] = pd.to_datetime(text[rem], errors="coerce")
    return result


def find_column(df: pd.DataFrame, canonical_name: str) -> Optional[str]:
    preferred = EXACT_PREFERRED.get(canonical_name, [])
    normalized_preferred = {normalize_text(x) for x in preferred}
    for col in df.columns:
        if normalize_text(col) in normalized_preferred:
            return col
    canonical_norm = normalize_text(canonical_name)
    for col in df.columns:
        if normalize_text(col) == canonical_norm:
            return col
    aliases = [normalize_text(a) for a in ALIASES.get(canonical_name, [])]
    alias_set = set(aliases)
    for col in df.columns:
        if normalize_text(col) in alias_set:
            return col
    for col in df.columns:
        ncol = normalize_text(col)
        if canonical_name == "Período para nómina":
            if ncol == "periodo para nomina":
                return col
        elif canonical_name == "Período En cálc.nóm.":
            if ncol == "periodo en calc nom":
                return col
        else:
            if any(alias in ncol for alias in aliases):
                return col
    return None


def add_canonical_columns(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    for canonical in ALIASES:
        source = find_column(df, canonical)
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


def get_uploads_hash(files: Sequence) -> Optional[str]:
    if not files:
        return None
    h = hashlib.md5()
    for f in files:
        data = f.getvalue()
        h.update(f.name.encode("utf-8", errors="ignore"))
        h.update(str(len(data)).encode())
        h.update(data[:4096])
    return h.hexdigest()


def get_upload_hash(uploaded_file) -> Optional[str]:
    if uploaded_file is None:
        return None
    data = uploaded_file.getvalue()
    h = hashlib.md5()
    h.update(uploaded_file.name.encode("utf-8", errors="ignore"))
    h.update(str(len(data)).encode())
    h.update(data[:4096])
    return h.hexdigest()


def read_text_bytes(data: bytes, filename: str) -> pd.DataFrame:
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


def read_excel_bytes(data: bytes, filename: str) -> pd.DataFrame:
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


def read_uploaded_table(uploaded_file) -> pd.DataFrame:
    filename = uploaded_file.name
    data = uploaded_file.getvalue()
    ext = Path(filename).suffix.lower()

    if ext in {".xlsx", ".xls", ".xlsb", ".ods"}:
        df = read_excel_bytes(data, filename)
    elif ext in {".txt", ".csv"}:
        df = read_text_bytes(data, filename)
        df["__archivo_origen__"] = filename
        df["__hoja_origen__"] = "TXT/CSV"
    else:
        raise ValueError(f"Formato no soportado: {filename}")

    return add_canonical_columns(df)


def combine_uploaded_files(uploaded_files, label: str) -> Tuple[pd.DataFrame, pd.DataFrame]:
    frames = []
    logs = []
    progress = st.progress(0, text=f"Leyendo archivos de {label}...")
    for idx, file in enumerate(uploaded_files, start=1):
        try:
            df = read_uploaded_table(file)
            frames.append(df)
            logs.append({"archivo": file.name, "estado": "OK", "registros": int(len(df)), "detalle": "Archivo leído correctamente"})
        except Exception as exc:
            logs.append({"archivo": file.name, "estado": "ERROR", "registros": 0, "detalle": str(exc)})
        progress.progress(idx / len(uploaded_files), text=f"Leyendo archivos de {label}... {idx}/{len(uploaded_files)}")
    progress.empty()

    if not frames:
        raise ValueError(f"No se pudo leer ningún archivo de {label}")
    return pd.concat(frames, ignore_index=True, sort=False), pd.DataFrame(logs)


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


def load_concepts_from_upload(uploaded_file) -> pd.DataFrame:
    df = read_uploaded_table(uploaded_file)
    for col in ["CC-nómina", "Texto expl.CC-nómina", "Tipo"]:
        if col not in df.columns:
            raise ValueError(f"El archivo de conceptos no tiene la columna {col}")

    concepts = df[["CC-nómina", "Texto expl.CC-nómina", "Tipo"]].copy()
    concepts["CC-nómina"] = concepts["CC-nómina"].astype(str).replace("nan", "").str.strip()
    concepts["Texto expl.CC-nómina"] = concepts["Texto expl.CC-nómina"].astype(str).replace("nan", "").str.strip()
    concepts["Tipo"] = concepts["Tipo"].map(normalize_tipo)
    concepts = concepts[concepts["CC-nómina"].ne("")].drop_duplicates(subset=["CC-nómina"], keep="first").reset_index(drop=True)

    invalid = concepts[concepts["Tipo"].isna()]
    if not invalid.empty:
        sample = invalid[["CC-nómina", "Texto expl.CC-nómina"]].head(10)
        raise ValueError("Hay conceptos con Tipo inválido. Usa únicamente: Salarial, Beneficio o No aplica.\n" + sample.to_string(index=False))
    return concepts


def require_columns(df: pd.DataFrame, required: Sequence[str], label: str) -> None:
    missing = [c for c in required if c not in df.columns]
    if missing:
        raise ValueError(f"En {label} faltan columnas obligatorias: {', '.join(missing)}")


def get_available_years(df_cc: pd.DataFrame) -> List[int]:
    if "Fecha de pago_año" not in df_cc.columns:
        return []
    years = [int(x) for x in df_cc["Fecha de pago_año"].dropna().unique().tolist()]
    return sorted(years)


def get_periods_for_year(df_cc: pd.DataFrame, year: int) -> List[str]:
    work = df_cc[df_cc["Fecha de pago"].dt.year.eq(year)].copy()
    if "Período para nómina" not in work.columns:
        return []
    periods = [str(x).strip() for x in work["Período para nómina"].dropna().astype(str).unique().tolist() if str(x).strip()]
    return sorted(periods)


def filter_cc(df_cc: pd.DataFrame, year: int, periods: Sequence[str]) -> pd.DataFrame:
    work = df_cc[df_cc["Fecha de pago"].dt.year.eq(year)].copy()
    work = work[work["Período para nómina"].astype(str).isin([str(p) for p in periods])].copy()
    return work


def filter_pcp0(df_pcp0: pd.DataFrame, periods: Sequence[str]) -> pd.DataFrame:
    periods_set = {str(p).strip() for p in periods if str(p).strip()}
    if not periods_set:
        return df_pcp0.iloc[0:0].copy()

    mask = pd.Series(False, index=df_pcp0.index)
    if "Período para nómina" in df_pcp0.columns:
        mask = mask | df_pcp0["Período para nómina"].astype(str).isin(periods_set)
    if "Período En cálc.nóm." in df_pcp0.columns:
        mask = mask | df_pcp0["Período En cálc.nóm."].astype(str).isin(periods_set)
    return df_pcp0[mask].copy()


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

    resumen = (
        merged.groupby(["Número de personal", "Nombre del empleado o candidato"], dropna=False, as_index=False)[["Salariales", "Beneficios adicionales"]]
        .sum()
        .sort_values(["Número de personal", "Nombre del empleado o candidato"])
        .reset_index(drop=True)
    )
    resumen["Importe total"] = resumen["Salariales"] + resumen["Beneficios adicionales"]

    detalle = (
        merged.groupby(["Número de personal", "Nombre del empleado o candidato", "CC-nómina", "Texto expl.CC-nómina"], dropna=False, as_index=False)[["Salariales", "Beneficios adicionales"]]
        .sum()
        .sort_values(["Número de personal", "CC-nómina"])
        .reset_index(drop=True)
    )
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

    resumen = (
        merged.groupby(["Número de personal"], dropna=False, as_index=False)[["Salariales", "Beneficios adicionales"]]
        .sum()
        .sort_values(["Número de personal"])
        .reset_index(drop=True)
    )
    resumen["Importe total"] = resumen["Salariales"] + resumen["Beneficios adicionales"]

    detalle = (
        merged.groupby(["Número de personal", "CC-nómina", "Texto expl.CC-nómina"], dropna=False, as_index=False)[["Salariales", "Beneficios adicionales"]]
        .sum()
        .sort_values(["Número de personal", "CC-nómina"])
        .reset_index(drop=True)
    )
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


def build_period_summary(df_cc: pd.DataFrame, df_pcp0: pd.DataFrame, concepts: pd.DataFrame) -> pd.DataFrame:
    cc_period = df_cc.merge(concepts[["CC-nómina", "Tipo"]], on="CC-nómina", how="left", suffixes=("", "_map"))
    cc_period["Tipo_final"] = cc_period["Tipo_map"] if "Tipo_map" in cc_period.columns else cc_period["Tipo"]
    cc_period = cc_period[cc_period["Tipo_final"].isin(["Salarial", "Beneficio"])].copy()
    cc_period["Importe"] = cc_period["Importe"].map(parse_number)
    cc_period["Periodo"] = cc_period["Período para nómina"].astype(str)
    cc_by_period = cc_period.groupby("Periodo", as_index=False)["Importe"].sum().rename(columns={"Importe": "Pagado CC"})

    pcp0_period = df_pcp0.merge(concepts[["CC-nómina", "Tipo"]], on="CC-nómina", how="left", suffixes=("", "_map"))
    pcp0_period["Tipo_final"] = pcp0_period["Tipo_map"] if "Tipo_map" in pcp0_period.columns else pcp0_period["Tipo"]
    pcp0_period = pcp0_period[pcp0_period["Tipo_final"].isin(["Salarial", "Beneficio"])].copy()
    pcp0_period["Importe"] = pcp0_period["Importe"].map(parse_number)

    parts = []
    if "Período para nómina" in pcp0_period.columns:
        parts.append(
            pcp0_period.assign(Periodo=pcp0_period["Período para nómina"].astype(str))
            .groupby("Periodo", as_index=False)["Importe"]
            .sum()
            .rename(columns={"Importe": "Contabilizado PCP0"})
        )
    if "Período En cálc.nóm." in pcp0_period.columns:
        parts.append(
            pcp0_period.assign(Periodo=pcp0_period["Período En cálc.nóm."].astype(str))
            .groupby("Periodo", as_index=False)["Importe"]
            .sum()
            .rename(columns={"Importe": "Contabilizado PCP0"})
        )

    pcp0_by_period = pd.concat(parts, ignore_index=True) if parts else pd.DataFrame(columns=["Periodo", "Contabilizado PCP0"])
    if not pcp0_by_period.empty:
        pcp0_by_period = pcp0_by_period.groupby("Periodo", as_index=False)["Contabilizado PCP0"].sum()

    resumen = cc_by_period.merge(pcp0_by_period, on="Periodo", how="outer").fillna(0.0)
    resumen["Diferencia"] = resumen["Pagado CC"] - resumen["Contabilizado PCP0"]
    return resumen.sort_values("Periodo").reset_index(drop=True)


def chunk_dataframe(df: pd.DataFrame, max_rows: int = MAX_EXCEL_ROWS - 1) -> List[pd.DataFrame]:
    if len(df) <= max_rows:
        return [df]
    parts = math.ceil(len(df) / max_rows)
    return [df.iloc[i * max_rows : (i + 1) * max_rows].copy() for i in range(parts)]


def write_sheet_with_format(writer: pd.ExcelWriter, sheet_name: str, df: pd.DataFrame) -> None:
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
                write_sheet_with_format(writer, base_name[:31], parts[0])
            else:
                for i, part in enumerate(parts, start=1):
                    write_sheet_with_format(writer, f"{base_name}_{i}"[:31], part)
    output.seek(0)
    return output.getvalue()


def to_csv_bytes(df: pd.DataFrame) -> bytes:
    return df.to_csv(index=False, encoding="utf-8-sig").encode("utf-8-sig")


def build_zip_bytes(files: Dict[str, bytes]) -> bytes:
    output = io.BytesIO()
    with zipfile.ZipFile(output, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        for name, data in files.items():
            zf.writestr(name, data)
    output.seek(0)
    return output.getvalue()


def render_log(df_log: pd.DataFrame, title: str, key: str) -> None:
    st.markdown(f"**{title}**")
    if df_log is None or df_log.empty:
        st.info("Aún no hay registros.")
        return
    c1, c2 = st.columns(2)
    c1.metric("OK", int((df_log["estado"] == "OK").sum()) if "estado" in df_log.columns else 0)
    c2.metric("Con error", int((df_log["estado"] == "ERROR").sum()) if "estado" in df_log.columns else 0)
    st.dataframe(df_log, width="stretch", height=220)
    st.download_button(
        f"Descargar log {title}",
        data=to_csv_bytes(df_log),
        file_name=f"{key}.csv",
        mime="text/csv",
        key=f"dl_{key}",
    )


def show_metric_cards(df: pd.DataFrame, total_col: str, label: str) -> None:
    c1, c2, c3 = st.columns(3)
    c1.metric(f"{label} registros", f"{len(df):,}".replace(",", "."))
    c2.metric(f"{label} SAP únicos", f"{df['Número de personal'].nunique():,}".replace(",", "."))
    c3.metric(f"{label} total", f"{df[total_col].sum():,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))


left, right = st.columns([1.2, 1], gap="large")

with left:
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.subheader("1) Cargar CC-nóminas")
    cc_files = st.file_uploader(
        "Carga archivos de CC-nóminas / acumulados",
        type=["txt", "csv", "xls", "xlsx", "xlsb", "ods"],
        accept_multiple_files=True,
        key="cc_files_uploader",
    )

    if st.button("Leer CC-nóminas", type="primary", key="btn_read_cc"):
        if not cc_files:
            st.warning("Carga al menos un archivo de CC-nóminas.")
        else:
            current_hash = get_uploads_hash(cc_files)
            try:
                cc_raw, cc_log = combine_uploaded_files(cc_files, "CC-nóminas")
                require_columns(
                    cc_raw,
                    ["Número de personal", "Nombre del empleado o candidato", "CC-nómina", "Importe", "Fecha de pago", "Período para nómina"],
                    "CC-nóminas",
                )
                st.session_state["cc_raw"] = cc_raw
                st.session_state["cc_log"] = cc_log
                st.session_state["cc_files_hash"] = current_hash
                st.session_state["selected_year"] = None
                st.session_state["selected_periods"] = []
                st.success("CC-nóminas cargadas correctamente.")
            except Exception as exc:
                st.error(f"No fue posible cargar CC-nóminas: {exc}")

    cc_raw = st.session_state.get("cc_raw")
    if isinstance(cc_raw, pd.DataFrame) and not cc_raw.empty:
        years = get_available_years(cc_raw)
        preview_cols = [c for c in ["__archivo_origen__", "__hoja_origen__", "Número de personal", "Nombre del empleado o candidato", "CC-nómina", "Importe", "Fecha de pago_display", "Período para nómina"] if c in cc_raw.columns]
        preview = cc_raw[preview_cols].copy()
        if "Fecha de pago_display" in preview.columns:
            preview = preview.rename(columns={"Fecha de pago_display": "Fecha de pago"})

        st.markdown("**Vista previa CC-nóminas**")
        st.dataframe(preview.head(200), width="stretch", height=260)

        with st.form("form_selection_cc"):
            year_default = st.session_state["selected_year"] if st.session_state["selected_year"] in years else (years[0] if years else None)
            selected_year = st.selectbox("Año según Fecha de pago", options=years, index=years.index(year_default) if year_default in years else 0)
            period_options = get_periods_for_year(cc_raw, selected_year)
            current_periods = [p for p in st.session_state.get("selected_periods", []) if p in period_options]
            selected_periods = st.multiselect(
                "Período para nómina",
                options=period_options,
                default=current_periods if current_periods else period_options,
            )
            apply_selection = st.form_submit_button("Guardar selección de año y períodos", type="primary")

        if apply_selection:
            st.session_state["selected_year"] = selected_year
            st.session_state["selected_periods"] = selected_periods
            st.success("Selección guardada.")

        if st.session_state["selected_year"] is not None:
            st.info(
                f"Año seleccionado: {st.session_state['selected_year']} | "
                f"Períodos seleccionados: {', '.join(st.session_state['selected_periods']) if st.session_state['selected_periods'] else 'Ninguno'}"
            )

    render_log(st.session_state.get("cc_log"), "Log CC-nóminas", "log_cc")
    st.markdown("</div>", unsafe_allow_html=True)

    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.subheader("2) Cargar conceptos")
    concept_file = st.file_uploader(
        "Archivo de conceptos con columnas: CC-nómina, Texto expl.CC-nómina, Tipo",
        type=["txt", "csv", "xls", "xlsx", "xlsb", "ods"],
        accept_multiple_files=False,
        key="concepts_uploader",
    )
    if st.button("Leer conceptos", key="btn_read_concepts"):
        if concept_file is None:
            st.warning("Carga el archivo de conceptos.")
        else:
            try:
                concepts = load_concepts_from_upload(concept_file)
                st.session_state["conceptos_df"] = concepts
                st.session_state["conceptos_hash"] = get_upload_hash(concept_file)
                st.success("Conceptos cargados correctamente.")
            except Exception as exc:
                st.error(f"No fue posible cargar conceptos: {exc}")

    concepts_df = st.session_state.get("conceptos_df")
    if isinstance(concepts_df, pd.DataFrame) and not concepts_df.empty:
        st.dataframe(concepts_df.head(200), width="stretch", height=220)
        st.caption(f"Conceptos válidos cargados: {len(concepts_df):,}".replace(",", "."))
    st.markdown("</div>", unsafe_allow_html=True)

with right:
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.subheader("3) Cargar PCP0 / contabilizaciones")
    pcp0_files = st.file_uploader(
        "Carga archivos PCP0",
        type=["txt", "csv", "xls", "xlsx", "xlsb", "ods"],
        accept_multiple_files=True,
        key="pcp0_files_uploader",
    )

    if st.button("Leer PCP0", type="primary", key="btn_read_pcp0"):
        if not pcp0_files:
            st.warning("Carga al menos un archivo de PCP0.")
        else:
            current_hash = get_uploads_hash(pcp0_files)
            try:
                pcp0_raw, pcp0_log = combine_uploaded_files(pcp0_files, "PCP0")
                require_columns(
                    pcp0_raw,
                    ["Número de personal", "CC-nómina", "Importe"],
                    "PCP0",
                )
                st.session_state["pcp0_raw"] = pcp0_raw
                st.session_state["pcp0_log"] = pcp0_log
                st.session_state["pcp0_files_hash"] = current_hash
                st.success("PCP0 cargado correctamente.")
            except Exception as exc:
                st.error(f"No fue posible cargar PCP0: {exc}")

    pcp0_raw = st.session_state.get("pcp0_raw")
    if isinstance(pcp0_raw, pd.DataFrame) and not pcp0_raw.empty:
        preview_cols = [c for c in ["__archivo_origen__", "__hoja_origen__", "Número de personal", "CC-nómina", "Importe", "Período para nómina", "Período En cálc.nóm."] if c in pcp0_raw.columns]
        st.markdown("**Vista previa PCP0**")
        st.dataframe(pcp0_raw[preview_cols].head(200), width="stretch", height=260)

    render_log(st.session_state.get("pcp0_log"), "Log PCP0", "log_pcp0")
    st.markdown("</div>", unsafe_allow_html=True)

    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.subheader("4) Generar validación")
    st.caption("La comparación usa: CC-nóminas = pagado | PCP0 = contabilizado")

    if st.button("Procesar validación", type="primary", key="btn_process_all"):
        cc_raw = st.session_state.get("cc_raw")
        pcp0_raw = st.session_state.get("pcp0_raw")
        concepts = st.session_state.get("conceptos_df")
        year = st.session_state.get("selected_year")
        periods = st.session_state.get("selected_periods", [])

        if not isinstance(cc_raw, pd.DataFrame) or cc_raw.empty:
            st.warning("Primero carga CC-nóminas.")
        elif not isinstance(pcp0_raw, pd.DataFrame) or pcp0_raw.empty:
            st.warning("Primero carga PCP0.")
        elif not isinstance(concepts, pd.DataFrame) or concepts.empty:
            st.warning("Primero carga conceptos.")
        elif year is None:
            st.warning("Primero selecciona el año.")
        elif not periods:
            st.warning("Selecciona al menos un Período para nómina.")
        else:
            try:
                progress = st.progress(0, text="Filtrando CC-nóminas...")
                cc_filtered = filter_cc(cc_raw, year, periods)
                progress.progress(0.2, text="Filtrando PCP0...")

                pcp0_filtered = filter_pcp0(pcp0_raw, periods)
                progress.progress(0.4, text="Procesando CC-nóminas...")

                cc_resumen, cc_detalle = process_cc(cc_filtered, concepts)
                progress.progress(0.6, text="Procesando PCP0...")

                pcp0_resumen, pcp0_detalle = process_pcp0(pcp0_filtered, concepts)
                progress.progress(0.8, text="Generando comparativos...")

                comparativo_resumen, comparativo_detalle = compare_cc_vs_pcp0(
                    cc_resumen, cc_detalle, pcp0_resumen, pcp0_detalle
                )
                comparativo_periodo = build_period_summary(cc_filtered, pcp0_filtered, concepts)

                parametros_df = pd.DataFrame(
                    {
                        "Parametro": ["Año seleccionado", "Períodos seleccionados"],
                        "Valor": [str(year), ", ".join(periods)],
                    }
                )

                st.session_state["cc_resumen"] = cc_resumen
                st.session_state["cc_detalle"] = cc_detalle
                st.session_state["pcp0_resumen"] = pcp0_resumen
                st.session_state["pcp0_detalle"] = pcp0_detalle
                st.session_state["comparativo_resumen"] = comparativo_resumen
                st.session_state["comparativo_detalle"] = comparativo_detalle
                st.session_state["comparativo_periodo"] = comparativo_periodo
                st.session_state["parametros_df"] = parametros_df

                progress.progress(1.0, text="Proceso terminado")
                progress.empty()
                st.success("Validación generada correctamente.")
            except Exception as exc:
                st.error(f"No fue posible generar la validación: {exc}")
    st.markdown("</div>", unsafe_allow_html=True)

results_tabs = st.tabs(["Resumen", "Detalle", "Descargas"])

with results_tabs[0]:
    comp = st.session_state.get("comparativo_resumen")
    cc_res = st.session_state.get("cc_resumen")
    pcp0_res = st.session_state.get("pcp0_resumen")
    comp_periodo = st.session_state.get("comparativo_periodo")

    if isinstance(comp, pd.DataFrame):
        c1, c2, c3 = st.columns(3)
        c1.metric("SAP comparados", f"{comp['Número de personal'].nunique():,}".replace(",", "."))
        c2.metric("SAP con diferencia", f"{(comp['Estado'] != 'OK').sum():,}".replace(",", "."))
        c3.metric("Diferencia total", f"{comp['Diferencia'].sum():,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))

    if isinstance(cc_res, pd.DataFrame):
        st.markdown("### Pagado en CC-nóminas")
        show_metric_cards(cc_res, "Importe total", "CC")
        st.dataframe(cc_res.head(300), width="stretch", height=280)

    if isinstance(pcp0_res, pd.DataFrame):
        st.markdown("### Contabilizado en PCP0")
        show_metric_cards(pcp0_res, "Importe total", "PCP0")
        st.dataframe(pcp0_res.head(300), width="stretch", height=280)

    if isinstance(comp, pd.DataFrame):
        st.markdown("### Comparativo por empleado")
        st.dataframe(comp.head(300), width="stretch", height=320)

    if isinstance(comp_periodo, pd.DataFrame):
        st.markdown("### Comparativo por período")
        st.dataframe(comp_periodo.head(300), width="stretch", height=280)

with results_tabs[1]:
    cc_det = st.session_state.get("cc_detalle")
    pcp0_det = st.session_state.get("pcp0_detalle")
    comp_det = st.session_state.get("comparativo_detalle")

    if isinstance(cc_det, pd.DataFrame):
        st.markdown("### Detalle CC-nóminas")
        st.dataframe(cc_det.head(300), width="stretch", height=300)

    if isinstance(pcp0_det, pd.DataFrame):
        st.markdown("### Detalle PCP0")
        st.dataframe(pcp0_det.head(300), width="stretch", height=300)

    if isinstance(comp_det, pd.DataFrame):
        st.markdown("### Detalle de diferencias")
        st.dataframe(comp_det.head(300), width="stretch", height=320)

with results_tabs[2]:
    cc_res = st.session_state.get("cc_resumen")
    cc_det = st.session_state.get("cc_detalle")
    pcp0_res = st.session_state.get("pcp0_resumen")
    pcp0_det = st.session_state.get("pcp0_detalle")
    comp = st.session_state.get("comparativo_resumen")
    comp_det = st.session_state.get("comparativo_detalle")
    comp_periodo = st.session_state.get("comparativo_periodo")
    params = st.session_state.get("parametros_df")
    concepts = st.session_state.get("conceptos_df")
    cc_log = st.session_state.get("cc_log")
    pcp0_log = st.session_state.get("pcp0_log")

    if all(isinstance(x, pd.DataFrame) for x in [cc_res, cc_det, pcp0_res, pcp0_det, comp, comp_det, comp_periodo, params, concepts]):
        excel_bytes = to_excel_bytes(
            {
                "Parametros": params,
                "Pagado_CC": cc_res,
                "Detalle_CC": cc_det,
                "Contabilizado_PCP0": pcp0_res,
                "Detalle_PCP0": pcp0_det,
                "Comparativo_Empleado": comp,
                "Detalle_Diferencias": comp_det,
                "Comparativo_Periodo": comp_periodo,
                "Conceptos_Utilizados": concepts,
                "Log_CC": cc_log if isinstance(cc_log, pd.DataFrame) else pd.DataFrame(),
                "Log_PCP0": pcp0_log if isinstance(pcp0_log, pd.DataFrame) else pd.DataFrame(),
            }
        )

        zip_bytes = build_zip_bytes(
            {
                "resultado_medios_magneticos.xlsx": excel_bytes,
                "parametros.csv": to_csv_bytes(params),
                "pagado_cc.csv": to_csv_bytes(cc_res),
                "detalle_cc.csv": to_csv_bytes(cc_det),
                "contabilizado_pcp0.csv": to_csv_bytes(pcp0_res),
                "detalle_pcp0.csv": to_csv_bytes(pcp0_det),
                "comparativo_empleado.csv": to_csv_bytes(comp),
                "detalle_diferencias.csv": to_csv_bytes(comp_det),
                "comparativo_periodo.csv": to_csv_bytes(comp_periodo),
                "conceptos_utilizados.csv": to_csv_bytes(concepts),
                "log_cc.csv": to_csv_bytes(cc_log if isinstance(cc_log, pd.DataFrame) else pd.DataFrame()),
                "log_pcp0.csv": to_csv_bytes(pcp0_log if isinstance(pcp0_log, pd.DataFrame) else pd.DataFrame()),
            }
        )

        c1, c2 = st.columns(2)
        c1.download_button(
            "Descargar Excel",
            data=excel_bytes,
            file_name="resultado_medios_magneticos.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="dl_excel_final",
        )
        c2.download_button(
            "Descargar ZIP",
            data=zip_bytes,
            file_name="resultado_medios_magneticos.zip",
            mime="application/zip",
            key="dl_zip_final",
        )
    else:
        st.info("Cuando proceses la validación, aquí aparecerán las descargas.")

st.markdown('<div class="footer-credit">Creado por Andrés Huérfano Dávila - Nómina JMC</div>', unsafe_allow_html=True)
