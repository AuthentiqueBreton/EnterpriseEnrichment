import io
import json
import math
import os
import re
import time
import unicodedata
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

import pandas as pd
import requests
import streamlit as st


st.set_page_config(layout="wide", page_title="Validation fournisseurs")

API_URL = "https://recherche-entreprises.api.gouv.fr/search"
MIN_DELAY_BETWEEN_CALLS = 0.2

WORK_DIR = Path("workdir")
WORK_DIR.mkdir(exist_ok=True)

AUTOSAVE_XLSX = WORK_DIR / "fournisseurs_autosave.xlsx"
AUTOSAVE_META = WORK_DIR / "fournisseurs_autosave_meta.json"

ENRICHMENT_COLUMNS = [
    "api_match_status",
    "api_query",
    "api_score_adresse",
    "api_nom_reel",
    "api_adresse_trouvee",
    "api_siren",
    "api_siret",
    "api_intracom",
    "api_tel_1",
    "api_tel_2",
]


ABBREVIATIONS = {
    "AVENUE": "AV",
    "BOULEVARD": "BD",
    "ROUTE": "RTE",
    "CHEMIN": "CHE",
    "IMPASSE": "IMP",
    "ALLEE": "ALL",
    "LOTISSEMENT": "LOT",
    "ZA": "ZA",
    "ZI": "ZI",
    "ZAC": "ZAC",
    "SAINT": "ST",
    "SAINTE": "STE",
}

STOP_WORDS = {"BP", "CEDEX", "CS"}


def is_nan(value: Any) -> bool:
    return value is None or (isinstance(value, float) and math.isnan(value))


def clean_str(value: Any) -> str:
    if is_nan(value):
        return ""
    return str(value).strip()


def strip_accents(text: str) -> str:
    return "".join(
        c for c in unicodedata.normalize("NFD", text)
        if unicodedata.category(c) != "Mn"
    )


def normalize_text(text: str) -> str:
    text = strip_accents(clean_str(text).upper())
    text = text.replace("'", " ").replace("’", " ")
    text = re.sub(r"[^A-Z0-9\s]", " ", text)
    text = re.sub(r"\s+", " ", text).strip()
    return text


def normalize_address(text: str) -> str:
    text = normalize_text(text)
    text = re.sub(r"\b(BP|CS)\s*\d+\b", " ", text)

    words = []
    for word in text.split():
        if word in STOP_WORDS:
            continue
        words.append(ABBREVIATIONS.get(word, word))

    return re.sub(r"\s+", " ", " ".join(words)).strip()


def extract_cp_city(cp_ville: str) -> Tuple[str, str]:
    cp_ville = normalize_text(cp_ville)
    match = re.match(r"^(\d{5})\s+(.*)$", cp_ville)
    if not match:
        return "", cp_ville

    cp = match.group(1)
    ville = match.group(2).strip()
    ville = re.sub(r"\bCEDEX\b\s*\d*", "", ville).strip()
    ville = re.sub(r"\s+", " ", ville)
    return cp, ville


def similarity(a: str, b: str) -> float:
    if not a or not b:
        return 0.0
    from difflib import SequenceMatcher
    return SequenceMatcher(None, a, b).ratio()


def build_input_address(row: pd.Series) -> str:
    rue = clean_str(row.get("Rue complète"))
    cp_ville = clean_str(row.get("CP-Ville"))
    return " ".join(x for x in [rue, cp_ville] if x).strip()


def is_france_or_unspecified(country: Any) -> bool:
    country = clean_str(country)
    if not country:
        return True

    norm = normalize_text(country)
    return norm in {"FRANCE", "FR", "FRA"}


def get_nested(d: Dict[str, Any], *keys, default=None):
    cur = d
    for key in keys:
        if not isinstance(cur, dict):
            return default
        cur = cur.get(key)
        if cur is None:
            return default
    return cur


def compute_french_vat_from_siren(siren: str) -> Optional[str]:
    siren = re.sub(r"\D", "", clean_str(siren))
    if len(siren) != 9:
        return None

    key = (12 + 3 * (int(siren) % 97)) % 97
    return f"FR{key:02d}{siren}"


class RechercheEntrepriseClient:
    def __init__(self):
        self.session = requests.Session()
        self._last_call_ts = 0.0

    def _throttle(self):
        elapsed = time.time() - self._last_call_ts
        if elapsed < MIN_DELAY_BETWEEN_CALLS:
            time.sleep(MIN_DELAY_BETWEEN_CALLS - elapsed)

    def search(
        self,
        query: str,
        per_page: int = 10,
        page: int = 1,
        max_retries: int = 5,
    ) -> Optional[Dict[str, Any]]:
        params = {
            "q": query,
            "per_page": per_page,
            "page": page,
        }

        for attempt in range(max_retries + 1):
            try:
                self._throttle()
                response = self.session.get(API_URL, params=params, timeout=10)
                self._last_call_ts = time.time()

                if response.status_code == 200:
                    return response.json()

                if response.status_code == 429:
                    retry_after = response.headers.get("Retry-After")
                    wait_s = int(retry_after) if retry_after and retry_after.isdigit() else min(2 ** attempt, 30)
                    time.sleep(wait_s)
                    continue

                if 500 <= response.status_code < 600:
                    time.sleep(min(2 ** attempt, 10))
                    continue

                return None

            except requests.RequestException:
                time.sleep(min(2 ** attempt, 10))
                continue

        return None


def extract_result_address(result: Dict[str, Any]) -> str:
    siege = result.get("siege") or {}

    adresse = clean_str(siege.get("adresse"))
    cp = clean_str(siege.get("code_postal"))
    ville = clean_str(siege.get("libelle_commune")) or clean_str(siege.get("commune"))

    if adresse or cp or ville:
        return " ".join(x for x in [adresse, cp, ville] if x).strip()

    return (
        clean_str(result.get("adresse"))
        or clean_str(result.get("adresse_complete"))
        or clean_str(result.get("full_address"))
    )


def extract_company_name(result: Dict[str, Any]) -> str:
    return (
        clean_str(result.get("nom_complet"))
        or clean_str(result.get("nom_raison_sociale"))
        or clean_str(result.get("denomination"))
        or clean_str(result.get("nom"))
    )


def extract_siren(result: Dict[str, Any]) -> str:
    return clean_str(result.get("siren"))


def extract_siret(result: Dict[str, Any]) -> str:
    return clean_str(get_nested(result, "siege", "siret")) or clean_str(result.get("siret"))


def result_country_is_france(result: Dict[str, Any]) -> bool:
    siege = result.get("siege") or {}
    country = clean_str(siege.get("libelle_pays_etranger")) or clean_str(result.get("pays"))
    if not country:
        return True
    return is_france_or_unspecified(country)


def address_score(input_address: str, result_address: str) -> float:
    input_norm = normalize_address(input_address)
    result_norm = normalize_address(result_address)

    score = similarity(input_norm, result_norm)

    input_cp, input_city = extract_cp_city(input_address)
    result_cp, result_city = extract_cp_city(result_address)

    if input_cp and result_cp:
        if input_cp == result_cp:
            score += 0.15
        else:
            score -= 0.25

    if input_city and result_city:
        city_sim = similarity(normalize_address(input_city), normalize_address(result_city))
        if city_sim >= 0.9:
            score += 0.10
        elif city_sim < 0.5:
            score -= 0.10

    return max(0.0, min(score, 1.0))


def get_api_matches_for_row(row: pd.Series, client: RechercheEntrepriseClient, per_page: int = 10) -> List[Dict[str, Any]]:
    if not is_france_or_unspecified(row.get("Pays")):
        return []

    query = build_input_address(row)
    if not query:
        return []

    data = client.search(query, per_page=per_page, page=1)
    if not data:
        return []

    results = data.get("results", []) or []

    output = []
    for result in results:
        if not result_country_is_france(result):
            continue

        siren = extract_siren(result)
        siret = extract_siret(result)
        addr = extract_result_address(result)

        output.append({
            "nom": extract_company_name(result),
            "adresse": addr,
            "siren": siren,
            "siret": siret,
            "intracom": compute_french_vat_from_siren(siren),
            "score": address_score(query, addr),
            "raw": result,
        })

    return output


def atomic_write_json(path: Path, data: dict) -> None:
    tmp = path.with_suffix(path.suffix + ".tmp")
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    os.replace(tmp, path)


def atomic_write_excel(df: pd.DataFrame, path: Path) -> None:
    tmp = path.with_suffix(path.suffix + ".tmp.xlsx")
    with pd.ExcelWriter(tmp, engine="openpyxl") as writer:
        df.to_excel(writer, index=False)
    os.replace(tmp, path)


def persist_progress(df: pd.DataFrame, current_idx: int) -> None:
    atomic_write_excel(df, AUTOSAVE_XLSX)
    atomic_write_json(AUTOSAVE_META, {"current_idx": current_idx})


def load_progress() -> Tuple[Optional[pd.DataFrame], int]:
    if AUTOSAVE_XLSX.exists():
        df = pd.read_excel(AUTOSAVE_XLSX)
        idx = 0
        if AUTOSAVE_META.exists():
            with open(AUTOSAVE_META, "r", encoding="utf-8") as f:
                idx = int(json.load(f).get("current_idx", 0))
        return df, idx
    return None, 0


def reset_progress_files():
    if AUTOSAVE_XLSX.exists():
        AUTOSAVE_XLSX.unlink()
    if AUTOSAVE_META.exists():
        AUTOSAVE_META.unlink()


def ensure_enrichment_columns(df: pd.DataFrame) -> pd.DataFrame:
    for col in ENRICHMENT_COLUMNS:
        if col not in df.columns:
            df[col] = None
    return df


def apply_match(df: pd.DataFrame, row_idx, row: pd.Series, match: Dict[str, Any]) -> None:
    df.at[row_idx, "api_match_status"] = "MATCHED"
    df.at[row_idx, "api_query"] = build_input_address(row)
    df.at[row_idx, "api_score_adresse"] = match.get("score")
    df.at[row_idx, "api_nom_reel"] = match.get("nom")
    df.at[row_idx, "api_adresse_trouvee"] = match.get("adresse")
    df.at[row_idx, "api_siren"] = match.get("siren")
    df.at[row_idx, "api_siret"] = match.get("siret")
    df.at[row_idx, "api_intracom"] = match.get("intracom")
    df.at[row_idx, "api_tel_1"] = clean_str(row.get("Tél")) or None
    df.at[row_idx, "api_tel_2"] = clean_str(row.get("Tél2")) or None


def apply_no_match(df: pd.DataFrame, row_idx, row: pd.Series) -> None:
    df.at[row_idx, "api_match_status"] = "NO_MATCH_SELECTED"
    df.at[row_idx, "api_query"] = build_input_address(row)
    df.at[row_idx, "api_score_adresse"] = 0
    df.at[row_idx, "api_nom_reel"] = None
    df.at[row_idx, "api_adresse_trouvee"] = None
    df.at[row_idx, "api_siren"] = None
    df.at[row_idx, "api_siret"] = None
    df.at[row_idx, "api_intracom"] = None
    df.at[row_idx, "api_tel_1"] = clean_str(row.get("Tél")) or None
    df.at[row_idx, "api_tel_2"] = clean_str(row.get("Tél2")) or None


if "df_work" not in st.session_state:
    st.session_state.df_work = None

if "current_idx" not in st.session_state:
    st.session_state.current_idx = 0

if "loaded_once" not in st.session_state:
    st.session_state.loaded_once = False

if "api_client" not in st.session_state:
    st.session_state.api_client = RechercheEntrepriseClient()


st.title("Validation manuelle des correspondances fournisseurs")

top_col1, top_col2, top_col3 = st.columns([1.3, 1, 1])

with top_col1:
    uploaded = st.file_uploader("Charge un Excel ou CSV", type=["xlsx", "csv"])

with top_col2:
    if st.button("Reprendre l'autosave"):
        df_saved, idx_saved = load_progress()
        if df_saved is not None:
            st.session_state.df_work = ensure_enrichment_columns(df_saved.copy(deep=True))
            st.session_state.current_idx = idx_saved
            st.session_state.loaded_once = True
            st.rerun()
        else:
            st.warning("Aucun autosave trouvé.")

with top_col3:
    if st.button("Réinitialiser la reprise"):
        reset_progress_files()
        st.session_state.df_work = None
        st.session_state.current_idx = 0
        st.session_state.loaded_once = False
        st.rerun()


if uploaded is not None and not st.session_state.loaded_once:
    if uploaded.name.lower().endswith(".csv"):
        df_in = pd.read_csv(uploaded)
    else:
        df_in = pd.read_excel(uploaded)

    df_in = ensure_enrichment_columns(df_in.copy(deep=True))
    st.session_state.df_work = df_in
    st.session_state.current_idx = 0
    st.session_state.loaded_once = True

    persist_progress(st.session_state.df_work, st.session_state.current_idx)
    st.rerun()


df_work = st.session_state.df_work

if df_work is None:
    st.info("Charge un fichier ou reprends un autosave.")
    st.stop()


total = len(df_work)
current_idx = min(st.session_state.current_idx, max(total - 1, 0))

done_count = int((df_work["api_match_status"].notna()).sum()) if "api_match_status" in df_work.columns else 0

st.progress(done_count / total if total else 0.0)
st.write(f"Traité : {done_count} / {total}")
st.write(f"Position courante : {current_idx + 1} / {total}")

if total == 0:
    st.warning("Le fichier est vide.")
    st.stop()

if current_idx >= total:
    st.success("Tout est traité.")
else:
    row = df_work.iloc[current_idx]
    row_idx = df_work.index[current_idx]

    left, right = st.columns([1, 1.35])

    with left:
        st.subheader("Informations du fichier")

        source_info = {
            "Numero": row.get("Numero"),
            "Intitulé": row.get("Intitulé"),
            "Clé": row.get("Clé"),
            "Rue complète": row.get("Rue complète"),
            "CP-Ville": row.get("CP-Ville"),
            "Pays": row.get("Pays"),
            "Tél": row.get("Tél"),
            "Tél2": row.get("Tél2"),
            "Email": row.get("Email"),
        }
        st.json(source_info)

        nav1, nav2, nav3 = st.columns(3)

        with nav1:
            if st.button("⬅️ Précédent", disabled=(current_idx == 0)):
                st.session_state.current_idx -= 1
                st.rerun()

        with nav2:
            if st.button("Aucune correspondance", type="primary"):
                apply_no_match(df_work, row_idx, row)
                next_idx = min(current_idx + 1, total)
                st.session_state.current_idx = next_idx
                persist_progress(df_work, st.session_state.current_idx)
                st.rerun()

        with nav3:
            if st.button("➡️ Suivant sans choisir", disabled=(current_idx >= total - 1)):
                st.session_state.current_idx += 1
                st.rerun()

    with right:
        st.subheader("Résultats API")
        matches = get_api_matches_for_row(row, st.session_state.api_client, per_page=10)

        if not matches:
            st.warning("Aucun résultat API pour cette ligne.")
        else:
            for i, match in enumerate(matches):
                with st.container(border=True):
                    st.markdown(f"**{match.get('nom') or ''}**")
                    st.write(f"Adresse : {match.get('adresse') or ''}")
                    st.write(f"SIREN : {match.get('siren') or ''}")
                    st.write(f"SIRET : {match.get('siret') or ''}")
                    st.write(f"Intracom : {match.get('intracom') or ''}")
                    st.write(f"Proximité adresse : {match.get('score', 0):.4f}")

                    if st.button(f"Choisir ce résultat #{i + 1}", key=f"pick_{current_idx}_{i}"):
                        apply_match(df_work, row_idx, row, match)
                        next_idx = min(current_idx + 1, total)
                        st.session_state.current_idx = next_idx
                        persist_progress(df_work, st.session_state.current_idx)
                        st.rerun()


st.divider()
st.subheader("Téléchargement")

if AUTOSAVE_XLSX.exists():
    with open(AUTOSAVE_XLSX, "rb") as f:
        st.download_button(
            label="Télécharger le fichier enrichi courant",
            data=f.read(),
            file_name="fournisseurs_enrichis_autosave.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
else:
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df_work.to_excel(writer, index=False)
    buffer.seek(0)

    st.download_button(
        label="Télécharger le fichier enrichi courant",
        data=buffer.getvalue(),
        file_name="fournisseurs_enrichis.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )