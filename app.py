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

    parts = []
    for word in text.split():
        if word in STOP_WORDS:
            continue
        parts.append(ABBREVIATIONS.get(word, word))

    return re.sub(r"\s+", " ", " ".join(parts)).strip()


def extract_cp_city(value: str) -> Tuple[str, str]:
    value = normalize_text(value)
    match = re.match(r"^(\d{5})\s+(.*)$", value)
    if not match:
        return "", value

    cp = match.group(1)
    city = match.group(2).strip()
    city = re.sub(r"\bCEDEX\b\s*\d*", "", city).strip()
    city = re.sub(r"\s+", " ", city)
    return cp, city


def similarity(a: str, b: str) -> float:
    if not a or not b:
        return 0.0
    from difflib import SequenceMatcher
    return SequenceMatcher(None, a, b).ratio()


def source_address(row: pd.Series) -> str:
    rue = clean_str(row.get("Rue complète"))
    cp_ville = clean_str(row.get("CP-Ville"))
    return " ".join(x for x in [rue, cp_ville] if x).strip()


def search_query(row: pd.Series) -> str:
    return clean_str(row.get("Intitulé"))


def is_france_or_unspecified(country: Any) -> bool:
    country = clean_str(country)
    if not country:
        return True
    return normalize_text(country) in {"FRANCE", "FR", "FRA"}


def get_nested(data: Dict[str, Any], *keys, default=None):
    current = data
    for key in keys:
        if not isinstance(current, dict):
            return default
        current = current.get(key)
        if current is None:
            return default
    return current


def normalize_siret(value: Any) -> str:
    return re.sub(r"\D", "", clean_str(value))


def siren_from_siret(value: Any) -> Optional[str]:
    siret = normalize_siret(value)
    if len(siret) != 14:
        return None
    return siret[:9]


def compute_french_vat_from_siren(siren: str) -> Optional[str]:
    siren = re.sub(r"\D", "", clean_str(siren))
    if len(siren) != 9:
        return None
    key = (12 + 3 * (int(siren) % 97)) % 97
    return f"FR{key:02d}{siren}"


class RechercheEntrepriseClient:
    def __init__(self):
        self.session = requests.Session()
        self.last_call = 0.0

    def throttle(self):
        elapsed = time.time() - self.last_call
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
                self.throttle()
                response = self.session.get(API_URL, params=params, timeout=10)
                self.last_call = time.time()

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


def extract_result_address(result: Dict[str, Any]) -> str:
    siege = result.get("siege") or {}

    full_address = clean_str(siege.get("adresse"))
    if full_address:
        return full_address

    street_parts = []
    for key in ["numero_voie", "type_voie", "libelle_voie", "complement_adresse"]:
        value = clean_str(siege.get(key))
        if value:
            street_parts.append(value)

    street = " ".join(street_parts)
    cp = clean_str(siege.get("code_postal"))
    city = clean_str(siege.get("libelle_commune")) or clean_str(siege.get("commune"))

    if street:
        return " ".join(x for x in [street, cp, city] if x).strip()

    return (
        clean_str(result.get("adresse"))
        or clean_str(result.get("adresse_complete"))
        or clean_str(result.get("full_address"))
    )


def result_country_is_france(result: Dict[str, Any]) -> bool:
    siege = result.get("siege") or {}
    country = clean_str(siege.get("libelle_pays_etranger")) or clean_str(result.get("pays"))
    if not country:
        return True
    return is_france_or_unspecified(country)


def address_score(source_addr: str, candidate_addr: str) -> float:
    if not source_addr or not candidate_addr:
        return 0.0

    source_norm = normalize_address(source_addr)
    candidate_norm = normalize_address(candidate_addr)

    score = similarity(source_norm, candidate_norm)

    source_cp, source_city = extract_cp_city(source_addr)
    candidate_cp, candidate_city = extract_cp_city(candidate_addr)

    if source_cp and candidate_cp:
        if source_cp == candidate_cp:
            score += 0.15
        else:
            score -= 0.25

    if source_city and candidate_city:
        city_score = similarity(normalize_address(source_city), normalize_address(candidate_city))
        if city_score >= 0.9:
            score += 0.10
        elif city_score < 0.5:
            score -= 0.10

    return max(0.0, min(score, 1.0))


def build_match_from_result(result: Dict[str, Any], row: pd.Series) -> Dict[str, Any]:
    siren = extract_siren(result)
    adresse = extract_result_address(result)

    return {
        "nom": extract_company_name(result),
        "adresse": adresse,
        "siren": siren,
        "siret": extract_siret(result),
        "intracom": compute_french_vat_from_siren(siren),
        "score": address_score(source_address(row), adresse),
    }


def get_api_matches_for_row(row: pd.Series, client: RechercheEntrepriseClient, per_page: int = 10) -> List[Dict[str, Any]]:
    if not is_france_or_unspecified(row.get("Pays")):
        return []

    query = search_query(row)
    if not query:
        return []

    data = client.search(query, per_page=per_page, page=1)
    if not data:
        return []

    matches = []

    for result in data.get("results", []) or []:
        if not result_country_is_france(result):
            continue
        matches.append(build_match_from_result(result, row))

    return matches


def get_company_by_siret(
    siret: str,
    row: pd.Series,
    client: RechercheEntrepriseClient,
) -> Optional[Dict[str, Any]]:
    siret = normalize_siret(siret)
    if len(siret) != 14:
        return None

    data = client.search(siret, per_page=10, page=1)
    if not data:
        return None

    for result in data.get("results", []) or []:
        if extract_siret(result) == siret:
            return build_match_from_result(result, row)

    return None


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


def get_resume_index(df: pd.DataFrame) -> int:
    if "api_match_status" not in df.columns:
        return 0

    pending = df.index[df["api_match_status"].isna()].tolist()
    if not pending:
        return len(df)

    first_pending_label = pending[0]
    return int(df.index.get_loc(first_pending_label))


def persist_progress(df: pd.DataFrame, current_idx: int) -> None:
    safe_idx = min(max(current_idx, 0), len(df))
    atomic_write_excel(df, AUTOSAVE_XLSX)
    atomic_write_json(AUTOSAVE_META, {"current_idx": safe_idx})


def load_progress() -> Tuple[Optional[pd.DataFrame], int]:
    if not AUTOSAVE_XLSX.exists():
        return None, 0

    df = pd.read_excel(AUTOSAVE_XLSX)
    df = ensure_enrichment_columns(df)
    resume_idx = get_resume_index(df)

    if AUTOSAVE_META.exists():
        try:
            with open(AUTOSAVE_META, "r", encoding="utf-8") as f:
                saved_idx = int(json.load(f).get("current_idx", 0))
            if resume_idx < len(df):
                return df, resume_idx
            return df, min(saved_idx, len(df))
        except Exception:
            pass

    return df, resume_idx


def reset_progress_files() -> None:
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
    df.at[row_idx, "api_query"] = search_query(row)
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
    df.at[row_idx, "api_query"] = search_query(row)
    df.at[row_idx, "api_score_adresse"] = 0
    df.at[row_idx, "api_nom_reel"] = None
    df.at[row_idx, "api_adresse_trouvee"] = None
    df.at[row_idx, "api_siren"] = None
    df.at[row_idx, "api_siret"] = None
    df.at[row_idx, "api_intracom"] = None
    df.at[row_idx, "api_tel_1"] = clean_str(row.get("Tél")) or None
    df.at[row_idx, "api_tel_2"] = clean_str(row.get("Tél2")) or None


def mark_foreign_row(df: pd.DataFrame, row_idx, row: pd.Series) -> None:
    df.at[row_idx, "api_match_status"] = "FOREIGN_NOT_SEARCHED"
    df.at[row_idx, "api_query"] = None
    df.at[row_idx, "api_score_adresse"] = 0
    df.at[row_idx, "api_nom_reel"] = None
    df.at[row_idx, "api_adresse_trouvee"] = None
    df.at[row_idx, "api_siren"] = None
    df.at[row_idx, "api_siret"] = None
    df.at[row_idx, "api_intracom"] = None
    df.at[row_idx, "api_tel_1"] = clean_str(row.get("Tél")) or None
    df.at[row_idx, "api_tel_2"] = clean_str(row.get("Tél2")) or None


def apply_manual_siret(
    df: pd.DataFrame,
    row_idx,
    row: pd.Series,
    manual_siret: str,
    match: Optional[Dict[str, Any]],
) -> None:
    siret = normalize_siret(manual_siret)
    siren = siren_from_siret(siret)

    df.at[row_idx, "api_query"] = siret
    df.at[row_idx, "api_tel_1"] = clean_str(row.get("Tél")) or None
    df.at[row_idx, "api_tel_2"] = clean_str(row.get("Tél2")) or None

    if match is None:
        df.at[row_idx, "api_match_status"] = "MANUAL_SIRET_ONLY"
        df.at[row_idx, "api_score_adresse"] = 0
        df.at[row_idx, "api_nom_reel"] = None
        df.at[row_idx, "api_adresse_trouvee"] = None
        df.at[row_idx, "api_siren"] = siren
        df.at[row_idx, "api_siret"] = siret or None
        df.at[row_idx, "api_intracom"] = compute_french_vat_from_siren(siren) if siren else None
        return

    df.at[row_idx, "api_match_status"] = "MANUAL_SIRET_MATCHED"
    df.at[row_idx, "api_score_adresse"] = match.get("score")
    df.at[row_idx, "api_nom_reel"] = match.get("nom")
    df.at[row_idx, "api_adresse_trouvee"] = match.get("adresse")
    df.at[row_idx, "api_siren"] = match.get("siren")
    df.at[row_idx, "api_siret"] = match.get("siret")
    df.at[row_idx, "api_intracom"] = match.get("intracom")


if "df_work" not in st.session_state:
    st.session_state.df_work = None

if "current_idx" not in st.session_state:
    st.session_state.current_idx = 0

if "api_client" not in st.session_state:
    st.session_state.api_client = RechercheEntrepriseClient()

if "matches_cache" not in st.session_state:
    st.session_state.matches_cache = {}


st.title("Validation manuelle des correspondances fournisseurs")

c1, c2, c3 = st.columns([1.4, 1, 1])

with c1:
    uploaded = st.file_uploader("Charge un fichier Excel ou CSV", type=["xlsx", "csv"])
    if uploaded is not None and st.button("Charger ce fichier"):
        if uploaded.name.lower().endswith(".csv"):
            df_in = pd.read_csv(uploaded)
        else:
            df_in = pd.read_excel(uploaded)

        df_in = ensure_enrichment_columns(df_in.copy(deep=True))
        resume_idx = get_resume_index(df_in)

        st.session_state.df_work = df_in
        st.session_state.current_idx = resume_idx
        st.session_state.matches_cache = {}

        persist_progress(st.session_state.df_work, st.session_state.current_idx)
        st.rerun()

with c2:
    if st.button("Reprendre l'autosave"):
        df_saved, idx_saved = load_progress()
        if df_saved is not None:
            st.session_state.df_work = ensure_enrichment_columns(df_saved.copy(deep=True))
            st.session_state.current_idx = idx_saved
            st.session_state.matches_cache = {}
            st.rerun()
        else:
            st.warning("Aucun autosave trouvé.")

with c3:
    if st.button("Réinitialiser la reprise"):
        reset_progress_files()
        st.session_state.df_work = None
        st.session_state.current_idx = 0
        st.session_state.matches_cache = {}
        st.rerun()


df_work = st.session_state.df_work

if df_work is None:
    st.info("Charge un fichier ou reprends un autosave.")
    st.stop()

total = len(df_work)
current_idx = st.session_state.current_idx
done_count = int(df_work["api_match_status"].notna().sum())

st.progress(done_count / total if total else 0.0)
st.write(f"Traité : {done_count} / {total}")

if total == 0:
    st.warning("Le fichier est vide.")
    st.stop()

if current_idx >= total:
    st.success("Tout est traité.")
else:
    st.write(f"Position courante : {current_idx + 1} / {total}")

    row = df_work.iloc[current_idx]
    row_idx = df_work.index[current_idx]
    foreign_row = not is_france_or_unspecified(row.get("Pays"))

    left, right = st.columns([1, 1.35])

    with left:
        st.subheader("Informations du fichier")
        st.json({
            "Numero": row.get("Numero"),
            "Intitulé": row.get("Intitulé"),
            "Clé": row.get("Clé"),
            "Rue complète": row.get("Rue complète"),
            "CP-Ville": row.get("CP-Ville"),
            "Pays": row.get("Pays"),
            "Tél": row.get("Tél"),
            "Tél2": row.get("Tél2"),
            "Email": row.get("Email"),
        })

        n1, n2, n3 = st.columns(3)

        with n1:
            if st.button("⬅️ Précédent", disabled=(current_idx == 0)):
                st.session_state.current_idx -= 1
                st.rerun()

        with n2:
            if foreign_row:
                if st.button("Marquer comme étranger", type="primary"):
                    mark_foreign_row(df_work, row_idx, row)
                    st.session_state.current_idx += 1
                    persist_progress(df_work, st.session_state.current_idx)
                    st.rerun()
            else:
                if st.button("Aucune correspondance", type="primary"):
                    apply_no_match(df_work, row_idx, row)
                    st.session_state.current_idx += 1
                    persist_progress(df_work, st.session_state.current_idx)
                    st.rerun()

        with n3:
            if st.button("➡️ Suivant sans choisir", disabled=(current_idx >= total - 1)):
                st.session_state.current_idx += 1
                st.rerun()

        st.divider()
        st.subheader("Saisie manuelle")

        manual_siret = st.text_input(
            "SIRET manuel",
            key=f"manual_siret_{row_idx}",
            placeholder="12345678901234",
        )

        if st.button("Valider le SIRET manuel", use_container_width=True):
            siret_value = normalize_siret(manual_siret)

            if len(siret_value) != 14:
                st.error("Le SIRET doit contenir exactement 14 chiffres.")
            else:
                manual_match = get_company_by_siret(
                    siret_value,
                    row,
                    st.session_state.api_client,
                )

                apply_manual_siret(
                    df_work,
                    row_idx,
                    row,
                    siret_value,
                    manual_match,
                )

                st.session_state.current_idx += 1
                persist_progress(df_work, st.session_state.current_idx)
                st.rerun()

    with right:
        st.subheader("Résultats API")

        if foreign_row:
            st.info("Pays étranger détecté : la ligne restera dans le fichier de sortie, mais aucune recherche automatique ne sera lancée.")
        else:
            cache_key = f"{row_idx}|{search_query(row)}"
            if cache_key not in st.session_state.matches_cache:
                st.session_state.matches_cache[cache_key] = get_api_matches_for_row(
                    row,
                    st.session_state.api_client,
                    per_page=10,
                )

            matches = st.session_state.matches_cache[cache_key]

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
                            st.session_state.current_idx += 1
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