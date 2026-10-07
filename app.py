# ==========================================================
# TRANSFERT DES PRIX HEIDENHAIN VERS ODOO
# V1 - Base import des fichiers
# ==========================================================

import streamlit as st
import pandas as pd
from io import BytesIO


# ==========================================================
# CONFIGURATION
# ==========================================================

st.set_page_config(
    page_title="Transfert des prix HEIDENHAIN vers ODOO",
    page_icon="📦",
    layout="wide",
)

st.title("Transfert des prix HEIDENHAIN vers ODOO")
st.caption("Étape 1 — Import des fichiers")


# ==========================================================
# INITIALISATION SESSION
# ==========================================================

if "heidenhain_df" not in st.session_state:
    st.session_state.heidenhain_df = None

if "odoo_df" not in st.session_state:
    st.session_state.odoo_df = None

if "heidenhain_file_name" not in st.session_state:
    st.session_state.heidenhain_file_name = None

if "odoo_file_name" not in st.session_state:
    st.session_state.odoo_file_name = None


# ==========================================================
# FONCTIONS
# ==========================================================

def read_uploaded_file(uploaded_file):
    """
    Lit un fichier Excel ou CSV et retourne un dictionnaire
    contenant les DataFrames.

    Pour Excel :
        {
            "NomFeuille1": dataframe,
            "NomFeuille2": dataframe,
            ...
        }

    Pour CSV :
        {
            "CSV": dataframe
        }
    """

    if uploaded_file is None:
        return {}

    extension = uploaded_file.name.lower().split(".")[-1]

    try:

        # --------------------------------------------------
        # EXCEL
        # --------------------------------------------------

        if extension in ["xlsx", "xls"]:

            excel_file = pd.ExcelFile(uploaded_file)

            sheets = {}

            for sheet_name in excel_file.sheet_names:
                sheets[sheet_name] = pd.read_excel(
                    excel_file,
                    sheet_name=sheet_name,
                )

            return sheets

        # --------------------------------------------------
        # CSV
        # --------------------------------------------------

        elif extension == "csv":

            uploaded_file.seek(0)

            # Première tentative avec ;
            try:
                df = pd.read_csv(
                    uploaded_file,
                    sep=";",
                    encoding="utf-8-sig",
                )

                # Si une seule colonne, le séparateur était
                # probablement incorrect.
                if len(df.columns) == 1:
                    uploaded_file.seek(0)

                    df = pd.read_csv(
                        uploaded_file,
                        sep=",",
                        encoding="utf-8-sig",
                    )

            except Exception:

                uploaded_file.seek(0)

                df = pd.read_csv(
                    uploaded_file,
                    sep=",",
                    encoding="utf-8-sig",
                )

            return {"CSV": df}

        else:
            st.error(
                f"Format non supporté : {uploaded_file.name}"
            )
            return {}

    except Exception as e:

        st.error(
            f"Impossible de lire le fichier "
            f"**{uploaded_file.name}**.\n\n"
            f"Erreur : `{e}`"
        )

        return {}


def display_file_info(
    file_name,
    sheets,
    file_label,
):
    """
    Affiche les informations générales du fichier.
    """

    if not sheets:
        return

    st.success(
        f"✅ {file_label} chargé : **{file_name}**"
    )

    st.write(
        f"**Nombre de feuille(s) :** {len(sheets)}"
    )

    for sheet_name, df in sheets.items():

        st.markdown(
            f"### 📄 Feuille : `{sheet_name}`"
        )

        col1, col2, col3 = st.columns(3)

        with col1:
            st.metric(
                "Lignes",
                f"{len(df):,}".replace(",", " "),
            )

        with col2:
            st.metric(
                "Colonnes",
                len(df.columns),
            )

        with col3:
            st.metric(
                "Cellules",
                f"{df.shape[0] * df.shape[1]:,}".replace(",", " "),
            )

        st.dataframe(
            df.head(10),
            use_container_width=True,
        )


def combine_excel_sheets(sheets):
    """
    Pour l'instant, on utilise la première feuille.

    Cette fonction pourra être remplacée lorsque nous
    connaîtrons précisément la structure du fichier Heidenhain
    et du fichier Odoo.
    """

    if not sheets:
        return None

    first_sheet = next(iter(sheets))

    return sheets[first_sheet].copy()


# ==========================================================
# SIDEBAR
# ==========================================================

st.sidebar.header("⚙️ Configuration")

st.sidebar.info(
    """
Cette première version permet uniquement de charger :

1. Le fichier des prix HEIDENHAIN
2. Le fichier d'import ODOO

Le traitement des données sera effectué dans l'étape suivante.
"""
)


# ==========================================================
# IMPORT DES FICHIERS
# ==========================================================

st.header("📂 1. Import des fichiers")


col1, col2 = st.columns(2)


# ==========================================================
# FICHIER HEIDENHAIN
# ==========================================================

with col1:

    st.subheader("📘 Fichier 1 — Prix HEIDENHAIN")

    heidenhain_file = st.file_uploader(
        "Sélectionner le fichier des prix HEIDENHAIN",
        type=["xlsx", "xls", "csv"],
        key="heidenhain_upload",
    )

    if heidenhain_file is not None:

        heidenhain_sheets = read_uploaded_file(
            heidenhain_file
        )

        if heidenhain_sheets:

            st.session_state.heidenhain_file_name = (
                heidenhain_file.name
            )

            st.session_state.heidenhain_df = (
                combine_excel_sheets(
                    heidenhain_sheets
                )
            )

            display_file_info(
                heidenhain_file.name,
                heidenhain_sheets,
                "Fichier prix HEIDENHAIN",
            )


# ==========================================================
# FICHIER ODOO
# ==========================================================

with col2:

    st.subheader("📗 Fichier 2 — Import ODOO")

    odoo_file = st.file_uploader(
        "Sélectionner le fichier d'import ODOO",
        type=["xlsx", "xls", "csv"],
        key="odoo_upload",
    )

    if odoo_file is not None:

        odoo_sheets = read_uploaded_file(
            odoo_file
        )

        if odoo_sheets:

            st.session_state.odoo_file_name = (
                odoo_file.name
            )

            st.session_state.odoo_df = (
                combine_excel_sheets(
                    odoo_sheets
                )
            )

            display_file_info(
                odoo_file.name,
                odoo_sheets,
                "Fichier import ODOO",
            )


# ==========================================================
# VERIFICATION DES DEUX FICHIERS
# ==========================================================

st.divider()

st.header("🔎 2. Vérification des fichiers")


heidenhain_ok = (
    st.session_state.heidenhain_df is not None
)

odoo_ok = (
    st.session_state.odoo_df is not None
)


col1, col2 = st.columns(2)


with col1:

    if heidenhain_ok:

        st.success(
            "✅ Fichier HEIDENHAIN disponible"
        )

        st.write(
            f"**Fichier :** "
            f"{st.session_state.heidenhain_file_name}"
        )

        st.write(
            f"**Lignes :** "
            f"{len(st.session_state.heidenhain_df)}"
        )

        st.write(
            f"**Colonnes :** "
            f"{len(st.session_state.heidenhain_df.columns)}"
        )

    else:

        st.warning(
            "⏳ Fichier HEIDENHAIN non chargé"
        )


with col2:

    if odoo_ok:

        st.success(
            "✅ Fichier ODOO disponible"
        )

        st.write(
            f"**Fichier :** "
            f"{st.session_state.odoo_file_name}"
        )

        st.write(
            f"**Lignes :** "
            f"{len(st.session_state.odoo_df)}"
        )

        st.write(
            f"**Colonnes :** "
            f"{len(st.session_state.odoo_df.columns)}"
        )

    else:

        st.warning(
            "⏳ Fichier ODOO non chargé"
        )


# ==========================================================
# LES DEUX FICHIERS SONT DISPONIBLES
# ==========================================================

if heidenhain_ok and odoo_ok:

    st.divider()

    st.success(
        "🎯 Les deux fichiers sont correctement chargés."
    )

    st.header("📊 Données disponibles pour le traitement")

    tab1, tab2 = st.tabs(
        [
            "📘 Prix HEIDENHAIN",
            "📗 Import ODOO",
        ]
    )

    with tab1:

        st.write(
            f"**Source :** "
            f"{st.session_state.heidenhain_file_name}"
        )

        st.dataframe(
            st.session_state.heidenhain_df,
            use_container_width=True,
            height=400,
        )

    with tab2:

        st.write(
            f"**Source :** "
            f"{st.session_state.odoo_file_name}"
        )

        st.dataframe(
            st.session_state.odoo_df,
            use_container_width=True,
            height=400,
        )

    st.divider()

    st.info(
        """
        ℹ️ Les deux fichiers sont maintenant disponibles
        en mémoire.

        **Étape suivante :** définir les colonnes et les règles
        permettant de faire le rapprochement entre les prix
        HEIDENHAIN et les produits ODOO.
        """
    )

else:

    st.info(
        "👆 Importez les deux fichiers pour pouvoir commencer "
        "le traitement."
    )
