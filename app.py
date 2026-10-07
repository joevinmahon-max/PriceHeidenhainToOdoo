# ==========================================================
# TRANSFERT DES PRIX HEIDENHAIN VERS ODOO
# JMA
# ==========================================================
#
# Etape 1 :
#   - Import fichier 1 : Prix Heidenhain
#   - Import fichier 2 : Import Odoo
#   - Traitement du fichier Heidenhain
#   - Duplication des lignes VG / PG avec suffixe "_SAv"
#
# ==========================================================

import streamlit as st
import pandas as pd

from traitement import (
    process_heidenhain,
    get_sheet_names,
)


# ==========================================================
# CONFIGURATION STREAMLIT
# ==========================================================

st.set_page_config(
    page_title="Transfert des prix HEIDENHAIN vers ODOO",
    page_icon="📦",
    layout="wide",
)

st.title("📦 Transfert des prix HEIDENHAIN vers ODOO")

st.caption(
    "Étape 1 — Préparation du fichier des prix HEIDENHAIN"
)


# ==========================================================
# SESSION STATE
# ==========================================================

if "heidenhain_bytes" not in st.session_state:
    st.session_state.heidenhain_bytes = None

if "heidenhain_result" not in st.session_state:
    st.session_state.heidenhain_result = None

if "heidenhain_stats" not in st.session_state:
    st.session_state.heidenhain_stats = None

if "odoo_df" not in st.session_state:
    st.session_state.odoo_df = None


# ==========================================================
# PARAMETRES
# ==========================================================

st.sidebar.header("⚙️ Paramètres")


heidenhain_sheet = st.sidebar.text_input(
    "Feuille du fichier Heidenhain",
    value="Distributeurs",
    help="Nom de la feuille dans le fichier des prix Heidenhain.",
)


heidenhain_header_row = st.sidebar.number_input(
    "Ligne des noms de colonnes Heidenhain",
    min_value=1,
    value=4,
    step=1,
    help="Les noms des colonnes se trouvent sur cette ligne.",
)


heidenhain_data_start_row = st.sidebar.number_input(
    "Première ligne de données Heidenhain",
    min_value=1,
    value=5,
    step=1,
    help="Les données commencent à cette ligne.",
)


status_column = st.sidebar.text_input(
    "Colonne Statut",
    value="Statut",
    help="Nom exact de la colonne contenant VG / PG.",
)


id_column = st.sidebar.text_input(
    "Colonne ID",
    value="ID",
    help="Nom exact de la colonne contenant l'identifiant.",
)


st.sidebar.markdown("---")

st.sidebar.info(
    """
### Règle actuelle

Pour chaque ligne dont le statut est :

- `VG`
- `PG`

on crée une copie de la ligne avec :

`ID` → `ID_SAv`

La copie est créée uniquement si le nouvel ID
n'existe pas déjà dans le fichier.
"""
)


# ==========================================================
# IMPORT DES FICHIERS
# ==========================================================

st.header("📂 1. Import des fichiers")

col1, col2 = st.columns(2)


# ==========================================================
# FICHIER 1 — HEIDENHAIN
# ==========================================================

with col1:

    st.subheader("📘 Fichier 1 — Prix Heidenhain")

    heidenhain_file = st.file_uploader(
        "Sélectionner le fichier des prix Heidenhain",
        type=["xlsx", "xlsm"],
        key="heidenhain_upload",
    )

    if heidenhain_file is not None:

        st.success(
            f"✅ Fichier chargé : `{heidenhain_file.name}`"
        )

        # ----------------------------------------------
        # Liste des feuilles
        # ----------------------------------------------

        try:

            sheet_names = get_sheet_names(
                heidenhain_file
            )

            st.write(
                "**Feuilles disponibles :**"
            )

            st.write(sheet_names)

            if heidenhain_sheet not in sheet_names:

                st.error(
                    f"❌ La feuille `{heidenhain_sheet}` "
                    f"n'existe pas dans le fichier."
                )

                st.info(
                    "Sélectionnez le bon nom de feuille "
                    "dans les paramètres à gauche."
                )

            else:

                st.success(
                    f"✅ Feuille `{heidenhain_sheet}` trouvée."
                )

        except Exception as e:

            st.error(
                f"Impossible de lire le fichier : {e}"
            )


# ==========================================================
# FICHIER 2 — ODOO
# ==========================================================

with col2:

    st.subheader("📗 Fichier 2 — Import Odoo")

    odoo_file = st.file_uploader(
        "Sélectionner le fichier d'import Odoo",
        type=["xlsx", "xls", "xlsm", "csv"],
        key="odoo_upload",
    )

    if odoo_file is not None:

        st.success(
            f"✅ Fichier chargé : `{odoo_file.name}`"
        )

        # ----------------------------------------------
        # Lecture provisoire du fichier Odoo
        # ----------------------------------------------

        try:

            extension = (
                odoo_file.name
                .lower()
                .split(".")[-1]
            )

            if extension == "csv":

                odoo_df = pd.read_csv(
                    odoo_file,
                    header=0,
                    sep=None,
                    engine="python",
                )

            else:

                odoo_df = pd.read_excel(
                    odoo_file,
                    sheet_name="Sheet1",
                    header=0,
                )

            st.session_state.odoo_df = odoo_df

            st.success(
                "✅ Feuille `Sheet1` chargée."
            )

            st.write(
                f"**Nombre de lignes :** {len(odoo_df)}"
            )

            st.write(
                f"**Nombre de colonnes :** "
                f"{len(odoo_df.columns)}"
            )

            st.dataframe(
                odoo_df.head(10),
                use_container_width=True,
            )

        except Exception as e:

            st.error(
                f"Impossible de lire le fichier Odoo : {e}"
            )


# ==========================================================
# APERCU FICHIER HEIDENHAIN
# ==========================================================

if heidenhain_file is not None:

    st.divider()

    st.header("🔎 2. Vérification du fichier Heidenhain")

    try:

        # ----------------------------------------------
        # Lecture avec Pandas uniquement pour aperçu
        # ----------------------------------------------

        heidenhain_file.seek(0)

        df_preview = pd.read_excel(
            heidenhain_file,
            sheet_name=heidenhain_sheet,
            header=heidenhain_header_row - 1,
        )

        st.write(
            f"**Feuille :** `{heidenhain_sheet}`"
        )

        st.write(
            f"**Première ligne de données :** "
            f"{heidenhain_data_start_row}"
        )

        st.write(
            f"**Nombre de lignes détectées :** "
            f"{len(df_preview)}"
        )

        # ----------------------------------------------
        # Vérification colonnes
        # ----------------------------------------------

        if status_column not in df_preview.columns:

            st.error(
                f"❌ La colonne `{status_column}` "
                f"n'existe pas."
            )

            st.write(
                "Colonnes disponibles :"
            )

            st.write(
                list(df_preview.columns)
            )

        elif id_column not in df_preview.columns:

            st.error(
                f"❌ La colonne `{id_column}` "
                f"n'existe pas."
            )

            st.write(
                "Colonnes disponibles :"
            )

            st.write(
                list(df_preview.columns)
            )

        else:

            st.success(
                f"✅ Colonnes `{status_column}` et "
                f"`{id_column}` trouvées."
            )

            st.dataframe(
                df_preview.head(10),
                use_container_width=True,
            )

            # ------------------------------------------
            # Statistiques VG / PG
            # ------------------------------------------

            status_series = (
                df_preview[status_column]
                .astype(str)
                .str.strip()
                .str.upper()
            )

            vg_count = (
                status_series == "VG"
            ).sum()

            pg_count = (
                status_series == "PG"
            ).sum()

            st.write(
                f"**Lignes VG :** {vg_count}"
            )

            st.write(
                f"**Lignes PG :** {pg_count}"
            )

    except Exception as e:

        st.error(
            f"Erreur lors de l'analyse du fichier "
            f"Heidenhain : {e}"
        )


# ==========================================================
# TRAITEMENT
# ==========================================================

st.divider()

st.header("⚙️ 3. Traitement du fichier Heidenhain")


if heidenhain_file is None:

    st.info(
        "👆 Importez le fichier des prix Heidenhain "
        "pour commencer."
    )

elif (
    st.session_state.get("odoo_df") is None
):

    st.warning(
        "⚠️ Le fichier Odoo n'est pas encore chargé."
    )

else:

    st.write(
        "Les deux fichiers sont disponibles."
    )

    st.markdown(
        """
        ### Traitement prévu

        Le programme va :

        1. Parcourir les lignes à partir de la ligne 5.
        2. Lire la colonne **Statut**.
        3. Sélectionner uniquement `VG` et `PG`.
        4. Récupérer l'**ID**.
        5. Créer `ID_SAv`.
        6. Vérifier si `ID_SAv` existe déjà.
        7. Si l'ID n'existe pas, copier la ligne complète.
        8. Ajouter la copie à la première ligne libre.
        9. Générer un nouveau fichier Excel.
        """
    )

    if st.button(
        "🚀 Traiter le fichier Heidenhain",
        type="primary",
        use_container_width=True,
    ):

        try:

            # ------------------------------------------
            # Remettre le fichier au début
            # ------------------------------------------

            heidenhain_file.seek(0)

            # ------------------------------------------
            # Traitement
            # ------------------------------------------

            result_bytes, stats = process_heidenhain(
                uploaded_file=heidenhain_file,
                sheet_name=heidenhain_sheet,
                header_row=heidenhain_header_row,
                data_start_row=heidenhain_data_start_row,
                status_column=status_column,
                id_column=id_column,
            )

            # ------------------------------------------
            # Sauvegarde en session
            # ------------------------------------------

            st.session_state.heidenhain_result = (
                result_bytes
            )

            st.session_state.heidenhain_stats = (
                stats
            )

            # ------------------------------------------
            # Résultat
            # ------------------------------------------

            st.success(
                "✅ Traitement terminé."
            )

            st.subheader("📊 Résultat")

            c1, c2, c3 = st.columns(3)

            with c1:

                st.metric(
                    "VG / PG trouvés",
                    stats["vg_pg_found"],
                )

            with c2:

                st.metric(
                    "Nouvelles lignes créées",
                    stats["rows_created"],
                )

            with c3:

                st.metric(
                    "IDs déjà existants",
                    stats["ids_already_existing"],
                )

            if stats["rows_created"] > 0:

                st.success(
                    f"✅ {stats['rows_created']} "
                    f"ligne(s) ajoutée(s)."
                )

            else:

                st.warning(
                    "Aucune nouvelle ligne n'a été créée."
                )

            # ------------------------------------------
            # IDs créés
            # ------------------------------------------

            if stats["created_ids"]:

                st.subheader(
                    "🆕 Nouveaux IDs créés"
                )

                st.dataframe(
                    pd.DataFrame(
                        {
                            "Nouveau ID": (
                                stats["created_ids"]
                            )
                        }
                    ),
                    use_container_width=True,
                )

            # ------------------------------------------
            # Téléchargement
            # ------------------------------------------

            st.download_button(
                label="⬇️ Télécharger le fichier Heidenhain mis à jour",
                data=result_bytes,
                file_name=(
                    "Prix_Heidenhain_mis_a_jour.xlsx"
                ),
                mime=(
                    "application/vnd.openxmlformats-officedocument"
                    ".spreadsheetml.sheet"
                ),
                use_container_width=True,
            )

        except Exception as e:

            st.error(
                "❌ Le traitement a échoué."
            )

            st.exception(e)


# ==========================================================
# ETAPE SUIVANTE
# ==========================================================

if (
    st.session_state.get("heidenhain_result")
    is not None
):

    st.divider()

    st.header("➡️ Étape suivante")

    st.info(
        """
        Le fichier Heidenhain est maintenant préparé.

        **Prochaine étape :** utiliser ce fichier avec
        le fichier d'import Odoo afin de rechercher les
        produits correspondants et modifier le fichier
        d'import Odoo.
        """
    )
