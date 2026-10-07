# ==========================================================
# TRANSFERT DES PRIX HEIDENHAIN VERS ODOO
# JMA
# ==========================================================
#
# Fichier unique : app.py
#
# Fichier 1 : Prix Heidenhain
# Fichier 2 : Import Odoo
#
# ETAPE ACTUELLE :
#   - Charger les deux fichiers
#   - Traiter le fichier Heidenhain
#   - Chercher les statuts VG / PG
#   - Créer les nouveaux ID avec suffixe _SAv
#   - Copier les lignes complètes
#   - Ne pas créer de doublons
#   - Télécharger le fichier Heidenhain modifié
#
# ==========================================================

import streamlit as st
import pandas as pd

from io import BytesIO
from copy import copy

import openpyxl


# ==========================================================
# CONFIGURATION STREAMLIT
# ==========================================================

st.set_page_config(
    page_title="Transfert des prix HEIDENHAIN vers ODOO",
    page_icon="📦",
    layout="wide",
)

st.title("📦 Transfert des prix HEIDENHAIN vers ODOO")


# ==========================================================
# PARAMETRES
# ==========================================================

st.sidebar.header("⚙️ Paramètres")


# ----------------------------------------------------------
# Fichier Heidenhain
# ----------------------------------------------------------

st.sidebar.subheader("📘 Fichier Prix Heidenhain")

heidenhain_sheet = st.sidebar.text_input(
    "Nom de la feuille",
    value="Distributeurs",
    help="Nom de la feuille à traiter dans le fichier Heidenhain.",
)

heidenhain_header_row = st.sidebar.number_input(
    "Ligne des noms de colonnes",
    min_value=1,
    value=4,
    step=1,
    help="Les noms des colonnes sont sur cette ligne.",
)

heidenhain_data_start_row = st.sidebar.number_input(
    "Première ligne des données",
    min_value=1,
    value=5,
    step=1,
    help="Les données commencent à cette ligne.",
)

status_column = st.sidebar.text_input(
    "Colonne Statut",
    value="Statut",
    help="Nom de la colonne contenant VG / PG.",
)

id_column = st.sidebar.text_input(
    "Colonne ID",
    value="ID",
    help="Nom de la colonne contenant l'identifiant.",
)


# ----------------------------------------------------------
# Fichier Odoo
# ----------------------------------------------------------

st.sidebar.subheader("📗 Fichier Import Odoo")

odoo_sheet = st.sidebar.text_input(
    "Feuille Odoo",
    value="Sheet1",
    help="Nom de la feuille du fichier d'import Odoo.",
)


# ==========================================================
# FONCTIONS
# ==========================================================

def normalize_value(value):
    """
    Transforme une valeur en texte comparable.
    """

    if value is None:
        return ""

    return str(value).strip().upper()


def get_excel_sheet_names(uploaded_file):
    """
    Retourne la liste des feuilles d'un fichier Excel.
    """

    uploaded_file.seek(0)

    workbook = openpyxl.load_workbook(
        uploaded_file,
        read_only=True,
        data_only=False,
    )

    return workbook.sheetnames


def find_column_by_header(
    worksheet,
    header_row,
    column_name,
):
    """
    Recherche le numéro de colonne à partir du nom
    présent sur la ligne d'en-tête.
    """

    target = normalize_value(column_name)

    for cell in worksheet[header_row]:

        if normalize_value(cell.value) == target:
            return cell.column

    return None


def copy_cell(
    source_cell,
    target_cell,
):
    """
    Copie une cellule avec sa valeur et sa mise en forme.
    """

    target_cell.value = source_cell.value

    if source_cell.has_style:
        target_cell._style = copy(
            source_cell._style
        )

    if source_cell.number_format:
        target_cell.number_format = (
            source_cell.number_format
        )

    if source_cell.alignment:
        target_cell.alignment = copy(
            source_cell.alignment
        )

    if source_cell.protection:
        target_cell.protection = copy(
            source_cell.protection
        )

    if source_cell.comment:
        target_cell.comment = copy(
            source_cell.comment
        )

    if source_cell.hyperlink:
        target_cell._hyperlink = copy(
            source_cell.hyperlink
        )


def copy_row(
    worksheet,
    source_row,
    target_row,
):
    """
    Copie toute une ligne Excel.
    """

    # ------------------------------------------------------
    # Hauteur de ligne
    # ------------------------------------------------------

    source_dimension = worksheet.row_dimensions[
        source_row
    ]

    target_dimension = worksheet.row_dimensions[
        target_row
    ]

    target_dimension.height = (
        source_dimension.height
    )

    target_dimension.hidden = (
        source_dimension.hidden
    )

    target_dimension.outlineLevel = (
        source_dimension.outlineLevel
    )

    target_dimension.collapsed = (
        source_dimension.collapsed
    )

    # ------------------------------------------------------
    # Cellules
    # ------------------------------------------------------

    for column in range(
        1,
        worksheet.max_column + 1,
    ):

        source_cell = worksheet.cell(
            row=source_row,
            column=column,
        )

        target_cell = worksheet.cell(
            row=target_row,
            column=column,
        )

        copy_cell(
            source_cell,
            target_cell,
        )


def process_heidenhain(
    uploaded_file,
    sheet_name,
    header_row,
    data_start_row,
    status_column,
    id_column,
):
    """
    Traite le fichier Heidenhain.

    Pour chaque ligne ayant le statut VG ou PG :

        ID original
            ↓
        ID_SAv

    Si le nouvel ID existe déjà :
        aucune copie.

    Sinon :
        copie complète de la ligne
        + remplacement de l'ID.
    """

    # ======================================================
    # CHARGEMENT DU FICHIER
    # ======================================================

    uploaded_file.seek(0)

    workbook = openpyxl.load_workbook(
        uploaded_file,
        data_only=False,
    )

    # ======================================================
    # VERIFICATION FEUILLE
    # ======================================================

    if sheet_name not in workbook.sheetnames:

        raise ValueError(
            f"La feuille '{sheet_name}' "
            f"n'existe pas dans le fichier."
        )

    worksheet = workbook[sheet_name]

    # ======================================================
    # RECHERCHE DES COLONNES
    # ======================================================

    status_col = find_column_by_header(
        worksheet,
        header_row,
        status_column,
    )

    if status_col is None:

        raise ValueError(
            f"La colonne '{status_column}' "
            f"n'a pas été trouvée sur la ligne "
            f"{header_row}."
        )

    id_col = find_column_by_header(
        worksheet,
        header_row,
        id_column,
    )

    if id_col is None:

        raise ValueError(
            f"La colonne '{id_column}' "
            f"n'a pas été trouvée sur la ligne "
            f"{header_row}."
        )

    # ======================================================
    # RECUPERATION DE TOUS LES IDS EXISTANTS
    # ======================================================

    existing_ids = set()

    for row in range(
        data_start_row,
        worksheet.max_row + 1,
    ):

        value = worksheet.cell(
            row=row,
            column=id_col,
        ).value

        normalized = normalize_value(value)

        if normalized:
            existing_ids.add(normalized)

    # ======================================================
    # RECHERCHE VG / PG
    # ======================================================

    rows_to_create = []

    for row in range(
        data_start_row,
        worksheet.max_row + 1,
    ):

        # --------------------------------------------------
        # Statut
        # --------------------------------------------------

        status_value = worksheet.cell(
            row=row,
            column=status_col,
        ).value

        status = normalize_value(
            status_value
        )

        # On ne traite que VG et PG
        if status not in {"VG", "PG"}:
            continue

        # --------------------------------------------------
        # ID
        # --------------------------------------------------

        original_id = worksheet.cell(
            row=row,
            column=id_col,
        ).value

        if original_id is None:
            continue

        original_id = str(
            original_id
        ).strip()

        if not original_id:
            continue

        # --------------------------------------------------
        # Nouveau ID
        # --------------------------------------------------

        new_id = f"{original_id}_SAv"

        rows_to_create.append(
            {
                "source_row": row,
                "status": status,
                "original_id": original_id,
                "new_id": new_id,
            }
        )

    # ======================================================
    # CREATION DES LIGNES
    # ======================================================

    created_ids = []
    existing_count = 0

    next_row = worksheet.max_row + 1

    for item in rows_to_create:

        source_row = item["source_row"]
        new_id = item["new_id"]

        normalized_new_id = normalize_value(
            new_id
        )

        # --------------------------------------------------
        # Le nouvel ID existe déjà
        # --------------------------------------------------

        if normalized_new_id in existing_ids:

            existing_count += 1

            continue

        # --------------------------------------------------
        # Copier la ligne complète
        # --------------------------------------------------

        copy_row(
            worksheet,
            source_row,
            next_row,
        )

        # --------------------------------------------------
        # Modifier uniquement l'ID
        # --------------------------------------------------

        worksheet.cell(
            row=next_row,
            column=id_col,
        ).value = new_id

        # --------------------------------------------------
        # Ajouter le nouvel ID aux IDs existants
        # --------------------------------------------------

        existing_ids.add(
            normalized_new_id
        )

        created_ids.append(
            new_id
        )

        next_row += 1

    # ======================================================
    # SAUVEGARDE
    # ======================================================

    output = BytesIO()

    workbook.save(output)

    output.seek(0)

    # ======================================================
    # STATISTIQUES
    # ======================================================

    stats = {
        "vg_pg_found": len(rows_to_create),
        "rows_created": len(created_ids),
        "ids_already_existing": existing_count,
        "created_ids": created_ids,
        "status_column_number": status_col,
        "id_column_number": id_col,
        "original_max_row": (
            next_row - len(created_ids) - 1
        ),
        "new_max_row": worksheet.max_row,
    }

    return output.getvalue(), stats


# ==========================================================
# IMPORT DES FICHIERS
# ==========================================================

st.header("📂 1. Import des fichiers")

col1, col2 = st.columns(2)


# ==========================================================
# FICHIER 1
# ==========================================================

with col1:

    st.subheader("📘 Fichier 1 — Prix Heidenhain")

    heidenhain_file = st.file_uploader(
        "Importer le fichier des prix Heidenhain",
        type=["xlsx", "xlsm"],
        key="heidenhain_file",
    )

    if heidenhain_file is not None:

        st.success(
            f"✅ {heidenhain_file.name}"
        )

        try:

            sheet_names = get_excel_sheet_names(
                heidenhain_file
            )

            st.write(
                "**Feuilles disponibles :**"
            )

            for sheet in sheet_names:
                st.write(f"- `{sheet}`")

            if heidenhain_sheet not in sheet_names:

                st.error(
                    f"❌ La feuille "
                    f"`{heidenhain_sheet}` "
                    f"n'existe pas."
                )

            else:

                st.success(
                    f"✅ Feuille `{heidenhain_sheet}` trouvée."
                )

        except Exception as e:

            st.error(
                f"Erreur de lecture : {e}"
            )


# ==========================================================
# FICHIER 2
# ==========================================================

with col2:

    st.subheader("📗 Fichier 2 — Import Odoo")

    odoo_file = st.file_uploader(
        "Importer le fichier d'import Odoo",
        type=["xlsx", "xlsm", "xls", "csv"],
        key="odoo_file",
    )

    if odoo_file is not None:

        st.success(
            f"✅ {odoo_file.name}"
        )

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

                odoo_file.seek(0)

                workbook_odoo = openpyxl.load_workbook(
                    odoo_file,
                    read_only=True,
                    data_only=False,
                )

                if odoo_sheet not in workbook_odoo.sheetnames:

                    st.error(
                        f"❌ La feuille "
                        f"`{odoo_sheet}` "
                        f"n'existe pas dans le fichier Odoo."
                    )

                else:

                    odoo_file.seek(0)

                    odoo_df = pd.read_excel(
                        odoo_file,
                        sheet_name=odoo_sheet,
                        header=0,
                    )

                    st.success(
                        f"✅ Feuille `{odoo_sheet}` trouvée."
                    )

                    st.write(
                        f"**Lignes :** {len(odoo_df)}"
                    )

                    st.write(
                        f"**Colonnes :** "
                        f"{len(odoo_df.columns)}"
                    )

                    st.dataframe(
                        odoo_df.head(10),
                        use_container_width=True,
                    )

        except Exception as e:

            st.error(
                f"Erreur de lecture du fichier Odoo : {e}"
            )


# ==========================================================
# APERCU HEIDENHAIN
# ==========================================================

if heidenhain_file is not None:

    st.divider()

    st.header(
        "🔎 2. Vérification du fichier Heidenhain"
    )

    try:

        heidenhain_file.seek(0)

        df_preview = pd.read_excel(
            heidenhain_file,
            sheet_name=heidenhain_sheet,
            header=heidenhain_header_row - 1,
        )

        # --------------------------------------------------
        # Colonnes
        # --------------------------------------------------

        col_names = list(
            df_preview.columns
        )

        if status_column not in col_names:

            st.error(
                f"❌ Colonne `{status_column}` "
                f"non trouvée."
            )

            st.write(
                "Colonnes disponibles :"
            )

            st.write(col_names)

        elif id_column not in col_names:

            st.error(
                f"❌ Colonne `{id_column}` "
                f"non trouvée."
            )

            st.write(
                "Colonnes disponibles :"
            )

            st.write(col_names)

        else:

            st.success(
                "✅ Les colonnes nécessaires "
                "ont été trouvées."
            )

            # ----------------------------------------------
            # Statuts
            # ----------------------------------------------

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
                f"**VG :** {vg_count}"
            )

            st.write(
                f"**PG :** {pg_count}"
            )

            st.write(
                f"**VG + PG :** "
                f"{vg_count + pg_count}"
            )

            # ----------------------------------------------
            # Aperçu
            # ----------------------------------------------

            st.dataframe(
                df_preview.head(10),
                use_container_width=True,
            )

    except Exception as e:

        st.error(
            f"Erreur lors de l'analyse : {e}"
        )


# ==========================================================
# TRAITEMENT
# ==========================================================

st.divider()

st.header(
    "⚙️ 3. Traitement du fichier Heidenhain"
)


if heidenhain_file is None:

    st.info(
        "Importez d'abord le fichier "
        "des prix Heidenhain."
    )

else:

    if st.button(
        "🚀 Traiter le fichier Heidenhain",
        type="primary",
        use_container_width=True,
    ):

        try:

            heidenhain_file.seek(0)

            result_bytes, stats = process_heidenhain(
                uploaded_file=heidenhain_file,
                sheet_name=heidenhain_sheet,
                header_row=int(
                    heidenhain_header_row
                ),
                data_start_row=int(
                    heidenhain_data_start_row
                ),
                status_column=status_column,
                id_column=id_column,
            )

            # ------------------------------------------
            # Résultats
            # ------------------------------------------

            st.success(
                "✅ Traitement terminé avec succès."
            )

            st.subheader(
                "📊 Résultat du traitement"
            )

            c1, c2, c3 = st.columns(3)

            with c1:

                st.metric(
                    "VG / PG trouvés",
                    stats["vg_pg_found"],
                )

            with c2:

                st.metric(
                    "Lignes ajoutées",
                    stats["rows_created"],
                )

            with c3:

                st.metric(
                    "IDs déjà existants",
                    stats["ids_already_existing"],
                )

            # ------------------------------------------
            # Détail
            # ------------------------------------------

            if stats["rows_created"] > 0:

                st.success(
                    f"{stats['rows_created']} "
                    f"nouvelle(s) ligne(s) créée(s)."
                )

                st.subheader(
                    "🆕 IDs créés"
                )

                created_df = pd.DataFrame(
                    {
                        "Nouvel ID": (
                            stats["created_ids"]
                        )
                    }
                )

                st.dataframe(
                    created_df,
                    use_container_width=True,
                )

            else:

                st.warning(
                    "Aucune nouvelle ligne n'a été créée."
                )

            if stats["ids_already_existing"] > 0:

                st.info(
                    f"{stats['ids_already_existing']} "
                    f"ID(s) existaient déjà et "
                    f"n'ont pas été dupliqués."
                )

            # ------------------------------------------
            # Téléchargement
            # ------------------------------------------

            st.download_button(
                label=(
                    "⬇️ Télécharger le fichier "
                    "Heidenhain mis à jour"
                ),
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
                "❌ Une erreur est survenue "
                "pendant le traitement."
            )

            st.exception(e)


# ==========================================================
# ETAPE SUIVANTE
# ==========================================================

st.divider()

st.header("➡️ Étape suivante")

st.info(
    """
Le fichier Heidenhain est maintenant traité.

La prochaine étape sera de prendre ce fichier modifié
et le fichier d'import Odoo afin de rechercher les
produits correspondants et modifier le fichier Odoo.
"""
)
