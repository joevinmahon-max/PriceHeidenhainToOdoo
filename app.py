# ==========================================================
# PREPARATION DES PRIX HEIDENHAIN
# JMA
#
# FONCTIONNEMENT
# ----------------------------------------------------------
# - Fichier Heidenhain configurable
# - Feuille configurable
# - Ligne d'en-têtes configurable
# - Première ligne de données configurable
# - Conservation uniquement des colonnes demandées
# - Recherche des statuts VG / PG
# - Création des ID_SAV
# - Vérification des doublons
# - Ajout des lignes SAV à la fin
# - Pour "Prix (SAV)" :
#       récupération de la VALEUR CALCULEE
#       et non de la formule
# - Aucun aperçu du fichier
# - Résultat conservé après téléchargement
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
    page_title="Préparation des prix HEIDENHAIN",
    page_icon="📦",
    layout="wide",
)

st.title("📦 Préparation du fichier Prix HEIDENHAIN")


# ==========================================================
# SESSION STATE
# ==========================================================

if "result_bytes" not in st.session_state:
    st.session_state.result_bytes = None

if "result_stats" not in st.session_state:
    st.session_state.result_stats = None

if "result_file_name" not in st.session_state:
    st.session_state.result_file_name = None


# ==========================================================
# CONSTANTES
# ==========================================================

STATUS_TO_DUPLICATE = {
    "VG",
    "PG",
}

DEFAULT_OUTPUT_COLUMNS = [
    "ID",
    "Description",
    "Marque",
    "Statut",
    "Groupe Produit",
    "Prix (PPC)",
    "Prix (SAV)",
]


# ==========================================================
# PARAMETRES SIDEBAR
# ==========================================================

st.sidebar.header("⚙️ Paramètres")


# ==========================================================
# FICHIER HEIDENHAIN
# ==========================================================

st.sidebar.subheader(
    "📘 Fichier Prix Heidenhain"
)

heidenhain_sheet = st.sidebar.text_input(
    "Feuille Heidenhain",
    value="Distributeurs",
)

heidenhain_header_row = st.sidebar.number_input(
    "Ligne des noms de colonnes",
    min_value=1,
    value=4,
    step=1,
)

heidenhain_data_start_row = st.sidebar.number_input(
    "Première ligne de données",
    min_value=1,
    value=5,
    step=1,
)

status_column = st.sidebar.text_input(
    "Colonne Statut",
    value="Statut",
)

id_column = st.sidebar.text_input(
    "Colonne ID",
    value="ID",
)


# ==========================================================
# COLONNES A CONSERVER
# ==========================================================

st.sidebar.divider()

st.sidebar.subheader(
    "📋 Colonnes à conserver"
)

output_columns_text = st.sidebar.text_area(
    "Une colonne par ligne",
    value="\n".join(
        DEFAULT_OUTPUT_COLUMNS
    ),
    height=180,
)

OUTPUT_COLUMNS = [
    column.strip()
    for column in output_columns_text.splitlines()
    if column.strip()
]


# ==========================================================
# FICHIER ODOO
# ==========================================================

st.sidebar.divider()

st.sidebar.subheader(
    "📗 Fichier Odoo"
)

odoo_sheet = st.sidebar.text_input(
    "Feuille Odoo",
    value="Sheet1",
)


# ==========================================================
# UTILITAIRE : NORMALISATION
# ==========================================================

def normalize_value(value):
    """
    Normalise une valeur pour les comparaisons.
    """

    if value is None:
        return ""

    return str(value).strip().upper()


# ==========================================================
# UTILITAIRE : NOMS DES FEUILLES
# ==========================================================

def get_excel_sheet_names(uploaded_file):
    """
    Retourne uniquement les noms des feuilles.

    IMPORTANT :
    Cette fonction ne dépend d'aucune variable externe
    comme sheet_name.
    """

    uploaded_file.seek(0)

    workbook = openpyxl.load_workbook(
        uploaded_file,
        read_only=True,
        data_only=False,
    )

    sheet_names = list(
        workbook.sheetnames
    )

    workbook.close()

    return sheet_names


# ==========================================================
# UTILITAIRE : RECHERCHE DES COLONNES
# ==========================================================

def find_columns_by_headers(
    worksheet,
    header_row,
    requested_columns,
):
    """
    Recherche les colonnes demandées
    sur la ligne d'en-têtes.
    """

    requested_normalized = {
        normalize_value(column): column
        for column in requested_columns
    }

    found = {}

    for cell in worksheet[header_row]:

        normalized = normalize_value(
            cell.value
        )

        if normalized in requested_normalized:

            original_name = (
                requested_normalized[
                    normalized
                ]
            )

            if original_name not in found:

                found[
                    original_name
                ] = cell.column

    return found


# ==========================================================
# UTILITAIRE : COPIE DU STYLE
# ==========================================================

def copy_cell_style(
    source_cell,
    target_cell,
):
    """
    Copie le style Excel de la cellule source.
    """

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


# ==========================================================
# UTILITAIRE : COPIE D'UNE LIGNE
# ==========================================================

def copy_selected_row(
    source_ws,
    values_ws,
    target_ws,
    source_row,
    target_row,
    source_columns,
    output_columns,
    new_id=None,
):
    """
    Copie une ligne dans le nouveau fichier.

    Pour "Prix (SAV)", on prend la valeur calculée
    présente dans le classeur ouvert avec data_only=True.

    Pour les autres colonnes, on prend la valeur normale.
    """

    for output_col_index, column_name in enumerate(
        output_columns,
        start=1,
    ):

        source_col_index = (
            source_columns[
                column_name
            ]
        )

        source_cell = source_ws.cell(
            row=source_row,
            column=source_col_index,
        )

        target_cell = target_ws.cell(
            row=target_row,
            column=output_col_index,
        )

        # ==================================================
        # PRIX SAV
        # ==================================================

        if column_name == "Prix (SAV)":

            calculated_value = (
                values_ws.cell(
                    row=source_row,
                    column=source_col_index,
                ).value
            )

            target_cell.value = (
                calculated_value
            )

        else:

            target_cell.value = (
                source_cell.value
            )

        # ==================================================
        # STYLE
        # ==================================================

        copy_cell_style(
            source_cell,
            target_cell,
        )

    # ======================================================
    # REMPLACEMENT DE L'ID
    # ======================================================

    if new_id is not None:

        if "ID" not in output_columns:

            raise ValueError(
                "La colonne 'ID' doit être présente "
                "dans les colonnes à conserver."
            )

        id_output_column = (
            output_columns.index("ID")
            + 1
        )

        target_ws.cell(
            row=target_row,
            column=id_output_column,
        ).value = new_id


# ==========================================================
# TRAITEMENT PRINCIPAL
# ==========================================================

def process_heidenhain(
    uploaded_file,
    sheet_name,
    header_row,
    data_start_row,
    status_column,
    id_column,
    output_columns,
):
    """
    Traite complètement le fichier Heidenhain.
    """

    header_row = int(
        header_row
    )

    data_start_row = int(
        data_start_row
    )

    # ======================================================
    # VALIDATIONS
    # ======================================================

    if data_start_row <= header_row:

        raise ValueError(
            "La première ligne de données doit être "
            "supérieure à la ligne des en-têtes."
        )

    if not output_columns:

        raise ValueError(
            "Aucune colonne à conserver."
        )

    if "ID" not in output_columns:

        raise ValueError(
            "La colonne 'ID' doit obligatoirement "
            "être conservée."
        )

    if status_column not in output_columns:

        raise ValueError(
            f"La colonne '{status_column}' doit "
            "faire partie des colonnes conservées."
        )

    if id_column not in output_columns:

        raise ValueError(
            f"La colonne '{id_column}' doit "
            "faire partie des colonnes conservées."
        )

    # ======================================================
    # OUVERTURE DU FICHIER AVEC FORMULES
    # ======================================================

    uploaded_file.seek(0)

    workbook = openpyxl.load_workbook(
        uploaded_file,
        data_only=False,
    )

    if sheet_name not in workbook.sheetnames:

        available = ", ".join(
            workbook.sheetnames
        )

        workbook.close()

        raise ValueError(
            f"La feuille '{sheet_name}' "
            f"n'existe pas.\n\n"
            f"Feuilles disponibles : {available}"
        )

    source_ws = workbook[
        sheet_name
    ]

    # ======================================================
    # OUVERTURE DU MEME FICHIER AVEC VALEURS CALCULEES
    # ======================================================

    uploaded_file.seek(0)

    workbook_values = (
        openpyxl.load_workbook(
            uploaded_file,
            data_only=True,
        )
    )

    if sheet_name not in (
        workbook_values.sheetnames
    ):

        workbook.close()
        workbook_values.close()

        raise ValueError(
            f"La feuille '{sheet_name}' "
            "n'existe pas dans le fichier "
            "des valeurs calculées."
        )

    values_ws = (
        workbook_values[
            sheet_name
        ]
    )

    # ======================================================
    # RECHERCHE DES COLONNES
    # ======================================================

    required_columns = list(
        dict.fromkeys(
            output_columns
            + [
                status_column,
                id_column,
            ]
        )
    )

    source_columns = (
        find_columns_by_headers(
            source_ws,
            header_row,
            required_columns,
        )
    )

    missing_columns = [
        column
        for column in required_columns
        if column not in source_columns
    ]

    if missing_columns:

        workbook.close()
        workbook_values.close()

        raise ValueError(
            "Colonnes introuvables sur la ligne "
            f"{header_row} : "
            + ", ".join(
                missing_columns
            )
        )

    status_col = (
        source_columns[
            status_column
        ]
    )

    id_col = (
        source_columns[
            id_column
        ]
    )

    # ======================================================
    # NOMBRE DE LIGNES
    # ======================================================

    max_row = source_ws.max_row

    total_rows = max(
        0,
        max_row - data_start_row + 1,
    )

    # ======================================================
    # NOUVEAU CLASSEUR
    # ======================================================

    output_workbook = (
        openpyxl.Workbook()
    )

    output_ws = (
        output_workbook.active
    )

    output_ws.title = sheet_name

    # ======================================================
    # EN-TETES
    # ======================================================

    for output_col_index, column_name in enumerate(
        output_columns,
        start=1,
    ):

        source_header_cell = (
            source_ws.cell(
                row=header_row,
                column=source_columns[
                    column_name
                ],
            )
        )

        target_header_cell = (
            output_ws.cell(
                row=1,
                column=output_col_index,
            )
        )

        target_header_cell.value = (
            source_header_cell.value
        )

        copy_cell_style(
            source_header_cell,
            target_header_cell,
        )

    # ======================================================
    # PROGRESSION
    # ======================================================

    progress = st.progress(
        0,
        text="Préparation du fichier...",
    )

    # ======================================================
    # ETAPE 1
    # IDS EXISTANTS
    # ======================================================

    existing_ids = set()

    for index, row in enumerate(
        range(
            data_start_row,
            max_row + 1,
        )
    ):

        value = source_ws.cell(
            row=row,
            column=id_col,
        ).value

        normalized_id = (
            normalize_value(value)
        )

        if normalized_id:

            existing_ids.add(
                normalized_id
            )

        if (
            index % 5000 == 0
            or index == total_rows - 1
        ):

            percent = int(
                (
                    (index + 1)
                    / max(
                        1,
                        total_rows,
                    )
                )
                * 15
            )

            progress.progress(
                min(
                    percent,
                    15,
                ),
                text=(
                    "Lecture des IDs... "
                    f"{index + 1:,} / "
                    f"{total_rows:,}"
                ),
            )

    # ======================================================
    # ETAPE 2
    # COPIE DES LIGNES EXISTANTES
    # ======================================================

    rows_to_create = []

    vg_count = 0
    pg_count = 0
    empty_id_count = 0
    duplicate_count = 0

    output_data_row = 2

    for index, source_row in enumerate(
        range(
            data_start_row,
            max_row + 1,
        )
    ):

        # ==================================================
        # COPIE DE LA LIGNE ORIGINALE
        # ==================================================

        copy_selected_row(
            source_ws=source_ws,
            values_ws=values_ws,
            target_ws=output_ws,
            source_row=source_row,
            target_row=output_data_row,
            source_columns=source_columns,
            output_columns=output_columns,
        )

        # ==================================================
        # HAUTEUR
        # ==================================================

        output_ws.row_dimensions[
            output_data_row
        ].height = (
            source_ws.row_dimensions[
                source_row
            ].height
        )

        # ==================================================
        # STATUT
        # ==================================================

        status = normalize_value(
            source_ws.cell(
                row=source_row,
                column=status_col,
            ).value
        )

        # ==================================================
        # PAS VG / PG
        # ==================================================

        if status not in STATUS_TO_DUPLICATE:

            output_data_row += 1

            continue

        # ==================================================
        # COMPTEUR
        # ==================================================

        if status == "VG":

            vg_count += 1

        elif status == "PG":

            pg_count += 1

        # ==================================================
        # ID ORIGINAL
        # ==================================================

        original_id = (
            source_ws.cell(
                row=source_row,
                column=id_col,
            ).value
        )

        if original_id is None:

            empty_id_count += 1

            output_data_row += 1

            continue

        original_id = str(
            original_id
        ).strip()

        if not original_id:

            empty_id_count += 1

            output_data_row += 1

            continue

        # ==================================================
        # ID SAV
        # ==================================================

        new_id = (
            f"{original_id}_SAV"
        )

        normalized_new_id = (
            normalize_value(
                new_id
            )
        )

        # ==================================================
        # DOUBLON
        # ==================================================

        if (
            normalized_new_id
            in existing_ids
        ):

            duplicate_count += 1

            output_data_row += 1

            continue

        # ==================================================
        # RESERVATION
        # ==================================================

        existing_ids.add(
            normalized_new_id
        )

        rows_to_create.append(
            (
                source_row,
                new_id,
            )
        )

        output_data_row += 1

        # ==================================================
        # PROGRESSION
        # ==================================================

        if (
            index % 5000 == 0
            or index == total_rows - 1
        ):

            percent = (
                15
                + int(
                    (
                        (index + 1)
                        / max(
                            1,
                            total_rows,
                        )
                    )
                    * 45
                )
            )

            progress.progress(
                min(
                    percent,
                    60,
                ),
                text=(
                    "Traitement des VG / PG... "
                    f"{index + 1:,} / "
                    f"{total_rows:,}"
                ),
            )

    # ======================================================
    # ETAPE 3
    # CREATION DES LIGNES SAV
    # ======================================================

    rows_created = len(
        rows_to_create
    )

    created_ids = []

    next_output_row = (
        output_data_row
    )

    progress.progress(
        60,
        text=(
            f"Création de "
            f"{rows_created:,} "
            "ligne(s) SAV..."
        ),
    )

    for index, (
        source_row,
        new_id,
    ) in enumerate(
        rows_to_create
    ):

        copy_selected_row(
            source_ws=source_ws,
            values_ws=values_ws,
            target_ws=output_ws,
            source_row=source_row,
            target_row=next_output_row,
            source_columns=source_columns,
            output_columns=output_columns,
            new_id=new_id,
        )

        output_ws.row_dimensions[
            next_output_row
        ].height = (
            source_ws.row_dimensions[
                source_row
            ].height
        )

        created_ids.append(
            new_id
        )

        next_output_row += 1

        if (
            index % 500 == 0
            or index == rows_created - 1
        ):

            percent = (
                60
                + int(
                    (
                        (index + 1)
                        / max(
                            1,
                            rows_created,
                        )
                    )
                    * 30
                )
            )

            progress.progress(
                min(
                    percent,
                    90,
                ),
                text=(
                    "Création des lignes SAV... "
                    f"{index + 1:,} / "
                    f"{rows_created:,}"
                ),
            )

    # ======================================================
    # ETAPE 4
    # LARGEURS
    # ======================================================

    progress.progress(
        92,
        text="Mise en forme...",
    )

    for output_col_index, column_name in enumerate(
        output_columns,
        start=1,
    ):

        source_col_index = (
            source_columns[
                column_name
            ]
        )

        source_letter = (
            openpyxl.utils.get_column_letter(
                source_col_index
            )
        )

        target_letter = (
            openpyxl.utils.get_column_letter(
                output_col_index
            )
        )

        source_width = (
            source_ws.column_dimensions[
                source_letter
            ].width
        )

        if source_width is not None:

            output_ws.column_dimensions[
                target_letter
            ].width = source_width

    # ======================================================
    # FIGER L'EN-TETE
    # ======================================================

    output_ws.freeze_panes = "A2"

    # ======================================================
    # SAUVEGARDE
    # ======================================================

    progress.progress(
        95,
        text="Création du fichier Excel...",
    )

    output = BytesIO()

    output_workbook.save(
        output
    )

    output.seek(0)

    result_bytes = (
        output.getvalue()
    )

    # ======================================================
    # FERMETURE
    # ======================================================

    output_workbook.close()
    workbook.close()
    workbook_values.close()

    progress.progress(
        100,
        text="✅ Traitement terminé",
    )

    # ======================================================
    # STATISTIQUES
    # ======================================================

    stats = {
        "total_source_rows": total_rows,
        "vg_count": vg_count,
        "pg_count": pg_count,
        "vg_pg_found": (
            vg_count + pg_count
        ),
        "rows_created": rows_created,
        "duplicates": duplicate_count,
        "empty_ids": empty_id_count,
        "created_ids": created_ids,
        "output_total_rows": (
            next_output_row - 1
        ),
    }

    return (
        result_bytes,
        stats,
    )


# ==========================================================
# INTERFACE
# ==========================================================

st.header(
    "📂 1. Import des fichiers"
)

col1, col2 = st.columns(2)


# ==========================================================
# FICHIER HEIDENHAIN
# ==========================================================

with col1:

    st.subheader(
        "📘 Fichier Prix Heidenhain"
    )

    heidenhain_file = st.file_uploader(
        "Importer le fichier des prix Heidenhain",
        type=[
            "xlsx",
            "xlsm",
        ],
        key="heidenhain_upload",
    )

    if heidenhain_file is not None:

        st.success(
            f"✅ `{heidenhain_file.name}`"
        )

        try:

            sheet_names = (
                get_excel_sheet_names(
                    heidenhain_file
                )
            )

            if heidenhain_sheet in sheet_names:

                st.success(
                    f"✅ Feuille "
                    f"`{heidenhain_sheet}` trouvée."
                )

            else:

                st.error(
                    f"❌ La feuille "
                    f"`{heidenhain_sheet}` "
                    "n'existe pas."
                )

                st.write(
                    "Feuilles disponibles :"
                )

                for sheet in sheet_names:

                    st.write(
                        f"- `{sheet}`"
                    )

        except Exception as e:

            st.error(
                f"Erreur de lecture du fichier : {e}"
            )


# ==========================================================
# FICHIER ODOO
# ==========================================================

with col2:

    st.subheader(
        "📗 Fichier Odoo"
    )

    odoo_file = st.file_uploader(
        "Importer le fichier d'import Odoo",
        type=[
            "xlsx",
            "xlsm",
            "xls",
            "csv",
        ],
        key="odoo_upload",
    )

    if odoo_file is not None:

        st.success(
            f"✅ `{odoo_file.name}`"
        )

        st.info(
            "Le traitement Odoo sera ajouté "
            "dans l'étape suivante."
        )


# ==========================================================
# PREPARATION
# ==========================================================

st.divider()

st.header(
    "⚙️ 2. Préparation du fichier Heidenhain"
)


if heidenhain_file is None:

    st.info(
        "👆 Importez le fichier Heidenhain "
        "pour commencer."
    )

else:

    st.markdown(
        """
        Le traitement va :

        - conserver toutes les lignes du fichier source ;
        - conserver uniquement les colonnes configurées ;
        - rechercher les statuts **VG** et **PG** ;
        - créer les `ID_SAV` ;
        - vérifier les doublons ;
        - ajouter les nouvelles lignes SAV à la fin ;
        - récupérer la **valeur calculée** de `Prix (SAV)` ;
        - créer un nouveau fichier Excel.
        """
    )

    # ======================================================
    # BOUTON
    # ======================================================

    if st.button(
        "🚀 Préparer le fichier Heidenhain",
        type="primary",
        use_container_width=True,
        key="process_heidenhain_button",
    ):

        try:

            result_bytes, stats = (
                process_heidenhain(
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
                    output_columns=OUTPUT_COLUMNS,
                )
            )

            # ==================================================
            # SAUVEGARDE SESSION
            # ==================================================

            st.session_state.result_bytes = (
                result_bytes
            )

            st.session_state.result_stats = (
                stats
            )

            st.session_state.result_file_name = (
                "Prix_Heidenhain_prepare.xlsx"
            )

            st.success(
                "✅ Fichier Heidenhain préparé avec succès."
            )

        except Exception as e:

            st.error(
                "❌ Le traitement a rencontré une erreur."
            )

            st.exception(e)


# ==========================================================
# RESULTAT
# ==========================================================

if (
    st.session_state.result_bytes is not None
    and st.session_state.result_stats is not None
):

    stats = (
        st.session_state.result_stats
    )

    st.divider()

    st.header(
        "📊 Résultat"
    )

    # ======================================================
    # STATISTIQUES
    # ======================================================

    c1, c2, c3, c4, c5 = (
        st.columns(5)
    )

    with c1:

        st.metric(
            "Lignes source",
            f"{stats['total_source_rows']:,}",
        )

    with c2:

        st.metric(
            "VG",
            f"{stats['vg_count']:,}",
        )

    with c3:

        st.metric(
            "PG",
            f"{stats['pg_count']:,}",
        )

    with c4:

        st.metric(
            "Nouvelles lignes SAV",
            f"{stats['rows_created']:,}",
        )

    with c5:

        st.metric(
            "ID_SAV déjà existants",
            f"{stats['duplicates']:,}",
        )

    # ======================================================
    # INFOS
    # ======================================================

    st.info(
        f"🔎 {stats['vg_pg_found']:,} "
        "ligne(s) VG / PG trouvée(s) dans le fichier complet."
    )

    if stats["empty_ids"] > 0:

        st.warning(
            f"⚠️ {stats['empty_ids']:,} "
            "ligne(s) VG / PG ignorée(s) "
            "car leur ID était vide."
        )

    # ======================================================
    # IDS CREES
    # ======================================================

    if stats["created_ids"]:

        st.subheader(
            "🆕 Nouvelles références SAV"
        )

        ids_df = pd.DataFrame(
            {
                "Nouveau ID": (
                    stats[
                        "created_ids"
                    ]
                )
            }
        )

        st.dataframe(
            ids_df,
            use_container_width=True,
            height=400,
        )

    else:

        st.warning(
            "Aucune nouvelle référence SAV "
            "n'a été créée."
        )

    # ======================================================
    # TELECHARGEMENT
    # ======================================================

    st.subheader(
        "⬇️ Fichier résultant"
    )

    st.download_button(
        label=(
            "⬇️ Télécharger le fichier "
            "Heidenhain préparé"
        ),
        data=(
            st.session_state.result_bytes
        ),
        file_name=(
            st.session_state.result_file_name
            or "Prix_Heidenhain_prepare.xlsx"
        ),
        mime=(
            "application/vnd.openxmlformats-"
            "officedocument.spreadsheetml.sheet"
        ),
        use_container_width=True,
        key="download_heidenhain",
    )


# ==========================================================
# PROCHAINE ETAPE
# ==========================================================

st.divider()

st.header(
    "➡️ Prochaine étape : rapprochement Odoo"
)

st.info(
    """
    Le fichier Heidenhain est maintenant préparé.

    Le traitement du fichier Odoo sera ajouté
    dans une prochaine étape.
    """
)
