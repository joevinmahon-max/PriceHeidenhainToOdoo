# ==========================================================
# PREPARATION DES PRIX HEIDENHAIN
# JMA
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


# ==========================================================
# CONSTANTES
# ==========================================================

STATUS_TO_DUPLICATE = {"VG", "PG"}
PREVIEW_ROWS = 10

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
# PARAMETRES STREAMLIT
# ==========================================================

st.sidebar.header("⚙️ Paramètres")

# ----------------------------------------------------------
# FICHIER 1 — HEIDENHAIN
# ----------------------------------------------------------

st.sidebar.subheader("📘 Fichier 1 — Prix Heidenhain")

heidenhain_sheet = st.sidebar.text_input(
    "Feuille Heidenhain",
    value="Distributeurs",
    help="Nom de la feuille à traiter.",
)

heidenhain_header_row = st.sidebar.number_input(
    "Ligne des noms de colonnes",
    min_value=1,
    value=4,
    step=1,
    help="Les noms des colonnes se trouvent sur cette ligne.",
)

heidenhain_data_start_row = st.sidebar.number_input(
    "Première ligne de données",
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

# ----------------------------------------------------------
# COLONNES A CONSERVER
# ----------------------------------------------------------

st.sidebar.markdown("---")
st.sidebar.subheader("📋 Colonnes à conserver")

OUTPUT_COLUMNS_text = st.sidebar.text_area(
    "Colonnes",
    value="\n".join(DEFAULT_OUTPUT_COLUMNS),
    height=180,
    help=(
        "Une colonne par ligne. "
        "L'ordre saisi sera conservé dans le fichier résultant."
    ),
)

# Variable utilisée par tout le reste du programme
OUTPUT_COLUMNS = [
    column.strip()
    for column in OUTPUT_COLUMNS_text.splitlines()
    if column.strip()
]

# ----------------------------------------------------------
# FICHIER 2 — ODOO
# ----------------------------------------------------------

st.sidebar.markdown("---")
st.sidebar.subheader("📗 Fichier 2 — Import Odoo")

odoo_sheet = st.sidebar.text_input(
    "Feuille Odoo",
    value="Sheet1",
    help="Nom de la feuille du fichier d'import Odoo.",
)


# ==========================================================
# UTILITAIRES
# ==========================================================

def normalize_value(value):
    """
    Normalisation utilisée pour les comparaisons.
    """
    if value is None:
        return ""

    return str(value).strip().upper()


def get_excel_sheet_names(uploaded_file):
    """
    Retourne la liste des feuilles du fichier Excel.
    """
    uploaded_file.seek(0)

    workbook = openpyxl.load_workbook(
        uploaded_file,
        read_only=True,
        data_only=False,
    )

    names = workbook.sheetnames

    workbook.close()

    return names


def find_columns_by_headers(
    worksheet,
    header_row,
    requested_columns,
):
    """
    Recherche les colonnes demandées sur la ligne d'en-tête.

    Retour :
        {
            "ID": 3,
            "Description": 4,
            ...
        }
    """

    requested_normalized = {
        normalize_value(name): name
        for name in requested_columns
    }

    found = {}

    for cell in worksheet[header_row]:
        normalized = normalize_value(cell.value)

        if normalized in requested_normalized:

            original_name = requested_normalized[
                normalized
            ]

            if original_name not in found:
                found[original_name] = cell.column

    return found


def copy_cell_style(source_cell, target_cell):
    """
    Copie le style de la cellule source.
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


def copy_selected_row(
    source_ws,
    target_ws,
    source_row,
    target_row,
    source_columns,
    OUTPUT_COLUMNS,
    new_id=None,
):
    """
    Copie uniquement les colonnes demandées.

    new_id :
        Si renseigné, l'ID de la ligne copiée
        est remplacé par cette valeur.
    """

    for output_col_index, column_name in enumerate(
        OUTPUT_COLUMNS,
        start=1,
    ):

        source_col_index = source_columns[
            column_name
        ]

        source_cell = source_ws.cell(
            row=source_row,
            column=source_col_index,
        )

        target_cell = target_ws.cell(
            row=target_row,
            column=output_col_index,
        )

        target_cell.value = source_cell.value

        copy_cell_style(
            source_cell,
            target_cell,
        )

    # ------------------------------------------------------
    # Remplacement de l'ID
    # ------------------------------------------------------

    if new_id is not None:

        # On cherche la position de la colonne ID
        # dans le fichier résultant.

        id_output_index = (
            OUTPUT_COLUMNS.index("ID") + 1
        )

        target_ws.cell(
            row=target_row,
            column=id_output_index,
        ).value = new_id


# ==========================================================
# TRAITEMENT HEIDENHAIN
# ==========================================================

def process_heidenhain(
    uploaded_file,
    sheet_name,
    header_row,
    data_start_row,
    status_column,
    id_column,
    OUTPUT_COLUMNS,
):
    """
    Traitement du fichier Prix Heidenhain.

    Etapes :

    1. Lecture de la feuille source.
    2. Recherche des colonnes.
    3. Création d'un nouveau classeur.
    4. Copie uniquement des colonnes demandées.
    5. Recherche des statuts VG / PG.
    6. Création des ID_SAV absents.
    7. Ajout des lignes SAV à la fin.
    8. Sauvegarde du nouveau fichier.
    """

    header_row = int(header_row)
    data_start_row = int(data_start_row)

    # ------------------------------------------------------
    # Vérifications
    # ------------------------------------------------------

    if data_start_row <= header_row:
        raise ValueError(
            "La première ligne de données doit être "
            "supérieure à la ligne des en-têtes."
        )

    if "ID" not in OUTPUT_COLUMNS:
        raise ValueError(
            "La colonne 'ID' doit obligatoirement "
            "faire partie des colonnes conservées."
        )

    if status_column not in OUTPUT_COLUMNS:
        raise ValueError(
            f"La colonne '{status_column}' doit faire "
            "partie des colonnes conservées."
        )

    if id_column not in OUTPUT_COLUMNS:
        raise ValueError(
            f"La colonne '{id_column}' doit faire "
            "partie des colonnes conservées."
        )

    # ------------------------------------------------------
    # Chargement
    # ------------------------------------------------------

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
            f"La feuille '{sheet_name}' n'existe pas.\n\n"
            f"Feuilles disponibles : {available}"
        )

    source_ws = workbook[sheet_name]

    # ------------------------------------------------------
    # Recherche des colonnes
    # ------------------------------------------------------

    required_columns = list(
        dict.fromkeys(
            OUTPUT_COLUMNS
            + [
                status_column,
                id_column,
            ]
        )
    )

    source_columns = find_columns_by_headers(
        source_ws,
        header_row,
        required_columns,
    )

    missing_columns = [
        column
        for column in required_columns
        if column not in source_columns
    ]

    if missing_columns:

        workbook.close()

        raise ValueError(
            "Colonnes introuvables sur la ligne "
            f"{header_row} : "
            + ", ".join(missing_columns)
        )

    status_col = source_columns[
        status_column
    ]

    id_col = source_columns[
        id_column
    ]

    max_row = source_ws.max_row

    # ------------------------------------------------------
    # Nouveau classeur
    # ------------------------------------------------------

    output_workbook = openpyxl.Workbook()

    output_ws = (
        output_workbook.active
    )

    output_ws.title = sheet_name

    # ------------------------------------------------------
    # En-têtes
    # ------------------------------------------------------

    for output_col_index, column_name in enumerate(
        OUTPUT_COLUMNS,
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

    # ------------------------------------------------------
    # Nombre de lignes
    # ------------------------------------------------------

    total_rows = max(
        0,
        max_row - data_start_row + 1,
    )

    # ------------------------------------------------------
    # Barre de progression
    # ------------------------------------------------------

    progress = st.progress(
        0,
        text="Analyse du fichier Heidenhain...",
    )

    # ======================================================
    # ETAPE 1
    # RECUPERATION DES IDS EXISTANTS
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

        normalized = normalize_value(
            value
        )

        if normalized:
            existing_ids.add(
                normalized
            )

        if (
            index % 5000 == 0
            or index == total_rows - 1
        ):

            percent = 5 + int(
                (
                    (index + 1)
                    / max(1, total_rows)
                )
                * 15
            )

            progress.progress(
                min(percent, 20),
                text=(
                    "Recherche des IDs existants... "
                    f"{index + 1:,} / "
                    f"{total_rows:,}"
                ),
            )

    # ======================================================
    # ETAPE 2
    # COPIE DES LIGNES EXISTANTES
    # + DETECTION VG / PG
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

        # --------------------------------------------------
        # Copie de la ligne existante
        # --------------------------------------------------

        copy_selected_row(
            source_ws=source_ws,
            target_ws=output_ws,
            source_row=source_row,
            target_row=output_data_row,
            source_columns=source_columns,
            OUTPUT_COLUMNS=OUTPUT_COLUMNS,
        )

        # Conservation de la hauteur de ligne

        output_ws.row_dimensions[
            output_data_row
        ].height = (
            source_ws.row_dimensions[
                source_row
            ].height
        )

        # --------------------------------------------------
        # Lecture du statut
        # --------------------------------------------------

        status = normalize_value(
            source_ws.cell(
                row=source_row,
                column=status_col,
            ).value
        )

        # Si pas VG / PG :
        # aucune ligne supplémentaire.

        if status not in STATUS_TO_DUPLICATE:

            output_data_row += 1

            continue

        # --------------------------------------------------
        # Statistiques
        # --------------------------------------------------

        if status == "VG":
            vg_count += 1
        else:
            pg_count += 1

        # --------------------------------------------------
        # Lecture de l'ID
        # --------------------------------------------------

        original_id = source_ws.cell(
            row=source_row,
            column=id_col,
        ).value

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

        # --------------------------------------------------
        # Nouveau ID
        # --------------------------------------------------

        new_id = (
            f"{original_id}_SAV"
        )

        normalized_new_id = (
            normalize_value(new_id)
        )

        # --------------------------------------------------
        # Vérification doublon
        # --------------------------------------------------

        if (
            normalized_new_id
            in existing_ids
        ):

            duplicate_count += 1
            output_data_row += 1

            continue

        # On réserve immédiatement l'ID.
        #
        # Cela évite qu'un même ID_SAV
        # soit créé plusieurs fois.

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

        # --------------------------------------------------
        # Progression
        # --------------------------------------------------

        if (
            index % 5000 == 0
            or index == total_rows - 1
        ):

            percent = 20 + int(
                (
                    (index + 1)
                    / max(1, total_rows)
                )
                * 40
            )

            progress.progress(
                min(percent, 60),
                text=(
                    "Traitement des lignes... "
                    f"{index + 1:,} / "
                    f"{total_rows:,} "
                    f"| SAV à créer : "
                    f"{len(rows_to_create):,}"
                ),
            )

    # ======================================================
    # ETAPE 3
    # AJOUT DES LIGNES SAV
    # ======================================================

    rows_created = len(
        rows_to_create
    )

    progress.progress(
        60,
        text=(
            f"Création de "
            f"{rows_created:,} ligne(s) SAV..."
        ),
    )

    created_ids = []

    next_output_row = (
        output_data_row
    )

    for index, (
        source_row,
        new_id,
    ) in enumerate(
        rows_to_create
    ):

        copy_selected_row(
            source_ws=source_ws,
            target_ws=output_ws,
            source_row=source_row,
            target_row=next_output_row,
            source_columns=source_columns,
            OUTPUT_COLUMNS=OUTPUT_COLUMNS,
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

            percent = 60 + int(
                (
                    (index + 1)
                    / max(1, rows_created)
                )
                * 35
            )

            progress.progress(
                min(percent, 95),
                text=(
                    "Création des lignes SAV... "
                    f"{index + 1:,} / "
                    f"{rows_created:,}"
                ),
            )

    # ======================================================
    # ETAPE 4
    # LARGEURS DES COLONNES
    # ======================================================

    for output_col_index, column_name in enumerate(
        OUTPUT_COLUMNS,
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

        output_ws.column_dimensions[
            target_letter
        ].width = (
            source_ws.column_dimensions[
                source_letter
            ].width
        )

    # ------------------------------------------------------
    # Figer les en-têtes
    # ------------------------------------------------------

    output_ws.freeze_panes = "A2"

    # ======================================================
    # ETAPE 5
    # SAUVEGARDE
    # ======================================================

    progress.progress(
        97,
        text="Sauvegarde du fichier...",
    )

    output = BytesIO()

    output_workbook.save(
        output
    )

    output.seek(0)

    output_workbook.close()
    workbook.close()

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
        "source_max_row": max_row,
        "output_data_rows": (
            output_data_row - 1
        ),
        "output_total_rows": (
            next_output_row - 1
        ),
        "source_columns": source_columns,
        "status_column_number": status_col,
        "id_column_number": id_col,
    }

    return (
        output.getvalue(),
        stats,
    )


# ==========================================================
# INTERFACE
# ==========================================================

st.header("📂 1. Import des fichiers")

col1, col2 = st.columns(2)


# ==========================================================
# FICHIER 1
# ==========================================================

with col1:

    st.subheader(
        "📘 Fichier 1 — Prix Heidenhain"
    )

    heidenhain_file = st.file_uploader(
        "Importer le fichier des prix Heidenhain",
        type=[
            "xlsx",
            "xlsm",
        ],
        key="heidenhain_upload",
    )

    if heidenhain_file:

        st.success(
            f"✅ `{heidenhain_file.name}`"
        )

        try:

            sheet_names = (
                get_excel_sheet_names(
                    heidenhain_file
                )
            )

            if (
                heidenhain_sheet
                in sheet_names
            ):

                st.success(
                    f"✅ Feuille "
                    f"`{heidenhain_sheet}` trouvée."
                )

            else:

                st.error(
                    f"❌ La feuille "
                    f"`{heidenhain_sheet}` "
                    f"n'existe pas."
                )

                st.write(
                    "**Feuilles disponibles :**"
                )

                st.write(
                    sheet_names
                )

        except Exception as e:

            st.error(
                f"Erreur : {e}"
            )


# ==========================================================
# FICHIER 2
# ==========================================================

with col2:

    st.subheader(
        "📗 Fichier 2 — Import Odoo"
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

    if odoo_file:

        st.success(
            f"✅ `{odoo_file.name}`"
        )

        try:

            if odoo_file.name.lower().endswith(
                ".csv"
            ):

                odoo_preview = pd.read_csv(
                    odoo_file,
                    header=0,
                    sep=None,
                    engine="python",
                    nrows=PREVIEW_ROWS,
                )

            else:

                odoo_preview = pd.read_excel(
                    odoo_file,
                    sheet_name=odoo_sheet,
                    header=0,
                    nrows=PREVIEW_ROWS,
                )

            st.success(
                f"✅ Feuille "
                f"`{odoo_sheet}` trouvée."
            )

            st.write(
                "Aperçu du fichier Odoo :"
            )

            st.dataframe(
                odoo_preview,
                use_container_width=True,
            )

        except Exception as e:

            st.error(
                f"Erreur de lecture Odoo : {e}"
            )

def analyze_heidenhain(
    uploaded_file,
    sheet_name,
    header_row,
    data_start_row,
    status_column,
    id_column,
):
    """
    Analyse le fichier Heidenhain complet.

    Ne charge pas les données dans pandas.
    Parcourt directement la feuille Excel.

    Retourne :
        - nombre total de lignes
        - nombre de VG
        - nombre de PG
        - nombre d'IDs existants
        - nombre d'IDs vides
    """

    header_row = int(header_row)
    data_start_row = int(data_start_row)

    uploaded_file.seek(0)

    workbook = openpyxl.load_workbook(
        uploaded_file,
        read_only=True,
        data_only=False,
    )

    if sheet_name not in workbook.sheetnames:
        available = ", ".join(workbook.sheetnames)
        workbook.close()

        raise ValueError(
            f"La feuille '{sheet_name}' n'existe pas.\n\n"
            f"Feuilles disponibles : {available}"
        )

    ws = workbook[sheet_name]

    # ==========================================================
    # ANALYSE COMPLETE DU FICHIER
    # ==========================================================
    
    if heidenhain_file is not None:
    
        st.divider()
    
        st.header(
            "🔎 Analyse du fichier Heidenhain"
        )
    
        try:
    
            analysis = analyze_heidenhain(
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
    
            st.success(
                "✅ Fichier complet analysé."
            )
    
            c1, c2, c3, c4 = st.columns(4)
    
            with c1:
                st.metric(
                    "Lignes du fichier",
                    f"{analysis['total_rows']:,}",
                )
    
            with c2:
                st.metric(
                    "Statut VG",
                    f"{analysis['vg_count']:,}",
                )
    
            with c3:
                st.metric(
                    "Statut PG",
                    f"{analysis['pg_count']:,}",
                )
    
            with c4:
                st.metric(
                    "VG + PG",
                    f"{analysis['vg_pg_count']:,}",
                )
    
            if analysis["empty_ids"]:
    
                st.warning(
                    f"{analysis['empty_ids']:,} "
                    "ligne(s) ont un ID vide."
                )
    
        except Exception as e:
    
            st.error(
                f"Erreur lors de l'analyse du fichier : {e}"
            )
    
    
        # ------------------------------------------------------
        # Recherche des colonnes
        # ------------------------------------------------------
    
        source_columns = find_columns_by_headers(
            ws,
            header_row,
            [
                status_column,
                id_column,
            ],
        )
    
        missing_columns = [
            column
            for column in [
                status_column,
                id_column,
            ]
            if column not in source_columns
        ]
    
        if missing_columns:
            workbook.close()
    
            raise ValueError(
                "Colonnes introuvables sur la ligne "
                f"{header_row} : "
                + ", ".join(missing_columns)
            )
    
        status_col = source_columns[status_column]
        id_col = source_columns[id_column]
    
        # ------------------------------------------------------
        # Analyse complète
        # ------------------------------------------------------
    
        max_row = ws.max_row
    
        total_rows = max(
            0,
            max_row - data_start_row + 1,
        )
    
        vg_count = 0
        pg_count = 0
        empty_id_count = 0
        total_ids = 0
    
        existing_ids = set()
    
        progress = st.progress(
            0,
            text="Analyse complète du fichier Heidenhain...",
        )
    
        for index, row in enumerate(
            range(
                data_start_row,
                max_row + 1,
            )
        ):
    
            # --------------------------------------------------
            # ID
            # --------------------------------------------------
    
            id_value = ws.cell(
                row=row,
                column=id_col,
            ).value
    
            normalized_id = normalize_value(
                id_value
            )
    
            if normalized_id:
    
                existing_ids.add(
                    normalized_id
                )
    
                total_ids += 1
    
            else:
    
                empty_id_count += 1
    
            # --------------------------------------------------
            # STATUT
            # --------------------------------------------------
    
            status = normalize_value(
                ws.cell(
                    row=row,
                    column=status_col,
                ).value
            )
    
            if status == "VG":
    
                vg_count += 1
    
            elif status == "PG":
    
                pg_count += 1
    
            # --------------------------------------------------
            # Progression
            # --------------------------------------------------
    
            if (
                index % 5000 == 0
                or index == total_rows - 1
            ):
    
                percent = int(
                    (
                        (index + 1)
                        / max(1, total_rows)
                    )
                    * 100
                )
    
                progress.progress(
                    min(percent, 100),
                    text=(
                        "Analyse complète du fichier... "
                        f"{index + 1:,} / "
                        f"{total_rows:,}"
                    ),
                )
    
        workbook.close()
    
        progress.empty()
    
        return {
            "total_rows": total_rows,
            "vg_count": vg_count,
            "pg_count": pg_count,
            "vg_pg_count": (
                vg_count + pg_count
            ),
            "total_ids": total_ids,
            "empty_ids": empty_id_count,
            "existing_ids": existing_ids,
            "max_row": max_row,
            "status_column_number": status_col,
            "id_column_number": id_col,
        }


# ==========================================================
# TRAITEMENT
# ==========================================================

st.divider()

st.header(
    "⚙️ 3. Préparation du fichier Heidenhain"
)

if heidenhain_file is None:

    st.info(
        "👆 Importez le fichier Heidenhain "
        "pour commencer."
    )

else:

    st.markdown(
        """
        Le traitement effectue les opérations suivantes :

        - conserve uniquement les colonnes configurées ;
        - conserve toutes les lignes existantes ;
        - recherche les statuts **VG** et **PG** ;
        - récupère leur **ID** ;
        - crée `ID_SAV` ;
        - vérifie que `ID_SAV` n'existe pas déjà ;
        - ajoute uniquement les nouvelles références à la fin ;
        - crée un nouveau fichier Excel ;
        - ne modifie pas le fichier source.
        """
    )

    if st.button(
        "🚀 Préparer le fichier Heidenhain",
        type="primary",
        use_container_width=True,
    ):

        try:
            result_bytes, stats = process_heidenhain(
                uploaded_file=heidenhain_file,
                sheet_name=heidenhain_sheet,
                header_row=int(heidenhain_header_row),
                data_start_row=int(heidenhain_data_start_row),
                status_column=status_column,
                id_column=id_column,
                output_columns=OUTPUT_COLUMNS,
            )
            
            # Conserver le résultat après le rerun Streamlit
            st.session_state.result_bytes = result_bytes
            st.session_state.result_stats = stats


            st.success(
                "✅ Fichier Heidenhain préparé avec succès."
            )

            # ==================================================
            # STATISTIQUES
            # ==================================================

            st.subheader(
                "📊 Résultat"
            )

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
                    stats["vg_count"],
                )

            with c3:

                st.metric(
                    "PG",
                    stats["pg_count"],
                )

            with c4:

                st.metric(
                    "Nouvelles lignes",
                    stats["rows_created"],
                )

            with c5:

                st.metric(
                    "Déjà existantes",
                    stats["duplicates"],
                )

            if stats["empty_ids"]:

                st.warning(
                    f"{stats['empty_ids']:,} "
                    "ligne(s) VG/PG ignorée(s) "
                    "car l'ID était vide."
                )

            # ==================================================
            # IDS CREES
            # ==================================================

            if stats["created_ids"]:

                st.subheader(
                    "🆕 Nouvelles références"
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
                    "Aucune nouvelle ligne SAV "
                    "n'a été créée."
                )

            # ==================================================
            # TELECHARGEMENT
            # ==================================================

            st.subheader(
                "⬇️ Fichier résultant"
            )

            st.download_button(
                label=(
                    "⬇️ Télécharger le fichier "
                    "Heidenhain préparé"
                ),
                data=result_bytes,
                file_name=(
                    "Prix_Heidenhain_prepare.xlsx"
                ),
                mime=(
                    "application/vnd.openxmlformats-"
                    "officedocument.spreadsheetml.sheet"
                ),
                use_container_width=True,
            )

        except Exception as e:

            st.error(
                "❌ Le traitement a rencontré une erreur."
            )

            st.exception(e)


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

    À l'étape suivante, le fichier `Sheet1` d'Odoo
    sera traité séparément.

    Pour Odoo, nous utiliserons les **numéros de colonnes**
    plutôt que les noms de colonnes, conformément à la
    structure de ton fichier d'import.
    """
)
