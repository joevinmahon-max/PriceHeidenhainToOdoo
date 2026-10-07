# ==========================================================
# PREPARATION DES PRIX HEIDENHAIN
# JMA
#
# ETAPE 1 :
# - Fichier 1 = Prix Heidenhain
# - Feuille configurable
# - En-têtes sur ligne configurable
# - Données à partir d'une ligne configurable
# - Recherche des statuts VG / PG
# - Création des ID_SAV
# - Vérification des doublons
# - Création d'un nouveau fichier avec uniquement
#   les colonnes configurées
#
# ETAPE 2 ODOO :
# Prévue pour plus tard
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

STATUS_TO_DUPLICATE = {"VG", "PG"}

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
# PARAMETRES
# ==========================================================

st.sidebar.header("⚙️ Paramètres")


# ==========================================================
# FICHIER 1 — HEIDENHAIN
# ==========================================================

st.sidebar.subheader("📘 Fichier 1 — Prix Heidenhain")

heidenhain_sheet = st.sidebar.text_input(
    "Feuille Heidenhain",
    value="Distributeurs",
    help="Nom exact de la feuille à traiter.",
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


# ==========================================================
# COLONNES A CONSERVER
# ==========================================================

st.sidebar.divider()

st.sidebar.subheader("📋 Colonnes à conserver")

output_columns_text = st.sidebar.text_area(
    "Liste des colonnes",
    value="\n".join(DEFAULT_OUTPUT_COLUMNS),
    height=180,
    help=(
        "Une colonne par ligne. "
        "L'ordre sera conservé dans le fichier résultant."
    ),
)

OUTPUT_COLUMNS = [
    column.strip()
    for column in output_columns_text.splitlines()
    if column.strip()
]


# ==========================================================
# FICHIER 2 — ODOO
# ==========================================================

st.sidebar.divider()

st.sidebar.subheader("📗 Fichier 2 — Import Odoo")

odoo_sheet = st.sidebar.text_input(
    "Feuille Odoo",
    value="Sheet1",
    help=(
        "Nom de la feuille du fichier d'import Odoo. "
        "Le traitement Odoo sera ajouté ultérieurement."
    ),
)


# ==========================================================
# UTILITAIRES
# ==========================================================

def normalize_value(value):
    """
    Normalise une valeur pour les comparaisons.

    Exemple :
        ' vg ' -> 'VG'
        'abc123' -> 'ABC123'
        None -> ''
    """

    if value is None:
        return ""

    return str(value).strip().upper()


# ==========================================================
# LECTURE DES FEUILLES
# ==========================================================

def get_excel_sheet_names(uploaded_file):
    """
    Retourne les noms des feuilles du fichier Excel.
    """
    
    uploaded_file.seek(0)
    
    # Classeur principal : formules + styles
    workbook = openpyxl.load_workbook(
        uploaded_file,
        data_only=False,
    )
    
    if sheet_name not in workbook.sheetnames:
        available = ", ".join(workbook.sheetnames)
        workbook.close()
    
        raise ValueError(
            f"La feuille '{sheet_name}' n'existe pas.\n\n"
            f"Feuilles disponibles : {available}"
        )
    
    source_ws = workbook[sheet_name]
    
    # Deuxième ouverture : valeurs calculées par Excel
    uploaded_file.seek(0)
    
    workbook_values = openpyxl.load_workbook(
        uploaded_file,
        data_only=True,
    )
    
    if sheet_name not in workbook_values.sheetnames:
        workbook.close()
        workbook_values.close()
    
        raise ValueError(
            f"La feuille '{sheet_name}' n'existe pas dans "
            "le fichier des valeurs."
        )
    
    values_ws = workbook_values[sheet_name]



    sheet_names = workbook.sheetnames
    values_ws = workbook_values[sheet_name]


    workbook.close()

    return sheet_names


# ==========================================================
# RECHERCHE DES COLONNES
# ==========================================================

def find_columns_by_headers(
    worksheet,
    header_row,
    requested_columns,
):
    """
    Recherche les colonnes demandées sur une ligne donnée.

    Retourne :

        {
            "ID": 3,
            "Description": 4,
            ...
        }

    Les numéros de colonnes retournés sont ceux d'Excel
    et commencent donc à 1.
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

            original_name = requested_normalized[
                normalized
            ]

            if original_name not in found:

                found[original_name] = cell.column

    return found


# ==========================================================
# COPIE DU STYLE
# ==========================================================

def copy_cell_style(
    source_cell,
    target_cell,
):
    """
    Copie le style de la cellule source
    vers la cellule destination.
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
# COPIE D'UNE LIGNE
# ==========================================================

def copy_selected_row(
    source_ws,
    values_ws,
    target_ws,
    source_row,
    target_row,
    source_columns,
    OUTPUT_COLUMNS,
    new_id=None,
):

    """
    Copie uniquement les colonnes demandées
    d'une ligne source vers une ligne destination.

    Si new_id est fourni, l'ID de la nouvelle ligne
    est remplacé par new_id.
    """

    for output_col_index, column_name in enumerate(
        output_columns,
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

        # Pour Prix (SAV), on prend la valeur calculée
        # par Excel et non la formule.
        if column_name == "Prix (SAV)":
            target_cell.value = values_ws.cell(
                row=source_row,
                column=source_col_index,
            ).value
        else:
            target_cell.value = source_cell.value


        copy_cell_style(
            source_cell,
            target_cell,
        )

    # ------------------------------------------------------
    # Remplacement de l'ID
    # ------------------------------------------------------

    if new_id is not None:

        if "ID" not in output_columns:

            raise ValueError(
                "La colonne 'ID' doit être présente "
                "dans les colonnes à conserver."
            )

        id_output_column = (
            output_columns.index("ID") + 1
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
    Traite le fichier Heidenhain.

    Fonctionnement :

    1. Ouvre la feuille demandée.
    2. Recherche les colonnes sur header_row.
    3. Lit toutes les lignes à partir de data_start_row.
    4. Conserve toutes les lignes existantes.
    5. Recherche les statuts VG / PG.
    6. Pour chaque VG / PG :
       - récupère l'ID ;
       - crée ID_SAV ;
       - vérifie si ID_SAV existe déjà.
    7. Ajoute les nouvelles lignes SAV à la fin.
    8. Crée un nouveau fichier Excel.
    9. Retourne le fichier + les statistiques.
    """

    header_row = int(header_row)
    data_start_row = int(data_start_row)

    # ======================================================
    # VALIDATIONS
    # ======================================================

    if data_start_row <= header_row:

        raise ValueError(
            "La première ligne de données doit être "
            "supérieure à la ligne des noms de colonnes."
        )

    if not output_columns:

        raise ValueError(
            "Aucune colonne à conserver n'a été définie."
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
    # OUVERTURE DU FICHIER SOURCE
    # ======================================================

    uploaded_file.seek(0)

    workbook = openpyxl.load_workbook(
        uploaded_file,
        data_only=False,
    )

    if sheet_name not in workbook.sheetnames:

        available_sheets = ", ".join(
            workbook.sheetnames
        )

        workbook.close()

        raise ValueError(
            f"La feuille '{sheet_name}' n'existe pas.\n\n"
            f"Feuilles disponibles : {available_sheets}"
        )

    source_ws = workbook[
        sheet_name
    ]

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

    # ======================================================
    # NOMBRE TOTAL DE LIGNES
    # ======================================================

    max_row = source_ws.max_row

    total_rows = max(
        0,
        max_row - data_start_row + 1,
    )

    # ======================================================
    # CREATION DU NOUVEAU CLASSEUR
    # ======================================================

    output_workbook = (
        openpyxl.Workbook()
    )

    output_ws = (
        output_workbook.active
    )

    output_ws.title = sheet_name

    # ======================================================
    # CREATION DES EN-TETES
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
    # BARRE DE PROGRESSION
    # ======================================================

    progress = st.progress(
        0,
        text="Lecture du fichier Heidenhain...",
    )

    # ======================================================
    # ETAPE 1
    # RECUPERATION DE TOUS LES IDS EXISTANTS
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

        normalized_id = normalize_value(
            value
        )

        if normalized_id:

            existing_ids.add(
                normalized_id
            )

        # Progression tous les 5000 enregistrements.

        if (
            index % 5000 == 0
            or index == total_rows - 1
        ):

            percent = int(
                (
                    (index + 1)
                    / max(1, total_rows)
                )
                * 20
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
    # ======================================================

    rows_to_create = []

    vg_count = 0
    pg_count = 0

    empty_id_count = 0
    duplicate_count = 0

    output_data_row = 2

    progress.progress(
        20,
        text="Copie des lignes existantes...",
    )

    for index, source_row in enumerate(
        range(
            data_start_row,
            max_row + 1,
        )
    ):

        # --------------------------------------------------
        # COPIE DE LA LIGNE ORIGINALE
        # --------------------------------------------------

        copy_selected_row(
            source_ws=source_ws,
            target_ws=output_ws,
            source_row=source_row,
            target_row=output_data_row,
            source_columns=source_columns,
            output_columns=output_columns,
        )

        # --------------------------------------------------
        # HAUTEUR DE LIGNE
        # --------------------------------------------------

        output_ws.row_dimensions[
            output_data_row
        ].height = (
            source_ws.row_dimensions[
                source_row
            ].height
        )

        # --------------------------------------------------
        # STATUT
        # --------------------------------------------------

        status = normalize_value(
            source_ws.cell(
                row=source_row,
                column=status_col,
            ).value
        )

        # --------------------------------------------------
        # Si le statut n'est pas VG / PG
        # --------------------------------------------------

        if status not in STATUS_TO_DUPLICATE:

            output_data_row += 1

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
                        f"{total_rows:,}"
                    ),
                )

            continue

        # --------------------------------------------------
        # COMPTEUR VG / PG
        # --------------------------------------------------

        if status == "VG":

            vg_count += 1

        elif status == "PG":

            pg_count += 1

        # --------------------------------------------------
        # ID SOURCE
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
        # CREATION DU NOUVEL ID
        # --------------------------------------------------

        new_id = (
            f"{original_id}_SAV"
        )

        normalized_new_id = (
            normalize_value(new_id)
        )

        # --------------------------------------------------
        # VERIFICATION DU DOUBLON
        # --------------------------------------------------

        if normalized_new_id in existing_ids:

            duplicate_count += 1

            output_data_row += 1

            continue

        # --------------------------------------------------
        # RESERVATION DE L'ID
        # --------------------------------------------------

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
        # PROGRESSION
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
                    "Recherche des VG / PG... "
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

    next_output_row = (
        output_data_row
    )

    created_ids = []

    progress.progress(
        60,
        text=(
            f"Ajout de {rows_created:,} "
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
            target_ws=output_ws,
            source_row=source_row,
            target_row=next_output_row,
            source_columns=source_columns,
            output_columns=output_columns,
            new_id=new_id,
        )

        # --------------------------------------------------
        # HAUTEUR DE LIGNE
        # --------------------------------------------------

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
                * 30
            )

            progress.progress(
                min(percent, 90),
                text=(
                    "Création des lignes SAV... "
                    f"{index + 1:,} / "
                    f"{rows_created:,}"
                ),
            )

    # ======================================================
    # ETAPE 4
    # LARGEUR DES COLONNES
    # ======================================================

    progress.progress(
        92,
        text="Mise en forme du fichier...",
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
    # FIGER LES EN-TETES
    # ======================================================

    output_ws.freeze_panes = "A2"

    # ======================================================
    # ETAPE 5
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

    result = output.getvalue()

    # ======================================================
    # FERMETURE
    # ======================================================

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

        "output_total_rows": (
            next_output_row - 1
        ),

        "source_columns": source_columns,

        "status_column_number": status_col,

        "id_column_number": id_col,
    }

    return result, stats


# ==========================================================
# INTERFACE
# ==========================================================

st.header("📂 1. Import des fichiers")


col1, col2 = st.columns(2)


# ==========================================================
# FICHIER HEIDENHAIN
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

    if odoo_file is not None:

        st.success(
            f"✅ `{odoo_file.name}`"
        )

        st.info(
            "Le traitement du fichier Odoo sera "
            "ajouté à l'étape suivante."
        )


# ==========================================================
# PREPARATION DU FICHIER
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

        - lire la feuille configurée ;
        - utiliser la ligne configurée comme ligne d'en-têtes ;
        - lire les données à partir de la ligne configurée ;
        - conserver uniquement les colonnes configurées ;
        - compter tous les statuts **VG** et **PG** du fichier complet ;
        - créer un `ID_SAV` pour chaque VG / PG ;
        - vérifier que cet `ID_SAV` n'existe pas déjà ;
        - ajouter uniquement les nouvelles lignes SAV à la fin ;
        - générer un nouveau fichier Excel.
        """
    )

    # ======================================================
    # BOUTON DE TRAITEMENT
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

            # --------------------------------------------------
            # IMPORTANT :
            # Le résultat est conservé dans la session.
            # Un rerun Streamlit ne le fera donc pas disparaître.
            # --------------------------------------------------

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
# AFFICHAGE DU RESULTAT
#
# Cette partie est volontairement EN DEHORS
# du bouton de traitement.
#
# Ainsi, quand Streamlit fait un rerun,
# le résultat reste affiché.
# ==========================================================

if (
    st.session_state.result_bytes
    is not None
    and st.session_state.result_stats
    is not None
):

    stats = (
        st.session_state.result_stats
    )

    st.divider()

    st.header(
        "📊 Résultat du traitement"
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
    # INFORMATIONS COMPLEMENTAIRES
    # ======================================================

    if stats["empty_ids"] > 0:

        st.warning(
            f"⚠️ {stats['empty_ids']:,} ligne(s) "
            "VG/PG ont été ignorées car leur ID était vide."
        )

    st.info(
        f"🔎 {stats['vg_pg_found']:,} ligne(s) "
        "ont le statut VG ou PG dans le fichier complet."
    )

    # ======================================================
    # NOUVEAUX IDS
    # ======================================================

    if stats["created_ids"]:

        st.subheader(
            "🆕 Nouvelles références créées"
        )

        ids_df = pd.DataFrame(
            {
                "Nouveau ID": (
                    stats["created_ids"]
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
            "Aucune nouvelle référence SAV n'a été créée."
        )

    # ======================================================
    # TELECHARGEMENT
    #
    # IMPORTANT :
    # Le bouton utilise session_state.
    # Il ne relance PAS le traitement.
    # ======================================================

    st.subheader(
        "⬇️ Fichier résultant"
    )

    st.download_button(
        label=(
            "⬇️ Télécharger le fichier "
            "Heidenhain préparé"
        ),
        data=st.session_state.result_bytes,
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

    Le fichier `Sheet1` d'Odoo sera traité dans une
    deuxième étape, en utilisant les numéros de colonnes
    conformément à la structure de ton fichier d'import.
    """
)
