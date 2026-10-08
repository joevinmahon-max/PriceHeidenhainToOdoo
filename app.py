# ==========================================================
# PREPARATION PRIX HEIDENHAIN + MISE EN FORME ODOO
# JMA
#
# ETAPE 1 :
# - Préparation du fichier Heidenhain
#
# ETAPE 2 :
# - Mise à jour / création des références Odoo
# - Recherche catégorie optimisée en mémoire
# - Compatible avec de gros fichiers
#
# AMELIORATIONS :
# - Boutons verrouillés pendant les calculs
# - Impossible de lancer deux fois le même traitement
# - Affichage immédiat "Calcul en cours"
# - Barre de progression visible
# - Messages d'étapes en direct
# - Déverrouillage automatique en cas d'erreur
# ==========================================================

import streamlit as st
import openpyxl
import time

from io import BytesIO
from copy import copy


# ==========================================================
# CONFIGURATION STREAMLIT
# ==========================================================

st.set_page_config(
    page_title="Préparation Heidenhain --> Odoo",
    page_icon="📦",
    layout="wide",
)

st.title("📦 Préparation des prix HEIDENHAIN --> Odoo")


# ==========================================================
# SESSION STATE
# ==========================================================

DEFAULT_STATE = {
    "heidenhain_result": None,
    "heidenhain_stats": None,

    "odoo_result": None,
    "odoo_stats": None,

    # ------------------------------------------------------
    # VERROUS DE CALCUL
    # ------------------------------------------------------

    "heidenhain_processing": False,
    "odoo_processing": False,
}

for key, value in DEFAULT_STATE.items():

    if key not in st.session_state:
        st.session_state[key] = value


# ==========================================================
# CONSTANTES
# ==========================================================

STATUS_TO_DUPLICATE = {
    "VG",
    "PG",
}

HEIDENHAIN_AVAILABLE_COLUMNS = [
    "ID",
    "Description",
    "Marque",
    "Statut",
    "Groupe Produit",
    "Prix (PPC)",
    "Prix (SAV)",
    "Date expiration",
    "Pays d'origine",
    "Code TVA",
    "Poids Net",
    "Poids Brut",
    "ROHS",
    "LP1 2027",
    "Coef",
    "Prix (PPC) 2027",
    "Prix (SAV) 2027",
    "Augmentation",
    "COEF2",
    "Prix HA",
]

DEFAULT_OUTPUT_COLUMNS = [
    "ID",
    "Description",
    "Marque",
    "Statut",
    "Groupe Produit",
    "Prix (PPC)",
    "Prix (SAV)",
    "Prix HA",
]


# ==========================================================
# UTILITAIRES
# ==========================================================

def normalize(value):

    if value is None:
        return ""

    return str(value).strip().upper()


def clean_reference(value):

    return normalize(value)


def is_sav(reference):

    return clean_reference(
        reference
    ).endswith("_SAV")


# ==========================================================
# COPIE STYLE
# ==========================================================

def copy_style_safe(source, target):

    try:
        target.font = copy(source.font)
    except Exception:
        pass

    try:
        target.fill = copy(source.fill)
    except Exception:
        pass

    try:
        target.border = copy(source.border)
    except Exception:
        pass

    try:
        target.alignment = copy(source.alignment)
    except Exception:
        pass

    try:
        target.protection = copy(source.protection)
    except Exception:
        pass

    try:
        target.number_format = source.number_format
    except Exception:
        pass


def copy_row_style(
    ws,
    source_row,
    target_row,
    max_col,
):

    for col in range(
        1,
        max_col + 1,
    ):

        source = ws.cell(
            row=source_row,
            column=col,
        )

        target = ws.cell(
            row=target_row,
            column=col,
        )

        copy_style_safe(
            source,
            target,
        )

    try:

        ws.row_dimensions[
            target_row
        ].height = (
            ws.row_dimensions[
                source_row
            ].height
        )

    except Exception:
        pass


# ==========================================================
# RECHERCHE COLONNE
# ==========================================================

def find_column(
    ws,
    header_row,
    header_name,
):

    wanted = normalize(
        header_name
    )

    for cell in ws[header_row]:

        if normalize(
            cell.value
        ) == wanted:

            return cell.column

    return None


def find_columns(
    ws,
    header_row,
    names,
):

    result = {}

    headers = {}

    for cell in ws[header_row]:

        value = normalize(
            cell.value
        )

        if value and value not in headers:

            headers[value] = cell.column

    for name in names:

        normalized_name = normalize(
            name
        )

        if normalized_name in headers:

            result[name] = headers[
                normalized_name
            ]

    return result


# ==========================================================
# NOMS FEUILLES
# ==========================================================

def get_sheet_names(
    uploaded_file,
):

    uploaded_file.seek(0)

    wb = openpyxl.load_workbook(
        uploaded_file,
        read_only=True,
        data_only=False,
    )

    names = wb.sheetnames

    wb.close()

    return names


# ==========================================================
# ETAPE 1
# COPIE LIGNE HEIDENHAIN
# ==========================================================

def copy_heidenhain_row(
    source_ws,
    values_ws,
    output_ws,
    source_row,
    output_row,
    source_columns,
    output_columns,
    new_id=None,
):

    for output_col, column_name in enumerate(
        output_columns,
        start=1,
    ):

        source_col = source_columns[
            column_name
        ]

        source_cell = source_ws.cell(
            row=source_row,
            column=source_col,
        )

        target_cell = output_ws.cell(
            row=output_row,
            column=output_col,
        )

        # --------------------------------------------------
        # Prix SAV
        # Prix HA
        # --------------------------------------------------

        if column_name == "Prix (SAV)":

            target_cell.value = (
                values_ws.cell(
                    row=source_row,
                    column=source_col,
                ).value
            )

        elif column_name == "Prix HA":

            target_cell.value = (
                values_ws.cell(
                    row=source_row,
                    column=source_col,
                ).value
            )

        else:

            target_cell.value = (
                source_cell.value
            )

        copy_style_safe(
            source_cell,
            target_cell,
        )

    if new_id is not None:

        if "ID" not in output_columns:

            raise ValueError(
                "La colonne ID doit être présente."
            )

        id_output_col = (
            output_columns.index("ID")
            + 1
        )

        output_ws.cell(
            row=output_row,
            column=id_output_col,
        ).value = new_id


# ==========================================================
# ETAPE 1
# HEIDENHAIN
# ==========================================================

def process_heidenhain(
    uploaded_file,
    sheet_name,
    header_row,
    data_start_row,
    status_column,
    id_column,
    output_columns,
    progress=None,
    status_display=None,
):

    header_row = int(
        header_row
    )

    data_start_row = int(
        data_start_row
    )

    if data_start_row <= header_row:

        raise ValueError(
            "La ligne de données doit être "
            "supérieure à la ligne des en-têtes."
        )

    if "ID" not in output_columns:

        raise ValueError(
            "La colonne ID est obligatoire."
        )

    # ======================================================
    # MESSAGE INITIAL
    # ======================================================

    if status_display is not None:

        status_display.info(
            "📂 Ouverture du fichier Heidenhain..."
        )

    if progress is not None:

        progress.progress(
            2,
            text="📂 Ouverture du fichier Heidenhain..."
        )

    # ======================================================
    # OUVERTURE FORMULES
    # ======================================================

    uploaded_file.seek(0)

    wb_formula = openpyxl.load_workbook(
        uploaded_file,
        data_only=False,
    )

    if sheet_name not in wb_formula.sheetnames:

        names = ", ".join(
            wb_formula.sheetnames
        )

        wb_formula.close()

        raise ValueError(
            f"Feuille '{sheet_name}' introuvable.\n"
            f"Feuilles disponibles : {names}"
        )

    source_ws = wb_formula[
        sheet_name
    ]

    if status_display is not None:

        status_display.info(
            "📊 Lecture des valeurs calculées..."
        )

    if progress is not None:

        progress.progress(
            5,
            text="📊 Lecture des valeurs calculées..."
        )

    # ======================================================
    # OUVERTURE VALEURS
    # ======================================================

    uploaded_file.seek(0)

    wb_values = openpyxl.load_workbook(
        uploaded_file,
        data_only=True,
    )

    values_ws = wb_values[
        sheet_name
    ]

    # ======================================================
    # COLONNES
    # ======================================================

    required = list(
        dict.fromkeys(
            output_columns
            + [
                status_column,
                id_column,
            ]
        )
    )

    source_columns = find_columns(
        source_ws,
        header_row,
        required,
    )

    missing = [
        x
        for x in required
        if x not in source_columns
    ]

    if missing:

        wb_formula.close()
        wb_values.close()

        raise ValueError(
            "Colonnes introuvables : "
            + ", ".join(missing)
        )

    status_col = source_columns[
        status_column
    ]

    id_col = source_columns[
        id_column
    ]

    # ======================================================
    # DERNIERE LIGNE REELLEMENT REMPLIE
    # ======================================================

    if status_display is not None:

        status_display.info(
            "🔎 Recherche de la dernière ligne..."
        )

    if progress is not None:

        progress.progress(
            8,
            text="🔎 Recherche des lignes à traiter..."
        )

    max_row = data_start_row - 1

    for row in range(
        data_start_row,
        source_ws.max_row + 1,
    ):

        id_value = source_ws.cell(
            row=row,
            column=id_col,
        ).value

        if id_value in (None, ""):

            break

        max_row = row

    total_rows = max(
        0,
        max_row - data_start_row + 1,
    )

    # ======================================================
    # NOUVEAU CLASSEUR
    # ======================================================

    if status_display is not None:

        status_display.info(
            f"📄 Préparation du nouveau fichier "
            f"({total_rows:,} lignes)..."
        )

    output_wb = (
        openpyxl.Workbook()
    )

    output_ws = (
        output_wb.active
    )

    output_ws.title = sheet_name

    # ======================================================
    # EN-TETES
    # ======================================================

    for output_col, column_name in enumerate(
        output_columns,
        start=1,
    ):

        source_header = source_ws.cell(
            row=header_row,
            column=source_columns[
                column_name
            ],
        )

        target_header = output_ws.cell(
            row=1,
            column=output_col,
        )

        target_header.value = (
            source_header.value
        )

        copy_style_safe(
            source_header,
            target_header,
        )

    # ======================================================
    # IDS EXISTANTS
    # ======================================================

    existing_ids = set()

    for row in range(
        data_start_row,
        max_row + 1,
    ):

        value = source_ws.cell(
            row=row,
            column=id_col,
        ).value

        normalized = normalize(
            value
        )

        if normalized:

            existing_ids.add(
                normalized
            )

    # ======================================================
    # TRAITEMENT
    # ======================================================

    rows_to_create = []

    vg_count = 0
    pg_count = 0
    duplicate_count = 0
    empty_id_count = 0

    output_row = 2

    if progress is not None:

        progress.progress(
            10,
            text=(
                f"⚙️ Traitement Heidenhain : "
                f"0 / {total_rows:,}"
            ),
        )

    for index, source_row in enumerate(
        range(
            data_start_row,
            max_row + 1,
        )
    ):

        copy_heidenhain_row(
            source_ws=source_ws,
            values_ws=values_ws,
            output_ws=output_ws,
            source_row=source_row,
            output_row=output_row,
            source_columns=source_columns,
            output_columns=output_columns,
        )

        output_ws.row_dimensions[
            output_row
        ].height = (
            source_ws.row_dimensions[
                source_row
            ].height
        )

        status = normalize(
            source_ws.cell(
                row=source_row,
                column=status_col,
            ).value
        )

        if status == "VG":

            vg_count += 1

        elif status == "PG":

            pg_count += 1

        if status in STATUS_TO_DUPLICATE:

            original_id = source_ws.cell(
                row=source_row,
                column=id_col,
            ).value

            original_id = (
                str(original_id).strip()
                if original_id is not None
                else ""
            )

            if original_id:

                new_id = (
                    f"{original_id}_SAV"
                )

                normalized_new = normalize(
                    new_id
                )

                if normalized_new in existing_ids:

                    duplicate_count += 1

                else:

                    existing_ids.add(
                        normalized_new
                    )

                    rows_to_create.append(
                        (
                            source_row,
                            new_id,
                        )
                    )

            else:

                empty_id_count += 1

        output_row += 1

        if (
            index % 10 == 0
            or index == total_rows - 1
        ):

            pct = int(
                (
                    (index + 1)
                    / max(
                        1,
                        total_rows,
                    )
                )
                * 75
            )

            pct = 10 + pct

            if progress is not None:

                progress.progress(
                    min(
                        pct,
                        85,
                    ),
                    text=(
                        f"⚙️ Heidenhain : "
                        f"{index + 1:,} / "
                        f"{total_rows:,}"
                        f" | VG : {vg_count:,}"
                        f" | PG : {pg_count:,}"
                    ),
                )

    # ======================================================
    # LIGNES SAV
    # ======================================================

    if status_display is not None:

        status_display.info(
            f"➕ Création des lignes SAV "
            f"({len(rows_to_create):,})..."
        )

    if progress is not None:

        progress.progress(
            87,
            text=(
                f"➕ Création des lignes SAV : "
                f"{len(rows_to_create):,}"
            ),
        )

    created_ids = []

    for source_row, new_id in rows_to_create:

        copy_heidenhain_row(
            source_ws=source_ws,
            values_ws=values_ws,
            output_ws=output_ws,
            source_row=source_row,
            output_row=output_row,
            source_columns=source_columns,
            output_columns=output_columns,
            new_id=new_id,
        )

        output_ws.row_dimensions[
            output_row
        ].height = (
            source_ws.row_dimensions[
                source_row
            ].height
        )

        created_ids.append(
            new_id
        )

        output_row += 1

    # ======================================================
    # LARGEURS
    # ======================================================

    if status_display is not None:

        status_display.info(
            "📐 Mise en forme des colonnes..."
        )

    if progress is not None:

        progress.progress(
            90,
            text="📐 Mise en forme du fichier..."
        )

    for output_col, column_name in enumerate(
        output_columns,
        start=1,
    ):

        source_col = source_columns[
            column_name
        ]

        source_letter = (
            openpyxl.utils.get_column_letter(
                source_col
            )
        )

        target_letter = (
            openpyxl.utils.get_column_letter(
                output_col
            )
        )

        width = (
            source_ws.column_dimensions[
                source_letter
            ].width
        )

        if width:

            output_ws.column_dimensions[
                target_letter
            ].width = width

    output_ws.freeze_panes = "A2"

    # ======================================================
    # SAUVEGARDE
    # ======================================================

    if status_display is not None:

        status_display.info(
            "💾 Création du fichier Excel..."
        )

    if progress is not None:

        progress.progress(
            94,
            text="💾 Création du fichier Excel..."
        )

    result = BytesIO()

    output_wb.save(
        result
    )

    result.seek(0)

    data = result.getvalue()

    output_wb.close()
    wb_formula.close()
    wb_values.close()

    stats = {
        "total_rows": total_rows,
        "vg": vg_count,
        "pg": pg_count,
        "vg_pg": (
            vg_count
            + pg_count
        ),
        "created": len(
            created_ids
        ),
        "duplicates": duplicate_count,
        "empty_ids": empty_id_count,
        "created_ids": created_ids,
    }

    if progress is not None:

        progress.progress(
            100,
            text="✅ Étape 1 terminée",
        )

    if status_display is not None:

        status_display.success(
            "✅ Fichier Heidenhain créé avec succès."
        )

    return data, stats


# ==========================================================
# ETAPE 2
# CHARGEMENT UNIQUE DES CATEGORIES
# ==========================================================

def load_category_mapping(
    category_file,
    category_sheet,
    category_id_col,
    category_search_col,
):

    category_file.seek(0)

    wb = openpyxl.load_workbook(
        category_file,
        data_only=True,
        read_only=True,
    )

    if category_sheet not in wb.sheetnames:

        names = ", ".join(
            wb.sheetnames
        )

        wb.close()

        raise ValueError(
            f"Feuille catégorie "
            f"'{category_sheet}' introuvable.\n"
            f"Feuilles disponibles : {names}"
        )

    ws = wb[
        category_sheet
    ]

    mapping = {}

    rows_loaded = 0

    for row in ws.iter_rows(
        min_row=1,
        values_only=True,
    ):

        max_index = max(
            category_id_col,
            category_search_col,
        ) - 1

        if max_index >= len(row):

            continue

        category_id = row[
            category_id_col - 1
        ]

        category_text = row[
            category_search_col - 1
        ]

        if (
            category_id is None
            or category_text is None
        ):

            continue

        category_id = str(
            category_id
        ).strip()

        category_text = str(
            category_text
        ).strip()

        if (
            not category_id
            or not category_text
        ):

            continue

        rows_loaded += 1

        normalized_text = normalize(
            category_text
        )

        if normalized_text not in mapping:

            mapping[
                normalized_text
            ] = category_id

        if "-" in category_text:

            first_part = (
                category_text
                .split("-", 1)[0]
                .strip()
            )

            if first_part:

                normalized_first = normalize(
                    first_part
                )

                if normalized_first not in mapping:

                    mapping[
                        normalized_first
                    ] = category_id

        if "–" in category_text:

            first_part = (
                category_text
                .split("–", 1)[0]
                .strip()
            )

            if first_part:

                normalized_first = normalize(
                    first_part
                )

                if normalized_first not in mapping:

                    mapping[
                        normalized_first
                    ] = category_id

    wb.close()

    return mapping, rows_loaded


# ==========================================================
# RECHERCHE CATEGORIE RAPIDE
# ==========================================================

def find_category_id_fast(
    groupe_produit,
    category_mapping,
):

    if groupe_produit is None:

        return None

    group = str(
        groupe_produit
    ).strip()

    if not group:

        return None

    normalized_group = normalize(
        group
    )

    category_id = category_mapping.get(
        normalized_group
    )

    if category_id is not None:

        return category_id

    if "-" in group:

        first_part = (
            group
            .split("-", 1)[0]
            .strip()
        )

        if first_part:

            category_id = category_mapping.get(
                normalize(first_part)
            )

            if category_id is not None:

                return category_id

    if "–" in group:

        first_part = (
            group
            .split("–", 1)[0]
            .strip()
        )

        if first_part:

            category_id = category_mapping.get(
                normalize(first_part)
            )

            if category_id is not None:

                return category_id

    return None


# ==========================================================
# ETAPE 2
# TRAITEMENT ODOO
# ==========================================================

def process_odoo(
    heidenhain_file,
    odoo_file,
    category_file,

    heidenhain_sheet,
    odoo_sheet,
    category_sheet,

    heidenhain_header_row,
    odoo_header_row,

    category_id_col,
    category_search_col,

    heidenhain_id_column,
    heidenhain_status_column,
    heidenhain_group_column,
    heidenhain_ppc_column,
    heidenhain_sav_column,
    heidenhain_description,
    heidenhain_marque,
    heidenhain_prixHA,

    odoo_reference_column,
    odoo_sales_status_column,
    odoo_price_column,
    odoo_barcode_column,
    odoo_supplier_column,
    odoo_purchase_column,
    odoo_sale_column,
    odoo_type_column,
    odoo_invoice_policy_column,
    odoo_category_column,
    odoo_name,
    odoo_brand_description,
    odoo_prixHA,

    progress=None,
    status_display=None,
):

    # ======================================================
    # CHRONOMETRES
    # ======================================================

    start_total = time.perf_counter()

    timer_categories = st.empty()
    timer_heidenhain = st.empty()
    timer_odoo = st.empty()
    timer_index = st.empty()
    timer_main = st.empty()

    # ======================================================
    # 1. CATEGORIES
    # ======================================================

    start_categories = time.perf_counter()

    if status_display is not None:

        status_display.info(
            "📂 Chargement des catégories..."
        )

    if progress is not None:

        progress.progress(
            2,
            text="📂 Chargement des catégories..."
        )

    category_mapping, category_rows = (
        load_category_mapping(
            category_file=category_file,
            category_sheet=category_sheet,
            category_id_col=int(
                category_id_col
            ),
            category_search_col=int(
                category_search_col
            ),
        )
    )

    time_categories = (
        time.perf_counter()
        - start_categories
    )

    if progress is not None:

        progress.progress(
            10,
            text=(
                f"📂 Catégories chargées : "
                f"{category_rows:,}"
            ),
        )

    # ======================================================
    # 2. HEIDENHAIN
    # ======================================================

    start_heidenhain = time.perf_counter()

    if status_display is not None:

        status_display.info(
            "📘 Chargement du fichier Heidenhain..."
        )

    if progress is not None:

        progress.progress(
            15,
            text="📘 Chargement du fichier Heidenhain..."
        )

    heidenhain_file.seek(0)

    wb_h = openpyxl.load_workbook(
        heidenhain_file,
        data_only=True,
        read_only=True,
    )

    if heidenhain_sheet not in wb_h.sheetnames:

        names = ", ".join(
            wb_h.sheetnames
        )

        wb_h.close()

        raise ValueError(
            f"Feuille Heidenhain "
            f"'{heidenhain_sheet}' "
            f"introuvable.\n"
            f"Feuilles disponibles : {names}"
        )

    h_ws = wb_h[
        heidenhain_sheet
    ]

    time_heidenhain_load = (
        time.perf_counter()
        - start_heidenhain
    )

    # ======================================================
    # 3. COLONNES HEIDENHAIN
    # ======================================================

    h_column_names = [
        heidenhain_id_column,
        heidenhain_status_column,
        heidenhain_group_column,
        heidenhain_ppc_column,
        heidenhain_sav_column,
        heidenhain_description,
        heidenhain_marque,
        heidenhain_prixHA,
    ]

    h_columns = find_columns(
        h_ws,
        1,
        h_column_names,
    )

    missing_h = [
        x
        for x in h_column_names
        if x not in h_columns
    ]

    if missing_h:

        wb_h.close()

        raise ValueError(
            "Colonnes introuvables "
            "dans le fichier 3 : "
            + ", ".join(missing_h)
        )

    if progress is not None:

        progress.progress(
            20,
            text="📘 Fichier Heidenhain chargé..."
        )

    # ======================================================
    # 4. ODOO
    # ======================================================

    start_odoo_load = time.perf_counter()

    if status_display is not None:

        status_display.info(
            "📗 Chargement du fichier Odoo..."
        )

    if progress is not None:

        progress.progress(
            25,
            text="📗 Chargement du fichier Odoo..."
        )

    odoo_file.seek(0)

    wb_o = openpyxl.load_workbook(
        odoo_file,
        data_only=False,
    )

    if odoo_sheet not in wb_o.sheetnames:

        names = ", ".join(
            wb_o.sheetnames
        )

        wb_o.close()
        wb_h.close()

        raise ValueError(
            f"Feuille Odoo "
            f"'{odoo_sheet}' "
            f"introuvable.\n"
            f"Feuilles disponibles : {names}"
        )

    ws = wb_o[
        odoo_sheet
    ]

    time_odoo_load = (
        time.perf_counter()
        - start_odoo_load
    )

    # ======================================================
    # 5. COLONNES ODOO
    # ======================================================

    if status_display is not None:

        status_display.info(
            "🔎 Recherche des colonnes Odoo..."
        )

    odoo_column_names = [
        odoo_reference_column,
        odoo_sales_status_column,
        odoo_price_column,
        odoo_barcode_column,
        odoo_supplier_column,
        odoo_purchase_column,
        odoo_sale_column,
        odoo_type_column,
        odoo_invoice_policy_column,
        odoo_category_column,
        odoo_name,
        odoo_brand_description,
        odoo_prixHA,
    ]

    odoo_columns = find_columns(
        ws,
        int(odoo_header_row),
        odoo_column_names,
    )

    missing_odoo = [
        x
        for x in odoo_column_names
        if x not in odoo_columns
    ]

    if missing_odoo:

        wb_o.close()
        wb_h.close()

        raise ValueError(
            "Colonnes introuvables "
            "dans Odoo : "
            + ", ".join(missing_odoo)
        )

    if progress is not None:

        progress.progress(
            30,
            text="📗 Fichier Odoo chargé..."
        )

    # ======================================================
    # 6. INDEX ODOO
    # ======================================================

    start_index = time.perf_counter()

    if status_display is not None:

        status_display.info(
            "🔎 Création de l'index des références Odoo..."
        )

    if progress is not None:

        progress.progress(
            35,
            text="🔎 Indexation des références Odoo..."
        )

    reference_index = {}

    reference_col = odoo_columns[
        odoo_reference_column
    ]

    max_odoo_row = ws.max_row

    for row in range(
        int(odoo_header_row) + 1,
        max_odoo_row + 1,
    ):

        value = ws.cell(
            row=row,
            column=reference_col,
        ).value

        ref = clean_reference(
            value
        )

        if ref:

            if ref not in reference_index:

                reference_index[
                    ref
                ] = row

    time_index = (
        time.perf_counter()
        - start_index
    )

    if progress is not None:

        progress.progress(
            40,
            text=(
                "🔎 Index Odoo créé : "
                f"{len(reference_index):,} références"
            ),
        )

    # ======================================================
    # 7. COLONNES
    # ======================================================

    col_ref = odoo_columns[
        odoo_reference_column
    ]

    col_status = odoo_columns[
        odoo_sales_status_column
    ]

    col_price = odoo_columns[
        odoo_price_column
    ]

    col_barcode = odoo_columns[
        odoo_barcode_column
    ]

    col_supplier = odoo_columns[
        odoo_supplier_column
    ]

    col_purchase = odoo_columns[
        odoo_purchase_column
    ]

    col_sale = odoo_columns[
        odoo_sale_column
    ]

    col_type = odoo_columns[
        odoo_type_column
    ]

    col_invoice = odoo_columns[
        odoo_invoice_policy_column
    ]

    col_category = odoo_columns[
        odoo_category_column
    ]

    col_prixHA = odoo_columns[
        odoo_prixHA
    ]

    # ======================================================
    # 8. STATISTIQUES
    # ======================================================

    updated = 0
    created = 0

    category_found = 0
    category_missing = 0

    references_processed = 0
    empty_references = 0

    # ======================================================
    # 9. PARCOURS HEIDENHAIN
    # ======================================================

    heidenhain_data_start = 2

    total_h_rows = max(
        0,
        h_ws.max_row
        - heidenhain_data_start
        + 1,
    )

    if status_display is not None:

        status_display.info(
            f"⚙️ Traitement de "
            f"{total_h_rows:,} lignes Heidenhain..."
        )

    if progress is not None:

        progress.progress(
            45,
            text=(
                f"⚙️ Traitement : "
                f"0 / {total_h_rows:,}"
            ),
        )

    # ======================================================
    # COLONNES HEIDENHAIN
    # ======================================================

    h_id_col = h_columns[
        heidenhain_id_column
    ]

    h_status_col = h_columns[
        heidenhain_status_column
    ]

    h_group_col = h_columns[
        heidenhain_group_column
    ]

    h_ppc_col = h_columns[
        heidenhain_ppc_column
    ]

    h_sav_col = h_columns[
        heidenhain_sav_column
    ]

    h_prixHA_col = h_columns[
        heidenhain_prixHA
    ]

    # ======================================================
    # BOUCLE PRINCIPALE
    # ======================================================

    start_main_loop = time.perf_counter()

    next_empty_row = (
        int(odoo_header_row) + 1
    )

    while ws.cell(
        row=next_empty_row,
        column=reference_col,
    ).value not in (None, ""):

        next_empty_row += 1

    for index, row_values in enumerate(
        h_ws.iter_rows(
            min_row=heidenhain_data_start,
            values_only=True,
        )
    ):

        try:

            reference = row_values[
                h_id_col - 1
            ]

        except IndexError:

            continue

        if reference in (None, ""):

            break

        # --------------------------------------------------
        # Lecture rapide
        # --------------------------------------------------

        try:

            status = row_values[
                h_status_col - 1
            ]

            groupe = row_values[
                h_group_col - 1
            ]

            ppc = row_values[
                h_ppc_col - 1
            ]

            sav = row_values[
                h_sav_col - 1
            ]

            prixHA = row_values[
                h_prixHA_col - 1
            ]

        except IndexError:

            continue

        # --------------------------------------------------
        # Reference
        # --------------------------------------------------

        reference = str(
            reference
        ).strip()

        if not reference:

            empty_references += 1

            continue

        references_processed += 1

        normalized_ref = clean_reference(
            reference
        )

        # ==================================================
        # RECHERCHE ODOO
        # ==================================================

        if normalized_ref in reference_index:

            target_row = (
                reference_index[
                    normalized_ref
                ]
            )

            updated += 1

        else:

            target_row = next_empty_row

            next_empty_row += 1

            reference_index[
                normalized_ref
            ] = target_row

            created += 1

        # ==================================================
        # DESCRIPTION
        # ==================================================

        description = row_values[
            h_columns[
                heidenhain_description
            ] - 1
        ]

        name_cell = ws.cell(
            row=target_row,
            column=odoo_columns[
                odoo_name
            ],
        )

        if name_cell.value in (
            None,
            "",
        ):

            name_cell.value = description

        # ==================================================
        # MARQUE
        # ==================================================

        marque = row_values[
            h_columns[
                heidenhain_marque
            ] - 1
        ]

        brand_cell = ws.cell(
            row=target_row,
            column=odoo_columns[
                odoo_brand_description
            ],
        )

        if brand_cell.value in (
            None,
            "",
        ):

            brand_cell.value = marque

        # ==================================================
        # PRIX
        # ==================================================

        reference_is_sav = (
            normalized_ref.endswith("_SAV")
        )

        if reference_is_sav:

            price = sav

        else:

            price = ppc

        # ==================================================
        # CATEGORIE
        # ==================================================

        category_id = (
            find_category_id_fast(
                groupe,
                category_mapping,
            )
        )

        if category_id is not None:

            category_found += 1

        else:

            category_missing += 1

        # ==================================================
        # ECRITURE
        # ==================================================

        ws.cell(
            row=target_row,
            column=col_ref,
        ).value = reference

        ws.cell(
            row=target_row,
            column=col_status,
        ).value = status

        ws.cell(
            row=target_row,
            column=col_price,
        ).value = price

        ws.cell(
            row=target_row,
            column=col_prixHA,
        ).value = prixHA

        # ==================================================
        # CODE-BARRES
        # ==================================================

        if reference_is_sav:

            ws.cell(
                row=target_row,
                column=col_barcode,
            ).value = ""

        else:

            ws.cell(
                row=target_row,
                column=col_barcode,
            ).value = (
                f"I {reference}"
            )

        # ==================================================
        # AUTRES CHAMPS
        # ==================================================

        ws.cell(
            row=target_row,
            column=col_supplier,
        ).value = (
            "HEIDENHAIN FRANCE"
        )

        ws.cell(
            row=target_row,
            column=col_purchase,
        ).value = "VRAI"

        ws.cell(
            row=target_row,
            column=col_sale,
        ).value = "VRAI"

        ws.cell(
            row=target_row,
            column=col_type,
        ).value = "Consommable"

        ws.cell(
            row=target_row,
            column=col_invoice,
        ).value = (
            "Quantités livrées"
        )

        if category_id is not None:

            ws.cell(
                row=target_row,
                column=col_category,
            ).value = category_id

        # ==================================================
        # PROGRESSION
        # ==================================================

        if (
            index % 10 == 0
            or index == total_h_rows - 1
        ):

            pct = 45 + int(
                (
                    (index + 1)
                    / max(
                        1,
                        total_h_rows,
                    )
                )
                * 45
            )

            progress_value = min(
                pct,
                90,
            )

            if progress is not None:

                progress.progress(
                    progress_value,
                    text=(
                        f"⚙️ Odoo : "
                        f"{index + 1:,} / "
                        f"{total_h_rows:,}"
                        f" | Mise à jour : "
                        f"{updated:,}"
                        f" | Création : "
                        f"{created:,}"
                        f" | Catégories : "
                        f"{category_found:,}"
                    ),
                )

        # ==================================================
        # INFOS TEMPS
        # ==================================================

        if index % 100 == 0:

            now = time.perf_counter()

            timer_categories.info(
                f"⏱️ Catégories : "
                f"{time_categories:.2f} s"
            )

            timer_heidenhain.info(
                f"⏱️ Chargement Heidenhain : "
                f"{time_heidenhain_load:.2f} s"
            )

            timer_odoo.info(
                f"⏱️ Chargement Odoo : "
                f"{time_odoo_load:.2f} s"
            )

            timer_index.info(
                f"⏱️ Index Odoo : "
                f"{time_index:.2f} s "
                f"({len(reference_index):,} références)"
            )

            timer_main.info(
                f"⏱️ Boucle principale : "
                f"{now - start_main_loop:.2f} s"
            )

    # ======================================================
    # SAUVEGARDE
    # ======================================================

    if status_display is not None:

        status_display.info(
            "💾 Création du fichier Excel final..."
        )

    if progress is not None:

        progress.progress(
            92,
            text="💾 Création du fichier Excel..."
        )

    output = BytesIO()

    wb_o.save(
        output
    )

    output.seek(0)

    result = output.getvalue()

    # ======================================================
    # FERMETURE
    # ======================================================

    wb_o.close()
    wb_h.close()

    total_time = (
        time.perf_counter()
        - start_total
    )

    if progress is not None:

        progress.progress(
            100,
            text="✅ Étape 2 terminée",
        )

    if status_display is not None:

        status_display.success(
            "✅ Fichier Odoo préparé avec succès."
        )

    timer_main.success(
        f"⏱️ Temps total : {total_time:.2f} s"
    )

    # ======================================================
    # STATISTIQUES
    # ======================================================

    stats = {
        "processed": references_processed,
        "updated": updated,
        "created": created,
        "category_found": category_found,
        "category_missing": category_missing,
        "empty_references": empty_references,
        "category_rows": category_rows,
        "odoo_indexed": len(
            reference_index
        ),
    }

    return result, stats


# ==========================================================
# SIDEBAR
# ==========================================================

st.sidebar.header(
    "⚙️ Paramètres"
)


# ==========================================================
# HEIDENHAIN
# ==========================================================

st.sidebar.subheader(
    "📋 Colonnes à conserver à l'étape 1"
)

heidenhain_output_columns = st.sidebar.multiselect(
    "Sélectionner les colonnes à garder",
    options=HEIDENHAIN_AVAILABLE_COLUMNS,
    default=DEFAULT_OUTPUT_COLUMNS,
    help=(
        "Ces colonnes seront conservées dans le fichier "
        "Heidenhain préparé."
    ),
)

if "ID" not in heidenhain_output_columns:

    st.sidebar.error(
        "⚠️ La colonne ID est obligatoire."
    )

st.sidebar.subheader(
    "📘 Fichier Heidenhain"
)

heidenhain_sheet = st.sidebar.text_input(
    "Feuille Heidenhain",
    "Distributeurs",
)

heidenhain_header_row = st.sidebar.number_input(
    "Ligne des en-têtes Heidenhain",
    min_value=1,
    value=4,
)

status_column = "Statut"

id_column = "ID"

group_column = "Groupe Produit"

ppc_column = "Prix (PPC)"

sav_column = "Prix (SAV)"

heidenhain_prixHA = "Prix HA"


# ==========================================================
# ODOO
# ==========================================================

st.sidebar.divider()

st.sidebar.subheader(
    "📗 Colonnes Odoo"
)

odoo_sheet = st.sidebar.text_input(
    "Feuille Odoo",
    "Sheet1",
)

odoo_header_row = st.sidebar.number_input(
    "Ligne des en-têtes Odoo",
    min_value=1,
    value=1,
)

odoo_reference = st.sidebar.text_input(
    "Référence",
    "Référence interne",
)

odoo_name = st.sidebar.text_input(
    "Nom",
    "Nom",
)

odoo_brand_description = st.sidebar.text_input(
    "Marque",
    "Marque",
)

odoo_sales_status = st.sidebar.text_input(
    "Sales Status",
    "Sales Status",
)

odoo_price = st.sidebar.text_input(
    "Prix de vente",
    "Prix de vente",
)

odoo_barcode = st.sidebar.text_input(
    "Code-barres",
    "Code-barres",
)

odoo_supplier = st.sidebar.text_input(
    "Fournisseur",
    "Fournisseurs/Fournisseur",
)

odoo_purchase = st.sidebar.text_input(
    "Peut être acheté",
    "Peut être acheté",
)

odoo_sale = st.sidebar.text_input(
    "Peut être vendu",
    "Peut être vendu",
)

odoo_type = st.sidebar.text_input(
    "Type de produit",
    "Type de produit",
)

odoo_invoice_policy = st.sidebar.text_input(
    "Politique de facturation",
    "Politique de facturation",
)

odoo_category = st.sidebar.text_input(
    "Catégorie de produits/ID",
    "Catégorie de produits/ID",
)

odoo_prixHA = st.sidebar.text_input(
    "Fournisseurs/Prix",
    "Fournisseurs/Prix",
)


# ==========================================================
# CATEGORIE PRODUIT
# ==========================================================

st.sidebar.divider()

st.sidebar.subheader(
    "📂 Catégorie de produit"
)

category_sheet = st.sidebar.text_input(
    "Feuille catégorie",
    "Sheet1",
)

category_id_col = st.sidebar.number_input(
    "Colonne ID catégorie",
    min_value=1,
    value=2,
)

category_search_col = st.sidebar.number_input(
    "Colonne recherche catégorie",
    min_value=1,
    value=3,
)


# ==========================================================
# IMPORT FICHIERS
# ==========================================================

st.header(
    "📂 Fichiers"
)

col1, col2, col3 = st.columns(3)


with col1:

    st.subheader(
        "📘 Prix Heidenhain"
    )

    heidenhain_file = st.file_uploader(
        "Fichier Prix Heidenhain",
        type=[
            "xlsx",
            "xlsm",
        ],
        key="heidenhain_file",
    )


with col2:

    st.subheader(
        "📗 Fichier Odoo"
    )

    odoo_file = st.file_uploader(
        "Fichier d'import Odoo",
        type=[
            "xlsx",
            "xlsm",
        ],
        key="odoo_file",
    )


with col3:

    st.subheader(
        "📂 Catégorie de produit"
    )

    category_file = st.file_uploader(
        "Fichier Catégorie de produit",
        type=[
            "xlsx",
            "xlsm",
        ],
        key="category_file",
    )


# ==========================================================
# ETAPE 1
# ==========================================================

st.divider()

st.header(
    "1️⃣ Création du fichier Heidenhain"
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

- conserver toutes les lignes du fichier ;
- conserver uniquement les colonnes sélectionnées ;
- récupérer la **valeur calculée** de `Prix (SAV)` ;
- rechercher les statuts **VG** et **PG** ;
- créer les `ID_SAV` ;
- ignorer les `ID_SAV` déjà existants ;
- ajouter les nouvelles lignes SAV à la fin ;
- créer un nouveau fichier Excel propre.
        """
    )

    # ======================================================
    # CALCUL EN COURS
    # ======================================================

    if st.session_state.heidenhain_processing:

        # --------------------------------------------------
        # BOUTON VERROUILLE
        # --------------------------------------------------

        st.button(
            "⏳ Calcul Heidenhain en cours...",
            disabled=True,
            use_container_width=True,
            key="create_heidenhain_locked",
        )

        st.warning(
            "⏳ **Le calcul Heidenhain est en cours.** "
            "Veuillez patienter et ne pas relancer le traitement."
        )

        progress_heidenhain = st.progress(
            0,
            text="Initialisation du calcul Heidenhain..."
        )

        status_heidenhain = st.empty()

        try:

            result, stats = (
                process_heidenhain(
                    uploaded_file=heidenhain_file,
                    sheet_name=heidenhain_sheet,
                    header_row=heidenhain_header_row,
                    data_start_row=(
                        heidenhain_header_row
                        + 1
                    ),
                    status_column=status_column,
                    id_column=id_column,
                    output_columns=(
                        heidenhain_output_columns
                    ),
                    progress=progress_heidenhain,
                    status_display=status_heidenhain,
                )
            )

            # --------------------------------------------------
            # SAUVEGARDE
            # --------------------------------------------------

            st.session_state.heidenhain_result = (
                result
            )

            st.session_state.heidenhain_stats = (
                stats
            )

            # --------------------------------------------------
            # DEVERROUILLAGE
            # --------------------------------------------------

            st.session_state.heidenhain_processing = False

            st.success(
                "✅ Fichier Heidenhain créé."
            )

            st.rerun()

        except Exception as e:

            # --------------------------------------------------
            # TOUJOURS DEVERROUILLER EN CAS D'ERREUR
            # --------------------------------------------------

            st.session_state.heidenhain_processing = False

            st.error(
                f"❌ Erreur : {e}"
            )

            st.exception(e)

    # ======================================================
    # BOUTON NORMAL
    # ======================================================

    else:

        if st.button(
            "🚀 Créer le fichier Heidenhain",
            type="primary",
            use_container_width=True,
            key="create_heidenhain",
        ):

            # --------------------------------------------------
            # VERROUILLAGE
            # --------------------------------------------------

            st.session_state.heidenhain_processing = True

            # --------------------------------------------------
            # FORCE STREAMLIT A RAFRAICHIR L'INTERFACE
            # AVANT LE GROS CALCUL
            # --------------------------------------------------

            st.rerun()


# ==========================================================
# RESULTAT ETAPE 1
# ==========================================================

if st.session_state.heidenhain_result:

    stats = (
        st.session_state.heidenhain_stats
    )

    st.success(
        f"Fichier créé : "
        f"{stats['created']:,} "
        f"nouvelles lignes SAV."
    )

    st.download_button(
        "⬇️ Télécharger le fichier Heidenhain",
        data=(
            st.session_state.heidenhain_result
        ),
        file_name=(
            "ETAPE 1 -- Products_Heidenhain_Prepare.xlsx"
        ),
        mime=(
            "application/vnd.openxmlformats-"
            "officedocument.spreadsheetml.sheet"
        ),
        use_container_width=True,
        key="download_heidenhain",
    )


# ==========================================================
# ETAPE 2
# ==========================================================

st.divider()

st.header(
    "2️⃣ Mise en forme du fichier Odoo"
)

if (
    st.session_state.heidenhain_result is None
    or odoo_file is None
    or category_file is None
):

    st.info(
        "Crée d'abord le fichier Heidenhain à l'étape 1, "
        "puis importe les fichiers Odoo et Catégorie."
    )

else:

    st.markdown(
        """
Le traitement va parcourir le fichier Heidenhain
et mettre à jour ou créer les références dans Odoo.

**Règles appliquées :**

- Référence existante → mise à jour.
- Référence absente → création d'une nouvelle ligne.
- Prix normal → **Prix (PPC)**.
- Référence `_SAV` → **Prix (SAV)**.
- Code-barres → `I ` + référence.
- Fournisseur → `HEIDENHAIN FRANCE`.
- Peut être acheté → `VRAI`.
- Peut être vendu → `VRAI`.
- Type de produit → `Consommable`.
- Politique de facturation → `Quantités livrées`.
- Catégorie → recherche dans le fichier Catégorie de produit.
        """
    )

    # ======================================================
    # CALCUL EN COURS
    # ======================================================

    if st.session_state.odoo_processing:

        # --------------------------------------------------
        # BOUTON VERROUILLE
        # --------------------------------------------------

        st.button(
            "⏳ Calcul Odoo en cours...",
            disabled=True,
            use_container_width=True,
            key="prepare_odoo_locked",
        )

        st.warning(
            "⏳ **La préparation Odoo est en cours.** "
            "Veuillez patienter et ne pas relancer le traitement."
        )

        progress_odoo = st.progress(
            0,
            text="Initialisation du traitement Odoo..."
        )

        status_odoo = st.empty()

        try:

            # ==================================================
            # FICHIER HEIDENHAIN PREPARE
            # ==================================================

            prepared_heidenhain_file = BytesIO(
                st.session_state.heidenhain_result
            )

            # ==================================================
            # TRAITEMENT
            # ==================================================

            result, stats = process_odoo(

                heidenhain_file=(
                    prepared_heidenhain_file
                ),

                odoo_file=(
                    odoo_file
                ),

                category_file=(
                    category_file
                ),

                heidenhain_sheet=(
                    heidenhain_sheet
                ),

                odoo_sheet=(
                    odoo_sheet
                ),

                category_sheet=(
                    category_sheet
                ),

                heidenhain_header_row=int(
                    heidenhain_header_row
                ),

                odoo_header_row=int(
                    odoo_header_row
                ),

                category_id_col=int(
                    category_id_col
                ),

                category_search_col=int(
                    category_search_col
                ),

                heidenhain_id_column=(
                    id_column
                ),

                heidenhain_status_column=(
                    status_column
                ),

                heidenhain_group_column=(
                    group_column
                ),

                heidenhain_ppc_column=(
                    ppc_column
                ),

                heidenhain_sav_column=(
                    sav_column
                ),

                heidenhain_description=(
                    "Description"
                ),

                heidenhain_marque=(
                    "Marque"
                ),

                heidenhain_prixHA=(
                    heidenhain_prixHA
                ),

                odoo_reference_column=(
                    odoo_reference
                ),

                odoo_sales_status_column=(
                    odoo_sales_status
                ),

                odoo_price_column=(
                    odoo_price
                ),

                odoo_barcode_column=(
                    odoo_barcode
                ),

                odoo_supplier_column=(
                    odoo_supplier
                ),

                odoo_purchase_column=(
                    odoo_purchase
                ),

                odoo_sale_column=(
                    odoo_sale
                ),

                odoo_type_column=(
                    odoo_type
                ),

                odoo_invoice_policy_column=(
                    odoo_invoice_policy
                ),

                odoo_category_column=(
                    odoo_category
                ),

                odoo_name=(
                    odoo_name
                ),

                odoo_brand_description=(
                    odoo_brand_description
                ),

                odoo_prixHA=(
                    odoo_prixHA
                ),

                progress=(
                    progress_odoo
                ),

                status_display=(
                    status_odoo
                ),
            )

            # ==================================================
            # SAUVEGARDE
            # ==================================================

            st.session_state.odoo_result = (
                result
            )

            st.session_state.odoo_stats = (
                stats
            )

            # ==================================================
            # DEVERROUILLAGE
            # ==================================================

            st.session_state.odoo_processing = False

            st.success(
                "✅ Fichier Odoo préparé avec succès."
            )

            st.rerun()

        except Exception as e:

            # --------------------------------------------------
            # TOUJOURS DEVERROUILLER EN CAS D'ERREUR
            # --------------------------------------------------

            st.session_state.odoo_processing = False

            st.error(
                "❌ Erreur pendant la préparation Odoo."
            )

            st.exception(e)

    # ======================================================
    # BOUTON NORMAL
    # ======================================================

    else:

        if st.button(
            "🚀 Préparer le fichier Odoo",
            type="primary",
            use_container_width=True,
            key="prepare_odoo",
        ):

            # --------------------------------------------------
            # VERROUILLAGE IMMEDIAT
            # --------------------------------------------------

            st.session_state.odoo_processing = True

            # --------------------------------------------------
            # RAFRAICHISSEMENT AVANT LE CALCUL
            # --------------------------------------------------

            st.rerun()


# ==========================================================
# RESULTAT ODOO
# ==========================================================

if st.session_state.odoo_result:

    stats = (
        st.session_state.odoo_stats
    )

    st.divider()

    st.header(
        "📊 Résultat Odoo"
    )

    c1, c2, c3, c4, c5 = (
        st.columns(5)
    )

    with c1:

        st.metric(
            "Références traitées",
            f"{stats['processed']:,}",
        )

    with c2:

        st.metric(
            "Lignes mises à jour",
            f"{stats['updated']:,}",
        )

    with c3:

        st.metric(
            "Nouvelles lignes",
            f"{stats['created']:,}",
        )

    with c4:

        st.metric(
            "Catégories trouvées",
            f"{stats['category_found']:,}",
        )

    with c5:

        st.metric(
            "Catégories introuvables",
            f"{stats['category_missing']:,}",
        )

    if stats["category_missing"] > 0:

        st.warning(
            f"⚠️ {stats['category_missing']:,} "
            "référence(s) n'ont pas trouvé de catégorie."
        )

    if stats["empty_references"] > 0:

        st.warning(
            f"⚠️ {stats['empty_references']:,} "
            "ligne(s) Heidenhain sans référence."
        )

    st.info(
        f"📂 {stats['category_rows']:,} lignes "
        "de catégories chargées une seule fois."
    )

    st.info(
        f"🔎 {stats['odoo_indexed']:,} références "
        "Odoo indexées."
    )

    st.download_button(
        label=(
            "⬇️ Télécharger le fichier Odoo "
            "prêt à importer"
        ),
        data=(
            st.session_state.odoo_result
        ),
        file_name=(
            "ETAPE 2 -- File for import Odoo.xlsx"
        ),
        mime=(
            "application/vnd.openxmlformats-"
            "officedocument.spreadsheetml.sheet"
        ),
        use_container_width=True,
        key="download_odoo",
    )

    st.success(
        "🎯 Le fichier Odoo final peut maintenant être téléchargé."
    )
