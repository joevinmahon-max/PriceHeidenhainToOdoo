# ==========================================================
# PREPARATION PRIX HEIDENHAIN + MISE EN FORME ODOO
# JMA
#
# VERSION :
# - Détection automatique des colonnes
# - Correspondance Heidenhain -> Odoo configurable
# - Conservation des colonnes Odoo existantes
# - Création automatique des colonnes Odoo manquantes
# - Positionnement configurable des nouvelles colonnes
# - Gestion automatique PPC / SAV
# - Gestion automatique des catégories
# - Gestion ID Odoo
# - Compatible gros fichiers
# ==========================================================

import streamlit as st
import openpyxl
import time

from io import BytesIO
from copy import copy


# ==========================================================
# CONFIGURATION
# ==========================================================

st.set_page_config(
    page_title="Préparation Heidenhain → Odoo",
    page_icon="📦",
    layout="wide",
)

st.title("📦 Préparation des prix HEIDENHAIN → Odoo")


# ==========================================================
# SESSION STATE
# ==========================================================

DEFAULT_STATE = {
    "heidenhain_result": None,
    "heidenhain_stats": None,
    "odoo_result": None,
    "odoo_stats": None,
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


# ==========================================================
# CORRESPONDANCES LOGIQUES
# ==========================================================
#
# key = nom logique interne
# label = nom affiché
# default_target = nom de colonne Odoo attendu
#
# source = colonne Heidenhain utilisée
#
# kind :
#   direct       = copie directe
#   price        = PPC / SAV
#   category     = recherche catégorie
#   barcode      = génération
#   constant     = valeur fixe
#   odoo_id      = ID Odoo existant
#
# ==========================================================

ODOO_FIELD_DEFINITIONS = {

    "id": {
        "label": "ID Odoo",
        "default_target": "id",
        "kind": "odoo_id",
    },

    "default_code": {
        "label": "Référence interne",
        "default_target": "default_code",
        "kind": "direct",
        "source": "ID",
    },

    "name": {
        "label": "Nom",
        "default_target": "name",
        "kind": "direct",
        "source": "Description",
    },

    "brand": {
        "label": "Marque",
        "default_target": "x_studio_marque_1",
        "kind": "direct",
        "source": "Marque",
    },

    "category": {
        "label": "Catégorie",
        "default_target": "categ_id",
        "kind": "category",
        "source": "Groupe Produit",
    },

    "sales_status": {
        "label": "Sales Status",
        "default_target": "x_studio_sales_status",
        "kind": "direct",
        "source": "Statut",
    },

    "list_price": {
        "label": "Prix de vente",
        "default_target": "list_price",
        "kind": "price",
    },

    "partner_id": {
        "label": "Fournisseur",
        "default_target": "seller_ids/partner_id",
        "kind": "constant",
        "default": "HEIDENHAIN FRANCE",
    },

    "seller_price": {
        "label": "Prix fournisseur",
        "default_target": "seller_ids/price",
        "kind": "direct",
        "source": "Prix HA",
    },

    "purchase_ok": {
        "label": "Peut être acheté",
        "default_target": "purchase_ok",
        "kind": "constant",
        "default": "VRAI",
    },

    "sale_ok": {
        "label": "Peut être vendu",
        "default_target": "sale_ok",
        "kind": "constant",
        "default": "VRAI",
    },

    "type": {
        "label": "Type de produit",
        "default_target": "type",
        "kind": "constant",
        "default": "Consommable",
    },

    "invoice_policy": {
        "label": "Politique de facturation",
        "default_target": "invoice_policy",
        "kind": "constant",
        "default": "Quantités livrées",
    },

    "barcode": {
        "label": "Code-barres",
        "default_target": "barcode",
        "kind": "barcode",
    },
}


# ==========================================================
# OUTILS
# ==========================================================

def normalize(value):

    if value is None:
        return ""

    return str(value).strip().upper()


def clean_reference(value):

    return normalize(value)


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


def normalize_header(value):

    if value is None:
        return ""

    return " ".join(
        str(value)
        .strip()
        .split()
    ).upper()


# ==========================================================
# DETECTION DES COLONNES
# ==========================================================

def detect_columns(
    ws,
    header_row,
):

    result = {}

    for cell in ws[header_row]:

        value = cell.value

        if value in (None, ""):
            continue

        header = str(value).strip()

        normalized = normalize_header(header)

        if normalized and normalized not in result:
            result[normalized] = {
                "name": header,
                "column": cell.column,
            }

    return result


def find_column(
    detected_columns,
    requested_name,
):

    normalized = normalize_header(
        requested_name
    )

    item = detected_columns.get(
        normalized
    )

    if item:
        return item["column"]

    return None


# ==========================================================
# LECTURE DES NOMS DE COLONNES D'UN FICHIER
# ==========================================================

def get_file_columns(
    uploaded_file,
    sheet_name,
    header_row,
):

    uploaded_file.seek(0)

    wb = openpyxl.load_workbook(
        uploaded_file,
        read_only=True,
        data_only=True,
    )

    if sheet_name not in wb.sheetnames:

        names = ", ".join(
            wb.sheetnames
        )

        wb.close()

        raise ValueError(
            f"Feuille '{sheet_name}' introuvable. "
            f"Feuilles disponibles : {names}"
        )

    ws = wb[sheet_name]

    columns = []

    for cell in ws[header_row]:

        if cell.value in (None, ""):
            continue

        columns.append(
            str(cell.value).strip()
        )

    wb.close()

    return columns


# ==========================================================
# FEUILLES
# ==========================================================

def get_sheet_names(
    uploaded_file,
):

    uploaded_file.seek(0)

    wb = openpyxl.load_workbook(
        uploaded_file,
        read_only=True,
        data_only=True,
    )

    names = wb.sheetnames

    wb.close()

    return names


# ==========================================================
# CATEGORIES
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
            f"'{category_sheet}' introuvable. "
            f"Feuilles disponibles : {names}"
        )

    ws = wb[category_sheet]

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

        if not category_id or not category_text:
            continue

        rows_loaded += 1

        normalized = normalize(
            category_text
        )

        mapping.setdefault(
            normalized,
            category_id,
        )

        for separator in ["-", "–"]:

            if separator in category_text:

                first_part = (
                    category_text
                    .split(separator, 1)[0]
                    .strip()
                )

                if first_part:

                    mapping.setdefault(
                        normalize(first_part),
                        category_id,
                    )

    wb.close()

    return mapping, rows_loaded


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

    normalized = normalize(group)

    if normalized in category_mapping:
        return category_mapping[
            normalized
        ]

    for separator in ["-", "–"]:

        if separator in group:

            first_part = (
                group
                .split(separator, 1)[0]
                .strip()
            )

            if first_part:

                result = category_mapping.get(
                    normalize(first_part)
                )

                if result is not None:
                    return result

    return None


# ==========================================================
# ETAPE 1
# PREPARATION HEIDENHAIN
# ==========================================================

def process_heidenhain(
    uploaded_file,
    sheet_name,
    header_row,
    output_columns,
    status_column,
    id_column,
    progress=None,
    status_display=None,
):

    header_row = int(header_row)

    data_start_row = (
        header_row + 1
    )

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
            f"Feuille '{sheet_name}' introuvable. "
            f"Feuilles disponibles : {names}"
        )

    source_ws = wb_formula[
        sheet_name
    ]

    uploaded_file.seek(0)

    wb_values = openpyxl.load_workbook(
        uploaded_file,
        data_only=True,
    )

    values_ws = wb_values[
        sheet_name
    ]

    detected = detect_columns(
        source_ws,
        header_row,
    )

    source_columns = {}

    for column_name in output_columns:

        col = find_column(
            detected,
            column_name,
        )

        if col is None:

            wb_formula.close()
            wb_values.close()

            raise ValueError(
                f"Colonne Heidenhain "
                f"'{column_name}' introuvable."
            )

        source_columns[
            column_name
        ] = col

    status_col = find_column(
        detected,
        status_column,
    )

    id_col = find_column(
        detected,
        id_column,
    )

    if status_col is None:
        raise ValueError(
            f"Colonne '{status_column}' introuvable."
        )

    if id_col is None:
        raise ValueError(
            f"Colonne '{id_column}' introuvable."
        )

    # ------------------------------------------------------
    # Dernière ligne
    # ------------------------------------------------------

    max_row = data_start_row - 1

    for row in range(
        data_start_row,
        source_ws.max_row + 1,
    ):

        value = source_ws.cell(
            row=row,
            column=id_col,
        ).value

        if value in (None, ""):
            break

        max_row = row

    total_rows = max(
        0,
        max_row - data_start_row + 1,
    )

    # ------------------------------------------------------
    # Nouveau classeur
    # ------------------------------------------------------

    output_wb = openpyxl.Workbook()

    output_ws = output_wb.active

    output_ws.title = sheet_name

    # ------------------------------------------------------
    # En-têtes
    # ------------------------------------------------------

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

    # ------------------------------------------------------
    # IDs existants
    # ------------------------------------------------------

    existing_ids = set()

    for row in range(
        data_start_row,
        max_row + 1,
    ):

        value = source_ws.cell(
            row=row,
            column=id_col,
        ).value

        normalized = normalize(value)

        if normalized:
            existing_ids.add(
                normalized
            )

    # ------------------------------------------------------
    # Traitement
    # ------------------------------------------------------

    rows_to_create = []

    vg_count = 0
    pg_count = 0
    duplicate_count = 0
    empty_id_count = 0

    output_row = 2

    for index, source_row in enumerate(
        range(
            data_start_row,
            max_row + 1,
        )
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

            # Prix calculés :
            # on prend la valeur calculée
            # du classeur data_only.
            if column_name in (
                "Prix (SAV)",
                "Prix HA",
                "Prix (PPC)",
                "Prix (PPC) 2027",
                "Prix (SAV) 2027",
            ):

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

        if progress is not None and (
            index % 10 == 0
            or index == total_rows - 1
        ):

            pct = int(
                (
                    (index + 1)
                    / max(1, total_rows)
                )
                * 80
            )

            progress.progress(
                min(80, 10 + pct),
                text=(
                    f"⚙️ Heidenhain : "
                    f"{index + 1:,} / "
                    f"{total_rows:,}"
                ),
            )

    # ------------------------------------------------------
    # Lignes SAV
    # ------------------------------------------------------

    for source_row, new_id in rows_to_create:

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

            if column_name in (
                "Prix (SAV)",
                "Prix HA",
                "Prix (PPC)",
                "Prix (PPC) 2027",
                "Prix (SAV) 2027",
            ):

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

        id_output_col = (
            output_columns.index(
                id_column
            ) + 1
        )

        output_ws.cell(
            row=output_row,
            column=id_output_col,
        ).value = new_id

        output_row += 1

    # ------------------------------------------------------
    # Largeurs
    # ------------------------------------------------------

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

    # ------------------------------------------------------
    # Sauvegarde
    # ------------------------------------------------------

    result = BytesIO()

    output_wb.save(result)

    result.seek(0)

    data = result.getvalue()

    output_wb.close()
    wb_formula.close()
    wb_values.close()

    stats = {
        "total_rows": total_rows,
        "vg": vg_count,
        "pg": pg_count,
        "created": len(rows_to_create),
        "duplicates": duplicate_count,
        "empty_ids": empty_id_count,
    }

    if progress is not None:
        progress.progress(
            100,
            text="✅ Étape 1 terminée",
        )

    return data, stats


# ==========================================================
# CREATION / INSERTION DE COLONNES ODOO
# ==========================================================

def insert_missing_columns(
    ws,
    header_row,
    required_columns,
    placement_config,
):

    detected = detect_columns(
        ws,
        header_row,
    )

    existing_names = [
        item["name"]
        for item in detected.values()
    ]

    missing = []

    for column_name in required_columns:

        if find_column(
            detected,
            column_name,
        ) is None:

            missing.append(
                column_name
            )

    # ------------------------------------------------------
    # Ajout colonne par colonne
    # ------------------------------------------------------

    for column_name in missing:

        position = placement_config.get(
            column_name
        )

        if position is None:
            position = "END"

        if position == "END":

            new_col = (
                ws.max_column + 1
            )

            ws.cell(
                row=header_row,
                column=new_col,
            ).value = column_name

        else:

            # position = nom de la colonne
            # avant laquelle insérer

            reference_col = find_column(
                detect_columns(
                    ws,
                    header_row,
                ),
                position,
            )

            if reference_col is None:

                new_col = (
                    ws.max_column + 1
                )

                ws.cell(
                    row=header_row,
                    column=new_col,
                ).value = column_name

            else:

                ws.insert_cols(
                    reference_col,
                    1,
                )

                ws.cell(
                    row=header_row,
                    column=reference_col,
                ).value = column_name

    return missing


# ==========================================================
# VALEUR D'UNE SOURCE HEIDENHAIN
# ==========================================================

def get_h_value(
    row_values,
    h_columns,
    source_name,
):

    if source_name is None:
        return None

    col = h_columns.get(
        source_name
    )

    if col is None:
        return None

    index = col - 1

    if index >= len(row_values):
        return None

    return row_values[index]


# ==========================================================
# TRAITEMENT ODOO
# ==========================================================

def process_odoo(
    heidenhain_file,
    odoo_file,
    category_file,

    heidenhain_sheet,
    odoo_sheet,
    category_sheet,

    category_id_col,
    category_search_col,

    odoo_mapping,
    odoo_existing_id_column,
    odoo_import_compatible,

    progress=None,
    status_display=None,
):

    start_total = time.perf_counter()

    # ======================================================
    # CATEGORIES
    # ======================================================

    if status_display is not None:
        status_display.info(
            "📂 Chargement des catégories..."
        )

    category_mapping, category_rows = (
        load_category_mapping(
            category_file,
            category_sheet,
            int(category_id_col),
            int(category_search_col),
        )
    )

    # ======================================================
    # HEIDENHAIN
    # ======================================================

    if status_display is not None:
        status_display.info(
            "📘 Chargement du fichier Heidenhain..."
        )

    heidenhain_file.seek(0)

    wb_h = openpyxl.load_workbook(
        heidenhain_file,
        data_only=True,
        read_only=True,
    )

    if heidenhain_sheet not in wb_h.sheetnames:

        wb_h.close()

        raise ValueError(
            f"Feuille Heidenhain "
            f"'{heidenhain_sheet}' introuvable."
        )

    h_ws = wb_h[
        heidenhain_sheet
    ]

    # Le fichier étape 1 possède ses en-têtes en ligne 1.
    h_detected = detect_columns(
        h_ws,
        1,
    )

    h_columns = {
        item["name"]: item["column"]
        for item in h_detected.values()
    }

    # ======================================================
    # VERIFICATION DES SOURCES
    # ======================================================

    source_errors = []

    for key, config in ODOO_FIELD_DEFINITIONS.items():

        if config["kind"] not in (
            "direct",
            "price",
            "category",
            "barcode",
        ):
            continue

        source = config.get(
            "source"
        )

        if source and source not in h_columns:

            source_errors.append(
                f"{config['label']} → "
                f"colonne Heidenhain "
                f"'{source}' absente"
            )

    if source_errors:

        wb_h.close()

        raise ValueError(
            "Colonnes nécessaires absentes "
            "du fichier Heidenhain :\n- "
            + "\n- ".join(source_errors)
        )

    # ======================================================
    # ODOO
    # ======================================================

    if status_display is not None:
        status_display.info(
            "📗 Chargement du fichier Odoo..."
        )

    odoo_file.seek(0)

    wb_o = openpyxl.load_workbook(
        odoo_file,
        data_only=False,
    )

    if odoo_sheet not in wb_o.sheetnames:

        wb_o.close()
        wb_h.close()

        raise ValueError(
            f"Feuille Odoo "
            f"'{odoo_sheet}' introuvable."
        )

    ws = wb_o[
        odoo_sheet
    ]

    # ======================================================
    # COLONNES ODOO NECESSAIRES
    # ======================================================

    required_columns = []

    for key, config in ODOO_FIELD_DEFINITIONS.items():

        target = odoo_mapping.get(
            key
        )

        if target:
            required_columns.append(
                target
            )

    # ======================================================
    # COLONNES MANQUANTES
    # ======================================================

    current_detected = detect_columns(
        ws,
        int(
            st.session_state.get(
                "odoo_header_row",
                1,
            )
        ),
    )

    # Le header row réel est récupéré plus bas
    # dans session_state.
    odoo_header_row = int(
        st.session_state.odoo_header_row
    )

    missing_columns = []

    current_detected = detect_columns(
        ws,
        odoo_header_row,
    )

    for column_name in required_columns:

        if find_column(
            current_detected,
            column_name,
        ) is None:

            missing_columns.append(
                column_name
            )

    # ======================================================
    # AJOUT DES COLONNES MANQUANTES
    # ======================================================

    placement_config = (
        st.session_state.get(
            "odoo_column_placement",
            {},
        )
    )

    for column_name in missing_columns:

        position = placement_config.get(
            column_name,
            "END",
        )

        if position == "END":

            new_col = (
                ws.max_column + 1
            )

            ws.cell(
                row=odoo_header_row,
                column=new_col,
            ).value = column_name

        else:

            detected_now = detect_columns(
                ws,
                odoo_header_row,
            )

            reference_col = find_column(
                detected_now,
                position,
            )

            if reference_col is None:

                new_col = (
                    ws.max_column + 1
                )

                ws.cell(
                    row=odoo_header_row,
                    column=new_col,
                ).value = column_name

            else:

                ws.insert_cols(
                    reference_col,
                    1,
                )

                ws.cell(
                    row=odoo_header_row,
                    column=reference_col,
                ).value = column_name

    # ======================================================
    # RE-DETECTION DES COLONNES
    # ======================================================

    detected_odoo = detect_columns(
        ws,
        odoo_header_row,
    )

    odoo_columns = {}

    for key, config in ODOO_FIELD_DEFINITIONS.items():

        target = odoo_mapping.get(
            key
        )

        if not target:
            continue

        col = find_column(
            detected_odoo,
            target,
        )

        if col is None:

            wb_o.close()
            wb_h.close()

            raise ValueError(
                f"Impossible de trouver "
                f"la colonne Odoo '{target}'."
            )

        odoo_columns[key] = col

    # ======================================================
    # INDEX ODOO
    # ======================================================

    if progress is not None:
        progress.progress(
            25,
            text="🔎 Indexation des références Odoo..."
        )

    reference_index = {}

    reference_col = odoo_columns[
        "default_code"
    ]

    for row in range(
        odoo_header_row + 1,
        ws.max_row + 1,
    ):

        value = ws.cell(
            row=row,
            column=reference_col,
        ).value

        ref = clean_reference(
            value
        )

        if ref and ref not in reference_index:

            reference_index[
                ref
            ] = row

    # ======================================================
    # COLONNES HEIDENHAIN
    # ======================================================

    h_id_col = h_columns["ID"]

    total_h_rows = max(
        0,
        h_ws.max_row - 1,
    )

    updated = 0
    created = 0
    category_found = 0
    category_missing = 0
    processed = 0
    empty_references = 0

    # ======================================================
    # PROCHAINE LIGNE ODOO
    # ======================================================

    next_empty_row = (
        odoo_header_row + 1
    )

    while ws.cell(
        row=next_empty_row,
        column=reference_col,
    ).value not in (None, ""):

        next_empty_row += 1

    # ======================================================
    # BOUCLE HEIDENHAIN
    # ======================================================

    for index, row_values in enumerate(
        h_ws.iter_rows(
            min_row=2,
            values_only=True,
        )
    ):

        reference = get_h_value(
            row_values,
            h_columns,
            "ID",
        )

        if reference in (None, ""):
            break

        reference = str(
            reference
        ).strip()

        if not reference:

            empty_references += 1

            continue

        processed += 1

        normalized_ref = clean_reference(
            reference
        )

        # --------------------------------------------------
        # EXISTANT / CREATION
        # --------------------------------------------------

        if normalized_ref in reference_index:

            target_row = reference_index[
                normalized_ref
            ]

            updated += 1

        else:

            target_row = next_empty_row

            next_empty_row += 1

            reference_index[
                normalized_ref
            ] = target_row

            created += 1

        # ==================================================
        # ID ODOO
        # ==================================================
        #
        # IMPORTANT :
        #
        # id Odoo n'est PAS remplacé par l'ID Heidenhain.
        #
        # Si le mode compatible import est activé :
        # on conserve l'ID Odoo existant.
        #
        # Pour une nouvelle ligne :
        # on laisse vide.
        #
        # ==================================================

        if "id" in odoo_columns:

            id_cell = ws.cell(
                row=target_row,
                column=odoo_columns["id"],
            )

            if odoo_import_compatible:

                if id_cell.value in (
                    None,
                    "",
                ):

                    existing_id = (
                        odoo_existing_id_column
                    )

                    if existing_id:
                        id_cell.value = (
                            get_h_value(
                                row_values,
                                h_columns,
                                existing_id,
                            )
                        )

            # Sinon :
            # on ne touche pas à l'ID Odoo.

        # ==================================================
        # default_code
        # ==================================================

        ws.cell(
            row=target_row,
            column=odoo_columns[
                "default_code"
            ],
        ).value = reference

        # ==================================================
        # NAME
        # ==================================================

        description = get_h_value(
            row_values,
            h_columns,
            "Description",
        )

        ws.cell(
            row=target_row,
            column=odoo_columns[
                "name"
            ],
        ).value = description

        # ==================================================
        # MARQUE
        # ==================================================

        marque = get_h_value(
            row_values,
            h_columns,
            "Marque",
        )

        ws.cell(
            row=target_row,
            column=odoo_columns[
                "brand"
            ],
        ).value = marque

        # ==================================================
        # STATUT
        # ==================================================

        status = get_h_value(
            row_values,
            h_columns,
            "Statut",
        )

        ws.cell(
            row=target_row,
            column=odoo_columns[
                "sales_status"
            ],
        ).value = status

        # ==================================================
        # PRIX
        # ==================================================

        if normalized_ref.endswith("_SAV"):

            price = get_h_value(
                row_values,
                h_columns,
                "Prix (SAV)",
            )

        else:

            price = get_h_value(
                row_values,
                h_columns,
                "Prix (PPC)",
            )

        ws.cell(
            row=target_row,
            column=odoo_columns[
                "list_price"
            ],
        ).value = price

        # ==================================================
        # PRIX HA
        # ==================================================

        prix_ha = get_h_value(
            row_values,
            h_columns,
            "Prix HA",
        )

        ws.cell(
            row=target_row,
            column=odoo_columns[
                "seller_price"
            ],
        ).value = prix_ha

        # ==================================================
        # CATEGORIE
        # ==================================================

        groupe = get_h_value(
            row_values,
            h_columns,
            "Groupe Produit",
        )

        category_id = (
            find_category_id_fast(
                groupe,
                category_mapping,
            )
        )

        if category_id is not None:

            category_found += 1

            ws.cell(
                row=target_row,
                column=odoo_columns[
                    "category"
                ],
            ).value = category_id

        else:

            category_missing += 1

        # ==================================================
        # FOURNISSEUR
        # ==================================================

        ws.cell(
            row=target_row,
            column=odoo_columns[
                "partner_id"
            ],
        ).value = (
            "HEIDENHAIN FRANCE"
        )

        # ==================================================
        # ACHAT
        # ==================================================

        ws.cell(
            row=target_row,
            column=odoo_columns[
                "purchase_ok"
            ],
        ).value = "VRAI"

        # ==================================================
        # VENTE
        # ==================================================

        ws.cell(
            row=target_row,
            column=odoo_columns[
                "sale_ok"
            ],
        ).value = "VRAI"

        # ==================================================
        # TYPE
        # ==================================================

        ws.cell(
            row=target_row,
            column=odoo_columns[
                "type"
            ],
        ).value = "Consommable"

        # ==================================================
        # FACTURATION
        # ==================================================

        ws.cell(
            row=target_row,
            column=odoo_columns[
                "invoice_policy"
            ],
        ).value = (
            "Quantités livrées"
        )

        # ==================================================
        # BARCODE
        # ==================================================

        if normalized_ref.endswith("_SAV"):

            barcode = ""

        else:

            barcode = (
                f"I {reference}"
            )

        ws.cell(
            row=target_row,
            column=odoo_columns[
                "barcode"
            ],
        ).value = barcode

        # ==================================================
        # PROGRESSION
        # ==================================================

        if progress is not None and (
            index % 10 == 0
            or index == total_h_rows - 1
        ):

            pct = 35 + int(
                (
                    (index + 1)
                    / max(
                        1,
                        total_h_rows,
                    )
                )
                * 60
            )

            progress.progress(
                min(95, pct),
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

    # ======================================================
    # SAUVEGARDE
    # ======================================================

    if status_display is not None:
        status_display.info(
            "💾 Création du fichier Odoo final..."
        )

    output = BytesIO()

    wb_o.save(output)

    output.seek(0)

    result = output.getvalue()

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

    stats = {
        "processed": processed,
        "updated": updated,
        "created": created,
        "category_found": category_found,
        "category_missing": category_missing,
        "empty_references": empty_references,
        "category_rows": category_rows,
        "odoo_indexed": len(
            reference_index
        ),
        "missing_columns": missing_columns,
        "total_time": total_time,
    }

    return result, stats


# ==========================================================
# SIDEBAR
# ==========================================================

st.sidebar.header(
    "⚙️ Paramètres"
)


# ==========================================================
# FICHIER HEIDENHAIN
# ==========================================================

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
    step=1,
)


# ==========================================================
# FICHIERS
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
        type=["xlsx", "xlsm"],
        key="heidenhain_file",
    )

with col2:

    st.subheader(
        "📗 Fichier Odoo"
    )

    odoo_file = st.file_uploader(
        "Fichier d'import Odoo",
        type=["xlsx", "xlsm"],
        key="odoo_file",
    )

with col3:

    st.subheader(
        "📂 Catégorie de produit"
    )

    category_file = st.file_uploader(
        "Fichier Catégorie de produit",
        type=["xlsx", "xlsm"],
        key="category_file",
    )


# ==========================================================
# DETECTION HEIDENHAIN
# ==========================================================

heidenhain_output_columns = []

if heidenhain_file is not None:

    try:

        detected_h_columns = get_file_columns(
            heidenhain_file,
            heidenhain_sheet,
            int(heidenhain_header_row),
        )

        st.sidebar.divider()

        st.sidebar.subheader(
            "📋 Colonnes Heidenhain"
        )

        st.sidebar.caption(
            f"🔎 {len(detected_h_columns)} "
            "colonnes détectées automatiquement."
        )

        default_h_columns = [
            x
            for x in [
                "ID",
                "Description",
                "Marque",
                "Statut",
                "Groupe Produit",
                "Prix (PPC)",
                "Prix (SAV)",
                "Prix HA",
            ]
            if x in detected_h_columns
        ]

        heidenhain_output_columns = (
            st.sidebar.multiselect(
                "Colonnes à conserver",
                options=detected_h_columns,
                default=default_h_columns,
                key="heidenhain_output_columns",
                help=(
                    "Les colonnes sont détectées "
                    "automatiquement dans le fichier "
                    "Heidenhain."
                ),
            )
        )

        if "ID" not in heidenhain_output_columns:

            st.sidebar.error(
                "⚠️ ID est obligatoire."
            )

        if "Prix (PPC)" not in heidenhain_output_columns:

            st.sidebar.warning(
                "⚠️ Prix (PPC) n'est pas sélectionné."
            )

        if "Prix (SAV)" not in heidenhain_output_columns:

            st.sidebar.warning(
                "⚠️ Prix (SAV) n'est pas sélectionné."
            )

    except Exception as e:

        st.sidebar.error(
            f"Erreur détection Heidenhain : {e}"
        )


# ==========================================================
# PARAMETRES CATEGORIES
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
# PARAMETRES ODOO
# ==========================================================

st.sidebar.divider()

st.sidebar.subheader(
    "📗 Paramètres Odoo"
)

odoo_sheet = st.sidebar.text_input(
    "Feuille Odoo",
    "Sheet1",
)

odoo_header_row = st.sidebar.number_input(
    "Ligne des en-têtes Odoo",
    min_value=1,
    value=1,
    step=1,
)

st.session_state.odoo_header_row = (
    int(odoo_header_row)
)


# ==========================================================
# MODE IMPORT ODOO
# ==========================================================

odoo_import_compatible = st.sidebar.checkbox(
    "Odoo export compatible avec l'import",
    value=False,
    help=(
        "Active la gestion de l'ID Odoo existant. "
        "Dans ce mode, la colonne id Odoo n'est jamais "
        "remplacée par l'ID Heidenhain."
    ),
)

odoo_existing_id_column = None

if odoo_import_compatible:

    st.sidebar.info(
        "ℹ️ Mode import Odoo activé : "
        "l'ID Odoo existant est conservé."
    )

    odoo_existing_id_column = st.sidebar.text_input(
        "ID Odoo existant dans le fichier source",
        "id",
        help=(
            "Nom de la colonne contenant l'ID Odoo "
            "existant. Cet ID est conservé pour les "
            "références déjà présentes."
        ),
    )


# ==========================================================
# DETECTION ODOO
# ==========================================================

odoo_detected_columns = []

if odoo_file is not None:

    try:

        odoo_detected_columns = get_file_columns(
            odoo_file,
            odoo_sheet,
            int(odoo_header_row),
        )

        st.sidebar.caption(
            f"🔎 {len(odoo_detected_columns)} "
            "colonnes Odoo détectées."
        )

    except Exception as e:

        st.sidebar.error(
            f"Erreur détection Odoo : {e}"
        )


# ==========================================================
# MAPPING ODOO
# ==========================================================

odoo_mapping = {}

if odoo_detected_columns:

    st.sidebar.subheader(
        "🔗 Correspondance Odoo"
    )

    st.sidebar.caption(
        "Les colonnes Odoo sont recherchées "
        "automatiquement. Tu peux modifier le nom "
        "cible si nécessaire."
    )

    for key, config in ODOO_FIELD_DEFINITIONS.items():

        default_target = config[
            "default_target"
        ]

        # --------------------------------------------------
        # Cherche automatiquement le nom exact
        # --------------------------------------------------

        matching_column = None

        for detected_name in odoo_detected_columns:

            if normalize_header(
                detected_name
            ) == normalize_header(
                default_target
            ):

                matching_column = detected_name
                break

        # --------------------------------------------------
        # Pour les champs directs :
        # possibilité de choisir une colonne
        # --------------------------------------------------

        if matching_column:

            target_value = st.sidebar.selectbox(
                config["label"],
                options=(
                    ["Créer si absente"]
                    + odoo_detected_columns
                ),
                index=(
                    odoo_detected_columns.index(
                        matching_column
                    ) + 1
                ),
                key=f"odoo_target_{key}",
            )

        else:

            target_value = st.sidebar.text_input(
                config["label"],
                value=default_target,
                key=f"odoo_target_text_{key}",
            )

        if target_value != "Créer si absente":

            odoo_mapping[key] = (
                target_value
            )


# ==========================================================
# COLONNES ODOO MANQUANTES + POSITION
# ==========================================================

missing_odoo_columns = []

if odoo_detected_columns:

    for key, config in ODOO_FIELD_DEFINITIONS.items():

        target = odoo_mapping.get(
            key
        )

        if not target:
            continue

        found = False

        for existing in odoo_detected_columns:

            if normalize_header(
                existing
            ) == normalize_header(
                target
            ):

                found = True
                break

        if not found:

            missing_odoo_columns.append(
                target
            )


odoo_column_placement = {}

if missing_odoo_columns:

    st.sidebar.divider()

    st.sidebar.subheader(
        "➕ Colonnes Odoo à créer"
    )

    st.sidebar.warning(
        f"{len(missing_odoo_columns)} "
        "colonne(s) demandée(s) sont absentes "
        "du fichier Odoo."
    )

    placement_options = [
        "END"
    ] + odoo_detected_columns

    for column_name in dict.fromkeys(
        missing_odoo_columns
    ):

        placement = st.sidebar.selectbox(
            f"Où placer « {column_name} » ?",
            options=placement_options,
            format_func=lambda x: (
                "À la fin des colonnes"
                if x == "END"
                else f"Avant « {x} »"
            ),
            key=f"placement_{column_name}",
        )

        odoo_column_placement[
            column_name
        ] = placement

st.session_state.odoo_column_placement = (
    odoo_column_placement
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
        "👆 Importez le fichier Heidenhain."
    )

else:

    st.markdown(
        """
Le fichier Heidenhain est détecté automatiquement.

Le traitement conserve uniquement les colonnes
sélectionnées dans les paramètres.

Les lignes **VG** et **PG** génèrent automatiquement
leur ligne `_SAV`, comme dans la version précédente.
"""
    )

    if st.session_state.heidenhain_processing:

        st.button(
            "⏳ Calcul Heidenhain en cours...",
            disabled=True,
            use_container_width=True,
        )

        progress = st.progress(
            0,
            text="Initialisation..."
        )

        status = st.empty()

        try:

            result, stats = (
                process_heidenhain(
                    uploaded_file=heidenhain_file,
                    sheet_name=heidenhain_sheet,
                    header_row=int(
                        heidenhain_header_row
                    ),
                    output_columns=(
                        heidenhain_output_columns
                    ),
                    status_column="Statut",
                    id_column="ID",
                    progress=progress,
                    status_display=status,
                )
            )

            st.session_state.heidenhain_result = (
                result
            )

            st.session_state.heidenhain_stats = (
                stats
            )

            st.session_state.heidenhain_processing = False

            st.rerun()

        except Exception as e:

            st.session_state.heidenhain_processing = False

            st.error(
                f"❌ Erreur : {e}"
            )

            st.exception(e)

    else:

        if st.button(
            "🚀 Créer le fichier Heidenhain",
            type="primary",
            use_container_width=True,
        ):

            if not heidenhain_output_columns:

                st.error(
                    "Sélectionnez au moins une colonne."
                )

            elif "ID" not in heidenhain_output_columns:

                st.error(
                    "La colonne ID est obligatoire."
                )

            else:

                st.session_state.heidenhain_processing = True

                st.rerun()


# ==========================================================
# RESULTAT ETAPE 1
# ==========================================================

if st.session_state.heidenhain_result:

    stats = (
        st.session_state.heidenhain_stats
    )

    st.success(
        f"✅ Fichier Heidenhain créé — "
        f"{stats['created']:,} lignes SAV créées."
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
        "Créez d'abord le fichier Heidenhain "
        "puis importez le fichier Odoo et le fichier "
        "Catégorie."
    )

else:

    if not odoo_mapping:

        st.warning(
            "⚠️ Aucune correspondance Odoo configurée."
        )

    else:

        st.markdown(
            """
### Correspondances appliquées

- `default_code` ← `ID`
- `name` ← `Description`
- `x_studio_marque_1` ← `Marque`
- `categ_id` ← `Groupe Produit` via le fichier catégorie
- `x_studio_sales_status` ← `Statut`
- `list_price` ← `Prix (PPC)` ou `Prix (SAV)` pour les `_SAV`
- `seller_ids/partner_id` ← `HEIDENHAIN FRANCE`
- `seller_ids/price` ← `Prix HA`
- `purchase_ok` ← `VRAI`
- `sale_ok` ← `VRAI`
- `type` ← `Consommable`
- `invoice_policy` ← `Quantités livrées`
- `barcode` ← `I ` + `default_code`
- `id` ← ID Odoo existant uniquement lorsque le mode compatible import est activé
"""
        )

        if missing_odoo_columns:

            st.warning(
                "⚠️ Colonnes Odoo manquantes détectées : "
                + ", ".join(
                    dict.fromkeys(
                        missing_odoo_columns
                    )
                )
            )

        if st.session_state.odoo_processing:

            st.button(
                "⏳ Préparation Odoo en cours...",
                disabled=True,
                use_container_width=True,
            )

            progress = st.progress(
                0,
                text="Initialisation..."
            )

            status = st.empty()

            try:

                prepared_heidenhain = BytesIO(
                    st.session_state.heidenhain_result
                )

                result, stats = process_odoo(

                    heidenhain_file=(
                        prepared_heidenhain
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

                    category_id_col=int(
                        category_id_col
                    ),

                    category_search_col=int(
                        category_search_col
                    ),

                    odoo_mapping=(
                        odoo_mapping
                    ),

                    odoo_existing_id_column=(
                        odoo_existing_id_column
                    ),

                    odoo_import_compatible=(
                        odoo_import_compatible
                    ),

                    progress=progress,

                    status_display=status,
                )

                st.session_state.odoo_result = (
                    result
                )

                st.session_state.odoo_stats = (
                    stats
                )

                st.session_state.odoo_processing = False

                st.rerun()

            except Exception as e:

                st.session_state.odoo_processing = False

                st.error(
                    "❌ Erreur pendant la préparation Odoo."
                )

                st.exception(e)

        else:

            if st.button(
                "🚀 Préparer le fichier Odoo",
                type="primary",
                use_container_width=True,
            ):

                st.session_state.odoo_processing = True

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

    c1, c2, c3, c4, c5 = st.columns(5)

    with c1:

        st.metric(
            "Références traitées",
            f"{stats['processed']:,}",
        )

    with c2:

        st.metric(
            "Mises à jour",
            f"{stats['updated']:,}",
        )

    with c3:

        st.metric(
            "Créations",
            f"{stats['created']:,}",
        )

    with c4:

        st.metric(
            "Catégories trouvées",
            f"{stats['category_found']:,}",
        )

    with c5:

        st.metric(
            "Catégories absentes",
            f"{stats['category_missing']:,}",
        )

    if stats["category_missing"]:

        st.warning(
            f"⚠️ {stats['category_missing']:,} "
            "catégorie(s) non trouvée(s)."
        )

    if stats["missing_columns"]:

        st.info(
            "➕ Colonnes créées automatiquement : "
            + ", ".join(
                dict.fromkeys(
                    stats["missing_columns"]
                )
            )
        )

    st.info(
        f"📂 {stats['category_rows']:,} "
        "lignes de catégories chargées."
    )

    st.info(
        f"🔎 {stats['odoo_indexed']:,} "
        "références Odoo indexées."
    )

    st.info(
        f"⏱️ Temps total : "
        f"{stats['total_time']:.2f} secondes"
    )

    st.download_button(
        "⬇️ Télécharger le fichier Odoo prêt à importer",
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
    )
