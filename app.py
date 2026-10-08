# ==========================================================
# PREPARATION PRIX HEIDENHAIN + MISE EN FORME ODOO
# JMA
#
# VERSION :
# - Détection automatique des colonnes Heidenhain
# - Sélection des colonnes Heidenhain à conserver
# - Mapping automatique Heidenhain -> Odoo
# - Détection automatique des colonnes Odoo
# - Création automatique des colonnes Odoo manquantes
# - Conservation de l'ID Odoo existant
# - Nouveaux produits : ID laissé vide
# - Gestion PPC / SAV conservée
# - Gestion Prix HA conservée
# - Gestion catégories conservée
# - Gestion VG / PG / _SAV conservée
# - Index Odoo en mémoire
# - Compatible gros fichiers
# - Boutons verrouillés pendant calcul
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
# UTILITAIRES
# ==========================================================

def normalize(value):

    if value is None:
        return ""

    return str(value).strip().upper()


def clean_reference(value):

    return normalize(value)


def is_sav(reference):

    return clean_reference(reference).endswith("_SAV")


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


# ==========================================================
# RECHERCHE COLONNES
# ==========================================================

def find_columns(
    ws,
    header_row,
    names,
):

    result = {}
    headers = {}

    for cell in ws[header_row]:

        value = normalize(cell.value)

        if value and value not in headers:

            headers[value] = cell.column

    for name in names:

        normalized_name = normalize(name)

        if normalized_name in headers:

            result[name] = headers[
                normalized_name
            ]

    return result


def get_headers(
    ws,
    header_row,
):

    headers = []

    for cell in ws[header_row]:

        value = cell.value

        if value not in (
            None,
            "",
        ):

            headers.append(
                str(value).strip()
            )

    return headers


def find_header_column(
    ws,
    header_row,
    header_name,
):

    wanted = normalize(header_name)

    for cell in ws[header_row]:

        if normalize(cell.value) == wanted:

            return cell.column

    return None


# ==========================================================
# FEUILLES
# ==========================================================

def get_sheet_names(uploaded_file):

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
# DETECTION COLONNES HEIDENHAIN
# ==========================================================

def detect_heidenhain_columns(
    uploaded_file,
    sheet_name,
    header_row,
):

    uploaded_file.seek(0)

    wb = openpyxl.load_workbook(
        uploaded_file,
        read_only=True,
        data_only=False,
    )

    if sheet_name not in wb.sheetnames:

        names = ", ".join(wb.sheetnames)

        wb.close()

        raise ValueError(
            f"Feuille '{sheet_name}' introuvable. "
            f"Feuilles disponibles : {names}"
        )

    ws = wb[sheet_name]

    headers = get_headers(
        ws,
        int(header_row),
    )

    wb.close()

    return headers


# ==========================================================
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

        if column_name in (
            "Prix (SAV)",
            "Prix HA",
        ):

            target_cell.value = (
                values_ws.cell(
                    row=source_row,
                    column=source_col,
                ).value
            )

        else:

            target_cell.value = source_cell.value

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

    header_row = int(header_row)
    data_start_row = int(data_start_row)

    if data_start_row <= header_row:

        raise ValueError(
            "La ligne de données doit être "
            "supérieure à la ligne des en-têtes."
        )

    if "ID" not in output_columns:

        raise ValueError(
            "La colonne ID est obligatoire."
        )

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
    # FORMULES
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
            f"Feuille '{sheet_name}' introuvable. "
            f"Feuilles disponibles : {names}"
        )

    source_ws = wb_formula[sheet_name]

    # ======================================================
    # VALEURS CALCULEES
    # ======================================================

    uploaded_file.seek(0)

    wb_values = openpyxl.load_workbook(
        uploaded_file,
        data_only=True,
    )

    values_ws = wb_values[sheet_name]

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
            "Colonnes introuvables dans Heidenhain : "
            + ", ".join(missing)
        )

    status_col = source_columns[
        status_column
    ]

    id_col = source_columns[
        id_column
    ]

    # ======================================================
    # DERNIERE LIGNE
    # ======================================================

    max_row = data_start_row - 1

    for row in range(
        data_start_row,
        source_ws.max_row + 1,
    ):

        id_value = source_ws.cell(
            row=row,
            column=id_col,
        ).value

        if id_value in (
            None,
            "",
        ):

            break

        max_row = row

    total_rows = max(
        0,
        max_row - data_start_row + 1,
    )

    # ======================================================
    # NOUVEAU CLASSEUR
    # ======================================================

    output_wb = openpyxl.Workbook()

    output_ws = output_wb.active

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

        normalized = normalize(value)

        if normalized:

            existing_ids.add(normalized)

    # ======================================================
    # TRAITEMENT
    # ======================================================

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

        copy_heidenhain_row(
            source_ws=source_ws,
            values_ws=values_ws,
            output_ws=output_ws,
            source_row=source_row,
            output_row=output_row,
            source_columns=source_columns,
            output_columns=output_columns,
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

                new_id = f"{original_id}_SAV"

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
                    / max(1, total_rows)
                )
                * 75
            )

            pct = 10 + pct

            if progress is not None:

                progress.progress(
                    min(pct, 85),
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

        created_ids.append(new_id)

        output_row += 1

    # ======================================================
    # LARGEURS
    # ======================================================

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
        "vg_pg": vg_count + pg_count,
        "created": len(created_ids),
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

        names = ", ".join(wb.sheetnames)

        wb.close()

        raise ValueError(
            f"Feuille catégorie '{category_sheet}' "
            f"introuvable. Feuilles disponibles : {names}"
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

        normalized_text = normalize(
            category_text
        )

        if normalized_text not in mapping:

            mapping[
                normalized_text
            ] = category_id

        for separator in (
            "-",
            "–",
        ):

            if separator in category_text:

                first_part = (
                    category_text
                    .split(separator, 1)[0]
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
# RECHERCHE CATEGORIE
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

    normalized_group = normalize(group)

    category_id = category_mapping.get(
        normalized_group
    )

    if category_id is not None:
        return category_id

    for separator in (
        "-",
        "–",
    ):

        if separator in group:

            first_part = (
                group
                .split(separator, 1)[0]
                .strip()
            )

            if first_part:

                category_id = (
                    category_mapping.get(
                        normalize(first_part)
                    )
                )

                if category_id is not None:
                    return category_id

    return None


# ==========================================================
# AJOUT COLONNE ODOO
# ==========================================================

def ensure_odoo_column(
    ws,
    header_row,
    header_name,
):

    existing = find_header_column(
        ws,
        header_row,
        header_name,
    )

    if existing is not None:

        return existing, False

    new_col = ws.max_column + 1

    # Si max_column est vide mais Excel considère
    # quand même une colonne, on cherche réellement
    # la dernière colonne utilisée.

    last_used = 0

    for col in range(
        1,
        ws.max_column + 1,
    ):

        value = ws.cell(
            row=header_row,
            column=col,
        ).value

        if value not in (
            None,
            "",
        ):

            last_used = col

    new_col = last_used + 1

    # Copie du style de la dernière colonne
    # existante de l'en-tête.

    if last_used > 0:

        source_header = ws.cell(
            row=header_row,
            column=last_used,
        )

        target_header = ws.cell(
            row=header_row,
            column=new_col,
        )

        copy_style_safe(
            source_header,
            target_header,
        )

        try:

            source_letter = (
                openpyxl.utils.get_column_letter(
                    last_used
                )
            )

            target_letter = (
                openpyxl.utils.get_column_letter(
                    new_col
                )
            )

            width = (
                ws.column_dimensions[
                    source_letter
                ].width
            )

            if width:
                ws.column_dimensions[
                    target_letter
                ].width = width

        except Exception:
            pass

    ws.cell(
        row=header_row,
        column=new_col,
    ).value = header_name

    return new_col, True


# ==========================================================
# MAPPING ODOO
# ==========================================================

def build_odoo_mapping(
    ws,
    header_row,
    mapping_definitions,
):

    mapping = {}
    created_columns = []

    for target_column, definition in mapping_definitions.items():

        if definition is None:
            continue

        target_col, created = ensure_odoo_column(
            ws=ws,
            header_row=header_row,
            header_name=target_column,
        )

        mapping[
            target_column
        ] = target_col

        if created:

            created_columns.append(
                target_column
            )

    return mapping, created_columns


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

    odoo_header_row,

    category_id_col,
    category_search_col,

    mapping,

    default_supplier,
    default_purchase_ok,
    default_sale_ok,
    default_type,
    default_invoice_policy,

    progress=None,
    status_display=None,
):

    start_total = time.perf_counter()

    timer_categories = st.empty()
    timer_heidenhain = st.empty()
    timer_odoo = st.empty()
    timer_index = st.empty()
    timer_main = st.empty()

    # ======================================================
    # CATEGORIES
    # ======================================================

    start_categories = time.perf_counter()

    if status_display is not None:

        status_display.info(
            "📂 Chargement des catégories..."
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
    # HEIDENHAIN
    # ======================================================

    start_heidenhain = time.perf_counter()

    heidenhain_file.seek(0)

    wb_h = openpyxl.load_workbook(
        heidenhain_file,
        data_only=True,
        read_only=True,
    )

    if heidenhain_sheet not in wb_h.sheetnames:

        names = ", ".join(wb_h.sheetnames)

        wb_h.close()

        raise ValueError(
            f"Feuille Heidenhain "
            f"'{heidenhain_sheet}' introuvable. "
            f"Feuilles disponibles : {names}"
        )

    h_ws = wb_h[heidenhain_sheet]

    time_heidenhain_load = (
        time.perf_counter()
        - start_heidenhain
    )

    if progress is not None:

        progress.progress(
            20,
            text="📘 Fichier Heidenhain chargé..."
        )

    # ======================================================
    # ODOO
    # ======================================================

    start_odoo_load = time.perf_counter()

    odoo_file.seek(0)

    wb_o = openpyxl.load_workbook(
        odoo_file,
        data_only=False,
    )

    if odoo_sheet not in wb_o.sheetnames:

        names = ", ".join(wb_o.sheetnames)

        wb_o.close()
        wb_h.close()

        raise ValueError(
            f"Feuille Odoo "
            f"'{odoo_sheet}' introuvable. "
            f"Feuilles disponibles : {names}"
        )

    ws = wb_o[odoo_sheet]

    time_odoo_load = (
        time.perf_counter()
        - start_odoo_load
    )

    if progress is not None:

        progress.progress(
            30,
            text="📗 Fichier Odoo chargé..."
        )

    # ======================================================
    # COLONNES ODOO
    #
    # On vérifie les colonnes demandées.
    # Si absentes : création à droite.
    # ======================================================

    odoo_mapping, created_columns = (
        build_odoo_mapping(
            ws=ws,
            header_row=int(
                odoo_header_row
            ),
            mapping_definitions=mapping,
        )
    )

    if status_display is not None:

        if created_columns:

            status_display.info(
                "➕ Colonnes Odoo ajoutées : "
                + ", ".join(created_columns)
            )

        else:

            status_display.info(
                "📗 Toutes les colonnes Odoo "
                "nécessaires existent déjà."
            )

    # ======================================================
    # COLONNES HEIDENHAIN
    # ======================================================

    h_headers = get_headers(
        h_ws,
        1,
    )

    h_columns = find_columns(
        h_ws,
        1,
        h_headers,
    )

    required_h = [
        "ID",
        "Description",
        "Marque",
        "Statut",
        "Groupe Produit",
        "Prix (PPC)",
        "Prix (SAV)",
        "Prix HA",
    ]

    missing_h = [
        col
        for col in required_h
        if col not in h_columns
    ]

    if missing_h:

        wb_o.close()
        wb_h.close()

        raise ValueError(
            "Colonnes nécessaires absentes "
            "du fichier Heidenhain : "
            + ", ".join(missing_h)
        )

    # ======================================================
    # INDEX ODOO
    # ======================================================

    start_index = time.perf_counter()

    reference_col = odoo_mapping[
        "default_code"
    ]

    id_col = odoo_mapping.get(
        "id"
    )

    reference_index = {}

    for row in range(
        int(odoo_header_row) + 1,
        ws.max_row + 1,
    ):

        value = ws.cell(
            row=row,
            column=reference_col,
        ).value

        ref = clean_reference(value)

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
                f"🔎 Index Odoo créé : "
                f"{len(reference_index):,} références"
            ),
        )

    # ======================================================
    # COLONNES HEIDENHAIN
    # ======================================================

    h_id_col = h_columns["ID"]
    h_description_col = h_columns["Description"]
    h_marque_col = h_columns["Marque"]
    h_status_col = h_columns["Statut"]
    h_group_col = h_columns["Groupe Produit"]
    h_ppc_col = h_columns["Prix (PPC)"]
    h_sav_col = h_columns["Prix (SAV)"]
    h_prixHA_col = h_columns["Prix HA"]

    # ======================================================
    # COLONNES ODOO CIBLES
    # ======================================================

    col_default_code = odoo_mapping[
        "default_code"
    ]

    col_name = odoo_mapping[
        "name"
    ]

    col_marque = odoo_mapping[
        "x_studio_marque_1"
    ]

    col_category = odoo_mapping[
        "categ_id"
    ]

    col_status = odoo_mapping[
        "x_studio_sales_status"
    ]

    col_price = odoo_mapping[
        "list_price"
    ]

    col_supplier = odoo_mapping[
        "seller_ids/partner_id"
    ]

    col_supplier_price = odoo_mapping[
        "seller_ids/price"
    ]

    col_purchase = odoo_mapping[
        "purchase_ok"
    ]

    col_sale = odoo_mapping[
        "sale_ok"
    ]

    col_type = odoo_mapping[
        "type"
    ]

    col_invoice = odoo_mapping[
        "invoice_policy"
    ]

    col_barcode = odoo_mapping[
        "barcode"
    ]

    # ======================================================
    # STATISTIQUES
    # ======================================================

    updated = 0
    created = 0

    category_found = 0
    category_missing = 0

    references_processed = 0
    empty_references = 0

    existing_ids_preserved = 0

    # ======================================================
    # DERNIERE LIGNE ODOO
    # ======================================================

    next_empty_row = (
        int(odoo_header_row) + 1
    )

    while ws.cell(
        row=next_empty_row,
        column=reference_col,
    ).value not in (
        None,
        "",
    ):

        next_empty_row += 1

    # ======================================================
    # BOUCLE HEIDENHAIN
    # ======================================================

    total_h_rows = max(
        0,
        h_ws.max_row - 1,
    )

    start_main_loop = time.perf_counter()

    if progress is not None:

        progress.progress(
            45,
            text=(
                f"⚙️ Traitement : "
                f"0 / {total_h_rows:,}"
            ),
        )

    for index, row_values in enumerate(
        h_ws.iter_rows(
            min_row=2,
            values_only=True,
        )
    ):

        # --------------------------------------------------
        # ID
        # --------------------------------------------------

        try:

            reference = row_values[
                h_id_col - 1
            ]

        except IndexError:

            continue

        if reference in (
            None,
            "",
        ):

            break

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

            target_row = reference_index[
                normalized_ref
            ]

            updated += 1

            # IMPORTANT :
            # on ne touche jamais à l'id
            # d'une ligne Odoo existante.

            if id_col is not None:

                existing_ids_preserved += 1

        else:

            target_row = next_empty_row

            next_empty_row += 1

            reference_index[
                normalized_ref
            ] = target_row

            created += 1

        # ==================================================
        # VALEURS HEIDENHAIN
        # ==================================================

        description = row_values[
            h_description_col - 1
        ]

        marque = row_values[
            h_marque_col - 1
        ]

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

        # ==================================================
        # PRIX PPC / SAV
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

        # --------------------------------------------------
        # default_code
        # --------------------------------------------------

        ws.cell(
            row=target_row,
            column=col_default_code,
        ).value = reference

        # --------------------------------------------------
        # name
        # --------------------------------------------------

        ws.cell(
            row=target_row,
            column=col_name,
        ).value = description

        # --------------------------------------------------
        # marque
        # --------------------------------------------------

        ws.cell(
            row=target_row,
            column=col_marque,
        ).value = marque

        # --------------------------------------------------
        # catégorie
        # --------------------------------------------------

        if category_id is not None:

            ws.cell(
                row=target_row,
                column=col_category,
            ).value = category_id

        # --------------------------------------------------
        # statut
        # --------------------------------------------------

        ws.cell(
            row=target_row,
            column=col_status,
        ).value = status

        # --------------------------------------------------
        # prix de vente
        # --------------------------------------------------

        ws.cell(
            row=target_row,
            column=col_price,
        ).value = price

        # --------------------------------------------------
        # fournisseur
        # --------------------------------------------------

        ws.cell(
            row=target_row,
            column=col_supplier,
        ).value = default_supplier

        # --------------------------------------------------
        # prix achat
        # --------------------------------------------------

        ws.cell(
            row=target_row,
            column=col_supplier_price,
        ).value = prixHA

        # --------------------------------------------------
        # purchase_ok
        # --------------------------------------------------

        ws.cell(
            row=target_row,
            column=col_purchase,
        ).value = default_purchase_ok

        # --------------------------------------------------
        # sale_ok
        # --------------------------------------------------

        ws.cell(
            row=target_row,
            column=col_sale,
        ).value = default_sale_ok

        # --------------------------------------------------
        # type
        # --------------------------------------------------

        ws.cell(
            row=target_row,
            column=col_type,
        ).value = default_type

        # --------------------------------------------------
        # invoice_policy
        # --------------------------------------------------

        ws.cell(
            row=target_row,
            column=col_invoice,
        ).value = default_invoice_policy

        # --------------------------------------------------
        # BARCODE
        # --------------------------------------------------

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

            if progress is not None:

                progress.progress(
                    min(pct, 90),
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
        # TEMPS
        # ==================================================

        if index % 100 == 0:

            now = time.perf_counter()

            timer_categories.info(
                f"⏱️ Catégories : "
                f"{time_categories:.2f} s"
            )

            timer_heidenhain.info(
                f"⏱️ Heidenhain : "
                f"{time_heidenhain_load:.2f} s"
            )

            timer_odoo.info(
                f"⏱️ Odoo : "
                f"{time_odoo_load:.2f} s"
            )

            timer_index.info(
                f"⏱️ Index Odoo : "
                f"{time_index:.2f} s "
                f"({len(reference_index):,})"
            )

            timer_main.info(
                f"⏱️ Boucle : "
                f"{now - start_main_loop:.2f} s"
            )

    # ======================================================
    # SAUVEGARDE
    # ======================================================

    if progress is not None:

        progress.progress(
            92,
            text="💾 Création du fichier Odoo..."
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

    timer_main.success(
        f"⏱️ Temps total : {total_time:.2f} s"
    )

    stats = {
        "processed": references_processed,
        "updated": updated,
        "created": created,
        "category_found": category_found,
        "category_missing": category_missing,
        "empty_references": empty_references,
        "category_rows": category_rows,
        "odoo_indexed": len(reference_index),
        "created_columns": created_columns,
        "existing_ids_preserved": existing_ids_preserved,
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
)


# ==========================================================
# DETECTION AUTOMATIQUE DES COLONNES
# ==========================================================

detected_heidenhain_columns = []

if heidenhain_file := st.session_state.get(
    "_dummy_heidenhain_file",
    None,
):
    pass


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
# COLONNES HEIDENHAIN DETECTEES
# ==========================================================

if heidenhain_file is not None:

    try:

        detected_heidenhain_columns = (
            detect_heidenhain_columns(
                uploaded_file=heidenhain_file,
                sheet_name=heidenhain_sheet,
                header_row=heidenhain_header_row,
            )
        )

    except Exception as e:

        st.sidebar.error(
            f"Erreur détection Heidenhain : {e}"
        )


# ==========================================================
# PARAMETRES COLONNES HEIDENHAIN
# ==========================================================

st.sidebar.divider()

st.sidebar.subheader(
    "📋 Colonnes Heidenhain"
)

if detected_heidenhain_columns:

    default_columns = [
        col
        for col in (
            "ID",
            "Description",
            "Marque",
            "Statut",
            "Groupe Produit",
            "Prix (PPC)",
            "Prix (SAV)",
            "Prix HA",
        )
        if col in detected_heidenhain_columns
    ]

    heidenhain_output_columns = (
        st.sidebar.multiselect(
            "Colonnes à conserver",
            options=detected_heidenhain_columns,
            default=default_columns,
            help=(
                "Les colonnes sont détectées automatiquement "
                "depuis le fichier Heidenhain."
            ),
        )
    )

else:

    st.sidebar.info(
        "Importez le fichier Heidenhain "
        "pour détecter automatiquement ses colonnes."
    )

    heidenhain_output_columns = []


if "ID" not in heidenhain_output_columns:

    st.sidebar.error(
        "⚠️ La colonne ID est obligatoire."
    )


# ==========================================================
# PARAMETRES ODOO
# ==========================================================

st.sidebar.divider()

st.sidebar.subheader(
    "📗 Correspondance Odoo"
)

st.sidebar.caption(
    "Les colonnes Odoo sont détectées automatiquement. "
    "Si une colonne cible n'existe pas, elle sera créée "
    "à droite des colonnes existantes."
)

odoo_header_row = st.sidebar.number_input(
    "Ligne des en-têtes Odoo",
    min_value=1,
    value=1,
)


# ==========================================================
# MAPPING AUTOMATIQUE
# ==========================================================

st.sidebar.markdown(
    """
**Correspondance automatique :**

- `default_code` ← ID
- `name` ← Description
- `x_studio_marque_1` ← Marque
- `categ_id` ← Groupe Produit
- `x_studio_sales_status` ← Statut
- `list_price` ← PPC / SAV
- `seller_ids/partner_id` ← Fournisseur
- `seller_ids/price` ← Prix HA
- `purchase_ok` ← Paramètre
- `sale_ok` ← Paramètre
- `type` ← Paramètre
- `invoice_policy` ← Paramètre
- `barcode` ← ID
"""
)


# ==========================================================
# NOMS DES COLONNES ODOO
# ==========================================================

st.sidebar.subheader(
    "🔗 Noms des colonnes Odoo"
)

odoo_col_id = st.sidebar.text_input(
    "ID Odoo",
    "id",
)

odoo_col_default_code = st.sidebar.text_input(
    "Référence",
    "default_code",
)

odoo_col_name = st.sidebar.text_input(
    "Nom",
    "name",
)

odoo_col_marque = st.sidebar.text_input(
    "Marque",
    "x_studio_marque_1",
)

odoo_col_category = st.sidebar.text_input(
    "Catégorie",
    "categ_id",
)

odoo_col_status = st.sidebar.text_input(
    "Sales Status",
    "x_studio_sales_status",
)

odoo_col_price = st.sidebar.text_input(
    "Prix de vente",
    "list_price",
)

odoo_col_supplier = st.sidebar.text_input(
    "Fournisseur",
    "seller_ids/partner_id",
)

odoo_col_supplier_price = st.sidebar.text_input(
    "Prix fournisseur",
    "seller_ids/price",
)

odoo_col_purchase = st.sidebar.text_input(
    "Peut être acheté",
    "purchase_ok",
)

odoo_col_sale = st.sidebar.text_input(
    "Peut être vendu",
    "sale_ok",
)

odoo_col_type = st.sidebar.text_input(
    "Type",
    "type",
)

odoo_col_invoice = st.sidebar.text_input(
    "Politique de facturation",
    "invoice_policy",
)

odoo_col_barcode = st.sidebar.text_input(
    "Code-barres",
    "barcode",
)


# ==========================================================
# PARAMETRES VALEURS FIXES
# ==========================================================

st.sidebar.subheader(
    "⚙️ Valeurs par défaut"
)

default_supplier = st.sidebar.text_input(
    "Fournisseur par défaut",
    "HEIDENHAIN FRANCE",
)

default_purchase_ok = st.sidebar.text_input(
    "Peut être acheté",
    "VRAI",
)

default_sale_ok = st.sidebar.text_input(
    "Peut être vendu",
    "VRAI",
)

default_type = st.sidebar.text_input(
    "Type de produit",
    "Consommable",
)

default_invoice_policy = st.sidebar.text_input(
    "Politique de facturation",
    "Quantités livrées",
)


# ==========================================================
# CATEGORIES
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
Le programme détecte automatiquement les colonnes du fichier Heidenhain.

Vous choisissez uniquement les colonnes à conserver.

La logique des lignes SAV reste inchangée :
- VG / PG → création d'une référence `_SAV`
- doublon `_SAV` → ignoré
- Prix SAV → valeur calculée du fichier Excel
- Prix HA → valeur calculée du fichier Excel
        """
    )

    if st.session_state.heidenhain_processing:

        st.button(
            "⏳ Calcul Heidenhain en cours...",
            disabled=True,
            use_container_width=True,
            key="create_heidenhain_locked",
        )

        st.warning(
            "⏳ Le calcul Heidenhain est en cours."
        )

        progress_heidenhain = st.progress(
            0,
            text="Initialisation..."
        )

        status_heidenhain = st.empty()

        try:

            result, stats = (
                process_heidenhain(
                    uploaded_file=heidenhain_file,
                    sheet_name=heidenhain_sheet,
                    header_row=heidenhain_header_row,
                    data_start_row=(
                        heidenhain_header_row + 1
                    ),
                    status_column="Statut",
                    id_column="ID",
                    output_columns=(
                        heidenhain_output_columns
                    ),
                    progress=progress_heidenhain,
                    status_display=status_heidenhain,
                )
            )

            st.session_state.heidenhain_result = result
            st.session_state.heidenhain_stats = stats

            st.session_state.heidenhain_processing = False

            st.success(
                "✅ Fichier Heidenhain créé."
            )

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
            key="create_heidenhain",
        ):

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
        f"Fichier créé : "
        f"{stats['created']:,} "
        f"nouvelles lignes SAV."
    )

    st.download_button(
        "⬇️ Télécharger le fichier Heidenhain",
        data=st.session_state.heidenhain_result,
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
        "Créez d'abord le fichier Heidenhain "
        "puis importez le fichier Odoo "
        "et le fichier catégorie."
    )

else:

    st.markdown(
        """
### Correspondance automatique

| Odoo | Source |
|---|---|
| `id` | ID Odoo existant |
| `default_code` | ID Heidenhain |
| `name` | Description |
| `x_studio_marque_1` | Marque |
| `categ_id` | Groupe Produit + fichier catégorie |
| `x_studio_sales_status` | Statut |
| `list_price` | Prix PPC / Prix SAV |
| `seller_ids/partner_id` | Fournisseur par défaut |
| `seller_ids/price` | Prix HA |
| `purchase_ok` | Paramètre |
| `sale_ok` | Paramètre |
| `type` | Paramètre |
| `invoice_policy` | Paramètre |
| `barcode` | Généré depuis la référence |

Les colonnes demandées mais absentes du fichier Odoo sont ajoutées automatiquement à droite.
        """
    )

    if st.session_state.odoo_processing:

        st.button(
            "⏳ Calcul Odoo en cours...",
            disabled=True,
            use_container_width=True,
            key="prepare_odoo_locked",
        )

        st.warning(
            "⏳ La préparation Odoo est en cours."
        )

        progress_odoo = st.progress(
            0,
            text="Initialisation..."
        )

        status_odoo = st.empty()

        try:

            prepared_heidenhain_file = BytesIO(
                st.session_state.heidenhain_result
            )

            mapping = {
                odoo_col_id: "id",
                odoo_col_default_code: "ID",
                odoo_col_name: "Description",
                odoo_col_marque: "Marque",
                odoo_col_category: "Groupe Produit",
                odoo_col_status: "Statut",
                odoo_col_price: "list_price",
                odoo_col_supplier: "seller_ids/partner_id",
                odoo_col_supplier_price: "Prix HA",
                odoo_col_purchase: "purchase_ok",
                odoo_col_sale: "sale_ok",
                odoo_col_type: "type",
                odoo_col_invoice: "invoice_policy",
                odoo_col_barcode: "barcode",
            }

            # ==================================================
            # IMPORTANT :
            # L'ID Odoo est spécial.
            # On le détecte / crée mais on ne l'écrit jamais
            # avec une valeur Heidenhain.
            # ==================================================

            result, stats = process_odoo(

                heidenhain_file=(
                    prepared_heidenhain_file
                ),

                odoo_file=odoo_file,

                category_file=category_file,

                heidenhain_sheet=(
                    heidenhain_sheet
                ),

                odoo_sheet=(
                    "Sheet1"
                ),

                category_sheet=(
                    category_sheet
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

                mapping=mapping,

                default_supplier=(
                    default_supplier
                ),

                default_purchase_ok=(
                    default_purchase_ok
                ),

                default_sale_ok=(
                    default_sale_ok
                ),

                default_type=(
                    default_type
                ),

                default_invoice_policy=(
                    default_invoice_policy
                ),

                progress=progress_odoo,

                status_display=status_odoo,
            )

            st.session_state.odoo_result = result
            st.session_state.odoo_stats = stats

            st.session_state.odoo_processing = False

            st.success(
                "✅ Fichier Odoo préparé avec succès."
            )

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
            key="prepare_odoo",
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

    if stats["created_columns"]:

        st.success(
            "➕ Colonnes Odoo ajoutées automatiquement : "
            + ", ".join(
                stats["created_columns"]
            )
        )

    if stats["category_missing"] > 0:

        st.warning(
            f"⚠️ {stats['category_missing']:,} "
            "référence(s) sans catégorie."
        )

    if stats["empty_references"] > 0:

        st.warning(
            f"⚠️ {stats['empty_references']:,} "
            "ligne(s) sans référence."
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
        f"🆔 {stats['existing_ids_preserved']:,} "
        "ID Odoo existants conservés."
    )

    st.download_button(
        label=(
            "⬇️ Télécharger le fichier Odoo "
            "prêt à importer"
        ),
        data=st.session_state.odoo_result,
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
