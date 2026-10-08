# ==========================================================
# PREPARATION PRIX HEIDENHAIN + MISE EN FORME ODOO
# JMA
#
# VERSION :
# - Détection automatique des colonnes Heidenhain
# - Détection automatique des colonnes Odoo
# - Correspondances Heidenhain -> Odoo configurables
# - Colonnes Odoo existantes conservées dans leur ordre
# - Colonnes Odoo absentes ajoutées à l'emplacement choisi
# - Détection automatique du champ "id"
# - Conservation de l'ID Odoo existant
# - PPC / SAV conservés
# - Prix HA conservé
# - Catégories conservées
# - Création des références absentes
# - Mise à jour des références existantes
# - Compatible gros fichiers
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
# CORRESPONDANCES
# ==========================================================
#
# ODOO                       <- HEIDENHAIN
#
# Les noms Odoo sont les noms techniques/export.
#
# Pour les champs fixes, source = None.
#
# ==========================================================

ODOO_FIELD_CONFIGURATION = {

    "id": {
        "label": "ID Odoo",
        "source": None,
        "fixed": None,
        "description": "ID Odoo existant. Conservé automatiquement.",
    },

    "default_code": {
        "label": "Référence interne",
        "source": "ID",
        "fixed": None,
    },

    "name": {
        "label": "Nom",
        "source": "Description",
        "fixed": None,
    },

    "x_studio_marque_1": {
        "label": "Marque",
        "source": "Marque",
        "fixed": None,
    },

    "categ_id": {
        "label": "Catégorie",
        "source": "Groupe Produit",
        "fixed": None,
    },

    "x_studio_sales_status": {
        "label": "Sales Status",
        "source": "Statut",
        "fixed": None,
    },

    "list_price": {
        "label": "Prix de vente",
        "source": "__PRICE__",
        "fixed": None,
    },

    "seller_ids/partner_id": {
        "label": "Fournisseur",
        "source": None,
        "fixed": "HEIDENHAIN FRANCE",
    },

    "seller_ids/price": {
        "label": "Prix fournisseur",
        "source": "Prix HA",
        "fixed": None,
    },

    "purchase_ok": {
        "label": "Peut être acheté",
        "source": None,
        "fixed": "VRAI",
    },

    "sale_ok": {
        "label": "Peut être vendu",
        "source": None,
        "fixed": "VRAI",
    },

    "type": {
        "label": "Type de produit",
        "source": None,
        "fixed": "Consommable",
    },

    "invoice_policy": {
        "label": "Politique de facturation",
        "source": None,
        "fixed": "Quantités livrées",
    },

    "barcode": {
        "label": "Code-barres",
        "source": "__BARCODE__",
        "fixed": None,
    },
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
# DETECTION DES COLONNES
# ==========================================================

def get_headers(
    ws,
    header_row,
):

    result = []

    seen = set()

    for cell in ws[header_row]:

        value = cell.value

        if value in (None, ""):
            continue

        name = str(value).strip()

        if not name:
            continue

        normalized = normalize(name)

        if normalized in seen:
            continue

        seen.add(normalized)

        result.append(name)

    return result


def find_column(
    ws,
    header_row,
    header_name,
):

    wanted = normalize(header_name)

    for cell in ws[header_row]:

        if normalize(cell.value) == wanted:
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

        value = normalize(cell.value)

        if value and value not in headers:

            headers[value] = cell.column

    for name in names:

        normalized = normalize(name)

        if normalized in headers:

            result[name] = headers[normalized]

    return result


def detect_excel_headers(
    uploaded_file,
    sheet_name,
    header_row,
):

    if uploaded_file is None:
        return []

    uploaded_file.seek(0)

    wb = openpyxl.load_workbook(
        uploaded_file,
        read_only=True,
        data_only=False,
    )

    if sheet_name not in wb.sheetnames:

        wb.close()
        return []

    ws = wb[sheet_name]

    headers = get_headers(
        ws,
        int(header_row),
    )

    wb.close()

    return headers


# ==========================================================
# FEUILLES
# ==========================================================

def get_sheet_names(
    uploaded_file,
):

    if uploaded_file is None:
        return []

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

        source_col = source_columns[column_name]

        source_cell = source_ws.cell(
            row=source_row,
            column=source_col,
        )

        target_cell = output_ws.cell(
            row=output_row,
            column=output_col,
        )

        if column_name in {
            "Prix (SAV)",
            "Prix HA",
        }:

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
            output_columns.index("ID") + 1
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

    if "ID" not in output_columns:

        raise ValueError(
            "La colonne ID doit être conservée "
            "pour pouvoir créer les références SAV."
        )

    uploaded_file.seek(0)

    if status_display:
        status_display.info(
            "📂 Ouverture du fichier Heidenhain..."
        )

    if progress:
        progress.progress(
            2,
            text="📂 Ouverture du fichier Heidenhain..."
        )

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

    source_ws = wb_formula[sheet_name]

    uploaded_file.seek(0)

    wb_values = openpyxl.load_workbook(
        uploaded_file,
        data_only=True,
    )

    values_ws = wb_values[sheet_name]

    # ------------------------------------------------------
    # Détection colonnes
    # ------------------------------------------------------

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
            "Colonnes Heidenhain introuvables : "
            + ", ".join(missing)
        )

    status_col = source_columns[
        status_column
    ]

    id_col = source_columns[
        id_column
    ]

    # ------------------------------------------------------
    # Dernière ligne
    # ------------------------------------------------------

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

    # ------------------------------------------------------
    # Nouveau fichier
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
            column=source_columns[column_name],
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
            existing_ids.add(normalized)

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
            progress
            and (
                index % 10 == 0
                or index == total_rows - 1
            )
        ):

            pct = 10 + int(
                (
                    (index + 1)
                    / max(1, total_rows)
                )
                * 75
            )

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

    # ------------------------------------------------------
    # Lignes SAV
    # ------------------------------------------------------

    if status_display:

        status_display.info(
            f"➕ Création des lignes SAV "
            f"({len(rows_to_create):,})..."
        )

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
        "vg_pg": vg_count + pg_count,
        "created": len(rows_to_create),
        "duplicates": duplicate_count,
        "empty_ids": empty_id_count,
    }

    if progress:
        progress.progress(
            100,
            text="✅ Étape 1 terminée",
        )

    if status_display:
        status_display.success(
            "✅ Fichier Heidenhain créé."
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
            f"Feuille catégorie "
            f"'{category_sheet}' introuvable.\n"
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

        normalized_text = normalize(
            category_text
        )

        if normalized_text not in mapping:

            mapping[
                normalized_text
            ] = category_id

        for separator in ["-", "–"]:

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

                    if (
                        normalized_first
                        not in mapping
                    ):

                        mapping[
                            normalized_first
                        ] = category_id

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

    category_id = category_mapping.get(
        normalize(group)
    )

    if category_id is not None:
        return category_id

    for separator in ["-", "–"]:

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
# AJOUT DES COLONNES ODOO MANQUANTES
# ==========================================================

def ensure_odoo_columns(
    ws,
    header_row,
    required_columns,
    placement_config,
):

    existing_headers = get_headers(
        ws,
        header_row,
    )

    existing_normalized = {
        normalize(x): x
        for x in existing_headers
    }

    missing_columns = [
        x
        for x in required_columns
        if normalize(x) not in existing_normalized
    ]

    if not missing_columns:
        return

    # ------------------------------------------------------
    # On ajoute les colonnes une par une.
    #
    # Pour conserver l'ordre existant :
    # - aucune colonne existante n'est déplacée sauf
    #   lorsqu'une nouvelle colonne est explicitement
    #   demandée avant une colonne existante.
    # ------------------------------------------------------

    for new_column in missing_columns:

        placement = placement_config.get(
            new_column,
            "__END__",
        )

        # Toujours relire les colonnes après
        # chaque insertion.
        current_headers = get_headers(
            ws,
            header_row,
        )

        if placement == "__END__":

            new_col = ws.max_column + 1

        else:

            before_normalized = normalize(
                placement
            )

            target_col = None

            for cell in ws[header_row]:

                if normalize(
                    cell.value
                ) == before_normalized:

                    target_col = cell.column
                    break

            if target_col is None:

                new_col = ws.max_column + 1

            else:

                ws.insert_cols(
                    target_col,
                    1,
                )

                new_col = target_col

        ws.cell(
            row=header_row,
            column=new_col,
        ).value = new_column

        # Style depuis la colonne voisine
        if new_col > 1:

            copy_style_safe(
                ws.cell(
                    row=header_row,
                    column=new_col - 1,
                ),
                ws.cell(
                    row=header_row,
                    column=new_col,
                ),
            )


# ==========================================================
# PROCESS ODOO
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
    mapping_config,
    placement_config,
    odoo_header_row,
    progress=None,
    status_display=None,
):

    start_total = time.perf_counter()

    # ======================================================
    # CATEGORIES
    # ======================================================

    if status_display:

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

    if progress:
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
            f"'{heidenhain_sheet}' introuvable.\n"
            f"Feuilles disponibles : {names}"
        )

    h_ws = wb_h[
        heidenhain_sheet
    ]

    h_headers = get_headers(
        h_ws,
        1,
    )

    # ======================================================
    # VÉRIFICATION SOURCES
    # ======================================================

    required_sources = set()

    for odoo_field, config in mapping_config.items():

        source = config.get("source")

        if source in (
            None,
            "__PRICE__",
            "__BARCODE__",
        ):
            continue

        required_sources.add(source)

    missing_sources = [
        source
        for source in required_sources
        if normalize(source)
        not in {
            normalize(x)
            for x in h_headers
        }
    ]

    if missing_sources:

        wb_h.close()

        raise ValueError(
            "Colonnes source absentes du fichier "
            "Heidenhain : "
            + ", ".join(missing_sources)
        )

    h_columns = find_columns(
        h_ws,
        1,
        list(required_sources),
    )

    # ======================================================
    # ODOO
    # ======================================================

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
            f"'{odoo_sheet}' introuvable.\n"
            f"Feuilles disponibles : {names}"
        )

    ws = wb_o[
        odoo_sheet
    ]

    # ======================================================
    # COLONNES ODOO DEMANDEES
    # ======================================================

    required_odoo_columns = list(
        mapping_config.keys()
    )

    ensure_odoo_columns(
        ws=ws,
        header_row=int(odoo_header_row),
        required_columns=required_odoo_columns,
        placement_config=placement_config,
    )

    # ======================================================
    # RELECTURE COLONNES ODOO
    # ======================================================

    odoo_columns = find_columns(
        ws,
        int(odoo_header_row),
        required_odoo_columns,
    )

    missing_odoo = [
        x
        for x in required_odoo_columns
        if x not in odoo_columns
    ]

    if missing_odoo:

        wb_o.close()
        wb_h.close()

        raise ValueError(
            "Impossible de créer les colonnes Odoo : "
            + ", ".join(missing_odoo)
        )

    # ======================================================
    # DETECTION ID
    # ======================================================

    id_exists = "id" in odoo_columns

    # ======================================================
    # INDEX ODOO
    # ======================================================

    if status_display:

        status_display.info(
            "🔎 Création de l'index des références Odoo..."
        )

    reference_col = odoo_columns[
        "default_code"
    ]

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

        if ref and ref not in reference_index:

            reference_index[
                ref
            ] = row

    if progress:

        progress.progress(
            35,
            text=(
                f"🔎 Index Odoo créé : "
                f"{len(reference_index):,} références"
            ),
        )

    # ======================================================
    # COLONNES ODOO
    # ======================================================

    col = {
        name: odoo_columns[name]
        for name in required_odoo_columns
    }

    # ======================================================
    # COLONNES HEIDENHAIN
    # ======================================================

    def hcol(name):

        return h_columns.get(name)

    # ======================================================
    # STATISTIQUES
    # ======================================================

    updated = 0
    created = 0

    category_found = 0
    category_missing = 0

    processed = 0
    empty_references = 0

    # ======================================================
    # LIGNE VIDE ODOO
    # ======================================================

    next_empty_row = (
        int(odoo_header_row) + 1
    )

    while ws.cell(
        row=next_empty_row,
        column=reference_col,
    ).value not in (None, ""):

        next_empty_row += 1

    # ======================================================
    # PARCOURS HEIDENHAIN
    # ======================================================

    total_h_rows = max(
        0,
        h_ws.max_row - 1,
    )

    if status_display:

        status_display.info(
            f"⚙️ Traitement de "
            f"{total_h_rows:,} lignes Heidenhain..."
        )

    for index, row_values in enumerate(
        h_ws.iter_rows(
            min_row=2,
            values_only=True,
        )
    ):

        # --------------------------------------------------
        # ID HEIDENHAIN
        # --------------------------------------------------

        id_value = row_values[
            hcol("ID") - 1
        ]

        if id_value in (None, ""):
            break

        reference = str(
            id_value
        ).strip()

        if not reference:

            empty_references += 1
            continue

        processed += 1

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
        # ID ODOO
        # ==================================================
        #
        # IMPORTANT :
        # - si ligne existante : on ne touche pas à id
        # - si nouvelle ligne : id reste vide
        #
        # Cela permet de conserver l'ID Odoo existant.
        # ==================================================

        if id_exists and target_row >= next_empty_row - created:

            # Ne rien écrire.
            # Odoo attribuera un ID lors d'une création.

            pass

        # ==================================================
        # SOURCE VALUES
        # ==================================================

        status = row_values[
            hcol("Statut") - 1
        ]

        groupe = row_values[
            hcol("Groupe Produit") - 1
        ]

        ppc = row_values[
            hcol("Prix (PPC)") - 1
        ]

        sav = row_values[
            hcol("Prix (SAV)") - 1
        ]

        prix_ha = row_values[
            hcol("Prix HA") - 1
        ]

        description = row_values[
            hcol("Description") - 1
        ]

        marque = row_values[
            hcol("Marque") - 1
        ]

        # ==================================================
        # PRIX
        # ==================================================

        if is_sav(reference):

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
        # ÉCRITURE DES CHAMPS
        # ==================================================

        # default_code
        ws.cell(
            row=target_row,
            column=col["default_code"],
        ).value = reference

        # name
        ws.cell(
            row=target_row,
            column=col["name"],
        ).value = description

        # marque
        ws.cell(
            row=target_row,
            column=col["x_studio_marque_1"],
        ).value = marque

        # sales status
        ws.cell(
            row=target_row,
            column=col["x_studio_sales_status"],
        ).value = status

        # catégorie
        if category_id is not None:

            ws.cell(
                row=target_row,
                column=col["categ_id"],
            ).value = category_id

        # prix de vente
        ws.cell(
            row=target_row,
            column=col["list_price"],
        ).value = price

        # fournisseur
        ws.cell(
            row=target_row,
            column=col["seller_ids/partner_id"],
        ).value = (
            "HEIDENHAIN FRANCE"
        )

        # prix fournisseur
        ws.cell(
            row=target_row,
            column=col["seller_ids/price"],
        ).value = prix_ha

        # achat
        ws.cell(
            row=target_row,
            column=col["purchase_ok"],
        ).value = "VRAI"

        # vente
        ws.cell(
            row=target_row,
            column=col["sale_ok"],
        ).value = "VRAI"

        # type
        ws.cell(
            row=target_row,
            column=col["type"],
        ).value = "Consommable"

        # facturation
        ws.cell(
            row=target_row,
            column=col["invoice_policy"],
        ).value = (
            "Quantités livrées"
        )

        # ==================================================
        # CODE BARRE
        # ==================================================

        if is_sav(reference):

            barcode = ""

        else:

            barcode = f"I {reference}"

        ws.cell(
            row=target_row,
            column=col["barcode"],
        ).value = barcode

        # ==================================================
        # PROGRESSION
        # ==================================================

        if (
            progress
            and (
                index % 10 == 0
                or index == total_h_rows - 1
            )
        ):

            pct = 40 + int(
                (
                    (index + 1)
                    / max(1, total_h_rows)
                )
                * 50
            )

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

    # ======================================================
    # SAUVEGARDE
    # ======================================================

    if status_display:

        status_display.info(
            "💾 Création du fichier Odoo final..."
        )

    if progress:

        progress.progress(
            94,
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

    if progress:

        progress.progress(
            100,
            text="✅ Étape 2 terminée",
        )

    if status_display:

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
        "id_exists": id_exists,
        "total_time": total_time,
    }

    return result, stats


# ==========================================================
# FICHIERS
# ==========================================================

st.header("📂 Fichiers")

col1, col2, col3 = st.columns(3)

with col1:

    st.subheader("📘 Prix Heidenhain")

    heidenhain_file = st.file_uploader(
        "Fichier Prix Heidenhain",
        type=["xlsx", "xlsm"],
        key="heidenhain_file",
    )

with col2:

    st.subheader("📗 Fichier Odoo")

    odoo_file = st.file_uploader(
        "Fichier d'import Odoo",
        type=["xlsx", "xlsm"],
        key="odoo_file",
    )

with col3:

    st.subheader("📂 Catégorie de produit")

    category_file = st.file_uploader(
        "Fichier Catégorie de produit",
        type=["xlsx", "xlsm"],
        key="category_file",
    )


# ==========================================================
# SIDEBAR
# ==========================================================

st.sidebar.header("⚙️ Paramètres")


# ==========================================================
# PARAMETRES HEIDENHAIN
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
# DETECTION AUTOMATIQUE DES COLONNES HEIDENHAIN
# ==========================================================

heidenhain_detected_columns = []

if heidenhain_file is not None:

    heidenhain_detected_columns = (
        detect_excel_headers(
            heidenhain_file,
            heidenhain_sheet,
            heidenhain_header_row,
        )
    )

    if heidenhain_detected_columns:

        st.sidebar.success(
            f"📋 {len(heidenhain_detected_columns)} "
            "colonnes Heidenhain détectées."
        )

    else:

        st.sidebar.warning(
            "⚠️ Aucune colonne Heidenhain détectée."
        )


# ==========================================================
# COLONNES A CONSERVER
# ==========================================================

st.sidebar.subheader(
    "📋 Colonnes à conserver à l'étape 1"
)

if heidenhain_detected_columns:

    preferred_columns = [
        "ID",
        "Description",
        "Marque",
        "Statut",
        "Groupe Produit",
        "Prix (PPC)",
        "Prix (SAV)",
        "Prix HA",
    ]

    default_columns = [
        x
        for x in preferred_columns
        if x in heidenhain_detected_columns
    ]

    heidenhain_output_columns = (
        st.sidebar.multiselect(
            "Sélectionner les colonnes à garder",
            options=heidenhain_detected_columns,
            default=default_columns,
            help=(
                "Les colonnes sont détectées "
                "automatiquement depuis le fichier."
            ),
        )
    )

else:

    heidenhain_output_columns = []


# ==========================================================
# COLONNES HEIDENHAIN UTILISEES
# ==========================================================

status_column = "Statut"
id_column = "ID"
group_column = "Groupe Produit"
ppc_column = "Prix (PPC)"
sav_column = "Prix (SAV)"
prix_ha_column = "Prix HA"


if heidenhain_file is not None:

    required_step1 = [
        id_column,
        status_column,
        group_column,
        ppc_column,
        sav_column,
        prix_ha_column,
    ]

    missing_step1 = [
        x
        for x in required_step1
        if x not in heidenhain_detected_columns
    ]

    if missing_step1:

        st.sidebar.error(
            "⚠️ Colonnes nécessaires absentes : "
            + ", ".join(missing_step1)
        )


# ==========================================================
# ODOO
# ==========================================================

st.sidebar.divider()

st.sidebar.subheader(
    "📗 Fichier Odoo"
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


# ==========================================================
# DETECTION ODOO
# ==========================================================

odoo_detected_columns = []

if odoo_file is not None:

    odoo_detected_columns = detect_excel_headers(
        odoo_file,
        odoo_sheet,
        odoo_header_row,
    )

    if odoo_detected_columns:

        st.sidebar.success(
            f"📋 {len(odoo_detected_columns)} "
            "colonnes Odoo détectées."
        )

        if "id" in {
            normalize(x)
            for x in odoo_detected_columns
        }:

            st.sidebar.info(
                "ℹ️ Mode import Odoo activé : "
                "l'ID Odoo existant est conservé."
            )

    else:

        st.sidebar.warning(
            "⚠️ Aucune colonne Odoo détectée."
        )


# ==========================================================
# CONFIGURATION DES COLONNES ODOO MANQUANTES
# ==========================================================

st.sidebar.subheader(
    "🔗 Correspondances Odoo"
)

placement_config = {}

if odoo_file is not None and odoo_detected_columns:

    detected_normalized = {
        normalize(x)
        for x in odoo_detected_columns
    }

    missing_odoo_columns = [
        field
        for field in ODOO_FIELD_CONFIGURATION
        if normalize(field)
        not in detected_normalized
    ]

    if missing_odoo_columns:

        st.sidebar.warning(
            f"⚠️ {len(missing_odoo_columns)} "
            "colonne(s) Odoo absente(s)."
        )

        st.sidebar.caption(
            "Les colonnes existantes gardent leur ordre. "
            "Choisissez uniquement où placer les "
            "colonnes manquantes."
        )

        placement_options = [
            ("__END__", "➡️ À la fin")
        ]

        for existing in odoo_detected_columns:

            placement_options.append(
                (
                    existing,
                    f"Avant « {existing} »",
                )
            )

        placement_values = [
            value
            for value, label
            in placement_options
        ]

        placement_labels = {
            value: label
            for value, label
            in placement_options
        }

        for field in missing_odoo_columns:

            config = (
                ODOO_FIELD_CONFIGURATION[
                    field
                ]
            )

            label = config.get(
                "label",
                field,
            )

            st.sidebar.markdown(
                f"**{field}** — {label}"
            )

            selected = st.sidebar.selectbox(
                f"Où placer « {field} » ?",
                options=placement_values,
                format_func=lambda x: (
                    placement_labels[x]
                ),
                key=f"placement_{field}",
            )

            placement_config[field] = selected

    else:

        st.sidebar.success(
            "✅ Toutes les colonnes demandées "
            "existent déjà dans Odoo."
        )


# ==========================================================
# AFFICHAGE CORRESPONDANCES
# ==========================================================

if odoo_file is not None:

    with st.sidebar.expander(
        "👁️ Voir les correspondances",
        expanded=False,
    ):

        for field, config in (
            ODOO_FIELD_CONFIGURATION.items()
        ):

            source = config.get("source")
            fixed = config.get("fixed")

            if source == "__PRICE__":

                text = (
                    "Prix (SAV) si référence _SAV, "
                    "sinon Prix (PPC)"
                )

            elif source == "__BARCODE__":

                text = (
                    "Créé automatiquement à partir "
                    "du default_code"
                )

            elif source:

                text = (
                    f"← {source}"
                )

            elif fixed is not None:

                text = (
                    f"← valeur fixe : {fixed}"
                )

            else:

                text = (
                    "← valeur Odoo existante"
                )

            st.write(
                f"**{field}** {text}"
            )


# ==========================================================
# CATEGORIE
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
    step=1,
)

category_search_col = st.sidebar.number_input(
    "Colonne recherche catégorie",
    min_value=1,
    value=3,
    step=1,
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
        "👆 Importez le fichier Heidenhain pour commencer."
    )

else:

    st.markdown(
        """
Le fichier Heidenhain est analysé automatiquement.

Les colonnes détectées peuvent être sélectionnées
dans **⚙️ Paramètres → Colonnes à conserver**.

Le traitement conserve également la logique :

- Statuts VG / PG ;
- création des `_SAV` ;
- récupération du Prix (SAV) calculé ;
- récupération du Prix HA ;
- détection des doublons ;
- ajout des lignes SAV à la fin.
        """
    )

    if st.session_state.heidenhain_processing:

        st.button(
            "⏳ Calcul Heidenhain en cours...",
            disabled=True,
            use_container_width=True,
            key="create_heidenhain_locked",
        )

        progress_heidenhain = st.progress(
            0,
            text="Initialisation..."
        )

        status_heidenhain = st.empty()

        try:

            result, stats = process_heidenhain(

                uploaded_file=heidenhain_file,

                sheet_name=heidenhain_sheet,

                header_row=int(
                    heidenhain_header_row
                ),

                data_start_row=int(
                    heidenhain_header_row
                ) + 1,

                status_column=status_column,

                id_column=id_column,

                output_columns=(
                    heidenhain_output_columns
                ),

                progress=progress_heidenhain,

                status_display=status_heidenhain,
            )

            st.session_state.heidenhain_result = (
                result
            )

            st.session_state.heidenhain_stats = (
                stats
            )

            st.session_state.heidenhain_processing = (
                False
            )

            st.rerun()

        except Exception as e:

            st.session_state.heidenhain_processing = (
                False
            )

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

            if "ID" not in heidenhain_output_columns:

                st.error(
                    "⚠️ La colonne ID doit être conservée."
                )

            else:

                st.session_state.heidenhain_processing = (
                    True
                )

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
        "Créez d'abord le fichier Heidenhain, "
        "puis importez le fichier Odoo et le fichier catégorie."
    )

else:

    st.markdown(
        """
### Correspondances automatiques

- `id` ← ID Odoo existant
- `default_code` ← ID Heidenhain
- `name` ← Description
- `x_studio_marque_1` ← Marque
- `categ_id` ← Groupe Produit via le fichier catégorie
- `x_studio_sales_status` ← Statut
- `list_price` ← Prix (PPC) ou Prix (SAV)
- `seller_ids/partner_id` ← HEIDENHAIN FRANCE
- `seller_ids/price` ← Prix HA
- `purchase_ok` ← VRAI
- `sale_ok` ← VRAI
- `type` ← Consommable
- `invoice_policy` ← Quantités livrées
- `barcode` ← `I ` + ID, sauf `_SAV`
        """
    )

    if st.session_state.odoo_processing:

        st.button(
            "⏳ Calcul Odoo en cours...",
            disabled=True,
            use_container_width=True,
            key="prepare_odoo_locked",
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

                category_id_col=int(
                    category_id_col
                ),

                category_search_col=int(
                    category_search_col
                ),

                mapping_config=(
                    ODOO_FIELD_CONFIGURATION
                ),

                placement_config=(
                    placement_config
                ),

                odoo_header_row=int(
                    odoo_header_row
                ),

                progress=progress_odoo,

                status_display=status_odoo,
            )

            st.session_state.odoo_result = (
                result
            )

            st.session_state.odoo_stats = (
                stats
            )

            st.session_state.odoo_processing = (
                False
            )

            st.rerun()

        except Exception as e:

            st.session_state.odoo_processing = (
                False
            )

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

    if stats["id_exists"]:

        st.info(
            "ℹ️ Mode import Odoo activé : "
            "l'ID Odoo existant est conservé."
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
        f"⏱️ Temps total : "
        f"{stats['total_time']:.2f} secondes"
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
        "🎯 Le fichier Odoo final est prêt."
    )
