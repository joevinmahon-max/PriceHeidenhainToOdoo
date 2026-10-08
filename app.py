# ==========================================================
# PREPARATION PRIX HEIDENHAIN + MISE EN FORME ODOO
# JMA
#
# VERSION :
# - Détection automatique des colonnes Heidenhain
# - Sélection des colonnes à conserver
# - Détection automatique des colonnes Odoo
# - Correspondance Heidenhain -> Odoo configurable
# - ID Heidenhain -> colonne Odoo configurable
# - Les colonnes sélectionnées à l'étape 1 sont utilisées
#   pour construire les correspondances de l'étape 2
# - Conservation de la logique VG / PG / SAV
# - Index Odoo en mémoire
# - Recherche catégorie optimisée
# - Compatible gros fichiers
# - Boutons verrouillés pendant les calculs
# - Barre de progression
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

st.title(
    "📦 Préparation des prix HEIDENHAIN --> Odoo"
)


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


# ==========================================================
# RECHERCHE COLONNES
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
# DETECTION AUTOMATIQUE DES COLONNES
# ==========================================================

def detect_columns_from_file(
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

        names = ", ".join(
            wb.sheetnames
        )

        wb.close()

        raise ValueError(
            f"Feuille '{sheet_name}' introuvable.\n"
            f"Feuilles disponibles : {names}"
        )

    ws = wb[
        sheet_name
    ]

    columns = []

    for cell in ws[
        int(header_row)
    ]:

        if cell.value is None:
            continue

        name = str(
            cell.value
        ).strip()

        if not name:
            continue

        if name not in columns:

            columns.append(name)

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

        # Prix calculés
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
                "La colonne ID doit être présente "
                "pour créer les lignes SAV."
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
# TRAITEMENT HEIDENHAIN
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
            "La colonne ID doit être sélectionnée."
        )

    # ------------------------------------------------------
    # OUVERTURE FORMULES
    # ------------------------------------------------------

    if status_display is not None:
        status_display.info(
            "📂 Ouverture du fichier Heidenhain..."
        )

    if progress is not None:
        progress.progress(
            2,
            text="📂 Ouverture du fichier Heidenhain..."
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
            f"Feuille '{sheet_name}' introuvable.\n"
            f"Feuilles disponibles : {names}"
        )

    source_ws = wb_formula[
        sheet_name
    ]

    # ------------------------------------------------------
    # VALEURS CALCULEES
    # ------------------------------------------------------

    if status_display is not None:
        status_display.info(
            "📊 Lecture des valeurs calculées..."
        )

    uploaded_file.seek(0)

    wb_values = openpyxl.load_workbook(
        uploaded_file,
        data_only=True,
    )

    values_ws = wb_values[
        sheet_name
    ]

    # ------------------------------------------------------
    # COLONNES NECESSAIRES
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
            "Colonnes introuvables dans Heidenhain : "
            + ", ".join(missing)
        )

    status_col = source_columns[
        status_column
    ]

    id_col = source_columns[
        id_column
    ]

    # ------------------------------------------------------
    # DERNIERE LIGNE
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

    # ------------------------------------------------------
    # NOUVEAU CLASSEUR
    # ------------------------------------------------------

    output_wb = openpyxl.Workbook()

    output_ws = output_wb.active
    output_ws.title = sheet_name

    # ------------------------------------------------------
    # EN-TETES
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
    # IDS EXISTANTS
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

        normalized = normalize(
            value
        )

        if normalized:
            existing_ids.add(
                normalized
            )

    # ------------------------------------------------------
    # TRAITEMENT
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
    # LIGNES SAV
    # ------------------------------------------------------

    if status_display is not None:

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
    # LARGEURS
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
    # SAUVEGARDE
    # ------------------------------------------------------

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
        "vg_pg": vg_count + pg_count,
        "created": len(rows_to_create),
        "duplicates": duplicate_count,
        "empty_ids": empty_id_count,
        "created_ids": [
            x[1]
            for x in rows_to_create
        ],
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
                    .split(
                        separator,
                        1,
                    )[0]
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

    normalized_group = normalize(
        group
    )

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
                .split(
                    separator,
                    1,
                )[0]
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

    mappings,

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

    # ======================================================
    # COLONNES HEIDENHAIN
    # ======================================================

    selected_heidenhain_columns = [
        mapping["heidenhain"]
        for mapping in mappings
        if mapping.get("heidenhain")
    ]

    h_columns = find_columns(
        h_ws,
        1,
        selected_heidenhain_columns,
    )

    missing_h = [
        column
        for column in selected_heidenhain_columns
        if column not in h_columns
    ]

    if missing_h:

        wb_h.close()

        raise ValueError(
            "Colonnes introuvables dans le fichier "
            "Heidenhain préparé : "
            + ", ".join(missing_h)
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
    # DETECTION DES COLONNES ODOO
    # ======================================================

    odoo_column_names = []

    for mapping in mappings:

        odoo_column = mapping.get(
            "odoo"
        )

        if odoo_column:

            if odoo_column not in odoo_column_names:

                odoo_column_names.append(
                    odoo_column
                )

    odoo_columns = find_columns(
        ws,
        int(odoo_header_row),
        odoo_column_names,
    )

    missing_odoo = [
        column
        for column in odoo_column_names
        if column not in odoo_columns
    ]

    if missing_odoo:

        wb_o.close()
        wb_h.close()

        raise ValueError(
            "Colonnes introuvables dans Odoo : "
            + ", ".join(missing_odoo)
        )

    # ======================================================
    # INDEX ODOO
    # ======================================================

    reference_mapping = next(
        (
            m
            for m in mappings
            if m.get("role") == "reference"
        ),
        None,
    )

    if reference_mapping is None:

        wb_o.close()
        wb_h.close()

        raise ValueError(
            "Aucune correspondance de référence "
            "Heidenhain -> Odoo n'a été définie."
        )

    odoo_reference_column = (
        reference_mapping["odoo"]
    )

    heidenhain_reference_column = (
        reference_mapping["heidenhain"]
    )

    reference_col = odoo_columns[
        odoo_reference_column
    ]

    h_reference_col = h_columns[
        heidenhain_reference_column
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

        ref = clean_reference(
            value
        )

        if ref and ref not in reference_index:

            reference_index[
                ref
            ] = row

    # ======================================================
    # GROUPE PRODUIT
    # ======================================================

    group_mapping = next(
        (
            m
            for m in mappings
            if m.get("role") == "category"
        ),
        None,
    )

    # ======================================================
    # PRIX
    # ======================================================

    price_ppc_mapping = next(
        (
            m
            for m in mappings
            if m.get("role") == "price_ppc"
        ),
        None,
    )

    price_sav_mapping = next(
        (
            m
            for m in mappings
            if m.get("role") == "price_sav"
        ),
        None,
    )

    prixha_mapping = next(
        (
            m
            for m in mappings
            if m.get("role") == "prix_ha"
        ),
        None,
    )

    # ======================================================
    # AUTRES REGLES
    # ======================================================

    special_roles = {
        "status",
        "barcode",
        "supplier",
        "purchase",
        "sale",
        "type",
        "invoice",
        "name",
        "brand",
    }

    # ======================================================
    # STATISTIQUES
    # ======================================================

    updated = 0
    created = 0

    category_found = 0
    category_missing = 0

    references_processed = 0
    empty_references = 0

    total_h_rows = max(
        0,
        h_ws.max_row - 1,
    )

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

    for index, row_values in enumerate(
        h_ws.iter_rows(
            min_row=2,
            values_only=True,
        )
    ):

        try:

            reference = row_values[
                h_reference_col - 1
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

        # --------------------------------------------------
        # RECHERCHE ODOO
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
        # APPLICATION DES CORRESPONDANCES
        # ==================================================

        for mapping in mappings:

            role = mapping.get(
                "role"
            )

            h_name = mapping.get(
                "heidenhain"
            )

            odoo_name = mapping.get(
                "odoo"
            )

            if not h_name or not odoo_name:
                continue

            if h_name not in h_columns:
                continue

            if odoo_name not in odoo_columns:
                continue

            h_col = h_columns[
                h_name
            ]

            odoo_col = odoo_columns[
                odoo_name
            ]

            try:

                value = row_values[
                    h_col - 1
                ]

            except IndexError:

                value = None

            # --------------------------------------------------
            # PRIX PPC / SAV
            # --------------------------------------------------

            if role == "price_ppc":

                if normalized_ref.endswith(
                    "_SAV"
                ):

                    continue

            elif role == "price_sav":

                if normalized_ref.endswith(
                    "_SAV"
                ):

                    ws.cell(
                        row=target_row,
                        column=odoo_col,
                    ).value = value

                continue

            # --------------------------------------------------
            # REFERENCE
            # --------------------------------------------------

            elif role == "reference":

                value = reference

            # --------------------------------------------------
            # CATEGORIE
            # --------------------------------------------------

            elif role == "category":

                category_id = (
                    find_category_id_fast(
                        value,
                        category_mapping,
                    )
                )

                if category_id is not None:

                    category_found += 1

                    value = category_id

                else:

                    category_missing += 1

                    continue

            # --------------------------------------------------
            # BARCODE
            # --------------------------------------------------

            elif role == "barcode":

                if normalized_ref.endswith(
                    "_SAV"
                ):

                    value = ""

                else:

                    value = (
                        f"I {reference}"
                    )

            # --------------------------------------------------
            # FOURNISSEUR
            # --------------------------------------------------

            elif role == "supplier":

                value = "HEIDENHAIN FRANCE"

            # --------------------------------------------------
            # ACHAT
            # --------------------------------------------------

            elif role == "purchase":

                value = "VRAI"

            # --------------------------------------------------
            # VENTE
            # --------------------------------------------------

            elif role == "sale":

                value = "VRAI"

            # --------------------------------------------------
            # TYPE
            # --------------------------------------------------

            elif role == "type":

                value = "Consommable"

            # --------------------------------------------------
            # FACTURATION
            # --------------------------------------------------

            elif role == "invoice":

                value = "Quantités livrées"

            # --------------------------------------------------
            # ECRITURE
            # --------------------------------------------------

            ws.cell(
                row=target_row,
                column=odoo_col,
            ).value = value

        # ==================================================
        # PROGRESSION
        # ==================================================

        if (
            index % 10 == 0
            or index == total_h_rows - 1
        ):

            pct = 40 + int(
                (
                    (index + 1)
                    / max(
                        1,
                        total_h_rows,
                    )
                )
                * 50
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
                    ),
                )

    # ======================================================
    # SAUVEGARDE
    # ======================================================

    if status_display is not None:

        status_display.info(
            "💾 Création du fichier Odoo final..."
        )

    if progress is not None:

        progress.progress(
            95,
            text="💾 Création du fichier Odoo..."
        )

    output = BytesIO()

    wb_o.save(
        output
    )

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
        "total_time": total_time,
    }

    return result, stats


# ==========================================================
# FICHIERS
# ==========================================================

st.header(
    "📂 Fichiers"
)

file_col1, file_col2, file_col3 = st.columns(3)


with file_col1:

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


with file_col2:

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


with file_col3:

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
# SIDEBAR
# ==========================================================

st.sidebar.header(
    "⚙️ Paramètres"
)


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
# DETECTION AUTOMATIQUE
# ==========================================================

heidenhain_available_columns = []

if heidenhain_file is not None:

    try:

        heidenhain_available_columns = (
            detect_columns_from_file(
                uploaded_file=heidenhain_file,
                sheet_name=heidenhain_sheet,
                header_row=heidenhain_header_row,
            )
        )

        st.sidebar.success(
            f"📋 {len(heidenhain_available_columns)} "
            "colonnes Heidenhain détectées."
        )

    except Exception as e:

        st.sidebar.error(
            f"❌ Détection impossible : {e}"
        )


# ==========================================================
# COLONNES A CONSERVER
# ==========================================================

st.sidebar.subheader(
    "📋 Colonnes à conserver à l'étape 1"
)

if heidenhain_available_columns:

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
        column
        for column in preferred_columns
        if column in heidenhain_available_columns
    ]

    heidenhain_output_columns = (
        st.sidebar.multiselect(
            "Sélectionner les colonnes à garder",
            options=heidenhain_available_columns,
            default=default_columns,
            help=(
                "Les colonnes sont détectées "
                "automatiquement depuis le fichier "
                "Heidenhain."
            ),
        )
    )

else:

    heidenhain_output_columns = []

    st.sidebar.info(
        "Importez le fichier Heidenhain "
        "pour détecter ses colonnes."
    )


# ==========================================================
# COLONNES NECESSAIRES AU TRAITEMENT
# ==========================================================

status_column = "Statut"
id_column = "ID"

required_processing_columns = [
    id_column,
    status_column,
]

if heidenhain_available_columns:

    missing_processing = [
        column
        for column in required_processing_columns
        if column not in heidenhain_available_columns
    ]

    if missing_processing:

        st.sidebar.warning(
            "⚠️ Colonnes nécessaires absentes : "
            + ", ".join(
                missing_processing
            )
        )


# ==========================================================
# PARAMETRES ODOO
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
# DETECTION AUTOMATIQUE ODOO
# ==========================================================

odoo_available_columns = []

if odoo_file is not None:

    try:

        odoo_available_columns = (
            detect_columns_from_file(
                uploaded_file=odoo_file,
                sheet_name=odoo_sheet,
                header_row=odoo_header_row,
            )
        )

        st.sidebar.success(
            f"📗 {len(odoo_available_columns)} "
            "colonnes Odoo détectées."
        )

    except Exception as e:

        st.sidebar.error(
            f"❌ Détection Odoo impossible : {e}"
        )


# ==========================================================
# CORRESPONDANCES
# ==========================================================

st.sidebar.subheader(
    "🔗 Correspondances Heidenhain → Odoo"
)

mappings = []


if (
    heidenhain_output_columns
    and odoo_available_columns
):

    st.sidebar.caption(
        "Chaque colonne Heidenhain sélectionnée "
        "peut être associée à une colonne Odoo."
    )

    # ------------------------------------------------------
    # DETECTION AUTOMATIQUE DES CORRESPONDANCES
    # ------------------------------------------------------

    automatic_matches = {
        "ID": [
            "default_code",
            "Référence interne",
            "Reference interne",
            "Référence",
        ],
        "Description": [
            "name",
            "Nom",
        ],
        "Marque": [
            "x_studio_marque",
            "Marque",
        ],
        "Statut": [
            "Sales Status",
            "Sales status",
            "Statut",
        ],
        "Groupe Produit": [
            "categ_id",
            "Catégorie de produits/ID",
            "Catégorie",
        ],
        "Prix (PPC)": [
            "list_price",
            "Prix de vente",
        ],
        "Prix (SAV)": [
            "list_price",
            "Prix de vente",
        ],
        "Prix HA": [
            "standard_price",
            "Fournisseurs/Prix",
        ],
    }

    odoo_choices = [
        "— Ne pas importer —"
    ] + odoo_available_columns

    for heidenhain_column in (
        heidenhain_output_columns
    ):

        possible_matches = automatic_matches.get(
            heidenhain_column,
            [],
        )

        default_odoo = None

        for candidate in possible_matches:

            if candidate in odoo_available_columns:

                default_odoo = candidate
                break

        if default_odoo is not None:

            default_index = (
                odoo_choices.index(
                    default_odoo
                )
            )

        else:

            default_index = 0

        selected_odoo = st.sidebar.selectbox(
            f"{heidenhain_column} →",
            options=odoo_choices,
            index=default_index,
            key=(
                f"mapping_"
                f"{heidenhain_column}"
            ),
        )

        if selected_odoo != "— Ne pas importer —":

            role = None

            if heidenhain_column == "ID":
                role = "reference"

            elif heidenhain_column == "Description":
                role = "name"

            elif heidenhain_column == "Marque":
                role = "brand"

            elif heidenhain_column == "Statut":
                role = "status"

            elif heidenhain_column == "Groupe Produit":
                role = "category"

            elif heidenhain_column == "Prix (PPC)":
                role = "price_ppc"

            elif heidenhain_column == "Prix (SAV)":
                role = "price_sav"

            elif heidenhain_column == "Prix HA":
                role = "prix_ha"

            mappings.append(
                {
                    "heidenhain": heidenhain_column,
                    "odoo": selected_odoo,
                    "role": role,
                }
            )


# ==========================================================
# REGLES SPECIALES ODOO
# ==========================================================

if odoo_available_columns:

    st.sidebar.subheader(
        "⚙️ Champs Odoo automatiques"
    )

    special_fields = [
        (
            "Code-barres",
            "barcode",
            [
                "barcode",
                "Code-barres",
            ],
        ),
        (
            "Fournisseur",
            "supplier",
            [
                "seller_ids",
                "Fournisseurs/Fournisseur",
                "Fournisseur",
            ],
        ),
        (
            "Peut être acheté",
            "purchase",
            [
                "purchase_ok",
                "Peut être acheté",
            ],
        ),
        (
            "Peut être vendu",
            "sale",
            [
                "sale_ok",
                "Peut être vendu",
            ],
        ),
        (
            "Type de produit",
            "type",
            [
                "detailed_type",
                "Type de produit",
                "type",
            ],
        ),
        (
            "Politique de facturation",
            "invoice",
            [
                "invoice_policy",
                "Politique de facturation",
            ],
        ),
    ]

    for label, role, candidates in special_fields:

        found = None

        for candidate in candidates:

            if candidate in odoo_available_columns:

                found = candidate
                break

        if found:

            st.sidebar.caption(
                f"✓ {label} → {found}"
            )

            # Ajouter le champ à la liste des règles
            # avec le premier champ Heidenhain pertinent.
            source_column = None

            if role == "barcode":
                source_column = "ID"

            elif role == "supplier":
                source_column = "ID"

            elif role == "purchase":
                source_column = "ID"

            elif role == "sale":
                source_column = "ID"

            elif role == "type":
                source_column = "ID"

            elif role == "invoice":
                source_column = "ID"

            if (
                source_column
                and source_column
                in heidenhain_output_columns
            ):

                mappings.append(
                    {
                        "heidenhain": source_column,
                        "odoo": found,
                        "role": role,
                    }
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
# VALIDATION DES PARAMETRES
# ==========================================================

reference_mapping_exists = any(
    mapping.get("role") == "reference"
    for mapping in mappings
)

if (
    heidenhain_file is not None
    and heidenhain_output_columns
    and "ID" not in heidenhain_output_columns
):

    st.sidebar.error(
        "❌ ID doit être sélectionné."
    )

if (
    odoo_file is not None
    and heidenhain_output_columns
    and odoo_available_columns
    and not reference_mapping_exists
):

    st.sidebar.error(
        "❌ Il faut obligatoirement associer "
        "ID Heidenhain à une colonne Odoo "
        "pour identifier les références."
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

- détecter automatiquement les colonnes du fichier ;
- conserver uniquement les colonnes sélectionnées ;
- récupérer les valeurs calculées ;
- rechercher les statuts VG et PG ;
- créer les références `_SAV` ;
- éviter les doublons ;
- ajouter les nouvelles lignes SAV à la fin.
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
                header_row=heidenhain_header_row,
                data_start_row=(
                    heidenhain_header_row + 1
                ),
                status_column=status_column,
                id_column=id_column,
                output_columns=(
                    heidenhain_output_columns
                ),
                progress=progress_heidenhain,
                status_display=status_heidenhain,
            )

            st.session_state.heidenhain_result = result
            st.session_state.heidenhain_stats = stats
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
            key="create_heidenhain",
            disabled=(
                not heidenhain_output_columns
                or "ID"
                not in heidenhain_output_columns
            ),
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
        f"{stats['created']:,} nouvelles lignes SAV."
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
        "puis importez Odoo et le fichier catégorie."
    )

else:

    st.markdown(
        """
Le traitement utilise maintenant les correspondances
définies dans les paramètres :

**Exemple :**

`ID Heidenhain → default_code Odoo`

`Description → name`

`Prix (PPC) → list_price`

`Prix HA → standard_price`

Les colonnes non sélectionnées ne sont pas utilisées.
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
                odoo_file=odoo_file,
                category_file=category_file,

                heidenhain_sheet=(
                    heidenhain_sheet
                ),
                odoo_sheet=odoo_sheet,
                category_sheet=category_sheet,

                heidenhain_header_row=1,
                odoo_header_row=int(
                    odoo_header_row
                ),

                category_id_col=int(
                    category_id_col
                ),
                category_search_col=int(
                    category_search_col
                ),

                mappings=mappings,

                progress=progress_odoo,
                status_display=status_odoo,
            )

            st.session_state.odoo_result = result
            st.session_state.odoo_stats = stats
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
            key="prepare_odoo",
            disabled=(
                not reference_mapping_exists
            ),
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
        f"📂 {stats['category_rows']:,} lignes "
        "de catégories chargées."
    )

    st.info(
        f"🔎 {stats['odoo_indexed']:,} références "
        "Odoo indexées."
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
