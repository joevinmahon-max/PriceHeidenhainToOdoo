# ==========================================================
# PREPARATION PRIX HEIDENHAIN + MISE EN FORME ODOO
# JMA
# ==========================================================

import streamlit as st
import openpyxl

from io import BytesIO
from copy import copy


# ==========================================================
# CONFIGURATION STREAMLIT
# ==========================================================

st.set_page_config(
    page_title="Préparation Heidenhain / Odoo",
    page_icon="📦",
    layout="wide",
)

st.title("📦 Préparation des prix HEIDENHAIN / Odoo")


# ==========================================================
# SESSION STATE
# ==========================================================

DEFAULT_STATE = {
    "heidenhain_result": None,
    "heidenhain_stats": None,
    "odoo_result": None,
    "odoo_stats": None,
}

for key, value in DEFAULT_STATE.items():
    if key not in st.session_state:
        st.session_state[key] = value


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
# UTILITAIRES
# ==========================================================

def normalize(value):
    if value is None:
        return ""

    return str(value).strip().upper()


def clean_reference(value):
    """
    Nettoie une référence pour les comparaisons.
    """
    return normalize(value)


def is_sav(reference):
    """
    Détermine si une référence est une référence SAV.
    """
    return clean_reference(reference).endswith("_SAV")


def copy_style_safe(source, target):
    """
    Copie uniquement les éléments de style nécessaires.

    On évite les manipulations globales du styles.xml
    qui peuvent provoquer des réparations Excel.
    """

    try:
        if source.font:
            target.font = copy(source.font)
    except Exception:
        pass

    try:
        if source.fill:
            target.fill = copy(source.fill)
    except Exception:
        pass

    try:
        if source.border:
            target.border = copy(source.border)
    except Exception:
        pass

    try:
        if source.alignment:
            target.alignment = copy(source.alignment)
    except Exception:
        pass

    try:
        if source.protection:
            target.protection = copy(source.protection)
    except Exception:
        pass

    try:
        target.number_format = source.number_format
    except Exception:
        pass


def copy_row_style(ws, source_row, target_row, max_col):
    """
    Copie la mise en forme d'une ligne existante
    vers une nouvelle ligne.
    """

    for col in range(1, max_col + 1):

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

    # Hauteur de ligne
    try:
        ws.row_dimensions[target_row].height = (
            ws.row_dimensions[source_row].height
        )
    except Exception:
        pass


def find_column(ws, header_row, header_name):
    """
    Recherche une colonne par son nom.
    """

    wanted = normalize(header_name)

    for cell in ws[header_row]:

        if normalize(cell.value) == wanted:
            return cell.column

    return None


def find_columns(ws, header_row, names):
    """
    Recherche plusieurs colonnes.
    """

    result = {}

    for name in names:

        col = find_column(
            ws,
            header_row,
            name,
        )

        if col is not None:
            result[name] = col

    return result


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
# ETAPE 1
# HEIDENHAIN
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

        # ----------------------------------------------
        # Prix SAV :
        # on prend la VALEUR calculée d'Excel
        # et non la formule.
        # ----------------------------------------------

        if column_name == "Prix (SAV)":

            target_cell.value = values_ws.cell(
                row=source_row,
                column=source_col,
            ).value

        else:

            target_cell.value = source_cell.value

        copy_style_safe(
            source_cell,
            target_cell,
        )


    # ----------------------------------------------
    # ID SAV
    # ----------------------------------------------

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


def process_heidenhain(
    uploaded_file,
    sheet_name,
    header_row,
    data_start_row,
    status_column,
    id_column,
    output_columns,
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

    # ------------------------------------------------------
    # OUVERTURE FORMULES
    # ------------------------------------------------------

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

    source_ws = wb_formula[sheet_name]

    # ------------------------------------------------------
    # OUVERTURE VALEURS CALCULEES
    # ------------------------------------------------------

    uploaded_file.seek(0)

    wb_values = openpyxl.load_workbook(
        uploaded_file,
        data_only=True,
    )

    values_ws = wb_values[sheet_name]

    # ------------------------------------------------------
    # COLONNES
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
            "Colonnes introuvables : "
            + ", ".join(missing)
        )

    status_col = source_columns[
        status_column
    ]

    id_col = source_columns[
        id_column
    ]

    max_row = source_ws.max_row

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
            column=source_columns[column_name],
        )

        target_header = output_ws.cell(
            row=1,
            column=output_col,
        )

        target_header.value = source_header.value

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

        normalized = normalize(value)

        if normalized:
            existing_ids.add(normalized)

    # ------------------------------------------------------
    # TRAITEMENT
    # ------------------------------------------------------

    rows_to_create = []

    vg_count = 0
    pg_count = 0
    duplicate_count = 0
    empty_id_count = 0

    output_row = 2

    progress = st.progress(
        0,
        text="Préparation Heidenhain...",
    )

    for index, source_row in enumerate(
        range(
            data_start_row,
            max_row + 1,
        )
    ):

        # ----------------------------------------------
        # Ligne originale
        # ----------------------------------------------

        copy_heidenhain_row(
            source_ws=source_ws,
            values_ws=values_ws,
            output_ws=output_ws,
            source_row=source_row,
            output_row=output_row,
            source_columns=source_columns,
            output_columns=output_columns,
        )

        # ----------------------------------------------
        # Hauteur
        # ----------------------------------------------

        output_ws.row_dimensions[
            output_row
        ].height = (
            source_ws.row_dimensions[
                source_row
            ].height
        )

        # ----------------------------------------------
        # Statut
        # ----------------------------------------------

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

        # ----------------------------------------------
        # ID SAV
        # ----------------------------------------------

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
            index % 5000 == 0
            or index == total_rows - 1
        ):

            pct = int(
                (
                    (index + 1)
                    / max(1, total_rows)
                )
                * 100
            )

            progress.progress(
                min(pct, 100),
                text=(
                    f"Heidenhain : "
                    f"{index + 1:,} / "
                    f"{total_rows:,}"
                ),
            )

    # ------------------------------------------------------
    # LIGNES SAV
    # ------------------------------------------------------

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

        created_ids.append(new_id)

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

        width = source_ws.column_dimensions[
            source_letter
        ].width

        if width:
            output_ws.column_dimensions[
                target_letter
            ].width = width

    output_ws.freeze_panes = "A2"

    # ------------------------------------------------------
    # SAUVEGARDE
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
        "created": len(created_ids),
        "duplicates": duplicate_count,
        "empty_ids": empty_id_count,
        "created_ids": created_ids,
    }

    return data, stats


# ==========================================================
# ETAPE 2
# ODOO — VERSION OPTIMISEE
# ==========================================================


# ==========================================================
# CHARGEMENT DU FICHIER CATEGORIE
# ==========================================================

def load_category_mapping(
    category_file,
    category_sheet,
    category_id_col,
    category_search_col,
):
    """
    Charge le fichier Catégorie de produit UNE SEULE FOIS.

    Construit plusieurs index pour permettre une recherche
    très rapide ensuite.

    Exemple :

        Colonne B = 20
        Colonne C = 20 - Métrologie

    permettra de retrouver :

        "20"
        -> "20"

    """

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
            f"introuvable.\n\n"
            f"Feuilles disponibles : {names}"
        )

    ws = wb[category_sheet]

    # ------------------------------------------------------
    # INDEX PRINCIPAL
    # ------------------------------------------------------

    mapping = {}

    # ------------------------------------------------------
    # PARCOURS UNIQUE DU FICHIER
    # ------------------------------------------------------

    for row in ws.iter_rows(
        min_row=1,
        values_only=True,
    ):

        # --------------------------------------------------
        # Vérification des indices
        # --------------------------------------------------

        id_index = category_id_col - 1
        search_index = category_search_col - 1

        if (
            id_index >= len(row)
            or search_index >= len(row)
        ):
            continue

        category_id = row[id_index]
        category_text = row[search_index]

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

        normalized_text = normalize(
            category_text
        )

        # --------------------------------------------------
        # Recherche exacte du texte complet
        # --------------------------------------------------

        if normalized_text not in mapping:

            mapping[
                normalized_text
            ] = category_id

        # --------------------------------------------------
        # Recherche du premier élément avant "-"
        #
        # Exemple :
        #
        # "20 - METROLOGIE"
        #
        # crée aussi :
        #
        # "20" -> ID
        # --------------------------------------------------

        if "-" in category_text:

            first_part = (
                category_text
                .split("-", 1)[0]
                .strip()
            )

            if first_part:

                normalized_first_part = normalize(
                    first_part
                )

                if (
                    normalized_first_part
                    not in mapping
                ):

                    mapping[
                        normalized_first_part
                    ] = category_id

    wb.close()

    return mapping


# ==========================================================
# RECHERCHE CATEGORIE EN MEMOIRE
# ==========================================================

def find_category_id_fast(
    groupe_produit,
    category_mapping,
):
    """
    Recherche une catégorie dans le dictionnaire
    chargé en mémoire.

    Aucun accès Excel ici.

    Exemple :

        Groupe Produit :
            "20 - Métrologie"

        Recherche :
            "20 - METROLOGIE"
            puis
            "20"

    """

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

    # ------------------------------------------------------
    # 1. Recherche exacte
    # ------------------------------------------------------

    if normalized_group in category_mapping:

        return category_mapping[
            normalized_group
        ]

    # ------------------------------------------------------
    # 2. Recherche de la partie avant "-"
    # ------------------------------------------------------

    if "-" in group:

        first_part = (
            group
            .split("-", 1)[0]
            .strip()
        )

        if first_part:

            normalized_first = normalize(
                first_part
            )

            if normalized_first in category_mapping:

                return category_mapping[
                    normalized_first
                ]

    # ------------------------------------------------------
    # 3. Dernière sécurité :
    # recherche du terme dans les clés.
    #
    # Cette partie est beaucoup plus rare.
    # ------------------------------------------------------

    for key, category_id in category_mapping.items():

        if (
            normalized_group
            and normalized_group in key
        ):

            return category_id

    return None


# ==========================================================
# INDEX DES REFERENCES ODOO
# ==========================================================

def build_odoo_reference_index(
    ws,
    reference_column,
    header_row,
):
    """
    Parcourt Odoo UNE SEULE FOIS et crée :

        référence -> numéro de ligne

    """

    reference_index = {}

    max_row = ws.max_row

    for row in range(
        header_row + 1,
        max_row + 1,
    ):

        value = ws.cell(
            row=row,
            column=reference_column,
        ).value

        ref = clean_reference(
            value
        )

        if not ref:
            continue

        # Première occurrence conservée

        if ref not in reference_index:

            reference_index[
                ref
            ] = row

    return reference_index


# ==========================================================
# COPIE DU STYLE D'UNE LIGNE
# ==========================================================

def copy_odoo_row_style(
    ws,
    source_row,
    target_row,
    max_col,
):
    """
    Copie uniquement le style de la ligne modèle.

    Les valeurs ne sont PAS copiées.
    """

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

    # Hauteur

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
# ECRITURE DES DONNEES ODOO
# ==========================================================

def write_odoo_row(
    ws,
    row,
    columns,
    reference,
    status,
    price,
    category_id,
):
    """
    Ecrit toutes les valeurs d'une ligne Odoo.
    """

    # ------------------------------------------------------
    # Référence
    # ------------------------------------------------------

    ws.cell(
        row=row,
        column=columns["reference"],
    ).value = reference

    # ------------------------------------------------------
    # Sales Status
    # ------------------------------------------------------

    ws.cell(
        row=row,
        column=columns["sales_status"],
    ).value = status

    # ------------------------------------------------------
    # Prix
    # ------------------------------------------------------

    ws.cell(
        row=row,
        column=columns["price"],
    ).value = price

    # ------------------------------------------------------
    # Code-barres
    # ------------------------------------------------------

    ws.cell(
        row=row,
        column=columns["barcode"],
    ).value = f"I {reference}"

    # ------------------------------------------------------
    # Fournisseur
    # ------------------------------------------------------

    ws.cell(
        row=row,
        column=columns["supplier"],
    ).value = "HEIDENHAIN FRANCE"

    # ------------------------------------------------------
    # Peut être acheté
    # ------------------------------------------------------

    ws.cell(
        row=row,
        column=columns["purchase"],
    ).value = "VRAI"

    # ------------------------------------------------------
    # Peut être vendu
    # ------------------------------------------------------

    ws.cell(
        row=row,
        column=columns["sale"],
    ).value = "VRAI"

    # ------------------------------------------------------
    # Type de produit
    # ------------------------------------------------------

    ws.cell(
        row=row,
        column=columns["type"],
    ).value = "Consommable"

    # ------------------------------------------------------
    # Politique de facturation
    # ------------------------------------------------------

    ws.cell(
        row=row,
        column=columns["invoice_policy"],
    ).value = "Quantités livrées"

    # ------------------------------------------------------
    # Catégorie
    # ------------------------------------------------------

    if category_id is not None:

        ws.cell(
            row=row,
            column=columns["category"],
        ).value = category_id


# ==========================================================
# TRAITEMENT ODOO OPTIMISE
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
):

    # ======================================================
    # 1. CHARGEMENT DU FICHIER CATEGORIE
    # ======================================================

    progress = st.progress(
        0,
        text="Chargement des catégories...",
    )

    category_mapping = load_category_mapping(
        category_file=category_file,
        category_sheet=category_sheet,
        category_id_col=int(category_id_col),
        category_search_col=int(category_search_col),
    )

    progress.progress(
        10,
        text=(
            f"Catégories chargées : "
            f"{len(category_mapping):,}"
        ),
    )

    # ======================================================
    # 2. OUVERTURE FICHIER HEIDENHAIN
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
            f"'{heidenhain_sheet}' introuvable.\n\n"
            f"Feuilles disponibles : {names}"
        )

    h_ws = wb_h[
        heidenhain_sheet
    ]

    # ======================================================
    # 3. COLONNES HEIDENHAIN
    # ======================================================

    heidenhain_columns = [
        heidenhain_id_column,
        heidenhain_status_column,
        heidenhain_group_column,
        heidenhain_ppc_column,
        heidenhain_sav_column,
    ]

    h_columns = find_columns(
        h_ws,
        int(heidenhain_header_row),
        heidenhain_columns,
    )

    missing_h = [
        column
        for column in heidenhain_columns
        if column not in h_columns
    ]

    if missing_h:

        wb_h.close()

        raise ValueError(
            "Colonnes introuvables dans le fichier "
            "Heidenhain : "
            + ", ".join(missing_h)
        )

    # ======================================================
    # 4. OUVERTURE ODOO
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
            f"'{odoo_sheet}' introuvable.\n\n"
            f"Feuilles disponibles : {names}"
        )

    ws = wb_o[
        odoo_sheet
    ]

    # ======================================================
    # 5. RECHERCHE DES COLONNES ODOO
    # ======================================================

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
    ]

    odoo_found = find_columns(
        ws,
        int(odoo_header_row),
        odoo_column_names,
    )

    missing_odoo = [
        column
        for column in odoo_column_names
        if column not in odoo_found
    ]

    if missing_odoo:

        wb_o.close()
        wb_h.close()

        raise ValueError(
            "Colonnes introuvables dans Odoo : "
            + ", ".join(missing_odoo)
        )

    # ======================================================
    # 6. CREATION D'UN INDEX ODOO
    # ======================================================

    progress.progress(
        20,
        text="Indexation des références Odoo...",
    )

    reference_index = build_odoo_reference_index(
        ws=ws,
        reference_column=odoo_found[
            odoo_reference_column
        ],
        header_row=int(odoo_header_row),
    )

    progress.progress(
        30,
        text=(
            f"Références Odoo indexées : "
            f"{len(reference_index):,}"
        ),
    )

    # ======================================================
    # 7. PREPARATION DES COLONNES
    # ======================================================

    columns = {
        "reference": odoo_found[
            odoo_reference_column
        ],

        "sales_status": odoo_found[
            odoo_sales_status_column
        ],

        "price": odoo_found[
            odoo_price_column
        ],

        "barcode": odoo_found[
            odoo_barcode_column
        ],

        "supplier": odoo_found[
            odoo_supplier_column
        ],

        "purchase": odoo_found[
            odoo_purchase_column
        ],

        "sale": odoo_found[
            odoo_sale_column
        ],

        "type": odoo_found[
            odoo_type_column
        ],

        "invoice_policy": odoo_found[
            odoo_invoice_policy_column
        ],

        "category": odoo_found[
            odoo_category_column
        ],
    }

    # ======================================================
    # 8. PARAMETRES DE LECTURE HEIDENHAIN
    # ======================================================

    h_header_row = int(
        heidenhain_header_row
    )

    h_data_start = (
        h_header_row + 1
    )

    h_max_row = h_ws.max_row

    total_h_rows = max(
        0,
        h_max_row - h_data_start + 1,
    )

    # ======================================================
    # 9. STATISTIQUES
    # ======================================================

    references_processed = 0

    updated = 0

    created = 0

    category_found = 0

    category_missing = 0

    sav_count = 0

    normal_count = 0

    empty_reference_count = 0

    # ======================================================
    # 10. DERNIERE LIGNE ODOO
    # ======================================================

    next_row = max(
        ws.max_row + 1,
        int(odoo_header_row) + 1,
    )

    template_row = (
        next_row - 1
        if next_row > int(odoo_header_row) + 1
        else None
    )

    # ======================================================
    # 11. TRAITEMENT HEIDENHAIN
    # ======================================================

    progress.progress(
        30,
        text="Mise à jour du fichier Odoo...",
    )

    for index, h_row in enumerate(
        range(
            h_data_start,
            h_max_row + 1,
        ),
        start=1,
    ):

        # --------------------------------------------------
        # ID / REFERENCE
        # --------------------------------------------------

        reference_value = h_ws.cell(
            row=h_row,
            column=h_columns[
                heidenhain_id_column
            ],
        ).value

        if reference_value is None:

            empty_reference_count += 1

            continue

        reference = str(
            reference_value
        ).strip()

        if not reference:

            empty_reference_count += 1

            continue

        references_processed += 1

        normalized_reference = clean_reference(
            reference
        )

        # --------------------------------------------------
        # STATUT
        # --------------------------------------------------

        status = h_ws.cell(
            row=h_row,
            column=h_columns[
                heidenhain_status_column
            ],
        ).value

        # --------------------------------------------------
        # GROUPE PRODUIT
        # --------------------------------------------------

        groupe = h_ws.cell(
            row=h_row,
            column=h_columns[
                heidenhain_group_column
            ],
        ).value

        # --------------------------------------------------
        # PRIX PPC
        # --------------------------------------------------

        ppc = h_ws.cell(
            row=h_row,
            column=h_columns[
                heidenhain_ppc_column
            ],
        ).value

        # --------------------------------------------------
        # PRIX SAV
        # --------------------------------------------------

        sav = h_ws.cell(
            row=h_row,
            column=h_columns[
                heidenhain_sav_column
            ],
        ).value

        # --------------------------------------------------
        # DETERMINATION DU PRIX
        # --------------------------------------------------

        if is_sav(reference):

            price = sav

            sav_count += 1

        else:

            price = ppc

            normal_count += 1

        # --------------------------------------------------
        # RECHERCHE CATEGORIE
        #
        # ICI AUCUN ACCES AU FICHIER EXCEL.
        # --------------------------------------------------

        category_id = find_category_id_fast(
            groupe_produit=groupe,
            category_mapping=category_mapping,
        )

        if category_id is not None:

            category_found += 1

        else:

            category_missing += 1

        # --------------------------------------------------
        # RECHERCHE ODOO
        # --------------------------------------------------

        if normalized_reference in reference_index:

            target_row = reference_index[
                normalized_reference
            ]

            updated += 1

        else:

            target_row = next_row

            # ----------------------------------------------
            # Copie du style de la ligne précédente
            # ----------------------------------------------

            if template_row is not None:

                copy_odoo_row_style(
                    ws=ws,
                    source_row=template_row,
                    target_row=target_row,
                    max_col=ws.max_column,
                )

            # ----------------------------------------------
            # Ajout dans l'index
            # ----------------------------------------------

            reference_index[
                normalized_reference
            ] = target_row

            # ----------------------------------------------
            # Ligne suivante
            # ----------------------------------------------

            next_row += 1

            template_row = target_row

            created += 1

        # --------------------------------------------------
        # ECRITURE
        # --------------------------------------------------

        write_odoo_row(
            ws=ws,
            row=target_row,
            columns=columns,
            reference=reference,
            status=status,
            price=price,
            category_id=category_id,
        )

        # --------------------------------------------------
        # PROGRESSION
        # --------------------------------------------------

        if (
            index % 1000 == 0
            or index == total_h_rows
        ):

            pct = 30 + int(
                (
                    index
                    / max(
                        1,
                        total_h_rows,
                    )
                )
                * 60
            )

            progress.progress(
                min(pct, 90),
                text=(
                    f"Odoo : "
                    f"{index:,} / "
                    f"{total_h_rows:,} "
                    f"| Mise à jour : "
                    f"{updated:,} "
                    f"| Création : "
                    f"{created:,}"
                ),
            )

    # ======================================================
    # 12. SAUVEGARDE
    # ======================================================

    progress.progress(
        95,
        text="Sauvegarde du fichier Odoo...",
    )

    output = BytesIO()

    wb_o.save(
        output
    )

    output.seek(0)

    result = output.getvalue()

    # ======================================================
    # 13. FERMETURE
    # ======================================================

    wb_o.close()
    wb_h.close()

    progress.progress(
        100,
        text="✅ Préparation Odoo terminée",
    )

    # ======================================================
    # 14. STATISTIQUES
    # ======================================================

    stats = {

        "processed":
            references_processed,

        "updated":
            updated,

        "created":
            created,

        "category_found":
            category_found,

        "category_missing":
            category_missing,

        "sav_count":
            sav_count,

        "normal_count":
            normal_count,

        "empty_reference":
            empty_reference_count,

        "category_index_size":
            len(category_mapping),

        "odoo_index_size":
            len(reference_index),
    }

    return result, stats



# ==========================================================
# SIDEBAR
# ==========================================================

st.sidebar.header("⚙️ Paramètres")


# ==========================================================
# HEIDENHAIN
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

status_column = st.sidebar.text_input(
    "Colonne Statut",
    "Statut",
)

id_column = st.sidebar.text_input(
    "Colonne ID",
    "ID",
)

group_column = st.sidebar.text_input(
    "Colonne Groupe Produit",
    "Groupe Produit",
)

ppc_column = st.sidebar.text_input(
    "Colonne Prix PPC",
    "Prix (PPC)",
)

sav_column = st.sidebar.text_input(
    "Colonne Prix SAV",
    "Prix (SAV)",
)


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
    if st.button(
            "🚀 Créer le fichier Heidenhain",
            type="primary",
            use_container_width=True,
            key="create_heidenhain",
        ):
    
            try:
    
                result, stats = process_heidenhain(
                    uploaded_file=heidenhain_file,
                    sheet_name=heidenhain_sheet,
                    header_row=heidenhain_header_row,
                    data_start_row=heidenhain_header_row + 1,
                    status_column=status_column,
                    id_column=id_column,
                    output_columns=DEFAULT_OUTPUT_COLUMNS,
                )
    
                st.session_state.heidenhain_result = result
                st.session_state.heidenhain_stats = stats
    
                st.success(
                    "✅ Fichier Heidenhain créé."
                )
    
            except Exception as e:
    
                st.error(
                    f"❌ Erreur : {e}"
                )
    
                st.exception(e)


# ==========================================================
# RESULTAT ETAPE 1
# ==========================================================

if st.session_state.heidenhain_result:

    stats = st.session_state.heidenhain_stats

    st.success(
        f"Fichier créé : "
        f"{stats['created']:,} nouvelles lignes SAV."
    )

    st.download_button(
        "⬇️ Télécharger le fichier Heidenhain",
        data=st.session_state.heidenhain_result,
        file_name="Prix_Heidenhain_prepare.xlsx",
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
    heidenhain_file is None
    or odoo_file is None
    or category_file is None
):

    st.info(
        "Importe les 3 fichiers pour lancer "
        "la préparation Odoo."
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

    if st.button(
        "🚀 Préparer le fichier Odoo",
        type="primary",
        use_container_width=True,
        key="prepare_odoo",
    ):

        try:

            result, stats = process_odoo(
                heidenhain_file=heidenhain_file,
                odoo_file=odoo_file,
                category_file=category_file,

                heidenhain_sheet=heidenhain_sheet,

                odoo_sheet=odoo_sheet,

                category_sheet=category_sheet,

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

                heidenhain_id_column=id_column,

                heidenhain_status_column=status_column,

                heidenhain_group_column=group_column,

                heidenhain_ppc_column=ppc_column,

                heidenhain_sav_column=sav_column,

                odoo_reference_column=odoo_reference,

                odoo_sales_status_column=odoo_sales_status,

                odoo_price_column=odoo_price,

                odoo_barcode_column=odoo_barcode,

                odoo_supplier_column=odoo_supplier,

                odoo_purchase_column=odoo_purchase,

                odoo_sale_column=odoo_sale,

                odoo_type_column=odoo_type,

                odoo_invoice_policy_column=odoo_invoice_policy,

                odoo_category_column=odoo_category,
            )

            st.session_state.odoo_result = result
            st.session_state.odoo_stats = stats

            st.success(
                "✅ Fichier Odoo préparé avec succès."
            )

        except Exception as e:

            st.error(
                "❌ Erreur pendant la préparation Odoo."
            )

            st.exception(e)


# ==========================================================
# RESULTAT ODOO
# ==========================================================

if st.session_state.odoo_result:

    stats = st.session_state.odoo_stats

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
            "référence(s) n'ont pas trouvé de catégorie."
        )

    st.download_button(
        label="⬇️ Télécharger le fichier Odoo prêt à importer",
        data=st.session_state.odoo_result,
        file_name="Import_Odoo_Heidenhain_prepare.xlsx",
        mime=(
            "application/vnd.openxmlformats-"
            "officedocument.spreadsheetml.sheet"
        ),
        use_container_width=True,
        key="download_odoo",
    )


# ==========================================================
# FIN
# ==========================================================

st.divider()

st.success(
    "🎯 Le fichier Odoo final peut maintenant être téléchargé."
)
