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
# CONSTANTES
# ==========================================================

STATUS_TO_DUPLICATE = {"VG", "PG"}
PREVIEW_ROWS = 10

# Colonnes conservées dans le fichier résultant.
# L'ordre est volontairement fixe.
OUTPUT_COLUMNS = [
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

st.sidebar.subheader("📘 Fichier Prix Heidenhain")

heidenhain_sheet = st.sidebar.text_input(
    "Feuille Heidenhain",
    value="Distributeurs",
    help="Nom de la feuille à traiter dans le fichier Prix Heidenhain.",
)

heidenhain_header_row = st.sidebar.number_input(
    "Ligne des noms de colonnes",
    min_value=1,
    value=4,
    step=1,
    help="Les noms des colonnes sont sur cette ligne.",
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
    help="Nom exact de la colonne contenant les identifiants.",
)

st.sidebar.divider()

st.sidebar.subheader("📗 Fichier Import Odoo")

odoo_sheet = st.sidebar.text_input(
    "Feuille Odoo",
    value="Sheet1",
    help="Nom de la feuille du fichier d'import Odoo. "
         "Elle sera utilisée à l'étape suivante.",
)


# ==========================================================
# UTILITAIRES
# ==========================================================

def normalize_value(value):
    """Normalisation rapide pour les comparaisons."""
    if value is None:
        return ""

    return str(value).strip().upper()


def get_excel_sheet_names(uploaded_file):
    """Retourne les feuilles sans charger les données."""
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
    Recherche toutes les colonnes demandées en une seule passe
    sur la ligne d'en-tête.

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
            original_name = requested_normalized[normalized]

            # Première occurrence uniquement.
            if original_name not in found:
                found[original_name] = cell.column

    return found


def copy_cell_style(source_cell, target_cell):
    """
    Copie uniquement ce qui est utile pour conserver
    l'apparence de la cellule.
    """
    if source_cell.has_style:
        target_cell._style = copy(source_cell._style)

    if source_cell.number_format:
        target_cell.number_format = source_cell.number_format

    if source_cell.alignment:
        target_cell.alignment = copy(source_cell.alignment)

    if source_cell.protection:
        target_cell.protection = copy(source_cell.protection)

    if source_cell.comment:
        target_cell.comment = copy(source_cell.comment)

    if source_cell.hyperlink:
        target_cell._hyperlink = copy(source_cell.hyperlink)


def copy_selected_row(
    source_ws,
    target_ws,
    source_row,
    target_row,
    source_columns,
    output_columns,
    new_id=None,
):
    """
    Copie uniquement les colonnes nécessaires.

    source_columns :
        mapping nom -> numéro de colonne source.

    output_columns :
        ordre des colonnes de sortie.

    new_id :
        si renseigné, remplace l'ID de la ligne créée.
    """
    for output_col_index, column_name in enumerate(
        output_columns,
        start=1,
    ):
        source_col_index = source_columns[column_name]

        source_cell = source_ws.cell(
            row=source_row,
            column=source_col_index,
        )

        target_cell = target_ws.cell(
            row=target_row,
            column=output_col_index,
        )

        target_cell.value = source_cell.value
        copy_cell_style(source_cell, target_cell)

    # Remplacement de l'ID après copie.
    if new_id is not None:
        target_ws.cell(
            row=target_row,
            column=1,
        ).value = new_id


def load_odoo_preview(
    uploaded_file,
    sheet_name,
):
    """Aperçu léger du fichier Odoo."""
    uploaded_file.seek(0)

    extension = uploaded_file.name.lower().split(".")[-1]

    if extension == "csv":
        return pd.read_csv(
            uploaded_file,
            header=0,
            sep=None,
            engine="python",
            nrows=PREVIEW_ROWS,
        )

    workbook = openpyxl.load_workbook(
        uploaded_file,
        read_only=True,
        data_only=False,
    )

    if sheet_name not in workbook.sheetnames:
        workbook.close()
        raise ValueError(
            f"La feuille '{sheet_name}' n'existe pas dans le fichier Odoo."
        )

    workbook.close()
    uploaded_file.seek(0)

    return pd.read_excel(
        uploaded_file,
        sheet_name=sheet_name,
        header=0,
        nrows=PREVIEW_ROWS,
    )


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
):
    """
    Prépare le fichier Heidenhain.

    1. Lit la feuille source.
    2. Conserve uniquement OUTPUT_COLUMNS.
    3. Copie toutes les lignes existantes.
    4. Cherche les statuts VG / PG.
    5. Crée ID_SAV si absent.
    6. Ajoute les lignes SAV à la suite.

    Optimisation gros fichiers :
      - openpyxl uniquement ;
      - recherche des colonnes en une passe ;
      - set Python pour les IDs ;
      - lecture uniquement des colonnes nécessaires ;
      - pas de DataFrame pour le traitement ;
      - les lignes à créer sont mémorisées avant écriture ;
      - progression Streamlit.
    """

    header_row = int(header_row)
    data_start_row = int(data_start_row)

    if data_start_row <= header_row:
        raise ValueError(
            "La première ligne de données doit être supérieure "
            "à la ligne des en-têtes."
        )

    # ----------------------------------------------------------
    # Chargement
    # ----------------------------------------------------------

    uploaded_file.seek(0)

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

    # ----------------------------------------------------------
    # Recherche des colonnes
    # ----------------------------------------------------------

    required_columns = list(
        dict.fromkeys(
            OUTPUT_COLUMNS
            + [status_column, id_column]
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

    status_col = source_columns[status_column]
    id_col = source_columns[id_column]

    max_row = source_ws.max_row

    # ----------------------------------------------------------
    # Création du nouveau classeur
    # ----------------------------------------------------------

    output_workbook = openpyxl.Workbook()
    output_ws = output_workbook.active
    output_ws.title = sheet_name

    # ----------------------------------------------------------
    # En-têtes
    # ----------------------------------------------------------

    for output_col_index, column_name in enumerate(
        OUTPUT_COLUMNS,
        start=1,
    ):
        source_header_cell = source_ws.cell(
            row=header_row,
            column=source_columns[column_name],
        )

        target_header_cell = output_ws.cell(
            row=1,
            column=output_col_index,
        )

        target_header_cell.value = source_header_cell.value
        copy_cell_style(
            source_header_cell,
            target_header_cell,
        )

    # ----------------------------------------------------------
    # Dimensions
    # ----------------------------------------------------------

    total_rows = max(
        0,
        max_row - data_start_row + 1,
    )

    progress = st.progress(
        0,
        text="Préparation du fichier Heidenhain...",
    )

    # ==========================================================
    # ETAPE 1
    # COLLECTE DES IDS EXISTANTS
    # ==========================================================

    progress.progress(
        5,
        text="Analyse des IDs existants...",
    )

    existing_ids = set()

    for index, row in enumerate(
        range(data_start_row, max_row + 1)
    ):
        value = source_ws.cell(
            row=row,
            column=id_col,
        ).value

        normalized = normalize_value(value)

        if normalized:
            existing_ids.add(normalized)

        if (
            index % 5000 == 0
            or index == total_rows - 1
        ):
            percent = (
                5
                + int(
                    ((index + 1) / max(1, total_rows))
                    * 15
                )
            )

            progress.progress(
                min(percent, 20),
                text=(
                    "Analyse des IDs existants... "
                    f"{index + 1:,} / {total_rows:,}"
                ),
            )

    # ==========================================================
    # ETAPE 2
    # COPIE DES DONNEES + DETECTION VG / PG
    # ==========================================================

    rows_to_create = []

    vg_count = 0
    pg_count = 0
    empty_id_count = 0
    duplicate_count = 0

    progress.progress(
        20,
        text="Copie des données et recherche VG / PG...",
    )

    # Ligne 2 dans le nouveau fichier = première donnée.
    output_data_row = 2

    for index, source_row in enumerate(
        range(data_start_row, max_row + 1)
    ):
        # ------------------------------------------------------
        # Copie de la ligne existante.
        # ------------------------------------------------------

        copy_selected_row(
            source_ws=source_ws,
            target_ws=output_ws,
            source_row=source_row,
            target_row=output_data_row,
            source_columns=source_columns,
            output_columns=OUTPUT_COLUMNS,
        )

        # Hauteur de ligne.
        output_ws.row_dimensions[
            output_data_row
        ].height = source_ws.row_dimensions[
            source_row
        ].height

        # ------------------------------------------------------
        # Lecture du statut.
        # ------------------------------------------------------

        status = normalize_value(
            source_ws.cell(
                row=source_row,
                column=status_col,
            ).value
        )

        if status not in STATUS_TO_DUPLICATE:
            output_data_row += 1

            if (
                index % 5000 == 0
                or index == total_rows - 1
            ):
                percent = (
                    20
                    + int(
                        ((index + 1) / max(1, total_rows))
                        * 40
                    )
                )

                progress.progress(
                    min(percent, 60),
                    text=(
                        "Traitement des lignes... "
                        f"{index + 1:,} / {total_rows:,}"
                    ),
                )

            continue

        if status == "VG":
            vg_count += 1
        else:
            pg_count += 1

        # ------------------------------------------------------
        # ID source.
        # ------------------------------------------------------

        original_id = source_ws.cell(
            row=source_row,
            column=id_col,
        ).value

        if original_id is None:
            empty_id_count += 1
            output_data_row += 1
            continue

        original_id = str(original_id).strip()

        if not original_id:
            empty_id_count += 1
            output_data_row += 1
            continue

        new_id = f"{original_id}_SAV"
        normalized_new_id = normalize_value(new_id)

        # ------------------------------------------------------
        # Doublon.
        # ------------------------------------------------------

        if normalized_new_id in existing_ids:
            duplicate_count += 1
            output_data_row += 1
            continue

        # Réservation immédiate.
        existing_ids.add(normalized_new_id)

        rows_to_create.append(
            (
                source_row,
                new_id,
            )
        )

        output_data_row += 1

        if (
            index % 5000 == 0
            or index == total_rows - 1
        ):
            percent = (
                20
                + int(
                    ((index + 1) / max(1, total_rows))
                    * 40
                )
            )

            progress.progress(
                min(percent, 60),
                text=(
                    "Traitement des lignes... "
                    f"{index + 1:,} / {total_rows:,} "
                    f"| SAV à créer : {len(rows_to_create):,}"
                ),
            )

    # ==========================================================
    # ETAPE 3
    # CREATION DES LIGNES SAV
    # ==========================================================

    rows_created = len(rows_to_create)

    progress.progress(
        60,
        text=(
            f"Création de {rows_created:,} ligne(s) SAV..."
        ),
    )

    created_ids = []

    next_output_row = output_data_row

    if rows_created:
        for index, (source_row, new_id) in enumerate(
            rows_to_create
        ):
            copy_selected_row(
                source_ws=source_ws,
                target_ws=output_ws,
                source_row=source_row,
                target_row=next_output_row,
                source_columns=source_columns,
                output_columns=OUTPUT_COLUMNS,
                new_id=new_id,
            )

            output_ws.row_dimensions[
                next_output_row
            ].height = source_ws.row_dimensions[
                source_row
            ].height

            created_ids.append(new_id)

            next_output_row += 1

            if (
                index % 500 == 0
                or index == rows_created - 1
            ):
                percent = (
                    60
                    + int(
                        ((index + 1) / rows_created)
                        * 35
                    )
                )

                progress.progress(
                    min(percent, 95),
                    text=(
                        "Création des lignes SAV... "
                        f"{index + 1:,} / {rows_created:,}"
                    ),
                )

    # ----------------------------------------------------------
    # Ajustement des largeurs de colonnes.
    # On reprend les largeurs des colonnes source.
    # ----------------------------------------------------------

    for output_col_index, column_name in enumerate(
        OUTPUT_COLUMNS,
        start=1,
    ):
        source_col_index = source_columns[column_name]

        source_letter = openpyxl.utils.get_column_letter(
            source_col_index
        )

        target_letter = openpyxl.utils.get_column_letter(
            output_col_index
        )

        output_ws.column_dimensions[
            target_letter
        ].width = source_ws.column_dimensions[
            source_letter
        ].width

    # ----------------------------------------------------------
    # Figer la première ligne.
    # ----------------------------------------------------------

    output_ws.freeze_panes = "A2"

    # ==========================================================
    # ETAPE 4
    # SAUVEGARDE
    # ==========================================================

    progress.progress(
        97,
        text="Sauvegarde du fichier Excel...",
    )

    output = BytesIO()

    output_workbook.save(output)
    output.seek(0)

    output_workbook.close()
    workbook.close()

    progress.progress(
        100,
        text="✅ Traitement terminé",
    )

    # ==========================================================
    # STATISTIQUES
    # ==========================================================

    stats = {
        "total_source_rows": total_rows,
        "vg_count": vg_count,
        "pg_count": pg_count,
        "vg_pg_found": vg_count + pg_count,
        "rows_created": rows_created,
        "duplicates": duplicate_count,
        "empty_ids": empty_id_count,
        "created_ids": created_ids,
        "source_max_row": max_row,
        "output_data_rows": output_data_row - 1,
        "output_total_rows": next_output_row - 1,
        "source_columns": source_columns,
        "status_column_number": status_col,
        "id_column_number": id_col,
    }

    return output.getvalue(), stats


# ==========================================================
# INTERFACE
# ==========================================================

st.header("📂 1. Import des fichiers")

col1, col2 = st.columns(2)


# ==========================================================
# FICHIER HEIDENHAIN
# ==========================================================

with col1:
    st.subheader("📘 Fichier 1 — Prix Heidenhain")

    heidenhain_file = st.file_uploader(
        "Importer le fichier des prix Heidenhain",
        type=["xlsx", "xlsm"],
        key="heidenhain_upload",
    )

    if heidenhain_file is not None:
        st.success(f"✅ `{heidenhain_file.name}`")

        try:
            sheet_names = get_excel_sheet_names(
                heidenhain_file
            )

            if heidenhain_sheet in sheet_names:
                st.success(
                    f"✅ Feuille `{heidenhain_sheet}` trouvée."
                )
            else:
                st.error(
                    f"❌ La feuille `{heidenhain_sheet}` "
                    "n'existe pas."
                )

                st.write(
                    "**Feuilles disponibles :**"
                )
                st.write(sheet_names)

        except Exception as e:
            st.error(f"Erreur : {e}")


# ==========================================================
# FICHIER ODOO
# ==========================================================

with col2:
    st.subheader("📗 Fichier 2 — Import Odoo")

    odoo_file = st.file_uploader(
        "Importer le fichier d'import Odoo",
        type=["xlsx", "xlsm", "xls", "csv"],
        key="odoo_upload",
    )

    if odoo_file is not None:
        st.success(f"✅ `{odoo_file.name}`")

        try:
            odoo_preview = load_odoo_preview(
                odoo_file,
                odoo_sheet,
            )

            st.success(
                f"✅ Feuille `{odoo_sheet}` trouvée."
            )

            st.write("Aperçu des données :")

            st.dataframe(
                odoo_preview,
                use_container_width=True,
            )

        except Exception as e:
            st.error(
                f"Erreur de lecture Odoo : {e}"
            )


# ==========================================================
# APERCU HEIDENHAIN
# ==========================================================

if (
    heidenhain_file is not None
    and heidenhain_sheet
):
    st.divider()

    st.header("🔎 2. Vérification du fichier Heidenhain")

    try:
        heidenhain_file.seek(0)

        preview_df = pd.read_excel(
            heidenhain_file,
            sheet_name=heidenhain_sheet,
            header=int(heidenhain_header_row) - 1,
            nrows=PREVIEW_ROWS,
        )

        columns = list(preview_df.columns)

        c1, c2 = st.columns(2)

        with c1:
            if status_column in columns:
                st.success(
                    f"✅ Colonne `{status_column}` trouvée."
                )
            else:
                st.error(
                    f"❌ Colonne `{status_column}` absente."
                )

        with c2:
            if id_column in columns:
                st.success(
                    f"✅ Colonne `{id_column}` trouvée."
                )
            else:
                st.error(
                    f"❌ Colonne `{id_column}` absente."
                )

        missing_output_columns = [
            column
            for column in OUTPUT_COLUMNS
            if column not in columns
        ]

        if missing_output_columns:
            st.error(
                "❌ Colonnes nécessaires absentes : "
                + ", ".join(missing_output_columns)
            )
        else:
            st.success(
                "✅ Toutes les colonnes de sortie sont présentes."
            )

        if (
            status_column in columns
            and id_column in columns
        ):
            status_values = (
                preview_df[status_column]
                .astype(str)
                .str.strip()
                .str.upper()
            )

            st.write("### Aperçu des statuts")

            m1, m2 = st.columns(2)

            with m1:
                st.metric(
                    "VG dans l'aperçu",
                    int(
                        (status_values == "VG").sum()
                    ),
                )

            with m2:
                st.metric(
                    "PG dans l'aperçu",
                    int(
                        (status_values == "PG").sum()
                    ),
                )

        st.write("### Colonnes qui seront conservées")
        st.write(OUTPUT_COLUMNS)

        st.dataframe(
            preview_df,
            use_container_width=True,
        )

    except Exception as e:
        st.error(
            f"Erreur lors de l'aperçu : {e}"
        )


# ==========================================================
# TRAITEMENT
# ==========================================================

st.divider()
st.header("⚙️ 3. Traitement")


if heidenhain_file is None:
    st.info(
        "👆 Importez le fichier Heidenhain "
        "pour commencer."
    )

else:
    st.markdown(
        """
        Le traitement va :

        - conserver uniquement les colonnes demandées ;
        - conserver les lignes existantes ;
        - rechercher les statuts **VG** et **PG** ;
        - créer les références **`ID_SAV`** absentes ;
        - ajouter ces nouvelles lignes à la fin ;
        - ne pas modifier le fichier Odoo à cette étape.
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
                "✅ Fichier Heidenhain préparé avec succès."
            )

            # ==================================================
            # STATISTIQUES
            # ==================================================

            st.subheader("📊 Résultat")

            c1, c2, c3, c4, c5 = st.columns(5)

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
                    f"{stats['empty_ids']:,} ligne(s) VG/PG "
                    "ont été ignorées car l'ID était vide."
                )

            # ==================================================
            # IDS CREES
            # ==================================================

            if stats["created_ids"]:
                st.subheader("🆕 Nouvelles références")

                max_display = 500

                display_ids = stats["created_ids"][
                    :max_display
                ]

                ids_df = pd.DataFrame(
                    {
                        "Nouveau ID": display_ids
                    }
                )

                st.dataframe(
                    ids_df,
                    use_container_width=True,
                    height=400,
                )

                if len(stats["created_ids"]) > max_display:
                    st.info(
                        f"{len(stats['created_ids']):,} IDs "
                        f"créés. Seuls les {max_display} "
                        "premiers sont affichés."
                    )
            else:
                st.warning(
                    "Aucune nouvelle ligne SAV n'a été créée."
                )

            # ==================================================
            # TELECHARGEMENT
            # ==================================================

            st.subheader("⬇️ Fichier résultant")

            st.download_button(
                label="⬇️ Télécharger le fichier Heidenhain préparé",
                data=result_bytes,
                file_name="Prix_Heidenhain_prepare.xlsx",
                mime=(
                    "application/vnd.openxmlformats-officedocument."
                    "spreadsheetml.sheet"
                ),
                use_container_width=True,
            )

        except Exception as e:
            st.error(
                "❌ Le traitement a rencontré une erreur."
            )
            st.exception(e)


# ==========================================================
# ETAPE SUIVANTE
# ==========================================================

st.divider()

st.header("➡️ Prochaine étape")

st.info(
    """
    Le fichier Heidenhain est maintenant préparé.

    À l'étape suivante, nous utiliserons le fichier
    `Sheet1` d'Odoo et travaillerons avec les numéros
    de colonnes pour faire le rapprochement et modifier
    l'import Odoo.
    """
)
