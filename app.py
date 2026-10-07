# ==========================================================
# TRANSFERT DES PRIX HEIDENHAIN VERS ODOO
# JMA
# ==========================================================
#
# Fichier unique : app.py
#
# FICHIER 1 :
#   Prix Heidenhain
#
# FICHIER 2 :
#   Import Odoo
#
# TRAITEMENT ACTUEL :
#
#   1. Charger le fichier Heidenhain
#   2. Charger le fichier Odoo
#   3. Dans Heidenhain :
#        - feuille configurable
#        - en-têtes ligne configurable
#        - données ligne configurable
#        - chercher Statut = VG / PG
#        - récupérer ID
#        - créer ID_SAv
#        - vérifier si ID_SAv existe
#        - copier la ligne complète si nécessaire
#   4. Télécharger le fichier Heidenhain modifié
#
# OPTIMISATION :
#
#   - Pas de DataFrame pour traiter le fichier Heidenhain
#   - Recherche des IDs avec un set Python
#   - Une seule lecture du classeur pour le traitement
#   - Pas de recherche répétée des IDs
#   - Aperçus limités
#   - Barre de progression
#
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
    page_title="Transfert des prix HEIDENHAIN vers ODOO",
    page_icon="📦",
    layout="wide",
)

st.title("📦 Transfert des prix HEIDENHAIN vers ODOO")


# ==========================================================
# CONSTANTES
# ==========================================================

STATUS_TO_DUPLICATE = {"VG", "PG"}

PREVIEW_ROWS = 10


# ==========================================================
# PARAMETRES
# ==========================================================

st.sidebar.header("⚙️ Paramètres")


# ----------------------------------------------------------
# HEIDENHAIN
# ----------------------------------------------------------

st.sidebar.subheader("📘 Fichier Prix Heidenhain")

heidenhain_sheet = st.sidebar.text_input(
    "Nom de la feuille",
    value="Distributeurs",
    help=(
        "Nom de la feuille du fichier Heidenhain "
        "à traiter."
    ),
)

heidenhain_header_row = st.sidebar.number_input(
    "Ligne des noms de colonnes",
    min_value=1,
    value=4,
    step=1,
    help=(
        "Les noms des colonnes sont sur cette ligne."
    ),
)

heidenhain_data_start_row = st.sidebar.number_input(
    "Première ligne de données",
    min_value=1,
    value=5,
    step=1,
    help=(
        "Les données commencent à cette ligne."
    ),
)

status_column = st.sidebar.text_input(
    "Colonne Statut",
    value="Statut",
    help=(
        "Nom exact de la colonne contenant "
        "les statuts VG / PG."
    ),
)

id_column = st.sidebar.text_input(
    "Colonne ID",
    value="ID",
    help=(
        "Nom exact de la colonne contenant "
        "les identifiants."
    ),
)


# ----------------------------------------------------------
# ODOO
# ----------------------------------------------------------

st.sidebar.subheader("📗 Fichier Import Odoo")

odoo_sheet = st.sidebar.text_input(
    "Feuille Odoo",
    value="Sheet1",
)


# ==========================================================
# FONCTIONS UTILITAIRES
# ==========================================================

def normalize_value(value):
    """
    Normalise une valeur pour les comparaisons.

    Exemple :

        " VG " -> "VG"
        "vg"   -> "VG"
    """

    if value is None:
        return ""

    return str(value).strip().upper()


def get_excel_sheet_names(uploaded_file):
    """
    Retourne les noms des feuilles Excel.

    Lecture légère en read_only.
    """

    uploaded_file.seek(0)

    workbook = openpyxl.load_workbook(
        uploaded_file,
        read_only=True,
        data_only=False,
    )

    sheet_names = workbook.sheetnames

    workbook.close()

    return sheet_names


def find_column_by_header(
    worksheet,
    header_row,
    column_name,
):
    """
    Cherche le numéro de colonne correspondant
    au nom fourni.

    Exemple :

        ID -> 3
        Statut -> 8
    """

    target = normalize_value(
        column_name
    )

    for cell in worksheet[header_row]:

        if normalize_value(
            cell.value
        ) == target:

            return cell.column

    return None


def copy_cell(
    source_cell,
    target_cell,
):
    """
    Copie une cellule et sa mise en forme.

    Cette fonction n'est appelée que pour les lignes
    qui doivent réellement être créées.
    """

    target_cell.value = source_cell.value

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


def copy_row_fast(
    worksheet,
    source_row,
    target_row,
    max_column,
):
    """
    Copie une ligne complète.

    On utilise max_column fourni afin de ne pas
    recalculer worksheet.max_column à chaque appel.
    """

    # ------------------------------------------------------
    # Hauteur de ligne
    # ------------------------------------------------------

    source_dimension = worksheet.row_dimensions[
        source_row
    ]

    target_dimension = worksheet.row_dimensions[
        target_row
    ]

    target_dimension.height = (
        source_dimension.height
    )

    target_dimension.hidden = (
        source_dimension.hidden
    )

    target_dimension.outlineLevel = (
        source_dimension.outlineLevel
    )

    target_dimension.collapsed = (
        source_dimension.collapsed
    )

    # ------------------------------------------------------
    # Cellules
    # ------------------------------------------------------

    for column in range(
        1,
        max_column + 1,
    ):

        source_cell = worksheet.cell(
            row=source_row,
            column=column,
        )

        target_cell = worksheet.cell(
            row=target_row,
            column=column,
        )

        copy_cell(
            source_cell,
            target_cell,
        )


def load_odoo_preview(
    uploaded_file,
    sheet_name,
):
    """
    Charge uniquement un aperçu du fichier Odoo.

    Le fichier Odoo ne sera pas encore modifié.
    """

    uploaded_file.seek(0)

    extension = (
        uploaded_file.name
        .lower()
        .split(".")[-1]
    )

    if extension == "csv":

        df = pd.read_csv(
            uploaded_file,
            header=0,
            sep=None,
            engine="python",
            nrows=PREVIEW_ROWS,
        )

        return df

    workbook = openpyxl.load_workbook(
        uploaded_file,
        read_only=True,
        data_only=False,
    )

    if sheet_name not in workbook.sheetnames:

        workbook.close()

        raise ValueError(
            f"La feuille '{sheet_name}' "
            f"n'existe pas dans le fichier Odoo."
        )

    workbook.close()

    uploaded_file.seek(0)

    df = pd.read_excel(
        uploaded_file,
        sheet_name=sheet_name,
        header=0,
        nrows=PREVIEW_ROWS,
    )

    return df


# ==========================================================
# TRAITEMENT HEIDENHAIN OPTIMISE
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
    Traitement optimisé du fichier Heidenhain.

    Règle :

        Statut = VG ou PG

    alors :

        ID = ABC123

        nouveau ID =
        ABC123_SAv

    Si ABC123_SAv existe déjà :
        aucune copie.

    Sinon :
        copie complète de la ligne
        et remplacement de l'ID.

    ------------------------------------------------------

    Optimisation :

    1. Les IDs existants sont chargés dans un set.
    2. Les lignes VG/PG sont détectées en une passe.
    3. Les doublons sont vérifiés en mémoire.
    4. Le fichier Excel n'est modifié qu'après
       avoir identifié les lignes à créer.
    """

    # ======================================================
    # CHARGEMENT
    # ======================================================

    uploaded_file.seek(0)

    workbook = openpyxl.load_workbook(
        uploaded_file,
        data_only=False,
    )

    # ======================================================
    # FEUILLE
    # ======================================================

    if sheet_name not in workbook.sheetnames:

        workbook.close()

        raise ValueError(
            f"La feuille '{sheet_name}' "
            f"n'existe pas dans le fichier.\n\n"
            f"Feuilles disponibles : "
            f"{', '.join(workbook.sheetnames)}"
        )

    worksheet = workbook[
        sheet_name
    ]

    # ======================================================
    # PARAMETRES
    # ======================================================

    header_row = int(
        header_row
    )

    data_start_row = int(
        data_start_row
    )

    if data_start_row <= header_row:

        workbook.close()

        raise ValueError(
            "La première ligne de données doit "
            "être supérieure à la ligne d'en-tête."
        )

    # ======================================================
    # COLONNES
    # ======================================================

    status_col = find_column_by_header(
        worksheet,
        header_row,
        status_column,
    )

    if status_col is None:

        workbook.close()

        raise ValueError(
            f"La colonne '{status_column}' "
            f"n'a pas été trouvée sur la ligne "
            f"{header_row}."
        )

    id_col = find_column_by_header(
        worksheet,
        header_row,
        id_column,
    )

    if id_col is None:

        workbook.close()

        raise ValueError(
            f"La colonne '{id_column}' "
            f"n'a pas été trouvée sur la ligne "
            f"{header_row}."
        )

    # ======================================================
    # DIMENSIONS
    # ======================================================

    max_row = worksheet.max_row
    max_column = worksheet.max_column

    # ======================================================
    # ETAPE 1
    # CONSTRUCTION DU SET DES IDS
    # ======================================================

    existing_ids = set()

    progress = st.progress(
        0,
        text="Analyse des IDs existants...",
    )

    total_rows = max(
        1,
        max_row - data_start_row + 1,
    )

    # On lit uniquement la colonne ID.
    # Cela évite de parcourir inutilement toutes
    # les cellules du fichier.

    for index, row in enumerate(
        range(
            data_start_row,
            max_row + 1,
        )
    ):

        value = worksheet.cell(
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

        # Mise à jour de progression
        if (
            index % 5000 == 0
            or index == total_rows - 1
        ):

            percent = int(
                ((index + 1) / total_rows)
                * 25
            )

            progress.progress(
                percent,
                text=(
                    "Analyse des IDs existants... "
                    f"{index + 1:,} / "
                    f"{total_rows:,}"
                ),
            )

    # ======================================================
    # ETAPE 2
    # RECHERCHE VG / PG
    # ======================================================

    rows_to_create = []

    vg_count = 0
    pg_count = 0

    progress.progress(
        25,
        text="Recherche des statuts VG / PG...",
    )

    for index, row in enumerate(
        range(
            data_start_row,
            max_row + 1,
        )
    ):

        status_value = worksheet.cell(
            row=row,
            column=status_col,
        ).value

        status = normalize_value(
            status_value
        )

        if status == "VG":

            vg_count += 1

        elif status == "PG":

            pg_count += 1

        else:

            continue

        original_id = worksheet.cell(
            row=row,
            column=id_col,
        ).value

        if original_id is None:
            continue

        original_id = str(
            original_id
        ).strip()

        if not original_id:
            continue

        new_id = (
            f"{original_id}_SAv"
        )

        normalized_new_id = normalize_value(
            new_id
        )

        # --------------------------------------------------
        # Vérification immédiate du doublon
        # --------------------------------------------------

        if normalized_new_id in existing_ids:

            continue

        # --------------------------------------------------
        # Réserver immédiatement le nouvel ID.
        #
        # Important si plusieurs lignes du fichier
        # pourraient générer le même ID.
        # --------------------------------------------------

        existing_ids.add(
            normalized_new_id
        )

        rows_to_create.append(
            (
                row,
                new_id,
            )
        )

        if (
            index % 5000 == 0
        ):

            percent = 25 + int(
                (
                    (index + 1)
                    / total_rows
                )
                * 25
            )

            progress.progress(
                min(percent, 50),
                text=(
                    "Recherche VG / PG... "
                    f"{index + 1:,} / "
                    f"{total_rows:,}"
                ),
            )

    # ======================================================
    # ETAPE 3
    # CREATION DES LIGNES
    # ======================================================

    rows_created = len(
        rows_to_create
    )

    progress.progress(
        50,
        text=(
            f"{rows_created:,} ligne(s) "
            "à créer..."
        ),
    )

    created_ids = []

    next_row = (
        max_row + 1
    )

    if rows_created > 0:

        for index, (
            source_row,
            new_id,
        ) in enumerate(
            rows_to_create
        ):

            # ----------------------------------------------
            # Copie de la ligne
            # ----------------------------------------------

            copy_row_fast(
                worksheet,
                source_row,
                next_row,
                max_column,
            )

            # ----------------------------------------------
            # Modification de l'ID
            # ----------------------------------------------

            worksheet.cell(
                row=next_row,
                column=id_col,
            ).value = new_id

            created_ids.append(
                new_id
            )

            next_row += 1

            # ----------------------------------------------
            # Progression
            # ----------------------------------------------

            if (
                index % 100 == 0
                or index == rows_created - 1
            ):

                percent = (
                    50
                    + int(
                        (
                            (index + 1)
                            / rows_created
                        )
                        * 45
                    )
                )

                progress.progress(
                    min(percent, 95),
                    text=(
                        "Création des nouvelles lignes... "
                        f"{index + 1:,} / "
                        f"{rows_created:,}"
                    ),
                )

    # ======================================================
    # ETAPE 4
    # SAUVEGARDE
    # ======================================================

    progress.progress(
        97,
        text="Sauvegarde du fichier Excel...",
    )

    output = BytesIO()

    workbook.save(
        output
    )

    output.seek(0)

    workbook.close()

    progress.progress(
        100,
        text="✅ Traitement terminé",
    )

    # ======================================================
    # STATISTIQUES
    # ======================================================

    # Nombre de VG/PG réellement rencontrés
    vg_pg_found = (
        vg_count + pg_count
    )

    # Nombre de lignes VG/PG qui n'ont pas été créées
    # parce que l'ID existait déjà ou que l'ID était vide.
    skipped = max(
        0,
        vg_pg_found - rows_created,
    )

    stats = {

        "vg_count": vg_count,

        "pg_count": pg_count,

        "vg_pg_found": vg_pg_found,

        "rows_created": rows_created,

        "ids_skipped": skipped,

        "created_ids": created_ids,

        "original_max_row": max_row,

        "new_max_row": (
            max_row + rows_created
        ),

        "status_column_number": status_col,

        "id_column_number": id_col,

        "max_column": max_column,
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

            if heidenhain_sheet in sheet_names:

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

        try:

            odoo_preview = (
                load_odoo_preview(
                    odoo_file,
                    odoo_sheet,
                )
            )

            st.success(
                f"✅ Feuille `{odoo_sheet}` trouvée."
            )

            st.write(
                "Aperçu des données :"
            )

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

    st.header(
        "🔎 2. Vérification du fichier Heidenhain"
    )

    try:

        heidenhain_file.seek(0)

        preview_df = pd.read_excel(
            heidenhain_file,
            sheet_name=heidenhain_sheet,
            header=(
                int(
                    heidenhain_header_row
                ) - 1
            ),
            nrows=PREVIEW_ROWS,
        )

        columns = list(
            preview_df.columns
        )

        col_a, col_b = st.columns(2)

        with col_a:

            if status_column in columns:

                st.success(
                    f"✅ Colonne "
                    f"`{status_column}` trouvée."
                )

            else:

                st.error(
                    f"❌ Colonne "
                    f"`{status_column}` absente."
                )

        with col_b:

            if id_column in columns:

                st.success(
                    f"✅ Colonne "
                    f"`{id_column}` trouvée."
                )

            else:

                st.error(
                    f"❌ Colonne "
                    f"`{id_column}` absente."
                )

        # --------------------------------------------------
        # Statistiques aperçu
        # --------------------------------------------------

        if (
            status_column in columns
            and id_column in columns
        ):

            status_values = (
                preview_df[
                    status_column
                ]
                .astype(str)
                .str.strip()
                .str.upper()
            )

            st.write(
                "### Aperçu des statuts"
            )

            c1, c2 = st.columns(2)

            with c1:

                st.metric(
                    "VG dans l'aperçu",
                    int(
                        (
                            status_values
                            == "VG"
                        ).sum()
                    ),
                )

            with c2:

                st.metric(
                    "PG dans l'aperçu",
                    int(
                        (
                            status_values
                            == "PG"
                        ).sum()
                    ),
                )

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

st.header(
    "⚙️ 3. Traitement"
)


if heidenhain_file is None:

    st.info(
        "👆 Importez le fichier Heidenhain "
        "pour commencer."
    )

else:

    st.markdown(
        """
        Le traitement va analyser le fichier Heidenhain
        et créer les lignes `_SAv` nécessaires.

        **Le fichier Odoo est uniquement chargé pour
        l'instant. Il ne sera pas modifié à cette étape.**
        """
    )

    if st.button(
        "🚀 Lancer le traitement Heidenhain",
        type="primary",
        use_container_width=True,
    ):

        try:

            heidenhain_file.seek(0)

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
                )
            )

            st.success(
                "✅ Traitement terminé."
            )

            # ==================================================
            # STATISTIQUES
            # ==================================================

            st.subheader(
                "📊 Résultat"
            )

            c1, c2, c3, c4 = (
                st.columns(4)
            )

            with c1:

                st.metric(
                    "VG",
                    stats["vg_count"],
                )

            with c2:

                st.metric(
                    "PG",
                    stats["pg_count"],
                )

            with c3:

                st.metric(
                    "Nouvelles lignes",
                    stats["rows_created"],
                )

            with c4:

                st.metric(
                    "Ignorées",
                    stats["ids_skipped"],
                )

            # ==================================================
            # IDS CREES
            # ==================================================

            if stats["created_ids"]:

                st.subheader(
                    "🆕 Nouvelles références"
                )

                # Pour éviter de faire exploser l'interface
                # si plusieurs milliers d'IDs sont créés.

                max_display = 500

                display_ids = (
                    stats["created_ids"][
                        :max_display
                    ]
                )

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

                if (
                    len(
                        stats["created_ids"]
                    )
                    > max_display
                ):

                    st.info(
                        f"{len(stats['created_ids']):,} "
                        "IDs ont été créés. "
                        f"Seuls les {max_display} "
                        "premiers sont affichés."
                    )

            else:

                st.warning(
                    "Aucune nouvelle ligne n'a été créée."
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
                    "Heidenhain mis à jour"
                ),
                data=result_bytes,
                file_name=(
                    "Prix_Heidenhain_mis_a_jour.xlsx"
                ),
                mime=(
                    "application/vnd.openxmlformats-officedocument"
                    ".spreadsheetml.sheet"
                ),
                use_container_width=True,
            )

        except Exception as e:

            st.error(
                "❌ Le traitement a rencontré "
                "une erreur."
            )

            st.exception(e)


# ==========================================================
# ETAPE SUIVANTE
# ==========================================================

st.divider()

st.header(
    "➡️ Prochaine étape"
)

st.info(
    """
    Le fichier Heidenhain est maintenant préparé.

    La prochaine étape consistera à utiliser le fichier
    Heidenhain et le fichier `Sheet1` d'Odoo pour faire
    le rapprochement des produits et modifier l'import Odoo.
    """
)
