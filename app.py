# ==========================================================
# PREPARATION PRIX HEIDENHAIN -> ODOO
# JMA
#
# - Detection automatique des colonnes
# - Colonnes Heidenhain sélectionnables
# - Colonnes par défaut définies automatiquement
# - Création des lignes SAV pour VG / PG
# - Conservation des prix PPC / SAV / HA
# - Correspondances Odoo automatiques
# - Colonnes Odoo existantes conservées dans leur ordre
# - Ajout des colonnes manquantes à l'emplacement choisi
# - Transfert des colonnes supplémentaires de l'étape 1
# - Conservation des ID Odoo existants
# - Index des références pour accélérer le traitement
# ==========================================================

import time
from io import BytesIO
from copy import copy

import openpyxl
import streamlit as st


# ==========================================================
# CONFIGURATION
# ==========================================================

st.set_page_config(
    page_title="Préparation Heidenhain vers Odoo",
    page_icon="📦",
    layout="wide",
)

st.title("📦 Préparation des prix HEIDENHAIN → Odoo")

