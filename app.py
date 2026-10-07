# ==========================================================
# JMA
# V1 - 03/03/2026
# ==========================================================

import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import datetime
from datetime import time
from fpdf import FPDF
from io import BytesIO
from zoneinfo import ZoneInfo  # gestion fuseaux horaires standard
import matplotlib.dates as mdates
import json
import uuid




# ==========================================================
# SAUVEGARDE / CHARGEMENT DES PARAMÈTRES
# ==========================================================

PARAM_KEYS = [
    # Type de données
    "data_mode",
    "importExport_is_monthly",
    "unite",
    "export_is_monthly",

    # Mots-clés
    "MotCle",
    "custom_date_tokens",
    "custom_import_tokens",
    "custom_export_tokens",

    # Tarification
    "mode_tarif",
    "GRD_select",
    "tariff_importHP",
    "tariff_importHC",
    "tariff_export",
    "weekend_hc",
    "nb_plages",

    # Paramètres généraux
    "debug",
    "roundtrip_eff",
    "cap_min",
    "cap_max",
    "cap_step",
    "soc_min_pct",
    "p_min",
    "p_max",
    "p_step",
    "gain_threshold",
    "daily_percentile",

    # Capacité auto
    "capacite_auto",
    "facteur_auto",
    "max_auto_extensions",
    "cap_securite",
]

MONTHS = [
    "Janvier", "Février", "Mars", "Avril", "Mai", "Juin",
    "Juillet", "Août", "Septembre", "Octobre", "Novembre", "Décembre"
]


def get_simulation_params():
    """Construit un dictionnaire avec tous les paramètres à sauvegarder."""
    params = {}

    # Sauvegarde des clés simples
    for key in PARAM_KEYS:
        if key in st.session_state:
            params[key] = st.session_state[key]

    # Sauvegarde de l'état tarifaire
    if "tarif_state" in st.session_state:
        params["tarif_state"] = {
            "active_GRD": st.session_state.tarif_state.get("active_GRD"),
            "hp_ranges": st.session_state.tarif_state.get("hp_ranges", []),
            "nb_plages": st.session_state.tarif_state.get("nb_plages", 1),
            "weekend_hc": st.session_state.tarif_state.get("weekend_hc", False),
        }

    # Sauvegarde des widgets horaires HP
    hp_widget_values = {}
    for key in st.session_state.keys():
        if key.startswith("hp_start_") or key.startswith("hp_end_"):
            hp_widget_values[key] = st.session_state[key]

    if hp_widget_values:
        params["hp_widget_values"] = hp_widget_values

    # Sauvegarde des exports mensuels
    monthly_export_values = {}
    for month in MONTHS:
        if month in st.session_state:
            monthly_export_values[month] = st.session_state[month]

    if monthly_export_values:
        params["monthly_export_values"] = monthly_export_values

    return params


def make_json_serializable(obj):
    """Convertit les objets non JSON-compatibles."""
    if isinstance(obj, dict):
        return {k: make_json_serializable(v) for k, v in obj.items()}
    elif isinstance(obj, list):
        return [make_json_serializable(v) for v in obj]
    elif isinstance(obj, tuple):
        return [make_json_serializable(v) for v in obj]
    elif isinstance(obj, time):
        return obj.strftime("%H:%M:%S")
    else:
        return obj


def restore_simulation_params(params):
    """Recharge les paramètres dans st.session_state."""
    for key, value in params.items():
        if key == "tarif_state":
            if "hp_ranges" in value:
                value["hp_ranges"] = [tuple(x) for x in value["hp_ranges"]]
            st.session_state[key] = value

        elif key == "hp_widget_values":
            for subkey, subvalue in value.items():
                # ici normalement ce sont des int pour tes selectbox
                if isinstance(subvalue, str) and subvalue.count(":") == 2:
                    h, m, s = map(int, subvalue.split(":"))
                    st.session_state[subkey] = time(h, m, s)
                else:
                    st.session_state[subkey] = subvalue

        elif key == "monthly_export_values":
            for month, month_value in value.items():
                st.session_state[month] = month_value

        else:
            st.session_state[key] = value

# Fonction de gestion des dates
def remove_dst(dt_series):
    # Supprimer "DST"
    dt_series = dt_series.astype(str).str.replace(r'\sDST', '', regex=True)

    # Formats explicites (priorité Europe)
    formats = [
        "%d/%m/%Y %H:%M:%S",
        "%d/%m/%Y %H:%M",
        "%d/%m/%Y",

        "%d-%m-%Y %H:%M:%S",
        "%d-%m-%Y %H:%M",
        "%d-%m-%Y",

        "%d.%m.%Y %H:%M:%S",
        "%d.%m.%Y %H:%M",
        "%d.%m.%Y",

        "%Y-%m-%d %H:%M:%S",
        "%Y-%m-%d %H:%M",
        "%Y-%m-%d",
    ]

    for fmt in formats:
        try:
            result = pd.to_datetime(dt_series, format=fmt, errors='coerce')
            if result.notna().any():  # au moins une conversion réussie
                return result
        except:
            continue

    # Si aucun format ne fonctionne, laisse pandas deviner
    return pd.to_datetime(dt_series, errors='coerce')
    
# Date actuelle en GMT+1
now_gmt1 = datetime.datetime.now(ZoneInfo("Europe/Zurich"))

LongBase = 0

# ==========================================================
# PAGE CONFIG
# ==========================================================
st.set_page_config(page_title="Battery Sizer By JMN", layout="wide")
st.title("Battery Sizer - For Switzerland - By JMN")


    
# ==========================================================
# IMPORT / EXPORT CONFIG
# ==========================================================

st.sidebar.header("💾 Fichier de config simulation")

if "last_loaded_config_name" not in st.session_state:
    st.session_state.last_loaded_config_name = None

# ---------- IMPORT ----------
st.sidebar.markdown("**📂 Import**")

uploaded = st.sidebar.file_uploader(
    "Fichier JSON",
    type=["json"],
    label_visibility="collapsed",
    key="config_uploader"
)

if uploaded is not None:
    # Import seulement si nouveau fichier
    if st.session_state.last_loaded_config_name != uploaded.name:
        loaded_params = json.load(uploaded)
        restore_simulation_params(loaded_params)
        st.session_state.last_loaded_config_name = uploaded.name
        st.sidebar.success("✅ Configuration chargée")
        st.rerun()
else:
    st.session_state.last_loaded_config_name = None


# ---------- EXPORT ----------
st.sidebar.markdown("**⬇️ Export**")

params = get_simulation_params()

json_str = json.dumps(
    make_json_serializable(params),
    indent=2,
    ensure_ascii=False
)

st.sidebar.download_button(
    label="Télécharger fichier de config",
    data=json_str,
    file_name="config_simulation.json",
    mime="application/json",
    use_container_width=True
)

# ==========================================================
# SIDEBAR PARAMETERS
# ==========================================================
dt_hours = 0

st.sidebar.header("📂 Type de données")
data_mode = st.sidebar.radio(
    "Format des données",
    ["Fichier GRD (Excel/CSV unique)", "Fichiers mensuels (12 fichiers cumulés HUAWEI)"],
    key="data_mode"
)

if data_mode == "Fichier GRD (Excel/CSV unique)":
    st.sidebar.subheader("⚙️ Options :")
    importExport_is_monthly = st.sidebar.checkbox(
        "Import / Export fourni en total mensuel Excel (kWh/mois)",
        value=False, key="importExport_is_monthly"
    )
    if not importExport_is_monthly :
        unite = st.sidebar.selectbox(
        "Unité des valeurs Import / Export",
        ["kW", "kWh", "Wh"],
        index=1, key="unite")
        
        export_is_monthly = st.sidebar.checkbox(
            "Export fourni en total mensuel (kWh/mois)",
            value=False, key="export_is_monthly"
        )
        monthly_export_values = None            
        if export_is_monthly:
            st.sidebar.markdown("#### Saisir les 12 valeurs d'export (kWh)")
            months = [
                "Janvier", "Février", "Mars", "Avril", "Mai", "Juin",
                "Juillet", "Août", "Septembre", "Octobre", "Novembre", "Décembre"
            ]
            monthly_export_values = {}
            for month in months:
                monthly_export_values[month] = st.sidebar.number_input(
                    f"{month} (kWh)",
                    min_value=0.0,
                    value=0.0,
                    step=10.0, 
                    key=month
                )
            # Vérification qu'il y a bien 12 valeurs
            if len(monthly_export_values) != 12:
                st.error("❌ Les 12 mois doivent être renseignés.")
                st.stop()
            # Vérification qu'aucune valeur n'est vide ou nulle
            missing_months = [
                month for month, value in monthly_export_values.items()
                if value is None or value <= 0
            ]
            if missing_months:
                st.error(f"❌ Valeur manquante ou nulle pour : {', '.join(missing_months)}")
                st.stop()
    else:
        unite="kWh"
        export_is_monthly = False
else : 
    unite = st.sidebar.selectbox(
    "Unité des valeurs Import / Export",
    ["kW", "kWh", "Wh"],
    index=1, key="unite")
    
    importExport_is_monthly = False
    export_is_monthly = False
            
if unite == "kW":
    values_are_kw = True
else:
    values_are_kw = False

st.sidebar.subheader("🔎 Mots-clés personnalisés (optionnel)")
MotCle = st.sidebar.checkbox("Avec / Sans", value=False, key="MotCle")
if MotCle:
    custom_date_tokens = st.sidebar.text_input(
        "Mots-clés Date (séparés par virgule)",
        value="",
        key="custom_date_tokens"
    )
    
    custom_import_tokens = st.sidebar.text_input(
        "Mots-clés Import (séparés par virgule)",
        value="",
        key="custom_import_tokens"
    )
    
    custom_export_tokens = st.sidebar.text_input(
        "Mots-clés Export (séparés par virgule)",
        value="",
        key="custom_export_tokens"
    )

st.sidebar.header("💰 Paramètres des Tarifs GRD")
mode_tarif = st.sidebar.selectbox(
    "Type de tarification",
    ["Tarif unique", "HP/HC"],
    key="mode_tarif"
)

weekend_hc = False

# ==============================
# MODE HP / HC
# ==============================
if mode_tarif == "HP/HC":
    # -----------------------
    # Sélection GRD
    # -----------------------
    GRD = st.sidebar.selectbox(
        "Sélection du GRD",
        [
            "Groupe E (nouveau)",
            "Romande - Bas-Valais Energie SA",
            "Romande - Pully / Belmont",
            "Manuel"
        ],
        key="GRD_select"
    )

    # -----------------------
    # Définition defaults
    # -----------------------
    if GRD == "Groupe E (nouveau)":
        default_hp = [(7, 12), (17, 23)]
        default_weekend = False

    elif GRD == "Romande - Bas-Valais Energie SA":
        default_hp = [(17, 22)]
        default_weekend = True

    elif GRD == "Romande - Pully / Belmont":
        default_hp = [(6, 22)]
        default_weekend = False

    else:  # Manuel
        default_hp = [(6, 22)]
        default_weekend = False

    # ==============================
    # INITIALISATION & RESET PROPRE
    # ==============================

    if "tarif_state" not in st.session_state:
        st.session_state.tarif_state = {
            "active_GRD": GRD,
            "hp_ranges": default_hp.copy(),
            "nb_plages": len(default_hp),
            "weekend_hc": default_weekend
        }

    state = st.session_state.tarif_state

    # Reset si GRD change
    if GRD != state["active_GRD"]:
        state["active_GRD"] = GRD
        state["hp_ranges"] = default_hp.copy()
        state["nb_plages"] = len(default_hp)
        state["weekend_hc"] = default_weekend

        # Synchroniser les widgets
        st.session_state["nb_plages"] = state["nb_plages"]
        st.session_state["weekend_hc"] = state["weekend_hc"]
    
        # Suppression anciennes clés horaires
        for key in list(st.session_state.keys()):
            if key.startswith("hp_start_") or key.startswith("hp_end_"):
                del st.session_state[key]

    # ==============================
    # WEEK-END
    # ==============================
    st.sidebar.subheader("Plages horaires HP")

    state["weekend_hc"] = st.sidebar.checkbox(
        "Week-end entièrement en HC",
        value=state["weekend_hc"],
        key="weekend_hc"
    )

    # ==============================
    # NOMBRE DE PLAGES
    # ==============================
    state["nb_plages"] = st.sidebar.number_input(
        "Nombre de plages HP",
        min_value=1,
        max_value=5,
        value=state["nb_plages"],
         key="nb_plages"
    )

    # Synchronisation longueur
    while len(state["hp_ranges"]) < state["nb_plages"]:
        state["hp_ranges"].append((6, 22))

    while len(state["hp_ranges"]) > state["nb_plages"]:
        state["hp_ranges"].pop()

    # ==============================
    # ÉDITION DES PLAGES
    # ==============================
    
    hp_ranges = []

    for i in range(state["nb_plages"]):
    
        start_hour, end_hour = state["hp_ranges"][i]

        # Initialisation session_state si absent
        if f"hp_start_{i}" not in st.session_state:
            st.session_state[f"hp_start_{i}"] = time(start_hour, 0)
    
        if f"hp_end_{i}" not in st.session_state:
            st.session_state[f"hp_end_{i}"] = time(end_hour, 0)
    
        col1, col2 = st.sidebar.columns(2)
    
        with col1:
            start = st.selectbox(
                f"Début HP {i+1}",
                options=list(range(24)),
                index=start_hour,
                key=f"hp_start_{i}"
            )
    
        with col2:
            end = st.selectbox(
                f"Fin HP {i+1}",
                options=list(range(1, 25)),
                index=end_hour-1,
                key=f"hp_end_{i}"
            )
    
        state["hp_ranges"][i] = (start, end)
        hp_ranges.append((start, end))

    weekend_hc = st.session_state.tarif_state["weekend_hc"]

    # -----------------------
    # TARIFS IMPORT
    # -----------------------
    st.sidebar.subheader("Import réseau")

    tariff_importHP = st.sidebar.number_input(
        "Tarif import HP (CHF/kWh)",
        min_value=0.0,
        value=0.32,
        step=0.01,
        key="tariff_importHP"
    )

    tariff_importHC = st.sidebar.number_input(
        "Tarif import HC (CHF/kWh)",
        min_value=0.0,
        value=0.21,
        step=0.01,
        key="tariff_importHC"
    )

    # -----------------------
    # TARIF EXPORT
    # -----------------------
    st.sidebar.subheader("Export réseau")

    tariff_export = st.sidebar.number_input(
        "Tarif export (CHF/kWh)",
        min_value=0.0,
        value=0.08,
        step=0.01,
        key="tariff_export"
    )

# ==============================
# MODE TARIF UNIQUE
# ==============================
else:
    st.sidebar.subheader("Import / Export réseau")
    tariff_importHP = st.sidebar.number_input(
        "Tarif import (CHF/kWh)",
        min_value=0.0,
        value=0.32,
        step=0.01,
        key="tariff_importHP"
    )

    tariff_importHC = tariff_importHP  # identique

    tariff_export = st.sidebar.number_input(
        "Tarif export (CHF/kWh)",
        min_value=0.0,
        value=0.08,
        step=0.01,
        key="tariff_export"
    )

    hp_ranges = []  # pas utilisé
    weekend_hc = False


st.sidebar.header("⚙️ Paramètres")
debug = st.sidebar.checkbox("Avec / Sans DEBUG", value=False, key="debug")
roundtrip_eff = st.sidebar.slider("Rendement aller-retour", 0.5, 1.0, 0.96, key="roundtrip_eff")
cap_min = st.sidebar.number_input("Capacité min (kWh)", value=5, key="cap_min")
cap_max = st.sidebar.number_input("Capacité max (kWh)", value=30, key="cap_max")
cap_step = st.sidebar.number_input("Pas capacité (kWh)", value=1, key="cap_step")
soc_min_pct = st.sidebar.slider("SOC minimum batterie (%)", 0, 100, 5, key="soc_min_pct")
p_min = st.sidebar.number_input("Puissance min (kW)", value=1, key="p_min")
p_max = st.sidebar.number_input("Puissance max (kW)", value=10, key="p_max")
p_step = st.sidebar.number_input("Pas puissance (kW)", value=1, key="p_step")
gain_threshold = st.sidebar.slider("Seuil % du gain max", 0.5, 1.0, 0.95, key="gain_threshold")
daily_percentile = st.sidebar.slider("Percentile export journalier (Pxx)", 0.5, 0.99, 0.8, key="daily_percentile")
st.sidebar.header("⚙️ Capacité Auto")
capacite_auto = st.sidebar.checkbox("Capacité auto dynamique", value=True, key="capacite_auto")
if capacite_auto:
    facteur_auto = st.sidebar.slider("Facteur augmentation auto", 1.1, 3.0, 1.5, key="facteur_auto")
    max_auto_extensions = st.sidebar.number_input("Nombre max d'auto-extensions", 1, 10, 5, key="max_auto_extensions")
    cap_securite = st.sidebar.number_input("Capacité plafond sécurité (kWh)", 10, 10000, 10000, key="cap_securite")

if importExport_is_monthly:
    # Dates de fin de mois pour 2025
    dates = pd.date_range(start="2025-01-31", end="2025-12-31", freq='M')
    
    # DataFrame vide
    df = pd.DataFrame({
        "Date": dates,
        "Import (kWh)": [0]*12,
        "Export (kWh)": [0]*12
    })
    
    # Conversion en Excel en mémoire (sans writer.save())
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='ImportExport')
    excel_data = output.getvalue()
    
    # Bouton de téléchargement
    st.download_button(
        label="⬇️ Télécharger le fichier Excel type mensuel",
        data=excel_data,
        file_name="ImportExport_mensuel.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

# Fonction calul auto du gain HP/HC + Horaires
def compute_gain_with_time_of_use(
    imp_array,
    exp_array,
    imp_after,
    exp_after,
    hours,
    weekdays
):
    import_avoided = imp_array - imp_after
    export_avoided = exp_array - exp_after

    is_hp = np.zeros(len(hours), dtype=bool)
    
    for start, end in hp_ranges:
        is_hp |= (hours >= start) & (hours < end)

    if weekend_hc:
        timestamps_local = df.index
        is_hp[weekdays >= 5] = False

    import_tariffs = np.where(is_hp, tariff_importHP, tariff_importHC)

    gain = np.sum(import_avoided * import_tariffs - export_avoided * tariff_export)

    return gain

def compute_import_export_cashflow(
    imp_array,
    exp_array,
    hours,
    weekdays
):
    """
    Calcule :
    - le coût total d'import
    - le revenu total d'export
    """

    # Détermination heures pleines
    is_hp = np.zeros(len(hours), dtype=bool)

    for start, end in hp_ranges:
        is_hp |= (hours >= start) & (hours < end)

    # Week-end entièrement en HC
    if weekend_hc:
        is_hp[weekdays >= 5] = False  # 5 = samedi, 6 = dimanche

    # Sélection du tarif import
    import_tariffs = np.where(is_hp, tariff_importHP, tariff_importHC)

    # Calculs
    import_cost = np.sum(imp_array * import_tariffs)
    export_revenue = np.sum(exp_array * tariff_export)

    return import_cost, export_revenue

# Fonction de détection ligne d'en-tête
def find_header_row(df, date_tokens, import_tokens, export_tokens, max_rows=120):
    for r in range(min(max_rows, len(df))):
        row = [str(x).lower().strip() for x in df.iloc[r].values]
        row_text = " | ".join(row)      
        if any(t in row_text for t in date_tokens) \
           and any(t in row_text for t in import_tokens) \
           and any(t in row_text for t in export_tokens):
            return r
    return None

# Fonction de détection collones
def find_column(df, tokens):
    import re
    
    # Colonnes déjà sélectionnées (stockées dynamiquement)
    if not hasattr(find_column, "used_columns"):
        find_column.used_columns = set()

    for col in df.columns:
        if col in find_column.used_columns:
            continue

        col_lower = str(col).lower()

        for t in tokens:
            # match mot entier uniquement
            if re.search(rf"\b{re.escape(t)}\b", col_lower):
                find_column.used_columns.add(col)
                return col

    return None

# Fonction de SIMULATION BATTERIE -- COEUR DU PROG
def simulate_battery(exp_array, imp_array, cap_kwh, power_kw, soc_min_pct, eta, dt_hours):
    """
    - limitation des micro-cycles
    - ralentissement charge/décharge selon SOC
    - légère perte de veille
    """
    soc_min_kwh = cap_kwh * soc_min_pct / 100.0
    soc_max_kwh = cap_kwh
    soc_val = soc_min_kwh

    # Rendement unidirectionnel attendu ici
    eta = float(np.clip(eta, 1e-6, 1.0))

    # Paramètres anti micro-cycles
    standby_kw = 0.005                      # petite perte de veille
    standby_loss_kwh = standby_kw * dt_hours

    if importExport_is_monthly or export_is_monthly:
        deadband_kwh = 0.05 * dt_hours / 0.25
        min_action_kwh = 0.10 * dt_hours / 0.25
    else:
        deadband_kwh = 0.15 * dt_hours / 0.25
        min_action_kwh = 0.30 * dt_hours / 0.25

    p_step_nominal = power_kw * dt_hours

    soc_list = []
    imp_after = np.zeros_like(imp_array, dtype=float)
    exp_after = np.zeros_like(exp_array, dtype=float)
    charge_series = np.zeros_like(exp_array, dtype=float)
    discharge_series = np.zeros_like(imp_array, dtype=float)

    for i in range(len(exp_array)):
        soc_val = max(soc_min_kwh, soc_val - standby_loss_kwh)

        exp_orig = max(0.0, float(exp_array[i]))
        imp_orig = max(0.0, float(imp_array[i]))

        exp_i = exp_orig
        imp_i = imp_orig

        soc_pct = 100.0 * soc_val / cap_kwh if cap_kwh > 0 else 0.0

        # Réduction puissance en fin de charge / proche SOC mini
        if soc_pct >= 98:
            charge_factor = 0.05
        elif soc_pct >= 95:
            charge_factor = 0.20
        elif soc_pct >= 90:
            charge_factor = 0.50
        else:
            charge_factor = 1.0

        discharge_factor = 1.0

        p_step_charge = p_step_nominal * charge_factor
        p_step_discharge = p_step_nominal * discharge_factor

        # -----------------------------
        # Anti import/export simultanés artificiels
        # -----------------------------
        if exp_i > deadband_kwh and imp_i > deadband_kwh:
            # On neutralise d'abord les flux opposés
            common = min(exp_i, imp_i)
            exp_i -= common
            imp_i -= common

        # Valeurs par défaut
        exp_after_i = exp_orig
        imp_after_i = imp_orig

        # -----------------------------
        # CHARGE
        # -----------------------------
        if exp_i >= min_action_kwh:
            room_kwh = max(0.0, soc_max_kwh - soc_val)
            max_charge_input = min(
                p_step_charge,
                room_kwh / eta if eta > 0 else 0.0
            )

            charge_input = min(exp_i, max_charge_input)

            # zone morte finale
            if charge_input >= min_action_kwh:
                energy_into_battery = charge_input * eta
                soc_val += energy_into_battery
                charge_series[i] = energy_into_battery
                exp_after_i = exp_orig - charge_input

        # -----------------------------
        # DÉCHARGE
        # -----------------------------
        if imp_i >= min_action_kwh:
            available_from_soc = max(0.0, soc_val - soc_min_kwh)

            max_energy_from_soc = min(
                available_from_soc,
                p_step_discharge / eta if eta > 0 else 0.0
            )

            energy_from_battery = min(
                max_energy_from_soc,
                imp_i / eta if eta > 0 else 0.0
            )

            if energy_from_battery * eta >= min_action_kwh:
                discharge_to_load = energy_from_battery * eta
                soc_val -= energy_from_battery
                discharge_series[i] = energy_from_battery
                imp_after_i = imp_orig - discharge_to_load

        soc_val = min(max(soc_val, soc_min_kwh), soc_max_kwh)

        imp_after[i] = max(0.0, imp_after_i)
        exp_after[i] = max(0.0, exp_after_i)
        soc_list.append(soc_val)

    charge_total = float(np.sum(charge_series))
    discharge_total = float(np.sum(discharge_series))

    return (
        soc_list,
        imp_after,
        exp_after,
        charge_series,
        discharge_series,
        charge_total,
        discharge_total
    )

def eq_cycles_dod(charge_series, discharge_series, cap_kwh, soc_min_pct):
    usable_cap = cap_kwh * (1 - soc_min_pct / 100)

    if usable_cap <= 0:
        return 0.0

    total_discharge = np.sum(np.maximum(discharge_series, 0))
    return float(total_discharge / usable_cap)
    

# Choix du Type de fichier
if data_mode == "Fichier GRD (Excel/CSV unique)":
    uploaded_file = st.file_uploader(
    (
        "Choisir un fichier Excel ou CSV -- \n"
        "Mots clefs :\n"
        "date [date, datetime, horodatage, timestamp, date/heure, date heure] -- \n"
        "import [soutirage, import, achat, reseau, consommation] -- \n"
        "export [surplus, surplus solaire, export, excedent, reinjection, réinjection, injection]"
    ),
    type=["xlsx", "xls", "csv"],
    accept_multiple_files=False
)
else:
    uploaded_file = st.file_uploader(
        "Choisir les 12 fichiers mensuels -- \n"
        "Mots clefs :\n"
        "date [début, date, datetime, horodatage, timestamp, date/heure, date heure] -- \n"
        "import [soutirage, import, achat, reseau, consommation] -- \n"
        "export [surplus, surplus solaire, export, excedent, reinjection, réinjection, injection]",
        type=["xlsx", "xls"],
        accept_multiple_files=True
    )
if uploaded_file:
    # ==========================================================
    # VERIFICATION PARAMETRES ENTREE
    # ==========================================================
    warnings = []
    if cap_min >= cap_max:
        warnings.append("⚠️ Capacité min >= capacité max.")
    if p_min >= p_max:
        warnings.append("⚠️ Puissance min >= puissance max.")
    if gain_threshold <= 0 or gain_threshold > 1:
        warnings.append("⚠️ Le seuil de gain doit être entre 0 et 1.")
    if roundtrip_eff <= 0 or roundtrip_eff > 1:
        warnings.append("⚠️ Le rendement doit être entre 0 et 1.")
    if warnings:
        for w in warnings:
            st.error(w)
        st.stop()
    
    st.header("🔹 Recherche des lignes / collones")
    # ==========================================================
    # CAS 1 : FICHIER GRD
    # ==========================================================
    if data_mode == "Fichier GRD (Excel/CSV unique)":
        file_type = uploaded_file.name.split('.')[-1].lower()
        date_tokens = ["début", "date", "datetime", "horodatage", "timestamp", "date/heure", "date heure"]
        import_tokens = ["soutirage", "import", "achat", "reseau", "consommation", "négative"]
        export_tokens = ["surplus", "surplus solaire", "export", "excedent", "reinjection", "réinjection", "positive", "injection"]

        # Ajouter mots-clés personnalisés si fournis
        if MotCle:
            if custom_date_tokens:
                date_tokens += [t.strip().lower() for t in custom_date_tokens.split(",")]
            
            if custom_import_tokens:
                import_tokens += [t.strip().lower() for t in custom_import_tokens.split(",")]
            
            if custom_export_tokens:
                export_tokens += [t.strip().lower() for t in custom_export_tokens.split(",")]

        if file_type == "csv":
            # Lire sans header pour détecter la ligne d'en-tête
            df_full = pd.read_csv(uploaded_file, header=None, sep=';', engine='python')
            uploaded_file.seek(0)  # Remettre le curseur au début
        else:
            df_full = pd.read_excel(uploaded_file, header=None)
            uploaded_file.seek(0)  # Remettre le curseur au début

        header_row = find_header_row(df_full, date_tokens, import_tokens, export_tokens)
        
        if header_row is None:
            st.error(f"❌ Impossible de détecter la ligne d'en-tête dans {uploaded_file.name}")
            st.warning(f"Vérifier que TOUS les mots clés sont présent sur la même ligne OU essayer par Mots-clés personnalisés")
            st.info(f"date_tokens {date_tokens}")
            st.info(f"import_tokens {import_tokens}")
            st.info(f"export_tokens {export_tokens}")
            st.stop()
        else:
            st.success(f"Ligne d'en-tête détectée : {header_row + 1}")

        if file_type == "csv":
            df = pd.read_csv(uploaded_file, header=header_row, sep=None, engine='python')
        else:
            df = pd.read_excel(uploaded_file, header=header_row)

        st.write("Aperçu des 5 premières lignes du fichier :")
        st.write(f"📊 Nombre de lignes : {len(df)}")
        
        # Verification de la longueur
        LongBase = len(df)
        
        st.dataframe(df.head())
        
        # Réinitialiser l'état avant chaque fichier
        find_column.used_columns = set()

        if importExport_is_monthly:
            date_col = find_column(df, date_tokens)
            imp_col = find_column(df, import_tokens)
            exp_col = find_column(df, export_tokens)

            if date_col is None :
                st.error("Impossible de détecter automatiquement la colonne date.")
            elif imp_col is None :
                st.error("Impossible de détecter automatiquement la colonne import.")
            elif exp_col is None:
                st.error("Impossible de détecter automatiquement la colonne export.")
            if date_col is None or imp_col is None or exp_col is None:
                st.write("Colonnes détectées :", list(df.columns))
                st.stop()
            else:
                st.success(f"Colonnes détectées : date={date_col}, import={imp_col}, export={exp_col}")
        
            # Afficher lignes sans date
            missing_dates = df[df[date_col].isna()]
            if not missing_dates.empty:
                st.warning(f"⚠️ {len(missing_dates)} lignes n'ont pas de date et seront ignorées.")
                st.dataframe(missing_dates.head(10))
    
            # Supprimer uniquement les lignes sans date
            df = df.dropna(subset=[date_col]).reset_index(drop=True)

            st.header(" 🔹 Nettoyage et conversion")
            # Nettoyage
            df[imp_col] = pd.to_numeric(df[imp_col], errors='coerce').fillna(0)
            df[exp_col] = pd.to_numeric(df[exp_col], errors='coerce').fillna(0)
            # certaines dates peuvent avoir DST
            df[date_col] = remove_dst(df[date_col])
            
            # Conversion en datetime
            df[date_col] = pd.to_datetime(df[date_col], errors='coerce', utc=True)
            
            df[date_col] = df[date_col].dt.tz_convert(None)
    
            # Trier par date
            df = df.sort_values(date_col).reset_index(drop=True)

            # =====================================================
            # CAS : FICHIER CONTIENT UNIQUEMENT 12 TOTAUX MENSUELS
            # =====================================================
            
            if len(df) != 12:
                st.error("❌ Le fichier doit contenir exactement 12 lignes (1 par mois).")
                st.stop()
            
            # Année détectée
            year = pd.to_datetime(df[date_col]).dt.year.mode()[0]
            
            # Création index annuel complet
            date_range = pd.date_range(
                start=f"{year}-01-01 00:00:00",
                end=f"{year}-12-31 23:59:59",
                freq=f"{int(dt_hours*60)}min"
            )
            
            df_full = pd.DataFrame(index=date_range)
            df_full["month"] = df_full.index.month
            df_full["hour"] = df_full.index.hour
            
            df_full["import_kWh"] = 0.0
            df_full["export_kWh"] = 0.0
            df_full[imp_col] = df[imp_col]
            df_full[exp_col] = df[exp_col]
            
            # ===============================
            # Reconstruction mensuelle
            # ===============================
            for _, row in df.iterrows():
            
                month = pd.to_datetime(row[date_col]).month
                total_import = row[imp_col]
                total_export = row[exp_col]
            
                mask_month = df_full["month"] == month
            
                # ----------------------
                # IMPORT → réparti 24h
                # ----------------------
                steps_month = mask_month.sum()
                import_per_step = total_import / steps_month
                df_full.loc[mask_month, "import_kWh"] = import_per_step
            
                # ----------------------
                # EXPORT → courbe solaire réaliste
                # ----------------------
                solar_windows = {
                    1: (8, 16),
                    2: (8, 17),
                    3: (7, 18),
                    4: (6, 20),
                    5: (6, 21),
                    6: (5, 21),
                    7: (5, 21),
                    8: (6, 20),
                    9: (7, 19),
                    10: (7, 18),
                    11: (8, 16),
                    12: (8, 16),
                }
                
                sunrise, sunset = solar_windows[month]
                
                hour_decimal = df_full.loc[mask_month].index.hour + df_full.loc[mask_month].index.minute / 60.0
                
                solar_mask = mask_month.copy()
                solar_mask.loc[mask_month] = (hour_decimal >= sunrise) & (hour_decimal <= sunset)
                
                idx = df_full.index[solar_mask]
                
                if len(idx) == 0:
                    st.error(f"❌ Aucun créneau solaire détecté pour mois {month}.")
                    st.stop()
                
                hours_selected = np.array(idx.hour + idx.minute / 60.0, dtype=float)
                day_length = sunset - sunrise
                x = (hours_selected - sunrise) / day_length
                
                # Courbe PV en cloche
                pv_shape = np.sin(np.pi * x)
                pv_shape = np.clip(pv_shape, 0, None) ** 1.5
                
                shape_sum = np.sum(pv_shape)
                
                if shape_sum > 0:
                    df_full.loc[idx, "export_kWh"] = total_export * (pv_shape / shape_sum)
            
            # =====================================================
            # Remplacement propre par la reconstruction annuelle
            # =====================================================
            
            df_full = df_full.reset_index().rename(columns={"index": date_col})
            
            df = df_full.copy()
            df = df.sort_values(date_col).reset_index(drop=True)

            df["dt_h"] = df[date_col].diff().dt.total_seconds() / 3600
            unique_dt = df["dt_h"].round(2).value_counts()
            dt_hours = unique_dt.idxmax()  # le plus fréquent
            st.success(f"✅ dt_hours (mensuel reconstruit) = {dt_hours:.3f} h")
            
            st.success("✅ Année complète reconstruite avec contrainte Export 10h–16h et Import 24/24h.")

        else:
            date_col = find_column(df, date_tokens)
            imp_col = find_column(df, import_tokens)
            exp_col = find_column(df, export_tokens)
    
            if date_col is None :
                st.error("Impossible de détecter automatiquement la colonne date.")
            elif imp_col is None :
                st.error("Impossible de détecter automatiquement la colonne import.")
            elif exp_col is None:
                st.error("Impossible de détecter automatiquement la colonne export.")
            if date_col is None or imp_col is None or exp_col is None:
                st.write("Colonnes détectées :", list(df.columns))
                st.stop()
            else:
                st.success(f"Colonnes détectées : date={date_col}, import={imp_col}, export={exp_col}")
            
            # Afficher lignes sans date
            missing_dates = df[df[date_col].isna()]
            if not missing_dates.empty:
                st.warning(f"⚠️ {len(missing_dates)} lignes n'ont pas de date et seront ignorées.")
                st.dataframe(missing_dates.head(10))
    
            # Supprimer uniquement les lignes sans date
            df = df.dropna(subset=[date_col]).reset_index(drop=True)
        
            st.header(" 🔹 Nettoyage et conversion")
            # Nettoyage
            df[imp_col] = pd.to_numeric(df[imp_col], errors='coerce').fillna(0)
            df[exp_col] = pd.to_numeric(df[exp_col], errors='coerce').fillna(0)
            # certaines dates peuvent avoir DST
            df[date_col] = remove_dst(df[date_col])
            
            # Conversion en datetime
            df[date_col] = pd.to_datetime(df[date_col], errors='coerce', utc=True)
            
            df[date_col] = df[date_col].dt.tz_convert(None)
    
            # Trier par date
            df = df.sort_values(date_col).reset_index(drop=True)
        
            # Détection automatique si compteur cumulatif
            import_is_cumulative = (df[imp_col].diff().dropna() >= 0).mean() > 0.95
            export_is_cumulative = (df[exp_col].diff().dropna() >= 0).mean() > 0.95
    
            st.info(
            f"Détection automatique : Import cumulatif = {import_is_cumulative}, "
            f"Export cumulatif = {export_is_cumulative}"
            )
    
            # -------------------------------------------------
            # 1 : Données en kW instantané
            # -------------------------------------------------
            if values_are_kw and not import_is_cumulative:
        
                df["dt_h"] = df[date_col].diff().dt.total_seconds() / 3600
                df.loc[0, "dt_h"] = df["dt_h"].median()
                
                dt_hours = df["dt_h"].median()
                st.info(
                f"Détection automatique : dt_hours = {dt_hours}")
        
                df["import_kWh"] = df[imp_col] * df["dt_h"]
                df["export_kWh"] = df[exp_col] * df["dt_h"]
        
            # -------------------------------------------------
            # 2 : Compteur cumulatif
            # -------------------------------------------------
            elif import_is_cumulative:
        
                df["import_kWh"] = df[imp_col].diff()
                df["export_kWh"] = df[exp_col].diff()
        
                df.loc[0, "import_kWh"] = df.loc[0, imp_col]
                df.loc[0, "export_kWh"] = df.loc[0, exp_col]
    
                 # Première ligne → supprimer car diff invalide
                df = df.iloc[1:].reset_index(drop=True)
        
            # -------------------------------------------------
            # 3 : Déjà en kWh par intervalle
            # -------------------------------------------------
            else:
                if unite == "kWh": # En kWh
                    df["import_kWh"] = df[imp_col]
                    df["export_kWh"] = df[exp_col]
                else:             # En Wh
                    df["import_kWh"] = df[imp_col]/1000
                    df["export_kWh"] = df[exp_col]/1000
        
            # Sécurité finale
            df["import_kWh"] = df["import_kWh"].clip(lower=0)
            df["export_kWh"] = df["export_kWh"].clip(lower=0)
            
            # -----------------------------
            # Gestion doublons DST
            # -----------------------------
            # On identifie les doublons
            duplicates_mask = df.duplicated(subset=[date_col], keep=False)
            if duplicates_mask.any():
                # Masque des doublons en octobre
                october_mask = df[date_col].dt.month == 10
                october_duplicates = duplicates_mask & october_mask
            
                if october_duplicates.any():
                    st.warning(f"⚠️ {(october_duplicates.sum()/2)} doublons détectés en octobre (ignorés car DST).")
                else:
                    st.warning(f"⚠️Doublons détectés hors octobre (Fusionnés).")
                # Garder le premier doublon sauf ceux en octobre
                df = df[~(duplicates_mask & ~october_mask)].reset_index(drop=True)

        # -------------------------------------------------
        # Nettoyer le DataFrame : garder uniquement les colonnes utiles
        # -------------------------------------------------
        columns_to_keep = [date_col, imp_col, exp_col, "import_kWh", "export_kWh"]
        df = df[[c for c in columns_to_keep if c in df.columns]]
        
        st.write("✅ Aperçu des données converties :")
        st.write(f"📊 Nombre de lignes : {len(df)}")
        st.dataframe(df.head())

    # ==========================================================
    # CAS 2 : FICHIERS MENSUELS
    # ==========================================================
    else:
        if not uploaded_file or len(uploaded_file) != 12:
            st.error("Veuillez charger exactement 12 fichiers (1 par mois).")
            st.stop()

        df_list = []

        for file in uploaded_file:

            file_type = file.name.split('.')[-1].lower()

            if file_type == "csv":
                df_month = pd.read_csv(file, sep=None, header=None, engine='python')
            else:
                df_month = pd.read_excel(file, header=None)

            # =============================
            # Détection colonnes EXACTES
            # =============================
            date_col = None

            # tokens pour la détection automatique
            date_tokens = ["début", "date", "datetime", "horodatage", "timestamp", "date/heure", "date heure"]
            import_tokens = ["soutirage", "import", "achat", "reseau", "consommation", "négative"]
            export_tokens = ["surplus", "export", "excedent", "reinjection", "réinjection", "positive", "injection"]

            # chercher ligne d'en-tête automatiquement
            header_row = find_header_row(df_month, date_tokens, import_tokens, export_tokens)
            if header_row is not None:
                df_month = pd.read_excel(file, header=header_row) if file_type != "csv" else pd.read_csv(file, sep=';', header=header_row, engine='python')
            else:
                st.error(f"❌ Impossible de détecter la ligne d'en-tête dans {file.name}")
                st.warning(f"Vérifier que TOUS les mots clés sont présent sur la même ligne OU essayer par Mots-clés personnalisés")
                st.info(f"date_tokens {date_tokens}")
                st.info(f"import_tokens {import_tokens}")
                st.info(f"export_tokens {export_tokens}")
                st.stop()
            
            # Réinitialiser l'état avant chaque fichier
            find_column.used_columns = set()
            
            # détecter colonnes
            date_col = find_column(df_month, date_tokens)
            imp_col = find_column(df_month, import_tokens)
            exp_col = find_column(df_month, export_tokens)

            if date_col is None :
                st.error("Impossible de détecter automatiquement la colonne date.")
            elif imp_col is None :
                st.error("Impossible de détecter automatiquement la colonne import.")
            elif exp_col is None:
                st.error("Impossible de détecter automatiquement la colonne export.")
            if date_col is None or imp_col is None or exp_col is None:
                st.write("Colonnes détectées :", list(df.columns))
                st.stop()
            else:
                st.success(f"Colonnes détectées dans {file.name} : date={date_col}, import={imp_col}, export={exp_col}")

            # Verification de la longueur
            LongBase += len(df_month)
            
            # -----------------------------
            # Nettoyage & conversion
            # -----------------------------
            df_month[imp_col] = pd.to_numeric(df_month[imp_col], errors="coerce").fillna(0)
            df_month[exp_col] = pd.to_numeric(df_month[exp_col], errors="coerce").fillna(0)
            # certaines dates peuvent avoir DST
            df_month[date_col] = remove_dst(df_month[date_col])
            
            # Conversion en datetime
            df_month[date_col] = pd.to_datetime(df_month[date_col], errors='coerce', utc=True)
            
            df_month[date_col] = df_month[date_col].dt.tz_convert(None)

            # Trier par date
            df_month = df_month.sort_values(date_col).reset_index(drop=True)
            
            # Détection automatique si compteur cumulatif
            import_is_cumulative = (df_month[imp_col].diff().dropna() >= 0).mean() > 0.95
            export_is_cumulative = (df_month[exp_col].diff().dropna() >= 0).mean() > 0.95
    
            st.info(
                f"Détection automatique {file.name} : Import cumulatif = {import_is_cumulative}, "
                f"Export cumulatif = {export_is_cumulative}"
            )
    
            # -------------------------------------------------
            # 1 : Données en kW instantané
            # -------------------------------------------------
            if values_are_kw and not import_is_cumulative:
            
                df_month["dt_h"] = df_month[date_col].diff().dt.total_seconds() / 3600
                df_month.loc[0, "dt_h"] = df_month["dt_h"].median()
                
                dt_hours = df_month["dt_h"].median()
                st.info(
                f"Détection automatique : dt_hours = {dt_hours}")
            
                df_month["import_kWh"] = df_month[imp_col] * df_month["dt_h"]
                df_month["export_kWh"] = df_month[exp_col] * df_month["dt_h"]
            
            # -------------------------------------------------
            # 2 : Compteur cumulatif
            # -------------------------------------------------
            elif import_is_cumulative:
            
                df_month["import_kWh"] = df_month[imp_col].diff()
                df_month["export_kWh"] = df_month[exp_col].diff()
            
                df_month.loc[0, "import_kWh"] = df_month.loc[0, imp_col]
                df_month.loc[0, "export_kWh"] = df_month.loc[0, exp_col]
            
                 # Première ligne → supprimer car diff invalide
                df_month = df_month.iloc[1:].reset_index(drop=True)
            
            # -------------------------------------------------
            # 3 : Déjà en kWh par intervalle
            # -------------------------------------------------
            else:
                if unite == "kWh": # En kWh
                    df_month["import_kWh"] = df_month[imp_col]
                    df_month["export_kWh"] = df_month[exp_col]
                else:             # En Wh
                    df_month["import_kWh"] = df_month[imp_col]/1000
                    df_month["export_kWh"] = df_month[exp_col]/1000
            
            # Sécurité finale
            df_month["import_kWh"] = df_month["import_kWh"].clip(lower=0)
            df_month["export_kWh"] = df_month["export_kWh"].clip(lower=0)

            # -----------------------------
            # Gestion doublons DST
            # -----------------------------
            # On identifie les doublons
            duplicates_mask = df_month.duplicated(subset=[date_col], keep=False)
            if duplicates_mask.any():
                # Masque des doublons en octobre
                october_mask = df_month[date_col].dt.month == 10
                october_duplicates = duplicates_mask & october_mask
            
                if october_duplicates.any():
                    st.warning(f"⚠️ {(october_duplicates.sum()/2)} doublons détectés en octobre (ignorés car DST).")
                else:
                    st.warning(f"⚠️Doublons détectés hors octobre (Fusionnés).")
                # Garder le premier doublon sauf ceux en octobre
                df_month = df_month[~(duplicates_mask & ~october_mask)].reset_index(drop=True)
                
            # -----------------------------
            # Colonnes finales
            # -----------------------------
            columns_to_keep = [date_col, imp_col, exp_col, "import_kWh", "export_kWh"]
            
            # Sélectionner les colonnes à conserver
            df_month = df_month[[c for c in columns_to_keep if c in df_month.columns]]

            # 👇 AJOUTER ICI
            with st.expander(f"📄 {file.name}", expanded=False):
                st.write(f"Nombre de lignes : {len(df_month)}")
                st.write(f"Colonnes : {list(df_month.columns)}")
                st.dataframe(df_month.head(20))
            
            # Ajouter ce fichier à la liste pour fusion annuelle
            df_list.append(df_month)
    
        # Fusion annuelle
        df = pd.concat(df_list).sort_values(date_col).reset_index(drop=True)
        
        st.success("✅ 12 fichiers mensuels fusionnés et convertis avec succès")
        st.dataframe(df.head())

    if data_mode == "Fichier GRD (Excel/CSV unique)" and not importExport_is_monthly and export_is_monthly :
        st.header(" 🔹 Nettoyage et conversion")
        st.info("⚙️ Reconstruction d’un profil export à partir des totaux mensuels.")
        # -----------------------------------------
        # Reconstruction export mensuel EXACT
        # -----------------------------------------

        month_map = {
            1: monthly_export_values["Janvier"],
            2: monthly_export_values["Février"],
            3: monthly_export_values["Mars"],
            4: monthly_export_values["Avril"],
            5: monthly_export_values["Mai"],
            6: monthly_export_values["Juin"],
            7: monthly_export_values["Juillet"],
            8: monthly_export_values["Août"],
            9: monthly_export_values["Septembre"],
            10: monthly_export_values["Octobre"],
            11: monthly_export_values["Novembre"],
            12: monthly_export_values["Décembre"],
        }

        df["month"] = df[date_col].dt.month
        df["hour"] = df[date_col].dt.hour

        # Initialisation export
        df["export_kWh"] = 0.0
        
        solar_windows = {
            1: (8, 16),
            2: (8, 17),
            3: (7, 18),
            4: (6, 20),
            5: (6, 21),
            6: (5, 21),
            7: (5, 21),
            8: (6, 20),
            9: (7, 19),
            10: (7, 18),
            11: (8, 16),
            12: (8, 16),
        }
        
        for month in range(1, 13):
            mask_month = df["month"] == month
            sunrise, sunset = solar_windows[month]
        
            hour_decimal = df.loc[mask_month, date_col].dt.hour + df.loc[mask_month, date_col].dt.minute / 60.0
        
            solar_mask = mask_month.copy()
            solar_mask.loc[mask_month] = (hour_decimal >= sunrise) & (hour_decimal <= sunset)
        
            idx = df.index[solar_mask]
            if len(idx) == 0:
                continue
        
            hours_selected = np.array(
                df.loc[idx, date_col].dt.hour + df.loc[idx, date_col].dt.minute / 60.0,
                dtype=float
            )
            day_length = sunset - sunrise
            x = (hours_selected - sunrise) / day_length
            
            # Courbe PV en cloche
            pv_shape = np.sin(np.pi * x)
            pv_shape = np.clip(pv_shape, 0, None) ** 1.5
            
            shape_sum = np.sum(pv_shape)
            if shape_sum > 0:
                month_total = monthly_export_values[list(monthly_export_values.keys())[month - 1]]
                df.loc[idx, "export_kWh"] = month_total * (pv_shape / shape_sum)
                
        # Vérification des sommes mensuelles
        sums = df.groupby("month")["export_kWh"].sum()

        # Nettoyage colonnes temporaires
        df.drop(columns=["month", "hour", "solar_mask"], inplace=True, errors="ignore")
    
    st.header(" 🔹 Nettoyage + verif global")
    # Calculer l'année majoritaire
    year_counts = df[date_col].dt.year.value_counts()
    target_year = year_counts.idxmax()  # l'année qui apparaît le plus
    
    # Filtrer le DataFrame pour ne garder que cette année
    df = df[df[date_col].dt.year == target_year].reset_index(drop=True)

    df["dt_h"] = df[date_col].diff().dt.total_seconds() / 3600
    unique_dt = df["dt_h"].round(2).value_counts()
    dt_hours = unique_dt.idxmax()  # le plus fréquent
    st.warning(f"⚠️ dt_hours FINAL utilisé = {dt_hours:.3f} h (~{dt_hours*60:.0f} min)")
    
    st.info(f"Filtrage pour l'année majoritaire : {target_year}")
        
    st.subheader("Aperçu du fichier reconstitué")
    st.write(f"Nombre de lignes : {len(df)}")
    LongApres = len(df)
    st.dataframe(df.head(31000))  # Affiche les 10 premières lignes

    # =========================================
    # Vérification de la qualité des données
    # =========================================
   
    # Colonnes utilisées
    cols_needed = ["import_kWh", "export_kWh"]

     # 0 Vérifier nombre de lignes avant après
    ratio = LongApres / LongBase if LongBase > 0 else 0
    
    ratio = LongApres / LongBase if LongBase > 0 else 0
    perte_lignes = LongBase - LongApres
    perte_pct = (perte_lignes / LongBase * 100) if LongBase > 0 else 0
    
   
    st.info(
        f"🔎 Vérification du nombre de lignes :\n\n"
        f"- Nombre de lignes avant conversion : {LongBase}\n"
        f"- Nombre de lignes après conversion : {LongApres}\n"
        f"- Lignes perdues : {perte_lignes}\n"
        f"- Perte relative : {perte_pct:.4f} %\n"
        f"- Ratio conservation : {ratio:.6f}"
    )
    
    # Identifier les lignes perdues
    lost_rows = LongBase - LongApres
    if ratio < 0.95:
        st.error("❌ Nombre de lignes après conversion incohérent - Veuillez verifier que le fichier source ne contienne par trop de date d'annee differente (Tolerance 95%)")
    
        st.write(f"➡️ Lignes perdues : {lost_rows}")
    
        # Vérifier trous temporels
        if isinstance(df[date_col], pd.Series):
            diffs = df[date_col].diff().dt.total_seconds()
            st.write("⏱️ Pas de temps min / max (s) :", diffs.min(), diffs.max())
    
        st.stop()

    # 1️ Vérifier qu'il n'y a pas de NaN
    nan_rows = df[df[cols_needed].isna().any(axis=1)]
    if not nan_rows.empty:
        st.error("⚠️ Certaines lignes contiennent des valeurs manquantes (NaN) :")
        st.dataframe(nan_rows)

    # 2️ Vérifier qu'il y a assez de données non nulles pour exploiter
    import_nonzero = (df["import_kWh"] > 0).sum()
    export_nonzero = (df["export_kWh"] > 0).sum()

    if import_nonzero == 0:
        st.error("❌ Toutes les valeurs d'import sont nulles. Impossible de simuler.")
        st.stop()
    if export_nonzero == 0:
        st.error("❌ Toutes les valeurs d'export sont nulles. Impossible de simuler.")
        st.info("Cocher dans la barre latérale « Export fourni en total mensuel » et entrer les valeurs manuellement.")
        st.stop()

    # 3️ Vérifier que les timestamps sont bien ordonnés
    if not df[date_col].is_monotonic_increasing:
        st.warning("⚠️ Les dates ne sont pas strictement croissantes, certaines anomalies sont possibles.")

   

    # 4 Vérifier qu'il y a un nombre minimal de points de données
    if len(df) < 1000:
        st.warning("⚠️ Le nombre total de lignes est très faible (<1000). La simulation risque d'être peu fiable.")

    # 5 Résumé des données
    import_total_before = df[imp_col].sum()
    export_total_before = df[exp_col].sum()
    import_total_after = df["import_kWh"].sum()
    export_total_after = df["export_kWh"].sum()
    
    import_nonzero = (df["import_kWh"] > 0).sum()
    export_nonzero = (df["export_kWh"] > 0).sum()
    
    st.info(
        f"📊 Résumé :\n"
        f"- Total de lignes : {len(df)}\n"
        f"- Import non nul : {import_nonzero} lignes\n"
        f"- Export non nul : {export_nonzero} lignes\n"
        f"- Valeurs min/max import : {df['import_kWh'].min():.2f} / {df['import_kWh'].max():.2f}\n"
        f"- Valeurs min/max export : {df['export_kWh'].min():.2f} / {df['export_kWh'].max():.2f}\n"
        f"- Total import avant conversion : {import_total_before:.2f} \n"
        f"- Total import après conversion : {import_total_after:.2f} \n"
        f"- Total export avant conversion : {export_total_before:.2f} \n"
        f"- Total export après conversion : {export_total_after:.2f} "
    )

    st.info("📥 Télécharger le fichier Excel reconstitue")

    st.write("dt_hours (valeur utilisée) :", dt_hours)

    # Créer un buffer Excel
    excel_buffer = BytesIO()
    with pd.ExcelWriter(excel_buffer, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name="Reconstitue")
    excel_buffer.seek(0)
    
    # Repositionner le curseur au début du buffer
    excel_buffer.seek(0)
    
    # Bouton Streamlit pour télécharger le fichier
    st.download_button(
        label="Télécharger le fichier Excel reconstitué",
        data=excel_buffer,
        file_name="fichier_reconstitue.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    st.header(" 🔹 SIMULATION VECTORISÉE")
    min_action_kwh = 0.01 * dt_hours / 0.25
    import_total_kWh = float(df["import_kWh"].sum())
    export_total_kWh = float(df["export_kWh"].sum())
    import_max_step = float(df["import_kWh"].max())
    export_max_step = float(df["export_kWh"].max())
    import_mean_step = float(df["import_kWh"].mean())
    export_mean_step = float(df["export_kWh"].mean())
    nb_import_nonzero = int((df["import_kWh"] > 0).sum())
    nb_export_nonzero = int((df["export_kWh"] > 0).sum())
    nb_export_ge_min_action = int((df["export_kWh"] >= min_action_kwh).sum())
    nb_import_ge_min_action = int((df["import_kWh"] >= min_action_kwh).sum())
    
    st.info(
        f"🔎 Contrôle avant simulation :\n\n"
        f"- Total import : {import_total_kWh:.2f} kWh\n"
        f"- Total export : {export_total_kWh:.2f} kWh\n"
        f"- Import non nul : {nb_import_nonzero} lignes\n"
        f"- Export non nul : {nb_export_nonzero} lignes\n"
        f"- Valeur max import par pas : {import_max_step:.2f} kWh\n"
        f"- Valeur max export par pas : {export_max_step:.2f} kWh\n"
        f"- Valeur moyenne import par pas : {import_mean_step:.2f} kWh\n"
        f"- Valeur moyenne export par pas : {export_mean_step:.2f} kWh\n"
        f"- Seuil min_action_kwh : {min_action_kwh:.2f} kWh\n"
        f"- Nombre de pas export ≥ seuil : {nb_export_ge_min_action}\n"
        f"- Nombre de pas import ≥ seuil : {nb_import_ge_min_action}"
    )

    # ==========================================================
    # CALCUL CAPACITÉ MAX DYNAMIQUE (avec auto-extension + auto-step)
    # ==========================================================
    with st.spinner("Calcul de la capacité maximale dynamique..."):

        daily_export = df.groupby(df[date_col].dt.date)["export_kWh"].sum()
        cap_dyn_raw = np.ceil(np.percentile(daily_export, daily_percentile*100))

        auto_extension_count = 0
        original_cap_max = cap_max
        original_cap_step = cap_step

        # ==================================================
        # NOUVELLE LOGIQUE : cap_max = cap_dyn_raw + 5
        # ==================================================
        if capacite_auto:
    
            cap_max = min(int(cap_dyn_raw + 5), cap_securite)
    
            st.info(
                f"⚙️ cap_max ajusté automatiquement à **cap_dyn_raw + 5** → "
                f"{cap_max} kWh"
            )
    
            if cap_max == cap_securite:
                st.warning("⚠️ Plafond cap_securite atteint.")
    
            # ==================================================
            # AJUSTEMENT AUTOMATIQUE DU PAS DE CAPACITÉ
            # ==================================================
            target_points = 40
            range_size = cap_max - cap_min
    
            if range_size > target_points:
                cap_step = max(1, int(np.ceil(range_size / target_points)))
    
                st.info(
                    f"⚙️ cap_step ajusté automatiquement de "
                    f"{original_cap_step} à {cap_step} kWh "
                    f"pour limiter le temps de calcul."
                )
    
        else:
            if cap_dyn_raw >= cap_max:
                st.warning(
                    "⚠️ Capacité dynamique plafonnée par cap_max → "
                    "Activer 'Capacité auto dynamique'."
                )
    
        cap_max_dyn = min(max(cap_dyn_raw, cap_min), cap_max)
    
        st.sidebar.markdown(f"Capacité max dynamique: **{cap_max_dyn} kWh**")
    
    # ==========================================================
    # SIMULATION VECTORISÉE
    # ==========================================================  
    timestamps = pd.to_datetime(df[date_col])
    hours = timestamps.dt.hour.to_numpy(dtype=int)
    weekdays = timestamps.dt.weekday.to_numpy(dtype=int)

    exp_array = df["export_kWh"].values
    imp_array = df["import_kWh"].values

    import_cost, export_revenue = compute_import_export_cashflow(
        imp_array,
        exp_array,
        hours,
        weekdays
    )
        
    with st.spinner("Simulation et recherche du meilleur choix..."):
        exp_array = df["export_kWh"].values
        imp_array = df["import_kWh"].values
        eta = np.sqrt(roundtrip_eff)

        caps = np.arange(cap_min, cap_max_dyn+1, cap_step)
        powers = np.arange(p_min, p_max+1, p_step)
        results = []

        for cap in caps:
            valid_powers = powers[powers <= cap]  # contrainte Power <= Capacity
            for p in valid_powers:
                soc_min_kWh = cap * soc_min_pct / 100

                p_step_val = p * dt_hours
                soc = np.zeros_like(exp_array)
                soc_val = 0
                exp_after = np.zeros_like(exp_array)
                imp_after = np.zeros_like(exp_array)

                charge = np.minimum(exp_array, p_step_val)
                discharge = np.minimum(imp_array, p_step_val)

                soc_list_tmp, imp_after_tmp, exp_after_tmp, charge_series_tmp, discharge_series_tmp, charge_total_tmp, discharge_total_tmp = simulate_battery(
                exp_array, imp_array, cap, p, soc_min_pct, eta, dt_hours)
            
                gain = compute_gain_with_time_of_use(
                imp_array,
                exp_array,
                imp_after_tmp,
                exp_after_tmp,
                hours,
                weekdays
                )

                eq_cycles = eq_cycles_dod(charge_series_tmp, discharge_series_tmp, cap, soc_min_pct)

                results.append([cap, p, gain, eq_cycles])

        results_df = pd.DataFrame(results, columns=["Cap_kWh","Power_kW","Gain_CHF","Cycles"])
        gain_max = results_df["Gain_CHF"].max()
        threshold = gain_threshold * gain_max
        candidates = results_df[results_df["Gain_CHF"] >= threshold]
        if not candidates.empty:
            best = (candidates.sort_values(["Cap_kWh", "Power_kW"], ignore_index=True).iloc[0])
        else:
            best = None
            st.error("❌ Aucune Configuration Trouvée.")
            st.info("⚙️ Revoir les paramètres de capacité et/ou de puissance à la baisse.")
            st.stop()

    st.success(f"🔋 Batterie optimale : {best.Cap_kWh} kWh / {best.Power_kW} kW")
    st.success(f"Gain annuel: {round(best.Gain_CHF,2)} CHF")
    
    # ===========================
    # TRACA
    # ===========================
    
    track_event(
    "battery_calculation_done",
    {
        "data_mode": data_mode,
        "tariff_mode": mode_tarif,
        "grd": GRD if mode_tarif == "HP/HC" else "Tarif unique",
        "year": int(target_year),
        "rows_before": int(LongBase),
        "rows_after": int(LongApres),
        "dt_hours": float(dt_hours),
        "import_total_kwh": float(import_total_kWh),
        "export_total_kwh": float(export_total_kWh),
        "cap_min": float(cap_min),
        "cap_max": float(cap_max),
        "cap_step": float(cap_step),
        "p_min": float(p_min),
        "p_max": float(p_max),
        "p_step": float(p_step),
        "roundtrip_eff": float(roundtrip_eff),
        "soc_min_pct": float(soc_min_pct),
        "gain_threshold": float(gain_threshold),
        "best_capacity_kwh": float(best.Cap_kWh),
        "best_power_kw": float(best.Power_kW),
        "gain_chf": float(best.Gain_CHF),
        "gain_max_chf": float(gain_max),
        "cycles": float(best.Cycles),
        "export_is_monthly": bool(export_is_monthly),
        "import_export_is_monthly": bool(importExport_is_monthly),
    }
    )

    # ===========================
    # Résumé de la batterie réelle
    # ===========================
    soc_list, imp_after, exp_after, charge_series, discharge_series, charge_total, discharge_total = simulate_battery(
    exp_array, imp_array, best.Cap_kWh, best.Power_kW, soc_min_pct, eta, dt_hours)

    soc_min_real = min(soc_list)
    soc_max_real = max(soc_list)
    #charge_total_real = charge_total
    #discharge_total_real = discharge_total
    charge_total_real = np.sum(exp_array - exp_after)      # énergie côté AC pour charger
    discharge_total_real = np.sum(imp_array - imp_after)  # énergie côté AC restituée
    
    if charge_total_real > 0:
        rendement_reel_batterie = discharge_total_real / charge_total_real
    else:
        rendement_reel_batterie = 0

    st.subheader("Résumé réel de la batterie")
    st.write(f"- SOC minimum atteint : {soc_min_real:.2f} kWh ({soc_min_real/best.Cap_kWh*100:.1f}%)")
    st.write(f"- SOC maximum atteint : {soc_max_real:.2f} kWh ({soc_max_real/best.Cap_kWh*100:.1f}%)")
    #st.write(f"- Energie totale chargée : {charge_total_real:.2f} kWh")
    #st.write(f"- Energie totale déchargée : {discharge_total_real:.2f} kWh")
    
    st.write(f"- Energie absorbée pour charge (AC) : {charge_total_real:.2f} kWh")
    st.write(f"- Energie restituée au bâtiment (AC) : {discharge_total_real:.2f} kWh")
    st.write(f"- Rendement réel observé : {rendement_reel_batterie:.3f} ({rendement_reel_batterie*100:.1f}%)")

    # Vérification gain réel

    gain_real = compute_gain_with_time_of_use(
    imp_array,
    exp_array,
    imp_after,
    exp_after,
    hours,
    weekdays
    )
    st.write(f"- Verification Gain réel simulation : {gain_real:.2f} CHF")

    # ==========================================================
    # SOC VECTORISÉ
    # ==========================================================

    soc_list, imp_after, exp_after, charge_series, discharge_series, charge_total, discharge_total = simulate_battery(
    exp_array, imp_array, best.Cap_kWh, best.Power_kW, soc_min_pct, eta, dt_hours)
    df["SOC_pct"] = [(s / best.Cap_kWh)*100 for s in soc_list]

    # ==========================================================
    # PRÉPARATION DES DONNÉES
    # ==========================================================
    df = df.sort_values(by=date_col)
    df[date_col] = pd.to_datetime(df[date_col])
    df = df.set_index(date_col)

    # Nettoyage
    df["import_kWh"] = df["import_kWh"].clip(lower=0).fillna(0)
    df["export_kWh"] = df["export_kWh"].clip(lower=0).fillna(0)

    imp_after = pd.Series(imp_after, index=df.index).clip(lower=0).fillna(0)
    exp_after = pd.Series(exp_after, index=df.index).clip(lower=0).fillna(0)

    st.header(" 🔹 Graphiques")
    
    # ==========================================================
    # CHOIX AGRÉGATION
    # ==========================================================
    aggregation_choice = st.selectbox(
        "📅 Niveau d'agrégation",
        ["Journalier", "Hebdomadaire", "Mensuel"],
        index=2
    )
    
    freq_map = {
    "Journalier": "D",
    "Hebdomadaire": "W-MON",
    "Mensuel": "MS"
    }
    
    freq = freq_map.get(aggregation_choice)
    
    if not isinstance(freq, str) or freq.strip() == "":
        st.error(f"❌ Fréquence invalide: {repr(freq)}")
        st.stop()
    
    # Sécurisation index
    df.index = pd.to_datetime(df.index, errors="coerce")
    df = df[~df.index.isna()]
    df = df.sort_index()

    # ==========================================================
    # AGRÉGATION
    # ==========================================================

    if freq is None:
        st.error(f"❌ Fréquence invalide : {aggregation_choice}")
        st.stop()
    
    if not isinstance(df.index, pd.DatetimeIndex):
        st.error("❌ Index n'est pas un DatetimeIndex")
        st.write(df.index[:5])
        st.stop()

    if df.index.isna().any():
        st.error("❌ Index contient des NaT")
        st.stop()
    
    # Avant optimisation : SOC en %
    before_agg = df[["import_kWh", "export_kWh", "SOC_pct"]].copy()
    before_agg = before_agg.resample(freq).agg({
        "import_kWh": "sum",
        "export_kWh": "sum",
        "SOC_pct": "mean" 
    })

    # Après optimisation : SOC en %
    after_agg = pd.DataFrame({
        "import_after": imp_after,
        "export_after": exp_after,
        "SOC_pct": df["SOC_pct"]  # SOC déjà en %
    }, index=df.index)
    after_agg = after_agg.resample(freq).agg({
        "import_after": "sum",
        "export_after": "sum",
        "SOC_pct": "mean"
    })

    # Détecter l'année automatiquement
    year_min = before_agg.index.min().year
    year_max = before_agg.index.max().year
    
    # ==========================================================
    # GRAPHIQUE SOC MATPLOTLIB 
    # ==========================================================
    st.subheader("📈 SOC (%) - Agrégation")
    fig_soc, ax_soc = plt.subplots(figsize=(12,5))
    
    ax_soc.plot(after_agg.index, after_agg["SOC_pct"], label="SOC après (%)", color='blue')
    
    ax_soc.set_xlabel("Mois")
    ax_soc.set_ylabel("SOC (%)")
    ax_soc.set_title("État de charge batterie")
    ax_soc.set_ylim(0, 100)  # échelle SOC fixe 0-100%    
    ax_soc.legend()
    ax_soc.grid(alpha=0.3)

    # Limiter à l'année complète automatiquement
    ax_soc.set_xlim(pd.Timestamp(f"{year_min}-01-01"), pd.Timestamp(f"{year_max}-12-31"))

    # Tick par mois
    ax_soc.xaxis.set_major_locator(mdates.MonthLocator())
    ax_soc.xaxis.set_major_formatter(mdates.DateFormatter('%b'))
    
    fig_soc.autofmt_xdate()
    st.pyplot(fig_soc)

    # Calcul de l'échelle commune
    max_import_export = max(
        before_agg[["import_kWh","export_kWh"]].max().max(),
        after_agg[["import_after","export_after"]].max().max()
    )

    # ===========================
    # Graphique AVANT
    # ===========================
    st.subheader("📊 Import / Export AVANT optimisation")
    fig_before, ax_before = plt.subplots(figsize=(12,5))

    ax_before.plot(before_agg.index, before_agg["import_kWh"], label="Import avant (kWh)")
    ax_before.plot(before_agg.index, before_agg["export_kWh"], label="Export avant (kWh)")

    ax_before.set_ylabel("Énergie (kWh)")
    ax_before.set_xlabel("Mois")
    ax_before.set_title("Import / Export AVANT optimisation")
    ax_before.set_ylim(0, max_import_export*1.05)
    ax_before.legend()
    ax_before.grid(alpha=0.3)

    # Limiter à l'année complète automatiquement
    ax_before.set_xlim(pd.Timestamp(f"{year_min}-01-01"), pd.Timestamp(f"{year_max}-12-31"))

    # Tick par mois
    ax_before.xaxis.set_major_locator(mdates.MonthLocator())
    ax_before.xaxis.set_major_formatter(mdates.DateFormatter('%b'))

    fig_before.autofmt_xdate()
    st.pyplot(fig_before)

    # ===========================
    # Graphique APRÈS
    # ===========================
    st.subheader("📊 Import / Export APRÈS optimisation")
    fig_after, ax_after = plt.subplots(figsize=(12,5))

    ax_after.plot(after_agg.index, after_agg["import_after"], label="Import après (kWh)")
    ax_after.plot(after_agg.index, after_agg["export_after"], label="Export après (kWh)")

    ax_after.set_ylabel("Énergie (kWh)")
    ax_after.set_xlabel("Mois")
    ax_after.set_title("Import / Export APRÈS optimisation")
    ax_after.set_ylim(0, max_import_export*1.05)
    ax_after.legend()
    ax_after.grid(alpha=0.3)

    # Limiter à l'année complète automatiquement
    ax_after.set_xlim(pd.Timestamp(f"{year_min}-01-01"), pd.Timestamp(f"{year_max}-12-31"))

    # Tick par mois
    ax_after.xaxis.set_major_locator(mdates.MonthLocator())
    ax_after.xaxis.set_major_formatter(mdates.DateFormatter('%b'))

    fig_after.autofmt_xdate()
    st.pyplot(fig_after)

    # ==========================================================
    # Courbe du gain annuel en fonction de la capacité
    # ==========================================================
    st.subheader("📊 Gain annuel vs Capacité batterie")

    # On moyenne le gain sur toutes les puissances pour chaque capacité
    gain_by_cap = results_df.groupby("Cap_kWh")["Gain_CHF"].max()  # ou .mean() si tu veux moyenne

    fig_gain, ax_gain = plt.subplots(figsize=(10,5))
    ax_gain.plot(gain_by_cap.index, gain_by_cap.values, marker='o', color='green')
    ax_gain.set_xlabel("Capacité batterie (kWh)")
    ax_gain.set_ylabel("Gain annuel (CHF)")
    ax_gain.set_title("Gain annuel en fonction de la capacité de la batterie")
    ax_gain.grid(alpha=0.3)
    
    # 👉 ticks tous les 2 kWh
    x_min = int(min(gain_by_cap.index))
    x_max = int(max(gain_by_cap.index))
    ax_gain.set_xticks(np.arange(x_min, x_max + 1, 2))


    # Affiche la figure dans Streamlit
    st.pyplot(fig_gain)


    # ==========================================================
    # CALCULS POUR LE RAPPORT
    # ==========================================================
    # Import/Export avant
    import_before = df["import_kWh"].sum()
    export_before = df["export_kWh"].sum()

    # Simulation SOC pour batterie optimale
    soc_list, imp_after, exp_after, charge_series, discharge_series, charge_total, discharge_total = simulate_battery(
    exp_array, imp_array, best.Cap_kWh, best.Power_kW, soc_min_pct, eta, dt_hours)

    df["SOC"] = soc_list

    # Calcul gains
    gain_net = compute_gain_with_time_of_use(
    imp_array,
    exp_array,
    imp_after,
    exp_after,
    hours,
    weekdays
    )
    
    eq_cycles = eq_cycles_dod(charge_series, discharge_series, best.Cap_kWh, soc_min_pct)
    
    if debug :
        # ==========================================================
        # VERIFICATION RESULTATS
        # ==========================================================
        st.header("🛠 DEBUG & VALIDATION DES DONNÉES")
        # ==========================================================
        # 1️ Vérification tailles des vecteurs
        # ==========================================================
        st.subheader("1️ Vérification dimensions")
        
        st.write("Longueur df :", len(df))
        st.write("Longueur imp_array :", len(imp_array))
        st.write("Longueur exp_array :", len(exp_array))
        st.write("Longueur imp_after :", len(imp_after))
        st.write("Longueur exp_after :", len(exp_after))
        st.write("Longueur soc_list :", len(soc_list))
        
        if not (
            len(df) == len(imp_array) == len(exp_array) == len(imp_after) == len(exp_after) == len(soc_list)
        ):
            st.error("❌ ERREUR : Les vecteurs n'ont pas la même longueur !")
        else:
            st.success("✅ Toutes les longueurs correspondent")
        
        # ==========================================================
        # 2️ Vérification index datetime
        # ==========================================================
        st.subheader("2️ Vérification index temporel")
        
        if not isinstance(df.index, pd.DatetimeIndex):
            st.error("❌ df n'a pas un DatetimeIndex")
        else:
            st.success("✅ df a un DatetimeIndex")
        
        if not df.index.is_monotonic_increasing:
            st.error("❌ Les dates ne sont pas triées")
        else:
            st.success("✅ Les dates sont triées")
        
        if df.index.has_duplicates:
            st.error("❌ Il y a des dates en doublon")
            duplicates = df[df.index.duplicated(keep=False)]
            st.write("Nombre doublons :", len(duplicates))
            st.dataframe(duplicates.head(1000))
        else:
            st.success("✅ Pas de doublons temporels")
        
        # ==========================================================
        # 3️ Vérification cohérence énergie AVANT / APRÈS
        # ==========================================================
        st.subheader("3️ Vérification énergie")
        
        import_before = df["import_kWh"].sum()
        export_before = df["export_kWh"].sum()
        
        import_after_total = imp_after.sum()
        export_after_total = exp_after.sum()
        
        st.write("Import total AVANT :", import_before)
        st.write("Import total APRÈS :", import_after_total)
        st.write("Export total AVANT :", export_before)
        st.write("Export total APRÈS :", export_after_total)
        
        if import_after_total > import_before + 1e-6:
            st.error("❌ Import après > import avant (impossible)")
        else:
            st.success("✅ Import cohérent")
        
        # ==========================================================
        # 4️ Vérification conservation énergétique batterie
        # ==========================================================
        st.subheader("4️ Vérification conservation énergie batterie")
        
        energy_delta_soc = soc_list[-1] - soc_list[0]
        battery_balance = charge_total - discharge_total
        
        st.write("ΔSOC total (kWh) :", energy_delta_soc)
        st.write("Charge - Décharge :", battery_balance)
        
        if abs(energy_delta_soc - battery_balance) > 0.01:
            st.error("❌ Incohérence dans le bilan batterie")
        else:
            st.success("✅ Bilan batterie cohérent")
        
        # ==========================================================
        # 5️ Vérification SOC
        # ==========================================================
        st.subheader("5️ Vérification SOC")
        
        soc_min = min(soc_list)
        soc_max = max(soc_list)
        
        st.write("SOC min (kWh) :", soc_min)
        st.write("SOC max (kWh) :", soc_max)
    
        st.write("Somme charge_series :", charge_series.sum())
        st.write("Somme discharge_series :", discharge_series.sum())
    
        st.write("charge_total :", charge_total)
        st.write("discharge_total :", discharge_total)
    
        # ==========================================================
        # DEBUG CONSERVATION ÉNERGIE BATTERIE
        # ==========================================================
        
        delta_soc = soc_list[-1] - soc_list[0]
        energy_balance = np.sum(charge_series) - np.sum(discharge_series)
        difference = delta_soc - energy_balance
        
        st.subheader("🔍 Debug conservation batterie")
        
        st.write(f"ΔSOC (kWh) : {delta_soc:.6f}")
        st.write(f"Somme(charge_series) - Somme(discharge_series) : {energy_balance:.6f}")
        st.write(f"Différence : {difference:.10f}")
        
        tolerance = 1e-6
        
        if abs(difference) < tolerance:
            st.success("✅ Conservation énergétique OK")
        else:
            st.error("❌ Problème de cohérence énergétique")
        
        if soc_min < 0:
            st.error("❌ SOC négatif")
        if soc_max > best.Cap_kWh + 1e-6:
            st.error("❌ SOC dépasse capacité batterie")
        else:
            st.success("✅ SOC dans les limites physiques")
        
        # ==========================================================
        # 6️ Vérification agrégation
        # ==========================================================
        st.subheader("6️ Vérification agrégation")
        
        sum_before_agg = before_agg["import_kWh"].sum()
        sum_after_agg = after_agg["import_after"].sum()
        
        st.write("Somme import AVANT (agrégé) :", sum_before_agg)
        st.write("Somme import APRÈS (agrégé) :", sum_after_agg)
        
        if abs(sum_before_agg - import_before) > 0.01:
            st.error("❌ Agrégation AVANT incorrecte")
        else:
            st.success("✅ Agrégation AVANT correcte")
        
        if abs(sum_after_agg - import_after_total) > 0.01:
            st.error("❌ Agrégation APRÈS incorrecte")
        else:
            st.success("✅ Agrégation APRÈS correcte")
        
        # ==========================================================
        # 7️ Vérification valeurs négatives
        # ==========================================================
        st.subheader("7️ Vérification valeurs négatives")
        
        if (imp_after < 0).any():
            st.error("❌ Valeurs négatives dans imp_after")
        
        if (exp_after < 0).any():
            st.error("❌ Valeurs négatives dans exp_after")
        
        if (df["import_kWh"] < 0).any():
            st.error("❌ Valeurs négatives dans import_kWh")
        
        if (df["export_kWh"] < 0).any():
            st.error("❌ Valeurs négatives dans export_kWh")
        
        st.success("✅ Aucune valeur négative détectée")

        st.subheader("🔍 DEBUG TARIFICATION HP/HC")

        # Reconstruction masque HP
        is_hp = np.zeros(len(hours), dtype=bool)
    
        for start, end in hp_ranges:
            is_hp |= (hours >= start) & (hours < end)
    
        # Gestion week-end
        if weekend_hc:
            weekdays = timestamps.dt.weekday.to_numpy()
            is_hp[weekdays >= 5] = False
    
        hp_count = np.sum(is_hp)
        hc_count = len(is_hp) - hp_count
    
        st.write("Nombre total points :", len(hours))
        st.write("HP count :", hp_count)
        st.write("HC count :", hc_count)
    
        if hp_count == 0:
            st.error("❌ AUCUNE heure HP détectée → problème plage horaire")

        
        st.subheader("🔍 Distribution des heures")
        unique_hours = np.unique(hours)
        st.write("Heures présentes :", unique_hours)

        st.subheader("🔍 Test sensibilité tarif")
    
        test_gain_hp_plus = np.sum((imp_array - imp_after) * (tariff_importHP + 0.05))
        st.write("Gain si HP + 0.05 CHF :", test_gain_hp_plus)

    
        st.subheader("🔍 Test FULL HP")
    
        gain_full_hp = np.sum((imp_array - imp_after) * tariff_importHP - 
                              (exp_array - exp_after) * tariff_export)
    
        st.write("Gain si 100% HP :", gain_full_hp)
        
        st.success("🎯 DEBUG TERMINÉ")

    st.header(" 🔹 Alertes")
    alerts = []

    # SOC limites physiques
    soc_max = np.max(df["SOC_pct"])
    soc_min = np.min(df["SOC_pct"])

    if soc_max > 100.1:
        alerts.append(f"⚠️ SOC dépasse 100% (max = {soc_max:.2f}%).")

    if soc_min < soc_min_pct - 0.1:
        alerts.append(f"⚠️ SOC descend sous 0% (min = {soc_min:.2f}%).")

    # Gain négatif
    if gain_net < 0:
        alerts.append("⚠️ Le gain net est négatif → batterie non rentable.")

    # Cycles excessifs (ex : > 365/an)
    if eq_cycles > 365:
        alerts.append(f"⚠️ Cycles élevés ({eq_cycles:.1f}/an) → usure importante.")

    # Batterie optimale en limite de plage
    if best.Cap_kWh == cap_min or best.Cap_kWh == cap_max:
        alerts.append("⚠️ Capacité optimale en limite de plage → élargir intervalle.")

    if (best.Power_kW == p_min) or (best.Power_kW == p_max):
        alerts.append("⚠️ Puissance optimale en limite de plage → élargir intervalle.")

    # Capacité dynamique bloquée par cap_max
    #if cap_max_dyn == cap_max:
        #alerts.append("⚠️ Capacité dynamique plafonnée par cap_max → augmenter cap_max pour analyse complète.")

    # Affichage alertes
    if alerts:
        st.warning(" ⚠️ Alertes de vérification")
        for a in alerts:
            st.write(a)
    else:
        st.success("✅ Vérification résultats : aucune anomalie détectée.")
    
    st.header(" 🔹 Export PDF Client")
    # ==========================================================
    # EXPORT PDF CLIENT PROFESSIONNEL
    # ==========================================================
    
    import_avoided = (imp_array - imp_after).sum()
    export_avoided = (exp_array - exp_after).sum()
    
    with st.spinner("Génération PDF client…"):
        pdf = FPDF()
        pdf.set_auto_page_break(auto=True, margin=15)

        # -------------------------
        # PAGE 1 : Couverture
        # -------------------------
        if data_mode == "Fichier GRD (Excel/CSV unique)":
            file_info = uploaded_file.name
        else:
            names = "\n".join([f.name for f in uploaded_file])
            file_info = names

        pdf.add_page()
        pdf.set_font("Arial", 'B', 18)
        pdf.cell(0, 10, "Simulation Batterie Optimisée", ln=True, align="C")
        pdf.set_font("Arial", '', 8)
        pdf.cell(0, 8, "Version : V0 - 03.03.2026 - JMN", ln=True)
        pdf.ln(8)
        pdf.set_font("Arial", '', 12)
        pdf.multi_cell(0, 8,
        f"Date : {now_gmt1.strftime('%d/%m/%Y %H:%M')}\n"
        f"Projet : Simulation client\n"
        f"Résumé : Capacité et puissance optimales, gain estimé et alertes éventuelles\n"
        f"Fichier de données : {file_info}")
        pdf.ln(5)

        # -------------------------
        # AJOUT INFO EXPORT MENSUEL
        # -------------------------
        if data_mode == "Fichier GRD (Excel/CSV unique)" and not importExport_is_monthly and export_is_monthly :
            pdf.set_font("Arial", 'I', 12)
            pdf.multi_cell(0, 8,
                "Profil export reconstitué à partir des totaux mensuels.")
            pdf.ln(5)

            # Liste des mois et valeurs export
            pdf.set_font("Arial", '', 12)
            pdf.multi_cell(0, 8, "Valeurs mensuelles (kWh) :", ln=True)
            for month in ["Janvier", "Février", "Mars", "Avril", "Mai", "Juin",
                          "Juillet", "Août", "Septembre", "Octobre", "Novembre", "Décembre"]:
                value = monthly_export_values.get(month, 0)
                pdf.cell(0, 6, f"{month} : {value:.1f} kWh", ln=True)
            pdf.ln(5)

        
        pdf.set_font("Arial", 'B', 12)
        pdf.cell(0, 8, "Paramètres de simulation - Non Commercial", ln=True)
        pdf.ln(5)
        pdf.set_font("Arial", '', 10)

        # -----------------------
        # TARIFS :     tariff_importHP, tariff_importHC, tariff_export
        # HORAIRES :   hp_ranges.append((start, end))
        # -----------------------
        # Transformation des plages horaires en texte
        if hp_ranges:  # s'il y a des plages HP définies
            hp_ranges_str = ", ".join([f"{start:02d}h-{end:02d}h" for start, end in hp_ranges])
        else:
            hp_ranges_str = "N/A"
            
        param_table = [
            ["Pas de temps (h)", dt_hours],
            ["Reconstruction du profil annuel Import/Export par totaux mensuels.", importExport_is_monthly],
            ["Reconstruction du profil annuel Export par totaux mensuels.", export_is_monthly],
            ["Unité des valeurs", unite],
            ["Rendement aller-retour", roundtrip_eff],
            ["SOC min (%)", soc_min_pct],
            ["Capacité min (kWh)", cap_min],
            ["Capacité max (kWh)", cap_max],
            ["Pas capacité (kWh)", cap_step],
            ["Puissance min (kW)", p_min],
            ["Puissance max (kW)", p_max],
            ["Pas puissance (kW)", p_step],
            ["Tarif import HP (CHF/kWh)", tariff_importHP],
            ["Tarif import HC (CHF/kWh)", tariff_importHC],
            ["Tarif export (CHF/kWh)", tariff_export],
            ["Week-end entièrement en HC", weekend_hc],
            ["Plages horaires HP", hp_ranges_str],
            ["Percentile export journalier", daily_percentile],
        ]

        # Dessiner tableau
        for row in param_table:
            pdf.cell(120, 6, str(row[0]), border=1)
            pdf.cell(40, 6, str(row[1]), border=1, ln=True)

        # -------------------------
        # PAGE 2 : 
        # -------------------------
        pdf.add_page()
        pdf.set_font("Arial", 'B', 12)
        pdf.cell(0, 8, "Résultats annuels estimés - Non Commercial", ln=True)
        pdf.ln(5)
        pdf.set_font("Arial", '', 10)

        results_table = [
            ["Capacité optimale (kWh)", best.Cap_kWh],
            ["Puissance optimale (kW)", best.Power_kW],
            ["Gain annuel net (CHF)", round(gain_net, 2)],
            ["Gain maximum (CHF)", round(gain_max, 2)],
            ["Seuil choisi (%)", gain_threshold*100],
            ["Cycles équivalents/an", round(eq_cycles, 2)],
            ["SOC min réel (%)", round(soc_min_real/best.Cap_kWh*100, 1)],
            ["SOC max réel (%)", round(soc_max_real/best.Cap_kWh*100, 1)],
            ["Énergie totale chargée (kWh)", round(charge_total_real, 2)],
            ["Énergie totale déchargée (kWh)", round(discharge_total_real, 2)],
            ["Import avant (kWh)", round(import_before, 2)],
            ["Import après (kWh)", round(imp_after.sum(), 2)],
            ["Export avant (kWh)", round(export_before, 2)],
            ["Export après (kWh)", round(exp_after.sum(), 2)],
            ["Import évité (kWh)", round(import_avoided, 2)],
            ["Export évité (kWh)", round(export_avoided, 2)],
            ["Capacité max dynamique (kWh)", cap_max_dyn]
        ]
        for row in results_table:
            pdf.cell(120, 6, str(row[0]), border=1)
            pdf.cell(40, 6, str(row[1]), border=1, ln=True)
        
        pdf.ln(5)
        # Courbe du gain annuel en fonction de la capacité
        img_gainAnn = BytesIO()
        fig_gain.savefig(img_gainAnn, format="png")
        img_gainAnn.seek(0)
        pdf.image(img_gainAnn, x=15, w=180)


        # -------------------------
        # PAGE 4 : Résultats annuels - COMMERCIAL
        # -------------------------
        pdf.add_page()
        pdf.set_font("Arial", 'B', 12)
        pdf.cell(0, 8, "Paramètres de simulation", ln=True)
        pdf.ln(5)
        pdf.set_font("Arial", '', 10)

        param_table = [
            ["Rendement aller-retour", roundtrip_eff],
            ["SOC min (%)", soc_min_pct],
            ["Capacité min (kWh)", cap_min],
            ["Capacité max (kWh)", cap_max],
            ["Puissance min (kW)", p_min],
            ["Puissance max (kW)", p_max],
            ["Tarif import HP (CHF/kWh)", tariff_importHP],
            ["Tarif import HC (CHF/kWh)", tariff_importHC],
            ["Tarif export (CHF/kWh)", tariff_export],
            ["Week-end entièrement en HC", weekend_hc],
            ["Plages horaires HP", hp_ranges_str],
        ]
        # Dessiner tableau
        for row in param_table:
            pdf.cell(120, 6, str(row[0]), border=1)
            pdf.cell(40, 6, str(row[1]), border=1, ln=True)
        pdf.ln(5)
        pdf.set_font("Arial", 'B', 12)
        pdf.cell(0, 8, "Résultats annuels estimés", ln=True)
        pdf.set_font("Arial", '', 10)
        pdf.ln(5)
        results_table = [
            ["Capacité optimale (kWh)", best.Cap_kWh],
            ["Puissance optimale (kW)", best.Power_kW],
            ["Seuil choisi (%)", gain_threshold*100],
            ["Cycles équivalents/an", round(eq_cycles, 2)],
            ["Import avant (kWh)", round(import_before, 2)],
            ["Import après (kWh)", round(imp_after.sum(), 2)],
            ["Export avant (kWh)", round(export_before, 2)],
            ["Export après (kWh)", round(exp_after.sum(), 2)],
            ["Import évité (kWh)", round(import_avoided, 2)],
            ["Export évité (kWh)", round(export_avoided, 2)],
        ]
        # Dessiner tableau
        for row in results_table:
            pdf.cell(120, 6, str(row[0]), border=1)
            pdf.cell(40, 6, str(row[1]), border=1, ln=True)

        # =========================
        # GLOSSAIRE / DEFINITIONS
        # =========================
        pdf.ln(10)
        pdf.set_font("Arial", 'B', 12)
        pdf.cell(0, 8, "Glossaire / Définitions", ln=True)
        pdf.ln(5)

        definitions = {
            "SOC (%)": "State of Charge - niveau de charge de la batterie en pourcentage de sa capacité totale.",
            "Capacité (kWh)": "Quantité maximale d'énergie que la batterie peut stocker.",
            "Puissance (kW)": "Vitesse maximale à laquelle la batterie peut charger ou décharger.",
            "Cycles équivalents": "Nombre de cycles complets de charge/décharge que la batterie réalise sur l'année.",
            "Import/Export (kWh)": "Énergie achetée (import) ou revendue (export) au réseau électrique.",
            "Rendement aller-retour": "Pourcentage de l'énergie restituée par la batterie après chargement et déchargement.",
            "Seuil choisi en % du gain max": "Pourcentage du gain maximal utilisé pour sélectionner la capacité et puissance optimales.",
        }

        pdf.set_font("Arial", '', 10)
        cell_width = 180  # largeur utile

        for term, definition in definitions.items():
            # terme en gras au début de la ligne, suivi de la définition en texte normal
            text = f"{term}: {definition}"
            pdf.set_font("Arial", '', 8)
            pdf.multi_cell(cell_width, 6, text)
            pdf.ln(2)  # espace entre définitions

        # -------------------------
        # PAGE 4 : Graphiques
        # -------------------------
        # Graph SOC
        pdf.add_page()
        pdf.set_font("Arial", 'B', 14)
        pdf.cell(0, 8, "Graphiques", ln=True)
        pdf.ln(2)

        # -------------------------
        # Graph SOC
        # -------------------------
        img_buf_soc = BytesIO()
        fig_soc.savefig(img_buf_soc, format="png")
        img_buf_soc.seek(0)
        pdf.image(img_buf_soc, x=15, w=180)
        pdf.ln(2)

        # Import / Export avant
        img_buf_before = BytesIO()
        fig_before.savefig(img_buf_before, format="png")
        img_buf_before.seek(0)
        pdf.image(img_buf_before, x=15, w=180)
        pdf.ln(2)

        # Import / Export après
        img_buf_after = BytesIO()
        fig_after.savefig(img_buf_after, format="png")
        img_buf_after.seek(0)
        pdf.image(img_buf_after, x=15, w=180)

        # -------------------------
        # EXPORT PDF Streamlit
        # -------------------------
        pdf_buffer = BytesIO()
        pdf.output(pdf_buffer)
        pdf_buffer.seek(0)
        st.download_button(
            "📥 Télécharger PDF client",
            pdf_buffer,
            file_name="simulation_batterie_client.pdf",
            mime="application/pdf"
        )

    st.success("PDF client généré ! ✅")
