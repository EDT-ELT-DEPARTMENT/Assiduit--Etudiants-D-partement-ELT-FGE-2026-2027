import streamlit as st
import pandas as pd
import os, random, io
from datetime import datetime
from supabase import create_client

# --- TITRE OFFICIEL OBLIGATOIRE ---
TITLE_PLATFORM = "Plateforme de gestion des EDTs-S2-2026-Département d'Électrotechnique-Faculté de génie électrique-UDL-SBA"

st.set_page_config(page_title="Gestion EDT ELT 2026", layout="wide")

# Affichage du titre réglementaire
st.title(TITLE_PLATFORM)

# Masquer les éléments du menu supérieur (Share, Star, Edit, etc.)
hide_st_style = """
            <style>
            #MainMenu {visibility: hidden;}
            header {visibility: hidden;}
            footer {visibility: hidden;}
            .stAppDeployButton {display:none;}
            #stDecoration {display:none;}
            </style>
            """
st.markdown(hide_st_style, unsafe_allow_html=True)

# ======================================================================================
# 1. CONFIGURATION & CONSTANTES GLOBALES
# ======================================================================================
NOM_SOURCE = "dataEDT-ELT-S2-2026.xlsx"
FILE_EMAILS = "Permanents-Vacataires-ELT2-2025-2026.xlsx"
TABLE_NAME = "surveillances_2026"

# Configuration Supabase
S_URL = "https://ajcbkidmcjtyomknijwa.supabase.co"
S_KEY = "sb_publishable_otn3XM8LPLV0OGw74LRhDw_F446jkpw"
supabase = create_client(S_URL, S_KEY)

# Ordre d'affichage réglementaire des colonnes
COLS_ORDRE = ['Enseignements', 'Code', 'Enseignants', 'Horaire', 'Jours', 'Lieu', 'Promotion']

# Variables de noms de fichiers spécifiques à T6
FILE_DATA_A = "DATA-ASSUIDUITE-2026.xlsx"
FILE_LISTE_A = "Liste des étudiants-2025-2026.xlsx"

DATA_AUTO = {
    "ING1": {"Effectif": 59, "Horaire": "08h00 – 10h00", "Salles": ["A12 (Promotion Entière)"]},
    "ING2": {"Effectif": 21, "Horaire": "10h00 – 12h00", "Salles": ["S08 (Promotion Entière)"]},
    "ING3EI": {"Effectif": 40, "Horaire": "08h00 – 10h00", "Salles": ["S04 (Promotion Entière)"]},
    "ING3RSE": {"Effectif": 20, "Horaire": "08h00 – 10h00", "Salles": ["S04 (Promotion Entière)"]},
    "ING4": {"Effectif": 15, "Horaire": "08h00 – 10h00", "Salles": ["S14 (Promotion Entière)"]},
    "L2ELT": {"Effectif": 94, "Horaire": "10h00 – 12h00", "Salles": ["A10 (G1, G2, G3)"]},
    "L3ELT": {"Effectif": 70, "Horaire": "08h00 – 10h00", "Salles": ["A10 (G1, G2, G3)"]},
    "L1MCIL": {"Effectif": 274, "Horaire": "08h00 – 10h00", "Salles": ["A01 (G1)", "A02 (G2)","A03 (G3)", "A05 (G4)","A08 (G5)"]},
    "L2MCIL": {"Effectif": 113, "Horaire": "10h00 – 12h00", "Salles": ["A04 (G1)", "A09 (G2)"]},
    "MCIL3": {"Effectif": 26, "Horaire": "08h00 – 10h00", "Salles": ["SN (Promotion Entière)"]},
    "M1CE": {"Effectif": 20, "Horaire": "08h00 – 10h00", "Salles": ["S01 (Promotion Entière)"]},
    "M1ER": {"Effectif": 19, "Horaire": "08h00 – 10h00", "Salles": ["S01 (Promotion Entière)"]},
    "M1ME": {"Effectif": 20, "Horaire": "08h00 – 10h00", "Salles": ["S01 (Promotion Entière)"]},
    "M1MCIL": {"Effectif": 38, "Horaire": "08h00 – 10h00", "Salles": ["S01 (Promotion Entière)"]},
    "M1RE": {"Effectif": 16, "Horaire": "08h00 – 10h00", "Salles": ["S01 (Promotion Entière)"]}
}

# ======================================================================================
# 2. FONCTIONS TECHNIQUES (CONSERVÉES POUR T6 & T7)
# ======================================================================================
def charger_donnees_locales(path):
    """Charge un fichier Excel localement et nettoie les colonnes (Utilisé pour T6)."""
    if os.path.exists(path):
        try:
            df = pd.read_excel(path)
            df.columns = [str(c).strip() for c in df.columns]
            return df
        except Exception as e:
            st.error(f"Erreur de lecture du fichier {path} : {e}")
            return pd.DataFrame()
    return pd.DataFrame()

def get_db():
    """Récupère les données de surveillance depuis Supabase (Utilisé pour T7)."""
    try:
        res = supabase.table(TABLE_NAME).select("*").execute()
        return pd.DataFrame(res.data) if res.data else pd.DataFrame(columns=COLS_ORDRE)
    except: 
        return pd.DataFrame(columns=COLS_ORDRE)

def generer_excel_bytes(df):
    """Génère le fichier Excel global téléchargeable en mémoire (Utilisé pour T7)."""
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Planning')
    return output.getvalue()

@st.cache_data
def charger_fichiers():
    """Charge la base de contacts et l'EDT source pour la cartographie des noms complets."""
    df_s = pd.DataFrame()
    map_nom_complet = {}
    d_em = {}

    if os.path.exists(FILE_EMAILS):
        try:
            df_c = pd.read_excel(FILE_EMAILS)
            df_c.columns = [str(c).strip().upper() for c in df_c.columns]
            for _, row in df_c.iterrows():
                n = str(row.get('NOM', '')).strip().upper()
                p = str(row.get('PRÉNOM', '')).strip().upper()
                m_val = row.get('EMAIL') if 'EMAIL' in df_c.columns else row.get('Email')
                m = str(m_val).strip().lower() if pd.notna(m_val) else ""
                
                if n and n != "NAN":
                    nom_complet = f"{n} {p}".strip()
                    map_nom_complet[n] = nom_complet
                    if "@" in m:
                        d_em[n] = m
                        d_em[nom_complet] = m
        except Exception as e:
            st.error(f"Erreur lors de la lecture du fichier {FILE_EMAILS} : {e}")

    if os.path.exists(NOM_SOURCE):
        try:
            df_f = pd.read_excel(NOM_SOURCE)
            df_f.columns = [str(c).strip() for c in df_f.columns]
            mask = df_f["Enseignements"].str.contains("Cours", case=False, na=False)
            df_s = df_f[mask].copy()
            for c in COLS_ORDRE:
                if c not in df_s.columns: df_s[c] = ""
            df_s = df_s[COLS_ORDRE]
        except Exception as e:
            st.error(f"Erreur source EDT : {e}")
            
    return df_s, map_nom_complet, d_em

# Initialisation des dictionnaires de correspondance de noms
_, map_noms, _ = charger_fichiers()
df_db_global = get_db()

# ======================================================================================
# 3. STRUCTURE DE L'INTERFACE UTILISATEUR (T6 & T7 ONLY)
# ======================================================================================
tab6, tab7 = st.tabs(["📊 T6 : Gestion de l'Assiduité", "📥 T7 : Téléchargements & EDTs"])

# --------------------------------------------------------------------------------------
# ONGLET T6 : GESTION DE L'ASSIDUITÉ
# --------------------------------------------------------------------------------------
with tab6:
    st.header("📊 T6 : Suivi et Gestion de l'Assiduité des Étudiants")
    st.info("Cet espace gère le chargement et le traitement des listes d'assiduité.")
    
    col_t6_1, col_t6_2 = st.columns(2)
    
    with col_t6_1:
        st.subheader("Fichier d'assiduité générale")
        if os.path.exists(FILE_DATA_A):
            df_assiduite = charger_donnees_locales(FILE_DATA_A)
            st.success(f"Fichier détecté : `{FILE_DATA_A}` ({len(df_assiduite)} lignes)")
            st.dataframe(df_assiduite.head(10), use_container_width=True)
        else:
            st.warning(f"Le fichier `{FILE_DATA_A}` est manquant à la racine.")
            
    with col_t6_2:
        st.subheader("Liste officielle des étudiants")
        if os.path.exists(FILE_LISTE_A):
            df_liste_etudiants = charger_donnees_locales(FILE_LISTE_A)
            st.success(f"Fichier détecté : `{FILE_LISTE_A}` ({len(df_liste_etudiants)} étudiants)")
            st.dataframe(df_liste_etudiants.head(10), use_container_width=True)
        else:
            st.warning(f"Le fichier `{FILE_LISTE_A}` est manquant à la racine.")

# --------------------------------------------------------------------------------------
# ONGLET T7 : TÉLÉCHARGEMENTS & EXPORTS EDTs
# --------------------------------------------------------------------------------------
with tab7:
    st.header("📥 T7 : Téléchargement des Documents & EDTs de Surveillance")
    
    if df_db_global.empty:
        st.warning("Aucune donnée de surveillance n'est actuellement disponible dans la base de données.")
    else:
        st.subheader("1. Documents Globaux du Département")
        col_g1, col_g2 = st.columns(2)
        
        with col_g1:
            st.download_button(
                label="📊 Télécharger le Planning Global (Excel)", 
                data=generer_excel_bytes(df_db_global[COLS_ORDRE]), 
                file_name="Planning_Global_S2_2026.xlsx", 
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
                key="dl_xlsx_global"
            )
            
        with col_g2:
            cols_global = COLS_ORDRE + ["Responsable"] if "Responsable" not in COLS_ORDRE else COLS_ORDRE
            html_g = f"""
            <html>
            <head><meta charset='UTF-8'></head>
            <body>
                <h2 style='text-align:center;'>{TITLE_PLATFORM}</h2>
                {df_db_global[cols_global].to_html(index=False)}
            </body>
            </html>
            """
            st.download_button(
                label="🌐 Télécharger le Planning Global (HTML)", 
                data=html_g, 
                file_name="Planning_Global_S2_2026.html", 
                mime="text/html", 
                use_container_width=True,
                key="dl_html_global"
            )
            
        st.divider()
        
        # 2. Section d'exportation par promotion spécifique (Code corrigé et complété)
        st.subheader("📅 2. Extraire et Télécharger les EDTs de surveillance par Promotion")
        
        promo_list = ["TOUTES LES PROMOTIONS (ALL)"] + sorted(list(DATA_AUTO.keys()))
        choice_promo = st.selectbox("Sélectionnez la promotion à exporter :", promo_list, key="sel_promo_tab7")
        
        if choice_promo == "TOUTES LES PROMOTIONS (ALL)":
            df_to_export = df_db_global.copy()
            suffix = "ALL"
        else:
            df_to_export = df_db_global[df_db_global["Promotion"].str.contains(choice_promo, na=False)].copy()
            suffix = choice_promo
            
        if not df_to_export.empty:
            # Tri chronologique des données pour le document final
            df_to_export = df_to_export.sort_values(by=["Jours", "Horaire"])
            df_display = df_to_export.copy()
            
            # Application du nom complet si disponible pour le Responsable de matière
            if "Responsable" in df_display.columns:
                df_display["Responsable"] = df_display["Responsable"].apply(
                    lambda x: map_noms.get(str(x).strip().upper(), x)
                )
            else:
                df_display["Responsable"] = "N/A"
            
            # Sélection et réorganisation finale des colonnes pour l'affichage de l'EDT
            cols_finales = ["Jours", "Horaire", "Promotion", "Enseignements", "Lieu", "Enseignants", "Responsable"]
            df_display = df_display[cols_finales].rename(columns={
                "Jours": "Date",
                "Enseignants": "Surveillants",
                "Responsable": "Chargé de Matière"
            })
            
            # Aperçu du tableau extrait
            st.write(f"**Aperçu de l'EDT extrait pour : {suffix}** ({len(df_display)} séances)")
            st.dataframe(df_display, use_container_width=True)
            
            # Génération du code HTML stylisé final (complétion de la section tronquée)
            html_surv = f"""
            <html>
            <head>
            <meta charset="UTF-8">
            <style>
                body {{ font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; margin: 25px; }}
                h2 {{ text-align: center; color: #2c3e50; text-transform: uppercase; font-size: 16px; border-bottom: 2px solid #2c3e50; padding-bottom: 10px; }}
                h3 {{ text-align: center; color: #1E88E5; margin-top: 10px; font-size: 14px; }}
                table {{ width: 100%; border-collapse: collapse; margin-top: 20px; }}
                th {{ background-color: #f8f9fa; color: #333; border: 1px solid #000; padding: 10px; font-size: 12px; font-weight: bold; text-align: left; }}
                td {{ border: 1px solid #000; padding: 8px; text-align: left; font-size: 11px; vertical-align: middle; }}
                tr:nth-child(even) {{ background-color: #f2f2f2; }}
                .footer {{ font-size: 10px; margin-top: 30px; text-align: right; color: #666; font-style: italic; }}
            </style>
            </head>
            <body>
                <h2>{TITLE_PLATFORM}</h2>
                <h3>PLANNING DE SURVEILLANCE - PROMOTION : {suffix}</h3>
                {df_display.to_html(index=False, escape=False)}
                <div class="footer">Document généré automatiquement via la plateforme le {datetime.now().strftime('%d/%m/%Y à %H:%M')}</div>
            </body>
            </html>
            """
            
            # Bouton de téléchargement du fichier HTML généré pour la promotion sélectionnée
            st.download_button(
                label=f"📅 Télécharger l'EDT Officiel ({suffix}) au format HTML",
                data=html_surv,
                file_name=f"EDT_Surveillance_{suffix}_S2_2026.html",
                mime="text/html",
                use_container_width=True,
                key="dl_html_promo_final"
            )
        else:
            st.info(f"Aucune séance de surveillance répertoriée pour la promotion : {choice_promo}")
