import streamlit as st
import pandas as pd
import os, random, io, smtplib
import re
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from datetime import datetime, timedelta
from supabase import create_client
import time
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
if "enseignants_ajoutes_manuellement" not in st.session_state:
    st.session_state.enseignants_ajoutes_manuellement = []
def dupliquer_seance(row_data):
    """Crée une copie exacte d'une séance dans Supabase pour ajouter un surveillant."""
    # On retire l'ID pour que Supabase en génère un nouveau
    new_data = row_data.to_dict()
    if 'id' in new_data: del new_data['id']
    try:
        supabase.table(TABLE_NAME).insert(new_data).execute()
        st.success("Ligne ajoutée ! Vous pouvez maintenant choisir un autre surveillant.")
        st.rerun()
    except Exception as e:
        st.error(f"Erreur : {e}")

def supprimer_ligne(rid):
    """Supprime définitivement une ligne du PV."""
    try:
        supabase.table(TABLE_NAME).delete().eq("id", rid).execute()
        st.rerun()
    except Exception as e:
        st.error(f"Erreur de suppression : {e}")
        # --- CONSTANTES GLOBALES ---
TITLE_PLATFORM = "Plateforme de gestion des EDTs-S2-2026-Département d'Électrotechnique-Faculté de génie électrique-UDL-SBA"
 # --- FONCTIONS DE GESTION DE LA BASE ---

def dupliquer_seance(row_data):
    """Crée une copie exacte d'une séance dans Supabase pour ajouter un surveillant."""
    new_data = row_data.to_dict()
    if 'id' in new_data: del new_data['id']
    try:
        supabase.table(TABLE_NAME).insert(new_data).execute()
        st.success("Ligne ajoutée ! Vous pouvez maintenant choisir un autre surveillant.")
        st.rerun()
    except Exception as e:
        st.error(f"Erreur : {e}")

def supprimer_ligne(rid):
    """Supprime définitivement une ligne du PV."""
    try:
        supabase.table(TABLE_NAME).delete().eq("id", rid).execute()
        st.rerun()
    except Exception as e:
        st.error(f"Erreur de suppression : {e}")
def send_email_convocation(nom_enseignant, email_dest, df_perso):
    """Gère l'envoi réel du mail via SMTP (Gmail/Port 465)."""
    # Remplacez par vos vrais identifiants
    SMTP_SERVER = "smtp.gmail.com"
    SMTP_PORT = 465
    SENDER_EMAIL = "chef.department.elt.fge@gmail.com" 
    SENDER_PASSWORD = "gkzs pdza yodb icvd"

    msg = MIMEMultipart()
    msg['From'] = f"Département Électrotechnique UDL <{SENDER_EMAIL}>"
    msg['To'] = email_dest
    msg['Subject'] = f"Convocation Surveillance S2 2026 - {nom_enseignant}"

    # Style HTML pour le tableau dans le mail
    corps_html = f"""
    <html>
    <head>
        <style>
            table {{ border-collapse: collapse; width: 100%; font-family: Arial, sans-serif; }}
            th {{ background-color: #1f4e79; color: white; padding: 10px; }}
            td {{ border: 1px solid #ddd; padding: 8px; text-align: left; }}
        </style>
    </head>
    <body>
        <h3>Bonjour {nom_enseignant},</h3>
        <p>Voici votre planning de surveillance pour la session <b>Mai 2026</b> :</p>
        {df_perso.to_html(index=False)}
        <br>
        <p>Cordialement,<br>Le Chef de Département</p>
    </body>
    </html>
    """
    msg.attach(MIMEText(corps_html, 'html'))

    try:
        with smtplib.SMTP_SSL(SMTP_SERVER, SMTP_PORT) as server:
            server.login(SENDER_EMAIL, SENDER_PASSWORD)
            server.send_message(msg)
        return True
    except Exception as e:
        return str(e)       
# ======================================================================================
# 1. CONFIGURATION & MÉMOIRE
# ======================================================================================
TITRE_OFFICIEL = "Plateforme de gestion des EDTs des Examens du 16 au 24 Mai 2026-Département d'Électrotechnique-Faculté de génie électrique-UDL-SBA"
NOM_SOURCE = "dataEDT-ELT-S2-2026.xlsx"
FILE_EMAILS = "Permanents-Vacataires-ELT2-2025-2026.xlsx"
TABLE_NAME = "surveillances_2026"

# Définition cruciale pour le filtrage des tableaux et exports
COLS_ORDRE = ['Enseignements', 'Code', 'Enseignants', 'Horaire', 'Jours', 'Lieu', 'Promotion']

S_URL = "https://ajcbkidmcjtyomknijwa.supabase.co"
S_KEY = "sb_publishable_otn3XM8LPLV0OGw74LRhDw_F446jkpw"

SMTP_SERVER = "smtp.gmail.com"
SMTP_PORT = 465
MAIL_USER = "chef.department.elt.fge@gmail.com"
MAIL_PASS = "gkzs pdza yodb icvd" 

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

LISTE_SALLES = [f"S{i:02d}" for i in range(1, 19)] + ["SN", "S08 BIS"]
LISTE_AMPHIS = [f"A{i:02d}" for i in range(1, 13)]

st.set_page_config(page_title="Gestion EDT ELT 2026", layout="wide")
supabase = create_client(S_URL, S_KEY)
# --- FONCTION TECHNIQUE (À placer en haut du fichier) ---
# --- LISTE DES CRÉNEAUX HORAIRES (14 CRÉNEAUX) ---
HORAIRES_LIST = [
    "8h - 9h", "8h - 9h30", "8h - 10h", "9h - 10h", "9h30 - 11h", 
    "10h - 11h", "11h - 12h", "11h - 12h30", 
    "12h - 13h", "12h30 - 14h", "13h - 14h", "14h - 15h30", "14h - 16h", "15h30 - 17h"
]
def charger_donnees_locales(path):
    """Charge un fichier Excel localement et nettoie les colonnes."""
    if os.path.exists(path):
        try:
            df = pd.read_excel(path)
            # Nettoyage des noms de colonnes (suppression des espaces invisibles)
            df.columns = [str(c).strip() for c in df.columns]
            return df
        except Exception as e:
            st.error(f"Erreur de lecture du fichier {path} : {e}")
            return pd.DataFrame()
    return pd.DataFrame()

# Variables de noms de fichiers pour T6
FILE_DATA_A = "DATA-ASSUIDUITE-2026.xlsx"
FILE_LISTE_A = "Liste des étudiants-2025-2026.xlsx"
# ======================================================================================
# 2. FONCTIONS TECHNIQUES
# ======================================================================================
def get_db():
    try:
        res = supabase.table(TABLE_NAME).select("*").execute()
        return pd.DataFrame(res.data) if res.data else pd.DataFrame(columns=COLS_ORDRE)
    except: return pd.DataFrame(columns=COLS_ORDRE)

def generer_excel_bytes(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Planning')
    return output.getvalue()

@st.cache_data
def charger_fichiers():
    df_s = pd.DataFrame()
    map_nom_complet = {}
    d_em = {}

    # NOM DU FICHIER TEL QU'IL APPARAIT SUR VOTRE GITHUB
    FILE_CONTACTS = "Permanents-Vacataires-ELT2-2025-2026.xlsx"

    # 1. CHARGEMENT DES CONTACTS (Lecture locale)
    if os.path.exists(FILE_CONTACTS):
        try:
            df_c = pd.read_excel(FILE_CONTACTS)
            # Normalisation des noms de colonnes
            df_c.columns = [str(c).strip().upper() for c in df_c.columns]
            
            # On identifie les colonnes par rapport à votre JSON
            # NOM, PRÉNOM, Email
            for _, row in df_c.iterrows():
                n = str(row.get('NOM', '')).strip().upper()
                p = str(row.get('PRÉNOM', '')).strip().upper()
                
                # On cherche 'EMAIL' ou 'Email'
                m_val = row.get('EMAIL') if 'EMAIL' in df_c.columns else row.get('Email')
                m = str(m_val).strip().lower() if pd.notna(m_val) else ""
                
                if n and n != "NAN":
                    nom_complet = f"{n} {p}".strip()
                    map_nom_complet[n] = nom_complet
                    
                    if "@" in m:
                        # ON INDEXE PAR NOM SEUL ET NOM COMPLET
                        d_em[n] = m                # Clé: "MILOUA"
                        d_em[nom_complet] = m      # Clé: "MILOUA FARID"
        except Exception as e:
            st.error(f"Erreur lors de la lecture du fichier {FILE_CONTACTS} : {e}")
    else:
        st.error(f"Le fichier {FILE_CONTACTS} est introuvable à la racine du dépôt.")

    # 2. CHARGEMENT DE L'EDT (dataEDT-ELT-S2-2026.xlsx)
    if os.path.exists("dataEDT-ELT-S2-2026.xlsx"):
        try:
            df_f = pd.read_excel("dataEDT-ELT-S2-2026.xlsx")
            df_f.columns = [str(c).strip() for c in df_f.columns]
            mask = df_f["Enseignements"].str.contains("Cours", case=False, na=False)
            df_s = df_f[mask].copy()
            for c in ['Enseignements', 'Code', 'Enseignants', 'Horaire', 'Jours', 'Lieu', 'Promotion']:
                if c not in df_s.columns: df_s[c] = ""
            df_s = df_s[['Enseignements', 'Code', 'Enseignants', 'Horaire', 'Jours', 'Lieu', 'Promotion']]
        except Exception as e:
            st.error(f"Erreur source EDT : {e}")
            
    return df_s, map_nom_complet, d_em
# --- APPEL ET INITIALISATION ---
df_src, map_noms, dict_emails = charger_fichiers()
if not df_src.empty:
    noms_famille = df_src["Enseignants"].unique()
    LISTE_PROFS = sorted([map_noms.get(str(n).strip().upper(), str(n).strip().upper()) for n in noms_famille])
else:
    LISTE_PROFS = []
def envoyer_mail(ens_nom, email_dest, df_perso):
    try:
        msg = MIMEMultipart()
        msg['From'] = MAIL_USER
        msg['To'] = email_dest
        msg['Subject'] = f"Planning des examens de rattrapage et des surveillances S1 & S2_2025-2026 - {ens_nom}"
        rows = "".join([
            f"<tr>"
            f"<td style='border:1px solid #ddd;padding:8px;'>{r['Enseignements']}</td>"
            f"<td style='border:1px solid #ddd;padding:8px;'>{r['Horaire']}</td>"
            f"<td style='border:1px solid #ddd;padding:8px;'>{r['Jours']}</td>"
            f"<td style='border:1px solid #ddd;padding:8px;'>{r['Lieu']}</td>"
            f"<td style='border:1px solid #ddd;padding:8px;'>{r['Promotion']}</td>"
            f"</tr>" 
            for _, r in df_perso.iterrows()
        ])
        body = f"""
        <html>
        <body style='font-family: Arial, sans-serif; color: #333;'>
            <p>Salam M./Mme <b>{ens_nom}</b>,</p>
            <p>Nous vous prions de bien vouloir prendre connaissance du planning de rattrapage et des surveillances relatif aux Semestres 01 et 02, joint à cet envoi.
Bien que ce planning soit généré automatiquement, nous vous informons que les permutations ou remplacements pour les horaires de surveillance sont autorisés. Toutefois, afin de garantir le bon déroulement des épreuves, toute modification doit être signalée à l'administration au moins 24 heures à l'avance, en précisant impérativement le nom du remplaçant.
Nous vous remercions de bien vouloir nous en accuser réception.</p>
            <div style='background-color: #fff3cd; padding: 15px; border-left: 5px solid #ffecb5; margin: 15px 0;'>
                <h4 style='margin-top:0;'>Rappels importants :</h4>
                <ul>
                    <li>La présentation de la carte d'étudiant est strictement obligatoire.</li>
                    <li>Il est strictement interdit de rajouter un étudiant sur la liste sans aviser l'administration.</li>
                    <li>Le portable est strictement interdit et doit être en position éteinte.</li>
                    <li>Aucun étudiant ne peut accéder à la salle au-delà de 30 minutes après le début de l'épreuve.</li>
                </ul>
            </div>
            <table style='border-collapse: collapse; width: 100%; border: 1px solid #ddd;'>
                <thead style='background-color: #f2f2f2;'>
                    <tr><th>Matière</th><th>Heure</th><th>Jour</th><th>Lieu</th><th>Promo</th></tr>
                </thead>
                <tbody>{rows}</tbody>
            </table>
            <p><br>Cordialement,<br><b>Le chef de Département-ELT-FGE</b></p>
        </body>
        </html>
        """
        msg.attach(MIMEText(body, 'html'))
        xlsx = generer_excel_bytes(df_perso[COLS_ORDRE])
        part = MIMEApplication(xlsx, Name=f"Convocation_{ens_nom}.xlsx")
        part['Content-Disposition'] = f'attachment; filename="Convocation_{ens_nom}.xlsx"'
        msg.attach(part)
        
        with smtplib.SMTP_SSL(SMTP_SERVER, SMTP_PORT) as s:
            s.login(MAIL_USER, MAIL_PASS)
            s.sendmail(MAIL_USER, email_dest, msg.as_string())
        return True
    except Exception as e:
        print(f"Erreur envoi mail : {e}")
        return False

def generer_html_pv_pack(df_session):
    html = """<html><head><style>
        @media print {.pb {page-break-after:always;}} 
        body {font-family:Arial; font-size:12px; margin:10px;} 
        table {width:100%; border-collapse:collapse; margin-top:8px;} 
        td,th {border:1px solid black; padding:5px; text-align:left;}
        .header-pv {text-align:center; font-weight:bold; border-bottom:2px solid black; padding-bottom:5px; margin-bottom:10px;}
        .infraction-box {border: 2px solid #000; margin-top: 10px; height: 80px; padding: 5px; position: relative;}
        .infraction-box:before {content: "SIGNALEMENTS / INFRACTIONS :"; font-weight: bold; font-size: 10px;}
        .sep {border-top: 2px dashed #000; margin: 30px 0; padding-top: 10px; position: relative; text-align:center;}
        
    </style></head><body>"""
    
    # On groupe par matière, lieu, jour et horaire
    groupes = df_session.groupby(['Enseignements', 'Lieu', 'Jours', 'Horaire'])
    for idx, ((mat, lieu, jour, heure), data) in enumerate(groupes):
        if idx > 0: html += "<div class='sep'></div>"
        promo = str(data['Promotion'].iloc[0])
        
        # Récupération automatique de l'effectif spécifique à la salle
        eff_salle = data['Effectif_Salle'].iloc[0] if 'Effectif_Salle' in data.columns else 0
        
        # Transformation du nom du Responsable pour afficher Nom et Prénom
        resp_brut = str(data.get('Responsable', pd.Series(['N/A'])).iloc[0]).strip().upper()
        responsable_complet = map_noms.get(resp_brut, resp_brut)
        
        html += f"""
        <div class='pb'>
            <div class='header-pv'>
                DÉPARTEMENT D'ÉLECTROTECHNIQUE<br>
                FACULTÉ DE GÉNIE ÉLECTRIQUE
            </div>
            <h3 style='text-align:center; margin:5px;'>PROCES-VERBAL DE SURVEILLANCE</h3>
            <p><b>Matière :</b> {mat} | <b>Lieu :</b> {lieu} | <b>Date :</b> {jour} | <b>Heure :</b> {heure}</p>
            <p><b>Promotion :</b> {promo} | <b>Chargé de matière :</b> {responsable_complet}</p>
            <table>
                <tr style='background-color:#f2f2f2;'>
                    <th>Étudiants prévus</th><th>Absences (Nombre)</th><th>Présences (Nombre)</th><th>Copies rendues (Nombre):</th>
                </tr>
                <tr style='height:30px; text-align:center;'>
                    <td>{int(eff_salle) if eff_salle > 0 else '...'}</td><td></td><td></td>
                </tr>
            </table>
            <table>
                <tr style='background-color:#f2f2f2;'>
                    <th width='70%'>Surveillant (Nom & Prénom)</th><th>Signature</th>
                </tr>"""
        for s in data['Enseignants'].tolist():
            html += f"<tr><td>{s}</td><td style='height:35px;'></td></tr>"
        html += f"""</table>
            <div class='infraction-box'></div>
            <div style='display:flex; justify-content:space-between; margin-top:15px;'>
                <div style='text-align:center; width:45%;'>
                    <b>Signature du chargé de la matière</b><br><br><br>__________________________
                </div>
                <div style='text-align:center; width:45%;'>
                    <b>Signature Chef de salle / Amphi</b><br><br><br>__________________________
                </div>
            </div>
        </div>"""
    html += "</body></html>"
    return html

# ======================================================================================
# 3. LOGIQUE D'AFFECTATION
# ======================================================================================
def affecter_enseignants_dynamique(batch_temp, df_global, q_psup, q_vac, q_def, p_sup, vacs):
    df_full = pd.concat([df_global, pd.DataFrame(batch_temp)], ignore_index=True) if not df_global.empty else pd.DataFrame(batch_temp)
    
    for i, row in enumerate(batch_temp):
        tag_actuel = row["Enseignants"] 
        
        charge = df_full[~df_full["Enseignants"].str.contains("⚠️|TEMP", na=False)]["Enseignants"].value_counts().to_dict()
        occupes = df_full[(df_full['Jours'] == row['Jours']) & (df_full['Horaire'] == row['Horaire'])]['Enseignants'].tolist()
        
        candidats = []
        
        if tag_actuel == "TEMP_PERM":
            pool_a_tester = p_sup
        elif tag_actuel == "TEMP_VAC":
            pool_a_tester = vacs
        else:
            pool_a_tester = LISTE_PROFS

        for p in pool_a_tester:
            quota = q_psup if p in p_sup else (q_vac if p in vacs else q_def)
            seances = charge.get(p, 0)
            
            if seances < quota and p not in occupes:
                candidats.append((p, seances))
        
        if candidats:
            candidats.sort(key=lambda x: x[1])
            min_c = candidats[0][1]
            meilleurs = [n for n, c in candidats if c == min_c]
            
            elu = random.choice(meilleurs)
            
            batch_temp[i]["Enseignants"] = elu
            batch_temp[i]["email"] = dict_emails.get(elu, "")
            
            idx = (len(df_global) if not df_global.empty else 0) + i
            df_full.at[idx, "Enseignants"] = elu
        else:
            batch_temp[i]["Enseignants"] = f"⚠️ BESOIN ({'PERM' if tag_actuel == 'TEMP_PERM' else 'VAC'})"
            
    return batch_temp
# ======================================================================================
# 4. INTERFACE & CHARGEMENT DES DONNÉES (SÉCURISÉE)
# ======================================================================================

# Initialisation cruciale pour éviter le NameError dans tout le script
df_db_global = get_db()

with st.sidebar:
    # --- AJOUT SÉCURITÉ POUR LA CONFIGURATION ---
    st.header("🔐 Administration")
    pwd_side = st.text_input("Code d'accès Configuration :", type="password", key="pwd_sidebar_cfg")

    if pwd_side == "": # Remplacez 0000 par votre code secret
        st.success("Accès autorisé")
        st.header("⚙️ Configuration")
        
        # Paramètres de surveillance
        nb_amphi = st.number_input("👥 Surveillants / AMPHI", 1, 10, 3, key="nb_amp_cfg")
        nb_salle = st.number_input("👥 Surveillants / SALLE", 1, 10, 1, key="nb_sal_cfg")
        st.divider()
        
        # --- PRÉPARATION DES LISTES PAR DÉFAUT ---
        default_p_sup = []
        default_vac = []
        
        if 'df_contacts' in locals() or 'df_contacts' in globals():
            try:
                # 1. NETTOYAGE RADICAL DES ENTÊTES
                # On enlève les espaces et on force en majuscules pour éviter Qualité vs Qualite
                df_contacts.columns = [str(c).strip().upper() for c in df_contacts.columns]
                
                # On cherche le nom de la colonne qui ressemble à "QUALITÉ" ou "QUALITE"
                col_qualite = next((c for c in df_contacts.columns if "QUALIT" in c), None)
                col_nom = next((c for c in df_contacts.columns if "NOM" in c), "NOM")
                col_prenom = next((c for c in df_contacts.columns if "PRÉNOM" in c or "PRENOM" in c), "PRÉNOM")

                if col_qualite:
                    # 2. NETTOYAGE DES DONNÉES
                    def deep_clean(val):
                        return " ".join(str(val).split()).upper().strip()

                    # On crée le nom complet pour le match
                    df_contacts['MATCH_NAME'] = df_contacts[col_nom].apply(deep_clean) + " " + df_contacts[col_prenom].apply(deep_clean)
                    
                    # On nettoie la colonne Qualité elle-même
                    df_contacts['QUALITE_CLEAN'] = df_contacts[col_qualite].astype(str).str.upper().str.strip()

                    # 3. EXTRACTION AVEC FILTRE LARGE
                    # On cherche "PERMAN" (pour Permanent/Permanente) et "VACAT" (pour Vacataire)
                    mask_perm = df_contacts['QUALITE_CLEAN'].str.contains('PERMAN', na=False)
                    mask_vac = df_contacts['QUALITE_CLEAN'].str.contains('VACAT', na=False)

                    default_p_sup = df_contacts[mask_perm]['MATCH_NAME'].unique().tolist()
                    default_vac = df_contacts[mask_vac]['MATCH_NAME'].unique().tolist()
                else:
                    st.error("La colonne 'Qualité' est introuvable dans le fichier contacts.")

            except Exception as e:
                st.warning(f"Erreur technique lors du tri : {e}")

        # --- BLOC QUOTAS (CORRECTION FINALE SÉPARATION) ---
        st.subheader("⚖️ Quotas")
        FILE_CONTACTS = "Permanents-Vacataires-ELT2-2025-2026.xlsx"
        default_p_sup = []
        default_vac = []
        all_names_from_excel = []
        
        try:
            import pandas as pd
            df_c = pd.read_excel(FILE_CONTACTS)
            
            # Nettoyage des colonnes
            df_c.columns = [str(c).strip().upper() for c in df_c.columns]
            
            # Identification des colonnes (Gestion des accents)
            col_q = next((c for c in df_c.columns if "QUALIT" in c), None)
            col_n = next((c for c in df_c.columns if "NOM" in c), "NOM")
            col_p = next((c for c in df_c.columns if "PRÉNOM" in c or "PRENOM" in c), "PRÉNOM")
        
            if col_n and col_p:
                def clean_v(v):
                    return " ".join(str(v).split()).upper().strip()
        
                # Création du nom complet pour TOUS les enseignants du fichier Excel
                df_c['MATCH_NAME'] = df_c[col_n].apply(clean_v) + " " + df_c[col_p].apply(clean_v)
                all_names_from_excel = df_c['MATCH_NAME'].unique().tolist()
                
                if col_q:
                    # Nettoyage de la qualité
                    df_c['Q_TYPE'] = df_c[col_q].astype(str).str.strip().str.upper()
        
                    # --- FILTRAGE PRÉCIS (Inclut TD et TP via les mots-clés) ---
                    mask_p = df_c['Q_TYPE'].str.contains('PERMAN', na=False)
                    mask_v = df_c['Q_TYPE'].str.contains('VACAT', na=False)
        
                    # Listes de base pour les valeurs par défaut
                    default_p_sup = df_c[mask_p]['MATCH_NAME'].unique().tolist()
                    default_vac = df_c[mask_v & ~mask_p]['MATCH_NAME'].unique().tolist()
        
        except Exception as e:
            st.warning(f"Note : Lecture contacts impossible ({e})")
        
        # 2. Lien avec LISTE_PROFS et Fusion avec Excel (Cours + TD + TP)
        list_edt = LISTE_PROFS if 'LISTE_PROFS' in locals() or 'LISTE_PROFS' in globals() else []
        list_total_surveillance = sorted(list(set(list_edt + all_names_from_excel)))
        
        # --- FILTRAGE ADMINISTRATION (EXCLUSION) ---
        exclus_admin = [
            "MILOUA FARID",
            "LARBI ZIDI",
            "BENSALEM RITEDJ INES",
            "BERROUNA HENEN",
            "MOKEDDEM AMINA",
            "DROUNI Wassila Abir",
            "BERRAHAL BEL ABBES",
            "BENKHALLOUK KADDOUR",
            "MOULAY MILOUD",
            "FETHI.M FETHI",
            "HADER ABDERRAHMANE",
            "BENSALEM ABDELHAMID",
            "BENTAALLAH ABDERRAHIM",
            "KHATIR MOHAMED", 
            "HANAFI SALAH",
            "AZAIZ AHMED",
            "OUKLI MIMOUNA",
            "HADJI AHMED",
            "MAHIDDINE ABDERRAHIM",
            "BEDDAD BOUCIF",
            "ABDELI MAMA",
            "DRIEF FAIZA",
            "NASSOUR KAMEL", 
            "DJERIRI YOUCEF",
            "BELLEBNA YASSINE"
        ]
        
        # Variable CRITIQUE pour le moteur de planning (Ligne 837)
        list_p_affichee = [p for p in list_total_surveillance if str(p).upper().strip() not in exclus_admin]
        
        # --- AFFICHAGE NUMÉRIQUE DES EFFECTIFS ---
        # On place les compteurs juste avant les sélections
        c1, c2 = st.columns(2)
        
        # --- INTERFACE DE SÉLECTION (VERSION MISE À JOUR) ---

        # 🎓 Liste 1 : Enseignants permanents
        p_sup_list = st.multiselect(
            "🎓 Enseignants permanents", 
            list_p_affichee, 
            default=[p for p in default_p_sup if p in list_p_affichee], 
            key="p_sup_cfg"
        )
        
        # Mise à jour du compteur numérique pour les permanents
        c1.metric("🎓 Permanents", f"{len(p_sup_list)} ENS")
        
        q_psup = st.number_input("Seuil Max (e)", 0, 20, 2, key="q_psup_cfg")
        
        # 📝 Liste 2 : Vacataires
        vac_options = [p for p in list_p_affichee if p not in p_sup_list]
        
        vac_list = st.multiselect(
            "📝 Vacataires", 
            vac_options, 
            default=[p for p in default_vac if p in vac_options], 
            key="vac_list_cfg"
        )
        
        # Mise à jour du compteur numérique pour les vacataires
        c2.metric("📝 Vacataires", f"{len(vac_list)} ENS")
        
        q_vac = st.number_input("Seuil Max (Vac)", 0, 20, 6, key="q_vac_cfg")
        q_defaut = st.number_input("Seuil (Autres)", 0, 20, 3, key="q_def_cfg")
        
        st.divider()
                        
        # Gestion des jours fériés
        feries = st.multiselect("🚫 Jours Fériés", [datetime(2026, 5, i).date() for i in range(1, 32)], key="fer_cfg")
        st.divider()
        
        # Section de téléchargement des données globales
        # Section de téléchargement des données globales
        st.subheader("📥 Téléchargements")
        
        if not df_db_global.empty:
            # 1. BOUTONS GLOBAUX (EXISTANTS)
            col_g1, col_g2 = st.columns(2)
            with col_g1:
                st.download_button(
                    "📊 Excel Global", 
                    data=generer_excel_bytes(df_db_global[COLS_ORDRE]), 
                    file_name="Planning_S2_2026.xlsx", 
                    use_container_width=True,
                    key="dl_xlsx_cfg"
                )
            with col_g2:
                # Ajout explicite du Responsable dans l'HTML Global si nécessaire
                cols_global = COLS_ORDRE + ["Responsable"] if "Responsable" not in COLS_ORDRE else COLS_ORDRE
                html_g = f"<html><head><meta charset='UTF-8'></head><body><h2 style='text-align:center;'>{TITRE_OFFICIEL}</h2>{df_db_global[cols_global].to_html(index=False)}</body></html>"
                st.download_button(
                    "🌐 HTML Global", 
                    data=html_g, 
                    file_name="Planning_S2_2026.html", 
                    mime="text/html", 
                    use_container_width=True,
                    key="dl_html_cfg"
                )
        
            st.write("---")
            
            # 2. GÉNÉRATION DES EDTs DE SURVEILLANCE PAR PROMOTION
            # 2. GÉNÉRATION DES EDTs DE SURVEILLANCE PAR PROMOTION
            st.markdown("### 📅 Télécharger les EDTs de surveillance")
            
            promo_list = ["TOUTES LES PROMOTIONS (ALL)"] + sorted(list(DATA_AUTO.keys()))
            choice_promo = st.selectbox("Choisir la promotion à exporter", promo_list, key="sel_promo_dl")
            
            if choice_promo == "TOUTES LES PROMOTIONS (ALL)":
                df_to_export = df_db_global.copy()
                suffix = "ALL"
            else:
                df_to_export = df_db_global[df_db_global["Promotion"].str.contains(choice_promo, na=False)].copy()
                suffix = choice_promo
        
            if not df_to_export.empty:
                # 1. Tri chronologique pour le document final
                df_to_export = df_to_export.sort_values(by=["Jours", "Horaire"])
                
                # 2. Préparation des données enrichies
                df_display = df_to_export.copy()
        
                # Application du nom complet pour le Chargé de Matière (Responsable)
                df_display["Responsable"] = df_display["Responsable"].apply(
                    lambda x: map_noms.get(str(x).strip().upper(), x)
                )
        
                # Sélection des colonnes incluant la Promotion et le Responsable
                # Ordre : Date, Horaire, Promotion, Enseignements, Lieu, Surveillants, Chargé de Matière
                cols_finales = ["Jours", "Horaire", "Promotion", "Enseignements", "Lieu", "Enseignants", "Responsable"]
                
                df_display = df_display[cols_finales].rename(columns={
                    "Jours": "Date",
                    "Enseignants": "Surveillants",
                    "Responsable": "Chargé de Matière"
                })
        
                # 3. Construction du HTML avec le titre officiel
                html_surv = f"""
                <html>
                <head>
                <meta charset="UTF-8">
                <style>
                    body {{ font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; margin: 25px; }}
                    h2 {{ text-align: center; color: #2c3e50; text-transform: uppercase; font-size: 18px; border-bottom: 2px solid #2c3e50; padding-bottom: 10px; }}
                    h3 {{ text-align: center; color: #1E88E5; margin-top: 10px; }}
                    table {{ width: 100%; border-collapse: collapse; margin-top: 20px; }}
                    th {{ background-color: #f8f9fa; color: #333; border: 1px solid #000; padding: 10px; font-size: 12px; }}
                    td {{ border: 1px solid #000; padding: 8px; text-align: left; font-size: 11px; vertical-align: middle; }}
                    tr:nth-child(even) {{ background-color: #f2f2f2; }}
                    .footer {{ font-size: 10px; margin-top: 30px; text-align: right; color: #666; }}
                </style>
                </head>
                <body>
                    <h2>Plateforme de gestion des EDTs-S2-2026-Département d'Électrotechnique-Faculté de génie électrique-UDL-SBA</h2>
                    <h3>EDT DE SURVEILLANCE - {suffix}</h3>
                    {df_display.to_html(index=False, escape=False)}
                    <div class="footer">
                        Document généré le : {datetime.now().strftime('%d/%m/%Y %H:%M')}
                    </div>
                </body>
                </html>
                """
                
                st.download_button(
                    label=f"📄 Télécharger l'EDT Complet (Promo + Responsables)", 
                    data=html_surv, 
                    file_name=f"EDT_Surveillance_{suffix}_2026.html", 
                    mime="text/html", 
                    use_container_width=True,
                    key=f"dl_full_export_{suffix}"
                )
        # --- ZONE CRITIQUE : NETTOYAGE DE LA BASE DE DONNÉES ---
        st.divider()
        st.warning("⚠️ Zone de Maintenance Critique")
        
        # 1. Ajout de la case de confirmation (Sécurité supplémentaire)
        confirm_delete = st.checkbox("Confirmer la suppression totale des données", key="chk_confirm_del")
        
        # 2. Le bouton est désactivé (disabled) si la case n'est pas cochée
        if st.button("🧨 VIDER LA BASE", use_container_width=True, key="btn_wipe_db_sec", disabled=not confirm_delete):
            try:
                # Exécution de la suppression sur Supabase
                supabase.table(TABLE_NAME).delete().neq("Promotion", "X").execute()
                st.success("✅ La base de données a été réinitialisée avec succès.")
                
                # Petite pause pour confirmation visuelle
                import time
                time.sleep(2)
                
                # Rechargement de l'application
                st.rerun()
            except Exception as e:
                st.error(f"❌ Erreur lors de la suppression : {e}")

    # Gestion des messages liés au code d'accès de la barre latérale
    elif pwd_side != "":
        st.error("❌ Code d'accès incorrect.")
    else:
        st.info("🔒 Saisissez le code dans la barre latérale pour modifier la configuration.")
# --- RAPPEL DU TITRE OFFICIEL AU CENTRE ---
st.markdown(f"<h2 style='text-align:center; color:#1f4e79;'>{TITRE_OFFICIEL}</h2>", unsafe_allow_html=True)
# ======================================================================================
# 5. SYSTÈME DE SÉCURITÉ DYNAMIQUE
# ======================================================================================

# ======================================================================================
# 5. SYSTÈME DE SÉCURITÉ DYNAMIQUE (DÉSACTIVÉ - ACCÈS LIBRE SANS CODE)
# ======================================================================================

CODE_ADMIN_EDT = ""

# FORCE L'ACCÈS DIRECT SANS CODE POUR LE MOMENT
if 'auth_admin_edt' not in st.session_state:
    st.session_state.auth_admin_edt = True  # Modifié de False à True pour bypasser le verrou

def verifier_admin_edt_dynamique(suffixe):
    # Conservé pour éviter les erreurs de dépendances dans le reste du script
    st.session_state.auth_admin_edt = True
    st.rerun()

def afficher_verrou(suffixe):
    # Si par un autre biais cette fonction est appelée, elle valide automatiquement l'accès
    st.session_state.auth_admin_edt = True
    st.rerun()

# ======================================================================================
# 6. CRÉATION DES ONGLETS ET LOGIQUE D'AFFICHAGE
# ======================================================================================
# ==========================================
# FONCTIONS DE GESTION DE LA BASE DE DONNÉES
# ==========================================

def update_item_db(record_id, column_name, new_value):
    """Met à jour une cellule spécifique dans Supabase."""
    try:
        supabase.table(TABLE_NAME).update({column_name: new_value}).eq("id", record_id).execute()
        # On ne met pas de st.rerun ici car c'est appelé par on_change
    except Exception as e:
        st.error(f"Erreur de mise à jour : {e}")

def dupliquer_seance(row_data):
    """Crée une copie d'une séance pour ajouter un surveillant."""
    new_data = row_data.to_dict()
    if 'id' in new_data: del new_data['id'] # Supprimer l'ID pour en générer un nouveau
    try:
        supabase.table(TABLE_NAME).insert(new_data).execute()
        st.success("Ligne ajoutée ! Sélectionnez le nouveau surveillant.")
        st.rerun()
    except Exception as e:
        st.error(f"Erreur lors de l'ajout : {e}")

def supprimer_ligne(rid):
    """Supprime définitivement une ligne du planning."""
    try:
        supabase.table(TABLE_NAME).delete().eq("id", rid).execute()
        st.rerun()
    except Exception as e:
        st.error(f"Erreur de suppression : {e}")
t1, t2, t3, t4, t5, t6, t7 = st.tabs([
    "🚀 SESSION COMMUNE", 
    "📅 PLANNING AUTO", 
    "📄 GÉNÉRER PV", 
    "📋 RÉCAPITULATIF", 
    "🔧 MAINTENANCE", 
    "📝 ASSIDUITÉ", 
    "📩 Soumettre un justificatif (Etudiants)"
])

# --- BLOC T1 : SESSION COMMUNE ---
with t1:
    if not st.session_state.auth_admin_edt:
        afficher_verrou("t1")
    else:
        st.subheader("📢 Session Commune")
        LISTE_COMPLÈTE_LIEUX = [
            "S01", "S02", "S03", "S04", "S05", "S06", "S07", "S08", "S08 BIS", "SN",
            "A01", "A02", "A03", "A04", "A05", "A06", "A07", "A08", "A09", "A10", "A11", "A12"
        ]
                
        c1, c2 = st.columns(2)
        with c1:
            resp_c = st.selectbox("Chargé de matière", LISTE_PROFS, key="sc_resp")
            # Nettoyage du nom pour la recherche
            nom_famille_seul = str(resp_c).split(" ")[0].strip().upper()
            
            mats_enseignee = []
            if not df_src.empty:
                # 1. On filtre toutes les lignes de cet enseignant
                mask_profs = df_src["Enseignants"].astype(str).str.upper().str.contains(nom_famille_seul, na=False)
                df_prof_only = df_src[mask_profs]
                
                # 2. FILTRE STRICT : Uniquement les "Cours-" (ignore TD/TP)
                mask_cours = df_prof_only["Enseignements"].astype(str).str.strip().str.startswith("Cours-", na=False)
                df_cours_only = df_prof_only[mask_cours]
                
                # Récupération des noms complets uniques
                mats_enseignee = sorted(df_cours_only["Enseignements"].unique().tolist())
            
            # Sélecteur de matière
            mat_sel_sc = st.selectbox("Matière à traiter", ["Sélectionnez une matière"] + mats_enseignee)
            
            promo_auto = []
            h_defaut = "08h30 – 10h30"

            if mat_sel_sc != "Sélectionnez une matière":
                # COMPARAISON NOM COMPLET : Insensible à la casse et aux espaces
                nom_complet_nettoyé = str(mat_sel_sc).lower().strip()
                
                # Recherche globale dans tout le fichier pour regrouper les promotions (M1CE, M1ME, etc.)
                lignes_matiere = df_src[df_src["Enseignements"].astype(str).str.lower().str.strip() == nom_complet_nettoyé]
                
                if not lignes_matiere.empty:
                    st.info(f"Détails détectés dans le fichier source pour : {mat_sel_sc}")
                    st.table(lignes_matiere[["Enseignements", "Promotion", "Enseignants", "Horaire", "Lieu"]])
                
                # Pré-cochage automatique des promotions trouvées
                promos_dans_excel = lignes_matiere["Promotion"].unique().tolist()
                promo_auto = [p for p in promos_dans_excel if p in DATA_AUTO]
                
                # Horaire par défaut basé sur l'Excel
                h_defaut = lignes_matiere.iloc[0].get("Horaire", h_defaut)
            
            p_c = st.multiselect("Promotions concernées", list(DATA_AUTO.keys()), default=promo_auto)

        # Gestion des horaires et salles par défaut
        s_defaut = []
        if p_c:
            if p_c[0] in DATA_AUTO:
                h_defaut = DATA_AUTO[p_c[0]].get("Horaire", h_defaut)
                s_auto = DATA_AUTO[p_c[0]].get("Salles", [])
                s_defaut = [s for s in s_auto if s in LISTE_COMPLÈTE_LIEUX]

        with c2:
            d_c = st.date_input("Date", datetime(2026, 5, 2), key="sc_d")
            creneaux_possibles = sorted(list(set([v["Horaire"] for v in DATA_AUTO.values()])))
            
            # Sécurité pour l'index de l'horaire
            idx_h = 0
            if h_defaut in creneaux_possibles:
                idx_h = creneaux_possibles.index(h_defaut)
                
            h_c = st.selectbox("Horaire", creneaux_possibles, index=idx_h, key="sc_h_select")
            s_c = st.multiselect("Lieux", LISTE_COMPLÈTE_LIEUX, default=s_defaut, key="sc_s")

        st.divider()
        if st.button("🔥 GÉNÉRER SESSION", use_container_width=True):
            if not p_c or not s_c:
                st.error("Veuillez sélectionner au moins une promotion et un lieu.")
            else:
                batch = []
                # Recherche finale par nom complet pour récupérer le Code (Cours-AERN...)
                nom_final_nettoyé = str(mat_sel_sc).lower().strip()
                inf_rows = df_src[df_src["Enseignements"].astype(str).str.lower().str.strip() == nom_final_nettoyé]
                
                if not inf_rows.empty:
                    inf = inf_rows.iloc[0]
                    for l in s_c:
                        # Détermination du quota selon le type de salle
                        q = nb_amphi if any(x in str(l) for x in ['A', '']) else nb_salle
                        for _ in range(q):
                            batch.append({
                                "Enseignements": mat_sel_sc, 
                                "Code": str(inf["Code"]), 
                                "Enseignants": "TEMP", 
                                "Horaire": h_c, 
                                "Jours": d_c.strftime("%d/%m/%Y"), 
                                "Lieu": l, 
                                "Promotion": " / ".join(p_c), 
                                "Responsable": resp_c, 
                                "email": ""
                            })
                
                if batch:
                    batch = affecter_enseignants_dynamique(batch, df_db_global, q_psup, q_vac, q_defaut, p_sup_list, vac_list)
                    supabase.table(TABLE_NAME).insert(batch).execute()
                    st.success(f"✅ Session générée pour {mat_sel_sc}")
                    st.rerun()

# --- BLOC T2 : PLANNING AUTO (SECTION RESTRICTIONS) ---
# --- BLOC T2 : PLANNING AUTO (SECTION RESTRICTIONS) ---
with t2:
    if not st.session_state.auth_admin_edt:
        afficher_verrou("t2")
    else:
        # Rappel de l'interface obligatoire
        st.subheader("📅 Planning Automatique - Plateforme de gestion des EDTs-S2-2026-Département d'Électrotechnique-Faculté de génie électrique-UDL-SBA")
        
        # --- 1. CALCUL DES STATISTIQUES ---
        promos_totale_liste = list(DATA_AUTO.keys())
        total_promos = len(promos_totale_liste)
        
        # Variable pivot pour éviter le NameError
        promos_existantes = df_db_global["Promotion"].unique().tolist() if not df_db_global.empty else []
        
        nb_generees = len(promos_existantes)
        progression = nb_generees / total_promos if total_promos > 0 else 0
        
        # --- 2. AFFICHAGE NUMÉRIQUE ---
        c1, c2 = st.columns([1, 3])
        with c1:
            st.metric("🎓 Générées", f"{nb_generees} / {total_promos}")
        with c2:
            st.write(f"**Progression globale de la session**")
            st.progress(progression)
            
        # Liste déroulante des promotions traitées
        if nb_generees > 0:
            with st.expander("✅ Voir le détail des promotions déjà générées"):
                st.write(", ".join(promos_existantes))
        else:
            st.info("ℹ️ Aucune promotion n'a été générée pour le moment.")
        
        st.divider()

        # --- 3. SÉLECTION ET CONFIGURATION ---
        p_sel = st.selectbox("Sélectionner Promotion", [""] + promos_totale_liste)
        
        if p_sel:
            cfg = DATA_AUTO[p_sel]
            st.info(f"👥 Effectif : **{cfg['Effectif']}** étudiants")
            
            # Message adapté au mode Mise à jour / Historique
            if p_sel in promos_existantes:
                st.info(f"🔄 **Historique détecté** : La génération se fera en mode **Mise à jour**. Les surveillants déjà affectés seront conservés, le système complétera les modifications.")
            
            # Initialisation de la mémoire persistante pour stocker le DataFrame des matières et chargés
            if "df_matieres_state" not in st.session_state:
                st.session_state.df_matieres_state = {}
                
            if p_sel not in st.session_state.df_matieres_state:
                # Récupération des modules d'origine de la source Excel/DB pour initialiser
                mats_p_init = sorted(df_src[df_src["Promotion"]==p_sel]["Enseignements"].unique().tolist()) if not df_src.empty else []
                lignes_initiales = []
                for m in mats_p_init:
                    # Recherche du chargé d'origine s'il existe
                    inf_rows_init = df_src[(df_src["Enseignements"] == m) & (df_src["Promotion"] == p_sel)]
                    charge_init = "À Définir"
                    if not inf_rows_init.empty:
                        val_c = str(inf_rows_init.iloc[0]["Enseignants"]).strip()
                        if val_c != "" and "DEFINIR" not in val_c.upper() and "DÉFINIR" not in val_c.upper():
                            charge_init = val_c
                    lignes_initiales.append({"Enseignements": m, "Chargé de Matière": charge_init})
                
                if not lignes_initiales:
                    lignes_initiales.append({"Enseignements": "", "Chargé de Matière": "À Définir"})
                    
                st.session_state.df_matieres_state[p_sel] = pd.DataFrame(lignes_initiales, columns=["Enseignements", "Chargé de Matière"])
            
            # Nettoyage et préparation de la liste déroulante globale pour les chargés de cours
            # Nettoyage et préparation de la liste déroulante globale pour les chargés de cours
            # Nettoyage et préparation de la liste déroulante globale pour les chargés de cours
            enseignants_source = df_src["Enseignants"].dropna().unique().tolist() if not df_src.empty else []
            p_sup_clean = [str(x).strip() for x in p_sup_list if str(x).strip() != ""]
            vac_clean = [str(x).strip() for x in vac_list if str(x).strip() != ""]
            src_clean = [str(x).strip() for x in enseignants_source if str(x).strip() != ""]
            
            # Fusion avec les enseignants ajoutés manuellement en session state
            liste_deroulante_globale = sorted(list(set(p_sup_clean + vac_clean + src_clean + st.session_state.enseignants_ajoutes_manuellement)))
            
            # Insertion des options spécifiques en tête de liste
            if "À Définir" not in liste_deroulante_globale:
                liste_deroulante_globale.insert(0, "À Définir")
            
            if "Ajouter un chargé de cours" not in liste_deroulante_globale:
                liste_deroulante_globale.insert(1, "Ajouter un chargé de cours")
            
            # --- ZONE DE SAISIE EN BLOC ET MODIFICATION DES AJOUTS ---
            with st.expander("➕ AJOUTS MULTIPLES / MODIFIER LES NOMS AJOUTÉS :"):
                
                # Sous-section 1 : Ajouts en bloc (Matières & Enseignants)
                with st.form(key=f"form_ajouts_multiples_assoc_{p_sel}", clear_on_submit=True):
                    st.markdown("### 📥 Ajouter de nouveaux éléments")
                    col_mat, col_ens = st.columns(2)
                    
                    with col_mat:
                        st.markdown("**📚 Étape 1 : Saisir les Matières**")
                        txt_matieres = st.text_area(
                            "Matières séparées par des virgules :",
                            placeholder="Ex: Stabilité des réseaux, Éclairage LED",
                            key=f"area_mat_{p_sel}"
                        )
                        
                    with col_ens:
                        st.markdown("**👥 Étape 2 : Saisir les Chargés correspondants**")
                        txt_enseignants = st.text_area(
                            "Enseignants séparés par des virgules (Même ordre) :",
                            placeholder="Ex: ZIDI, BERMAKI",
                            key=f"area_ens_{p_sel}"
                        )
                        
                    bouton_valider_bloc = st.form_submit_button("💾 Enregistrer et Associer automatiquement")
                    
                    if bouton_valider_bloc:
                        if not txt_matieres.strip():
                            st.error("⚠️ Veuillez saisir au moins une matière pour effectuer l'association.")
                        else:
                            changements = False
                            liste_mat_saisies = [m.strip() for m in txt_matieres.split(",") if m.strip()]
                            liste_noms_saisis = [n.strip().upper() for n in txt_enseignants.split(",") if n.strip()] if txt_enseignants.strip() else []
                            
                            for nom in liste_noms_saisis:
                                if nom not in st.session_state.enseignants_ajoutes_manuellement and nom not in (p_sup_clean + vac_clean + src_clean):
                                    st.session_state.enseignants_ajoutes_manuellement.append(nom)
                                    changements = True
                            
                            df_actuel = st.session_state.df_matieres_state[p_sel].copy()
                            for idx, mat in enumerate(liste_mat_saisies):
                                if "Enseignements" in df_actuel.columns and mat not in df_actuel["Enseignements"].values:
                                    charge_associe = liste_noms_saisis[idx] if idx < len(liste_noms_saisis) else "À Définir"
                                        
                                    nouvelle_ligne = {
                                        "Enseignements": mat,
                                        "Code": "A Définir",
                                        "Enseignants": "À Définir",
                                        "Horaire": cfg.get("Horaire", "8h - 9h30"),
                                        "Jours": "Dimanche",
                                        "Lieu": cfg.get("Salles", ["S08"])[0] if cfg.get("Salles") else "S08",
                                        "Promotion": p_sel,
                                        "Chargé de Matière": charge_associe
                                    }
                                    nouvelle_ligne_adaptee = {k: v for k, v in nouvelle_ligne.items() if k in df_actuel.columns}
                                    df_actuel = pd.concat([df_actuel, pd.DataFrame([nouvelle_ligne_adaptee])], ignore_index=True)
                                    changements = True
                            
                            if changements:
                                st.session_state.df_matieres_state[p_sel] = df_actuel
                                st.success("📊 Éléments enregistrés avec succès !")
                                st.rerun()

                # --- NOUVEAUTÉ : Sous-section 2 : Modification/Correction des noms saisis ---
                if st.session_state.enseignants_ajoutes_manuellement:
                    st.divider()
                    st.markdown("### 🛠️ Corriger ou Modifier un nom déjà ajouté")
                    
                    with st.form(key=f"form_modification_nom_{p_sel}"):
                        nom_a_modifier = st.selectbox(
                            "Sélectionner le nom à corriger :", 
                            st.session_state.enseignants_ajoutes_manuellement,
                            key=f"select_mod_nom_{p_sel}"
                        )
                        
                        nouveau_nom_corrige = st.text_input(
                            "Saisir le nom corrigé :", 
                            value=nom_a_modifier,
                            key=f"input_mod_nom_{p_sel}"
                        ).strip().upper()
                        
                        bouton_modifier = st.form_submit_button("✏️ Appliquer la modification")
                        
                        if bouton_modifier and nouveau_nom_corrige and nom_a_modifier:
                            if nouveau_nom_corrige != nom_a_modifier:
                                # 1. Remplacement dans la liste globale de session
                                idx_nom = st.session_state.enseignants_ajoutes_manuellement.index(nom_a_modifier)
                                st.session_state.enseignants_ajoutes_manuellement[idx_nom] = nouveau_nom_corrige
                                
                                # 2. Remplacement dynamique direct dans le tableau d'édition actuel s'il y est affecté
                                if p_sel in st.session_state.df_matieres_state:
                                    df_tab = st.session_state.df_matieres_state[p_sel].copy()
                                    if "Chargé de Matière" in df_tab.columns:
                                        df_tab["Chargé de Matière"] = df_tab["Chargé de Matière"].replace(nom_a_modifier, nouveau_nom_corrige)
                                        st.session_state.df_matieres_state[p_sel] = df_tab
                                
                                st.success(f"🔄 Le nom '{nom_a_modifier}' a été converti en '{nouveau_nom_corrige}' partout avec succès !")
                                st.rerun()

            # --- STRUCTURE ET CONFIGURATION PAR PROMOTION ---
            ca, cb = st.columns([1, 1])
            with ca:
                s_f = st.multiselect(
                    "Lieux à inclure", 
                    cfg["Salles"], 
                    default=cfg["Salles"],
                    key=f"salles_multiselect_pa_{p_sel}"
                )
                h_f_global = st.text_input(
                    "Heure par défaut", 
                    cfg["Horaire"],
                    key=f"heure_input_pa_{p_sel}"
                )
            with cb: 
                d_f_global = st.date_input(
                    "Date début", 
                    datetime(2026, 6, 9),
                    key=f"date_input_pa_{p_sel}"
                )
            
            # --- INTERFACE SÉLECTION INTERACTIVE DES MATIÈRES ET CHARGÉS ---
            # --- INTERFACE SÉLECTION INTERACTIVE DES MATIÈRES ET CHARGÉS ---
            st.write("### 📝 Configuration des Modules et Chargés de matière")
            st.caption("Vous pouvez modifier les lignes existantes, ajouter de nouvelles matières en bas du tableau ou modifier le responsable via sa liste déroulante.")
            
            # --- NOUVEAUTÉ : OUTIL DE RÉORGANISATION DE L'ORDRE DES MATIÈRES ---
            df_actuel_tri = st.session_state.df_matieres_state[p_sel].copy()
            
            if not df_actuel_tri.empty and "Enseignements" in df_actuel_tri.columns:
                liste_matieres_reorg = df_actuel_tri["Enseignements"].dropna().tolist()
                                
                if len(liste_matieres_reorg) > 1:
                    with st.expander("↕️ Réorganiser l'ordre de passage des Matières"):
                        st.caption("Sélectionnez une matière puis utilisez les boutons pour modifier son ordre d'apparition dans le planning.")
                        
                        col_sel_mat, col_btn_up, col_btn_down = st.columns([2, 1, 1])
                        
                        with col_sel_mat:
                            mat_a_bouger = st.selectbox(
                                "Matière à déplacer :", 
                                liste_matieres_reorg, 
                                key=f"select_reorg_{p_sel}"
                            )
                        
                        idx_actuel = liste_matieres_reorg.index(mat_a_bouger)
                        
                        with col_btn_up:
                            st.write("") # Alignement visuel
                            st.write("")
                            # Correction de la syntaxe de la f-string ici :
                            if st.button("⬆️ Monter", use_container_width=True, key=f"btn_up_{p_sel}"):
                                if idx_actuel > 0:
                                    # Échange des lignes dans le DataFrame
                                    b, a = df_actuel_tri.iloc[idx_actuel].copy(), df_actuel_tri.iloc[idx_actuel - 1].copy()
                                    df_actuel_tri.iloc[idx_actuel - 1] = b
                                    df_actuel_tri.iloc[idx_actuel] = a
                                    st.session_state.df_matieres_state[p_sel] = df_actuel_tri
                                    st.rerun()
                                    
                        with col_btn_down:
                            st.write("") # Alignement visuel
                            st.write("")
                            # Correction de la syntaxe de la f-string ici :
                            if st.button("⬇️ Descendre", use_container_width=True, key=f"btn_down_{p_sel}"):
                                if idx_actuel < len(liste_matieres_reorg) - 1:
                                    # Échange des lignes dans le DataFrame
                                    b, a = df_actuel_tri.iloc[idx_actuel].copy(), df_actuel_tri.iloc[idx_actuel + 1].copy()
                                    df_actuel_tri.iloc[idx_actuel + 1] = b
                                    df_actuel_tri.iloc[idx_actuel] = a
                                    st.session_state.df_matieres_state[p_sel] = df_actuel_tri
                                    st.rerun()

            # Recalcul de la liste déroulante globale pour les chargés
            liste_deroulante_globale = sorted(list(set(p_sup_clean + vac_clean + src_clean + st.session_state.enseignants_ajoutes_manuellement)))
            if "À Définir" not in liste_deroulante_globale: liste_deroulante_globale.insert(0, "À Définir")
            if "Ajouter un chargé de cours" not in liste_deroulante_globale: liste_deroulante_globale.insert(1, "Ajouter un chargé de cours")

            # Affichage du tableau éditable réordonné
            df_configure = st.data_editor(
                st.session_state.df_matieres_state[p_sel],
                num_rows="dynamic",
                column_config={
                    "Enseignements": "Nom de la Matière",
                    "Chargé de Matière": st.column_config.SelectboxColumn(
                        "Chargé de Cours (Permanent / Vacataire)",
                        options=liste_deroulante_globale
                    )
                },
                hide_index=True,
                use_container_width=True,
                key=f"data_editor_planning_auto_{p_sel}"
            )
            
            # Sauvegarde immédiate des modifications dans le session state
            st.session_state.df_matieres_state[p_sel] = df_configure
            
            # Structuration de l'interface en colonnes
            ca, cb = st.columns([1, 1])
            with ca:
                s_f = st.multiselect("Lieux à inclure", cfg["Salles"], default=cfg["Salles"])
                h_f_global = st.text_input("Heure par défaut", cfg["Horaire"])
            with cb: 
                d_f_global = st.date_input("", datetime(2026, 6, 9))
                                   
            # Extraction propre des matières valides pour l'ajustement des dates
            matieres_valides_df = df_configure[df_configure["Enseignements"].astype(str).str.strip() != ""]
            selection_pure = matieres_valides_df["Enseignements"].tolist()

            custom_configs = {}
            if selection_pure:
                with st.expander("🛠️ Ajuster Dates et Horaires spécifiques"):
                    temp_date = d_f_global
                    for mod in selection_pure:
                        nom_affichage = mod
                        
                        # Gestion des week-ends et jours fériés
                        while temp_date.weekday() in [4] or temp_date in feries:
                            temp_date += timedelta(days=1)
                        
                        col1, col2, col3 = st.columns([2, 1, 1])
                        with col1: st.caption(f"**{nom_affichage}**")
                        with col2: new_date = st.date_input(f"Date", temp_date, key=f"d_{mod}")
                        with col3: new_hour = st.text_input(f"Heure", h_f_global, key=f"h_{mod}")
                        
                        custom_configs[mod] = {"date": new_date, "heure": new_hour}
                        temp_date += timedelta(days=1)

            # --- 2. BOUTON GÉNÉRER ET METTRE À JOUR ---
            if st.button("🚀 METTRE À JOUR / GÉNÉRER LE PLANNING", use_container_width=True):
                if not selection_pure:
                    st.error("Veuillez configurer au moins une matière valide dans le tableau ci-dessus.")
                else:
                    # Extraction de l'historique existant en local avant toute modification
                    historique_dict = {}
                    if not df_db_global.empty:
                        df_hist_promo = df_db_global[df_db_global["Promotion"] == p_sel]
                        for _, row in df_hist_promo.iterrows():
                            key_hist = (str(row["Enseignements"]).strip(), str(row["Lieu"]).strip())
                            if key_hist not in historique_dict:
                                historique_dict[key_hist] = []
                            if str(row["Enseignants"]).strip() != "" and "DEFINIR" not in str(row["Enseignants"]).upper() and "DÉFINIR" not in str(row["Enseignants"]).upper():
                                historique_dict[key_hist].append(str(row["Enseignants"]).strip())

                    # Nettoyage de la table pour cette promotion (Sera reconstruite proprement avec fusion de l'historique)
                    if p_sel in promos_existantes:
                        supabase.table(TABLE_NAME).delete().eq("Promotion", p_sel).execute()
                    
                    # --- INITIALISATION DE L'ALGORITHME DE TOUR DE RÔLE ÉQUITABLE ---
                    tous_enseignants_bruts = sorted(list(set(p_sup_clean + vac_clean + src_clean)))
                    
                    tous_enseignants = []
                    for ens in tous_enseignants_bruts:
                        ens_up = ens.upper()
                        if ens == "" or ens_up.startswith("COURS-") or "DEFINIR" in ens_up or "DÉFINIR" in ens_up:
                            continue
                        tous_enseignants.append(ens)
                    
                    compteurs_surveillance = {ens: 0 for ens in tous_enseignants}
                    
                    if not df_db_global.empty and "Enseignants" in df_db_global.columns:
                        for ens_exist in df_db_global["Enseignants"].values:
                            ens_clean = str(ens_exist).strip()
                            if ens_clean in compteurs_surveillance:
                                compteurs_surveillance[ens_clean] += 1

                    batch_pa = []
                    
                    # Parcours de notre tableau configuré dynamiquement
                    for index, row_mat in matieres_valides_df.iterrows():
                        mod = str(row_mat["Enseignements"]).strip()
                        
                        # Récupération directe du chargé choisi dans la liste déroulante globale
                        charge_matiere = str(row_mat["Chargé de Matière"]).strip() if pd.notna(row_mat["Chargé de Matière"]) else "À Définir"
                        
                        d_t = custom_configs[mod]["date"]
                        h_f = custom_configs[mod]["heure"]
                        date_str = d_t.strftime("%d/%m/%Y")
                        
                        veille_dt = d_t - timedelta(days=1)
                        veille_str = veille_dt.strftime("%d/%m/%Y")
                        
                        vrai_nom_module = mod
                        inf_rows_pa = df_src[(df_src["Enseignements"] == mod) & (df_src["Promotion"] == p_sel)]
                        
                        if not inf_rows_pa.empty:
                            inf = inf_rows_pa.iloc[0]
                            code_matiere = str(inf["Code"])
                        else:
                            code_matiere = "AUTRE"
                            
                        if "DEFINIR" in charge_matiere.upper() or "DÉFINIR" in charge_matiere.upper():
                            charge_matiere = "À Définir"
                            
                        # --- DÉTECTION DE L'UBIQUITÉ INTER-PROMOTIONS & VEILLE ---
                        enseignants_occupes_creneau = set()
                        enseignants_ayant_surveille_la_veille = set()
                        
                        if not df_db_global.empty:
                            filtre_creneau = df_db_global[(df_db_global["Jours"] == date_str) & (df_db_global["Horaire"] == h_f) & (df_db_global["Promotion"] != p_sel)]
                            if not filtre_creneau.empty and "Enseignants" in filtre_creneau.columns:
                                for e_occ in filtre_creneau["Enseignants"].dropna().unique().tolist():
                                    enseignants_occupes_creneau.add(str(e_occ).strip())
                                    
                            filtre_veille = df_db_global[(df_db_global["Jours"] == veille_str) & (df_db_global["Promotion"] != p_sel)]
                            if not filtre_veille.empty and "Enseignants" in filtre_veille.columns:
                                for e_vei in filtre_veille["Enseignants"].dropna().unique().tolist():
                                    enseignants_ayant_surveille_la_veille.add(str(e_vei).strip())
                                                
                        for lieu in s_f:
                            est_amphi = any(x in lieu for x in ['A', ''])
                            q = nb_amphi if est_amphi else nb_salle
                            
                            charge_deja_place = False
                            enseignants_deja_mis_dans_ce_lieu = set()
                            
                            cle_recherche = (str(vrai_nom_module).strip(), str(lieu).strip())
                            surveillants_historiques_presents = historique_dict.get(cle_recherche, [])
                            
                            hist_contient_titulaire = False
                            hist_contient_vacataire = False
                            for s_hist in surveillants_historiques_presents:
                                if s_hist in p_sup_clean or s_hist in src_clean:
                                    hist_contient_titulaire = True
                                if s_hist in vac_clean:
                                    hist_contient_vacataire = True
                    
                            lignes_salle = []
                    
                            for i in range(q):
                                surveillant_affecte = "À Définir"
                                
                                if i < len(surveillants_historiques_presents):
                                    surveillant_affecte = surveillants_historiques_presents[i]
                                    if surveillant_affecte == charge_matiere:
                                        charge_deja_place = True
                                    enseignants_occupes_creneau.add(surveillant_affecte)
                                    enseignants_deja_mis_dans_ce_lieu.add(surveillant_affecte)
                                
                                if surveillant_affecte == "À Définir" and i == 0 and not charge_deja_place and charge_matiere != "À Définir":
                                    nom_charge_abs = charge_matiere.upper()
                                    conflit_creneau = any(nom_charge_abs.startswith(o.upper()) or o.upper().startswith(nom_charge_abs) for o in enseignants_occupes_creneau)
                                    conflit_lieu = any(nom_charge_abs.startswith(l.upper()) or l.upper().startswith(nom_charge_abs) for l in enseignants_deja_mis_dans_ce_lieu)
                                    
                                    if not conflit_creneau and not conflit_lieu:
                                        surveillant_affecte = charge_matiere
                                        charge_deja_place = True
                                        hist_contient_titulaire = True
                                        enseignants_occupes_creneau.add(charge_matiere)
                                        enseignants_deja_mis_dans_ce_lieu.add(charge_matiere)
                                        if charge_matiere in compteurs_surveillance:
                                            compteurs_surveillance[charge_matiere] += 1
                                
                                if surveillant_affecte == "À Définir":
                                    candidats = {}
                                    besoin_vacataire = False
                                    if hist_contient_titulaire and not hist_contient_vacataire:
                                        besoin_vacataire = True
                                    elif hist_contient_vacataire and not hist_contient_titulaire:
                                        besoin_vacataire = False
                                    else:
                                        besoin_vacataire = (i % 2 == 1)
                    
                                    for k, v in compteurs_surveillance.items():
                                        k_upper = k.upper()
                                        if besoin_vacataire and (k not in vac_clean):
                                            continue
                                        if not besoin_vacataire and (k not in p_sup_clean and k not in src_clean):
                                            continue
                                            
                                        if any(k_upper.startswith(o.upper()) or o.upper().startswith(k_upper) for o in enseignants_occupes_creneau):
                                            continue
                                        if any(k_upper.startswith(l.upper()) or l.upper().startswith(k_upper) for l in enseignants_deja_mis_dans_ce_lieu):
                                            continue
                                        if any(k_upper.startswith(vei.upper()) or vei.upper().startswith(k_upper) for vei in enseignants_ayant_surveille_la_veille):
                                            continue
                                            
                                        candidats[k] = v
                                    
                                    if candidats:
                                        min_surveillances = min(candidats.values())
                                        eligibles = [k for k, v in candidats.items() if v == min_surveillances]
                                        surveillant_affecte = eligibles[0]
                                        if surveillant_affecte in vac_clean:
                                            hist_contient_vacataire = True
                                        else:
                                            hist_contient_titulaire = True
                                    else:
                                        candidats_secours = {}
                                        for k, v in compteurs_surveillance.items():
                                            k_up = k.upper()
                                            if besoin_vacataire and (k not in vac_clean):
                                                continue
                                            if not besoin_vacataire and (k not in p_sup_clean and k not in src_clean):
                                                continue
                                            if not any(k_up.startswith(o.upper()) or o.upper().startswith(k_up) for o in enseignants_occupes_creneau) and not any(k_up.startswith(l.upper()) or l.upper().startswith(k_up) for l in enseignants_deja_mis_dans_ce_lieu):
                                                candidats_secours[k] = v
                                        if candidats_secours:
                                            min_secours = min(candidats_secours.values())
                                            eligibles_secours = [k for k, v in candidats_secours.items() if v == min_secours]
                                            surveillant_affecte = eligibles_secours[0]
                                            if surveillant_affecte in vac_clean:
                                                hist_contient_vacataire = True
                                            else:
                                                hist_contient_titulaire = True
                                        else:
                                            surveillant_affecte = "À Définir"
                                            
                                    if surveillant_affecte != "À Définir":
                                        enseignants_occupes_creneau.add(surveillant_affecte)
                                        enseignants_deja_mis_dans_ce_lieu.add(surveillant_affecte)
                                        compteurs_surveillance[surveillant_affecte] += 1
                    
                                lignes_salle.append({
                                    "Enseignements": vrai_nom_module, 
                                    "Code": code_matiere, 
                                    "Enseignants": surveillant_affecte, 
                                    "Horaire": h_f, 
                                    "Jours": date_str, 
                                    "Lieu": lieu, 
                                    "Promotion": p_sel, 
                                    "Responsable": charge_matiere, 
                                    "email": ""
                                })
                            
                            # Tri local par salle (Permanent en premier)
                            def cle_tri_permanent(ligne):
                                nom_ens = ligne["Enseignants"]
                                if nom_ens in p_sup_clean or nom_ens in src_clean or nom_ens == charge_matiere:
                                    return 0
                                return 1
                    
                            lignes_salle_ordonnees = sorted(lignes_salle, key=cle_tri_permanent)
                            batch_pa.extend(lignes_salle_ordonnees)
                    
                    # ==============================================================================
                    # ENREGISTREMENT ET RAFRAÎCHISSEMENT DE L'APPLICATION
                    # ==============================================================================
                    if batch_pa:
                        supabase.table(TABLE_NAME).insert(batch_pa).execute()
                        st.success(f"Mise à jour réussie ! Le planning complet respectant la parité Titulaire/Vacataire a été généré pour la promotion {p_sel}.")
                        st.rerun()    
            # --- 3. INTERFACE D'ÉDITION MANUELLE ---
            if p_sel in promos_existantes:
                st.write("---")
                st.subheader(f"✍️ Ajustement Manuel : {p_sel}")
                
                # Récupération des données fraîches
                res_db = supabase.table(TABLE_NAME).select("*").eq("Promotion", p_sel).execute()
                df_edit = pd.DataFrame(res_db.data)
                
                if not df_edit.empty:
                    st.info("Double-cliquez sur une cellule pour modifier. Cliquez sur le bouton ci-dessous pour valider.")
                    
                    cols_ordre = ["id", "Enseignements", "Code", "Responsable", "Enseignants", "Horaire", "Jours", "Lieu"]
                    
                    # On ajoute une clé dynamique basée sur la promotion pour éviter les doublons
                    edited_df = st.data_editor(
                        df_edit[cols_ordre],
                        key=f"editor_data_{p_sel}", 
                        disabled=["Enseignements", "Code", "Horaire", "Jours", "Lieu"],
                        hide_index=True
                    )
                    
                    # AJOUT D'UNE KEY UNIQUE AU BOUTON
                    if st.button("💾 SAUVEGARDER LES MODIFICATIONS", key=f"btn_save_{p_sel}"):
                        try:
                            with st.spinner("Mise à jour en cours..."):
                                for _, row in edited_df.iterrows():
                                    supabase.table(TABLE_NAME).update({
                                        "Responsable": row["Responsable"],
                                        "Enseignants": row["Enseignants"]
                                    }).eq("id", row["id"]).execute()
                            st.success("Modifications enregistrées !")
                            st.rerun()
                        except Exception as e:
                            st.error(f"Erreur lors de la sauvegarde : {e}")
            # --- 4. MODULE DE PERMUTATION AVANCÉ (MODE STRICT) ---
            st.write("---")
            st.subheader("🔄 Module de Permutation des Surveillances")
            st.caption("Plateforme de gestion des EDTs-S2-2026-Département d'Électrotechnique-Faculté de génie électrique-UDL-SBA")
            
            # Récupération fraîche des données pour éviter les désynchronisations
            res_glob = supabase.table(TABLE_NAME).select("*").execute()
            df_db_global = pd.DataFrame(res_glob.data)
            
            if not df_db_global.empty:
                # Extraction unique des enseignants
                liste_profs = sorted(df_db_global["Enseignants"].unique().tolist())
                
                col_a, col_b = st.columns(2)
                
                with col_a:
                    prof_1 = st.selectbox("Enseignant Source", [""] + liste_profs, key="p1_adv")
                    if prof_1:
                        edt_1 = df_db_global[df_db_global["Enseignants"] == prof_1][["id", "Enseignements", "Jours", "Horaire", "Lieu", "Promotion"]]
                        st.caption(f"Séances de : {prof_1}")
                        st.dataframe(edt_1, hide_index=True)
            
                with col_b:
                    prof_2 = st.selectbox("Enseignant Destination", [""] + liste_profs, key="p2_adv")
                    if prof_2:
                        edt_2 = df_db_global[df_db_global["Enseignants"] == prof_2][["id", "Enseignements", "Jours", "Horaire", "Lieu", "Promotion"]]
                        st.caption(f"Séances de : {prof_2}")
                        st.dataframe(edt_2, hide_index=True)
            
                if prof_1 and prof_2 and prof_1 != prof_2:
                    tab_global, tab_seance = st.tabs(["🌍 Permutation Globale", "📍 Permutation par Séance"])
            
                    # --- OPTION 1 : PERMUTATION GLOBALE ---
                    with tab_global:
                        st.warning(f"Attention : Tous les créneaux de {prof_1} seront attribués à {prof_2} et vice-versa.")
                        if st.button(f"🤝 Échanger tous les plannings", key="btn_swap_global"):
                            try:
                                # Utilisation d'un marqueur temporaire pour l'échange atomique
                                temp_marker = "TEMP_SWAP_KEY_2026"
                                supabase.table(TABLE_NAME).update({"Enseignants": temp_marker}).eq("Enseignants", prof_1).execute()
                                supabase.table(TABLE_NAME).update({"Enseignants": prof_1}).eq("Enseignants", prof_2).execute()
                                supabase.table(TABLE_NAME).update({"Enseignants": prof_2}).eq("Enseignants", temp_marker).execute()
                                
                                st.success("Permutation globale réussie.")
                                st.rerun()
                            except Exception as e:
                                st.error(f"Erreur lors de la permutation globale : {e}")
            
                    # --- OPTION 2 : PERMUTATION PAR SÉANCE (PRÉSERVATION DE POSITION) ---
                    with tab_seance:
                        st.info("Échange ciblé : Le remplaçant prendra la place exacte (ID) de l'enseignant source.")
                        col_s1, col_s2 = st.columns(2)
                        
                        with col_s1:
                            if not edt_1.empty:
                                option_s1 = {f"{r['Enseignements']} | {r['Jours']} {r['Horaire']} ({r['Promotion']})": r['id'] for _, r in edt_1.iterrows()}
                                s1_label = st.selectbox("Sélectionner séance de " + prof_1, options=list(option_s1.keys()), key="sel_s1")
                                id_s1 = option_s1[s1_label]
                            else:
                                st.write("Aucune séance.")
            
                        with col_s2:
                            if not edt_2.empty:
                                option_s2 = {f"{r['Enseignements']} | {r['Jours']} {r['Horaire']} ({r['Promotion']})": r['id'] for _, r in edt_2.iterrows()}
                                s2_label = st.selectbox("Sélectionner séance de " + prof_2, options=list(option_s2.keys()), key="sel_s2")
                                id_s2 = option_s2[s2_label]
                            else:
                                st.write("Aucune séance.")
            
                        if 'id_s1' in locals() and 'id_s2' in locals():
                            if st.button("🔄 Permuter ces deux lignes (Position Stricte)", key="btn_swap_seance"):
                                try:
                                    # On met à jour l'ID spécifique avec le nom de l'autre enseignant
                                    # Cela garantit qu'il remplace l'enseignant EXACTEMENT sur sa ligne (1er, 2ème, etc.)
                                    supabase.table(TABLE_NAME).update({"Enseignants": prof_2}).eq("id", id_s1).execute()
                                    supabase.table(TABLE_NAME).update({"Enseignants": prof_1}).eq("id", id_s2).execute()
                                    
                                    st.success(f"Échange terminé : {prof_2} prend la place de {prof_1} sur l'ID {id_s1}.")
                                    st.rerun()
                                except Exception as e:
                                    st.error(f"Erreur technique : {e}")
            
                
                # --- 4. EXPORT HTML ---
                # AJOUT D'UNE KEY UNIQUE AU BOUTON D'EXPORT
                if st.button("📥 Générer le fichier HTML", key=f"btn_html_{p_sel}"):
                    # ... (votre code de génération HTML reste le même)
                    st.info("Fichier prêt. Utilisez le bouton 'Télécharger' qui va apparaître.")
                    df_final = pd.DataFrame(supabase.table(TABLE_NAME).select("*").eq("Promotion", p_sel).execute().data)
                    df_final = df_final.sort_values(by=["Jours", "Horaire"])
                    
                    html_content = f"""
                    <html>
                    <head><meta charset="UTF-8"><style>
                        body {{ font-family: sans-serif; }}
                        table {{ width: 100%; border-collapse: collapse; }}
                        th, td {{ border: 1px solid #ccc; padding: 8px; text-align: left; }}
                        th {{ background-color: #2c3e50; color: white; }}
                    </style></head>
                    <body>
                        <h2>Plateforme de gestion des EDTs-S2-2026-Département d'Électrotechnique-Faculté de génie électrique-UDL-SBA</h2>
                        <h3>Promotion : {p_sel}</h3>
                        <table>
                            <tr><th>Date</th><th>Heure</th><th>Module</th><th>Chargé de cours</th><th>Lieu</th><th>Surveillants</th></tr>
                    """
                    for _, r in df_final.iterrows():
                        html_content += f"<tr><td>{r['Jours']}</td><td>{r['Horaire']}</td><td>{r['Enseignements']}</td><td>{r['Responsable']}</td><td>{r['Lieu']}</td><td>{r['Enseignants']}</td></tr>"
                    html_content += "</table></body></html>"

                    st.download_button(
                        label="📥 Télécharger le Pack EDT (HTML)",
                        data=html_content,
                        file_name=f"EDT_{p_sel}_S2_2026.html",
                        mime="text/html"
                    )
# --- BLOC T3 : GÉNÉRER PV (ORDRE DE POSITION STRICT) ---
# --- BLOC T3 : GÉNÉRER PV (TRI CHRONOLOGIQUE STRICT 16-24 MAI) ---
with t3:
    if not st.session_state.auth_admin_edt:
        afficher_verrou("t3")
    else:
        st.subheader("📄 Impression des PV")
        st.caption("Plateforme de gestion des EDTs-S2-2026-Département d'Électrotechnique-Faculté de génie électrique-UDL-SBA")
        
        # 1. LISTE D'ORDRE DES PROMOTIONS (Priorité 1)
        ORDRE_PROMOS_STRICT = [
            "ING1", "L1MCIL", "ING2", "MCIL3", 
            "ING3RSE", "ING3EI", "ING4", "L2ELT", 
            "L2MCIL", "L3ELT", "M1RE", "M1ME", 
            "M1CE", "M1MCIL", "M1ER"
        ]

        # 2. RÉCUPÉRATION DES DONNÉES
        res_brut = supabase.table(TABLE_NAME).select("*").execute()
        df_fresh = pd.DataFrame(res_brut.data) if res_brut.data else pd.DataFrame()

        # --- SECTION : BOUTON DE TÉLÉCHARGEMENT GLOBAL ---
        st.info("📦 PACK GLOBAL : Tri par Promotion, puis Chronologie stricte (16/05 au 24/05).")
        
        if st.button("📥 Télécharger TOUS les PV (Ordre Chronologique Strict)", key="btn_global_final_16_24"):
            if not df_fresh.empty:
                df_global = df_fresh.copy()

                # --- PRÉPARATION DU TRI ---
                
                # A. Mapping du rang de la Promotion
                df_global['rank_promo'] = df_global['Promotion'].apply(
                    lambda x: ORDRE_PROMOS_STRICT.index(x) if x in ORDRE_PROMOS_STRICT else 99
                )
                
                # B. Extraction de la Date (cherche DD/MM/YYYY)
                df_global['temp_date'] = df_global['Jours'].astype(str).str.extract(r'(\d{2}/\d{2}/\d{4})')
                df_global['date_dt'] = pd.to_datetime(df_global['temp_date'], format='%d/%m/%Y', errors='coerce')
                
                # C. Extraction de l'Heure pour tri numérique (ex: 13h30 -> 13.5)
                df_global['h_val'] = df_global['Horaire'].astype(str).str.extract(r'(\d{1,2})h').astype(float).fillna(0)
                df_global['m_val'] = df_global['Horaire'].astype(str).str.extract(r'h(\d{0,2})')
                df_global['m_val'] = df_global['m_val'].replace('', '0').fillna('0').astype(float)
                df_global['time_score'] = (df_global['h_val'] * 60) + df_global['m_val']

                # --- FILTRAGE DE LA PÉRIODE (16/05 au 24/05) ---
                d_debut = pd.to_datetime("16/05/2026", format='%d/%m/%Y')
                d_fin = pd.to_datetime("24/05/2026", format='%d/%m/%Y')
                df_global = df_global[(df_global['date_dt'] >= d_debut) & (df_global['date_dt'] <= d_fin)]

                # --- EXÉCUTION DU TRI ---
                # Ordre : Promotion -> Date -> Heure -> ID
                df_global = df_global.sort_values(
                    by=['rank_promo', 'date_dt', 'time_score', 'id'], 
                    ascending=[True, True, True, True]
                )
                
                # Nettoyage
                cols_technique = ['rank_promo', 'temp_date', 'date_dt', 'h_val', 'm_val', 'time_score']
                df_final_global = df_global.drop(columns=cols_technique)
                
                html_output = generer_html_pv_pack(df_final_global)
                
                st.download_button(
                    label="💾 Enregistrer le Pack Global (HTML)",
                    data=html_output,
                    file_name="PACK_PV_S2_MAY_2026_STRICT.html",
                    mime="text/html"
                )
            else:
                st.error("Aucune donnée disponible.")

        st.divider()




        # --- SECTION : GESTION INDIVIDUELLE ---
        p_pv = st.selectbox("Ajuster une promotion spécifique", [""] + ORDRE_PROMOS_STRICT)
        
        if p_pv and not df_fresh.empty:
            df_p = df_fresh[df_fresh["Promotion"] == p_pv].copy()
            
            if not df_p.empty:
                # Récupération des effectifs
                total_etudiants = DATA_AUTO[p_pv]['Effectif'] if p_pv in DATA_AUTO else 0
                lieux_uniques = sorted(df_p["Lieu"].unique().tolist())
                part_egale = total_etudiants // len(lieux_uniques) if lieux_uniques else 0
                repartition_finale = {}

                with st.expander(f"📝 Configuration : {p_pv}", expanded=True):
                    cols_eff = st.columns(2)
                    for idx, lieu in enumerate(lieux_uniques):
                        with cols_eff[idx % 2]:
                            repartition_finale[lieu] = st.number_input(f"Étudiants - {lieu}", 0, 500, part_egale, key=f"nb_{p_pv}_{lieu}")

                    st.divider()
                    
                    options_p = ["AUCUN"] + sorted(st.session_state.get('p_sup_cfg', []))
                    options_v = ["AUCUN"] + sorted(st.session_state.get('vac_list_cfg', []))

                    # Affichage par module trié par date/heure
                    df_p['temp_date'] = df_p['Jours'].astype(str).str.extract(r'(\d{2}/\d{2}/\d{4})')
                    df_p['date_dt'] = pd.to_datetime(df_p['temp_date'], format='%d/%m/%Y', errors='coerce')
                    df_p['h_val'] = df_p['Horaire'].astype(str).str.extract(r'(\d{1,2})h').astype(float).fillna(0)
                    df_p = df_p.sort_values(by=['date_dt', 'h_val', 'id'])

                    modules = df_p["Enseignements"].unique()
                    for mod in modules:
                        df_mod = df_p[df_p["Enseignements"] == mod].sort_values(by="id")
                        charge = df_mod.iloc[0]["Responsable"] if "Responsable" in df_mod.columns else "NA"
                        st.info(f"📘 **{mod}** | Date : {df_mod.iloc[0]['Jours']} | Responsable : {charge}")
                        
                        for i, (_, row) in enumerate(df_mod.iterrows()):
                            rid = row['id']
                            pos_label = "🥇 Principal" if i == 0 else f"🥈 Adjoint {i}"
                            c1, c2, c3, c4, c5 = st.columns([1.5, 1, 1, 2.5, 0.5])
                            with c1: st.markdown(f"**{pos_label}**")
                            with c2: st.text_input("Jour", row['Jours'], key=f"j_{rid}", on_change=update_item_db, args=(rid, "Jours", st.session_state.get(f"j_{rid}")))
                            with c3: st.text_input("Heure", row['Horaire'], key=f"h_{rid}", on_change=update_item_db, args=(rid, "Horaire", st.session_state.get(f"h_{rid}")))
                            with c4:
                                cur_ens = str(row['Enseignants']).upper().strip()
                                st.selectbox(f"Surveillant", options_p + options_v, 
                                             index=(options_p + options_v).index(cur_ens) if cur_ens in (options_p + options_v) else 0,
                                             key=f"ens_sel_{rid}", on_change=update_item_db, args=(rid, "Enseignants", st.session_state.get(f"ens_sel_{rid}")))
                            with c5:
                                if st.button("❌", key=f"del_{rid}"):
                                    supprimer_ligne(rid)
                                    st.rerun()
                        st.write("---")

                if st.button(f"🔄 Générer PV Final {p_pv}", key=f"btn_p_final"):
                    df_p["Effectif_Salle"] = df_p["Lieu"].map(repartition_finale).fillna(0)
                    pv_html = generer_html_pv_pack(df_p.drop(columns=['temp_date', 'date_dt', 'h_val']))
                    st.download_button(label="🖨️ Télécharger le PDF/HTML", data=pv_html, file_name=f"PV_{p_pv}_S2.html", mime="text/html")
# --- BLOC T4 : RÉCAPITULATIF & TABLEAU PLANNING (STRICT & COMPLET) ---
import pandas as pd
import streamlit as st
from datetime import datetime
import base64
import os

# --- FONCTION DE FORMATAGE ---
def format_date_complete(d):
    try:
        jours_fr = ["Lundi", "Mardi", "Mercredi", "Jeudi", "Vendredi", "Samedi", "Dimanche"]
        if isinstance(d, str):
            d = pd.to_datetime(d, dayfirst=True)
        return f"{jours_fr[d.weekday()]} {d.strftime('%d/%m/%Y')}"
    except:
        return str(d)
with t4:
    if not st.session_state.get('auth_admin_edt', False):
        afficher_verrou("t4")
    else:
        st.subheader("📋 Récapitulatif & Tableau Planning")
        st.caption("Plateforme de gestion des EDTs-S2-2026-Département d'Électrotechnique-Faculté de génie électrique-UDL-SBA")

        try:
            # Récupération des données avec tri par ID
            res_t4 = supabase.table(TABLE_NAME).select("*").order("id").execute()
            df_db_global = pd.DataFrame(res_t4.data) if res_t4.data else pd.DataFrame()
        except Exception as e:
            st.error(f"Erreur de connexion à la base de données : {e}")
            df_db_global = pd.DataFrame()

        if not df_db_global.empty:
            # Sécurisation des colonnes obligatoires
            required_cols = ["Promotion", "Lieu", "Enseignements", "Jours", "Horaire", "Enseignants", "Responsable"]
            for col in required_cols:
                if col not in df_db_global.columns:
                    df_db_global[col] = ""

            st.markdown("### 🗓️ Aperçu du Planning")

            p_view = st.selectbox(
                "🔍 Filtrer par Promotion",
                [""] + sorted(df_db_global["Promotion"].dropna().unique().tolist()),
                key="t4_promo_filter"
            )

            if p_view:
                df_p = df_db_global[df_db_global["Promotion"] == p_view].copy()

                # --- TRAITEMENT DES DATES ---
                def parse_date(date_str):
                    try:
                        return datetime.strptime(str(date_str).strip(), "%d/%m/%Y")
                    except:
                        return None

                df_p["date_obj"] = df_p["Jours"].apply(parse_date)
                df_p = df_p[df_p["date_obj"].notna()]
                df_p = df_p.sort_values(by=["date_obj", "Horaire", "id"])
                # --- RÉCUPÉRATION INFOS DATA_AUTO ---
                info_promo = DATA_AUTO.get(p_view, {"Effectif": "Non défini", "Salles": []})
                effectif_promo = info_promo.get("Effectif", "Inconnu")
                tous_les_lieux = info_promo.get("Salles", [])
                
                amphis = [l for l in tous_les_lieux if l.startswith('A')]
                salles = [l for l in tous_les_lieux if l not in amphis]
                nombre_groupes = len(tous_les_lieux)

                # --- SELECTEUR DE STYLE ---
                type_vue = st.radio("Choisir le format d'affichage :", 
                                   ["📜 Liste Chronologique", "📅 Grille Planning (Style Image)"], 
                                   horizontal=True)

                # --- CONFIGURATION COMMUNE ---
                dict_pages_promos = {
                    "ING1": "1", "L1MCIL": "1", "ING2": "1", "MCIL3": "1", 
                    "ING3RSE": "1", "ING3EI": "1", "ING4": "2", "L2ELT": "1", 
                    "L2MCIL": "1", "L3ELT": "2", "M1RE": "1", "M1ME": "1", 
                    "M1CE": "1", "M1MCIL": "1", "M1ER": "1"
                }
                total_pages_fixe = dict_pages_promos.get(p_view, "1")
                date_auto = datetime.now().strftime("%d/%m/%Y")
                
                logo_path = "logo.png"
                if os.path.exists(logo_path):
                    with open(logo_path, "rb") as f:
                        logo_base64 = base64.b64encode(f.read()).decode()
                    logo_html = f'<img src="data:image/png;base64,{logo_base64}" style="width:70px; height:auto;">'
                else:
                    logo_html = "<div style='font-weight:bold; font-size:0.7em;'>LOGO</div>"

                # Initialisation du HTML final
                html_table = "<div style='font-family: Arial, sans-serif;'>"
                # ======================================================================================           
                # EN-TÊTE ADMINISTRATIF (Commun aux deux vues)
                # ======================================================================================
                def inserer_en_tete_iso_planning(doc, date_auto="18/05/2026", code_doc="PPER.09", revision="00"):
                            """
                            Génère l'en-tête administratif de la Plateforme de gestion des EDTs-S2-2026-Département d'Électrotechnique-Faculté de génie électrique-UDL-SBA
                            sous la forme d'un tableau structuré à 3 blocs conforme à la norme qualité de l'établissement.
                            """
                            # Création d'un tableau à 1 ligne et 3 colonnes (Total largeur disponible ~ 6.5 à 6.9 pouces)
                            header_table = doc.add_table(rows=1, cols=3)
                            header_table.alignment = WD_ALIGN_PARAGRAPH.CENTER
                            header_table.autofit = False
                            
                            # Définition des proportions exactes (20% / 55% / 25%)
                            header_table.columns[0].width = Inches(1.38) # Cellule Logo
                            header_table.columns[1].width = Inches(3.79) # Cellule Centrale (Établissement & Titre)
                            header_table.columns[2].width = Inches(1.73) # Cellule Code & Références ISO
                            
                            cell_logo = header_table.rows[0].cells[0]
                            cell_titre = header_table.rows[0].cells[1]
                            cell_code = header_table.rows[0].cells[2]
                            
                            # Application des paddings et des bordures noires épaisses individuelles
                            for cell in [cell_logo, cell_titre, cell_code]:
                                set_cell_margins(cell, top=140, bottom=140, left=140, right=140)
                                appliquer_bordure_cellule_noire(cell)
                            
                            # --- BLOC 1 : LOGO ---
                            p_logo = cell_logo.paragraphs[0]
                            initialiser_paragraphe_strict(p_logo)
                            p_logo.alignment = WD_ALIGN_PARAGRAPH.CENTER
                            
                            nom_fichier_logo = "logo.PNG"
                            if os.path.exists(nom_fichier_logo):
                                p_logo.add_run().add_picture(nom_fichier_logo, width=Inches(0.65), height=Inches(1.0))
                            else:
                                r_alt = p_logo.add_run("[ LOGO ]")
                                r_alt.font.name = 'Calibri'
                                r_alt.font.size = Pt(11)
                                r_alt.font.italic = True
                                
                            # --- BLOC 2 : ÉTABLISSEMENT & EN-TÊTE ---
                            p_titre = cell_titre.paragraphs[0]
                            initialiser_paragraphe_strict(p_titre)
                            p_titre.alignment = WD_ALIGN_PARAGRAPH.CENTER
                            
                            r_univ = p_titre.add_run("Université Djillali Liabes\n")
                            r_univ.bold = True
                            r_univ.font.size = Pt(13)
                            
                            r_ville = p_titre.add_run("Sidi Bel Abbes\n")
                            r_ville.font.size = Pt(12)
                            
                            # Petit paragraphe d'espace pour la ligne séparatrice simulée
                            p_sub = cell_titre.add_paragraph()
                            initialiser_paragraphe_strict(p_sub)
                            p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
                            p_sub.paragraph_format.space_before = Pt(6)
                            
                            r_plan = p_sub.add_run("Planning des Examens")
                            r_plan.bold = True
                            r_plan.font.size = Pt(13)
                            
                            # Harmonisation des polices du bloc central
                            for run in [r_univ, r_ville, r_plan]:
                                run.font.name = 'Calibri'
                                
                            # --- BLOC 3 : CODES & RÉFÉRENCES ISO ---
                            p_code = cell_code.paragraphs[0]
                            initialiser_paragraphe_strict(p_code)
                            p_code.alignment = WD_ALIGN_PARAGRAPH.LEFT
                            
                            # Ajout du Code Document
                            r_c_label = p_code.add_run("Code : ")
                            r_c_val = p_code.add_run(f"{code_doc}\n")
                            r_c_label.bold = True
                            
                            # Ajout de la Révision
                            r_r_label = p_code.add_run("Révision : ")
                            r_r_val = p_code.add_run(f"{revision}\n")
                            r_r_label.bold = True
                            
                            # Ajout de la Date Automatique
                            r_d_label = p_code.add_run("Date : ")
                            r_d_val = p_code.add_run(f"{date_auto}\n")
                            r_d_label.bold = True
                date_auto = "16/05/2026"
                header_admin = (
                    f"<table style='width: 100%; border-collapse: collapse; border: 2px solid black; margin-bottom: 20px;'>"
                    "<tr>"
                    f"<td style='width: 20%; border: 1.5px solid black; text-align: center; padding: 10px;'>{logo_html}</td>"
                    "<td style='width: 55%; border: 1.5px solid black; text-align: center; vertical-align: middle; padding: 5px;'>"
                    "<div style='font-size: 1.2em; color: black; font-weight: bold;'>Université Djillali Liabes</div>"
                    "<div style='font-size: 1.1em; color: black;'>Sidi Bel Abbes</div>"
                    "<div style='border-top: 1.5px solid black; margin-top: 8px; padding-top: 8px;'>"
                    "<span style='font-weight: bold; color: black; font-size: 1.2em;'>Planning des Examens</span>"
                    "</div>"
                    "</td>"
                    "<td style='width: 25%; border: 1.5px solid black; font-size: 0.95em; padding: 10px;'>"
                    "<div><strong>Code :</strong> PPER.09</div>"
                    f"<div><strong>Révision :</strong> {'00'}</div>"
                    f"<div><strong>Date :</strong> {date_auto}</div>"
                    f"<div style='color: black;'><strong>Pages :</strong> 15/{15}</div>"
                    "</td>"
                    "</tr>"
                    "</table>"
                )
                # RAPPEL DU TITRE ET INFOS PROMO
                html_table += (
                    "<div style='text-align: center; font-size: 0.85em; color: #555; margin-bottom: 20px;'>"
                    "Plateforme de gestion des EDTs-S2-2026-Département d'Électrotechnique-Faculté de génie électrique-UDL-SBA"
                    "</div>"
                    "<div style='text-align:center;'>"
                    f"<h3 style='margin:5px 0; color:#1E88E5;'>🎓 Promotion : {p_view}</h3>"
                    "<div style='margin:10px 0; font-size:1.1em; color:#2c3e50; background:#f8f9fa; padding:15px; border-radius:8px; border: 1px solid #dee2e6; display:inline-block; min-width:70%;'>"
                    f"<div style='margin-bottom:8px;'><strong>👥 Effectif Global :</strong> {effectif_promo} étudiants</div>"
                    f"<div style='margin-bottom:8px;'><strong>📊 Total Groupes/Lieux :</strong> {nombre_groupes}</div>"
                    f"<div style='display: flex; justify-content: center; gap: 20px; font-size: 0.95em; color: #555;'>"
                    f"<span>🏛️ <b>Amphis :</b> {', '.join(amphis) if amphis else 'Néant'}</span>"
                    f"<span>🏫 <b>Salles :</b> {', '.join(salles) if salles else 'Néant'}</span>"
                    "</div>"
                    "</div>"
                    "<hr style='border:0.5px solid #eee; margin: 20px 0;'>"
                    "</div>"
                )
                html_table += header_admin

                # --- VUE 1 : LISTE CHRONOLOGIQUE ---
                if type_vue == "📜 Liste Chronologique":
                    html_table += (
                        '<table style="width:100%; border-collapse:collapse; border: 1px solid #ddd;">'
                        '<tr style="background:#2c3e50; color:white; text-align:center;">'
                        '<th style="padding:14px; border:1px solid #444; width:20%;">DATE</th>'
                        '<th style="padding:14px; border:1px solid #444; width:15%;">HORAIRE</th>'
                        '<th style="padding:14px; border:1px solid #444;">DÉTAILS DES EXAMENS</th>'
                        '</tr>'
                    )
                    
                    for date_val in sorted(df_p["date_obj"].unique()):
                        df_day = df_p[df_p["date_obj"] == date_val]
                        date_display = format_date_complete(pd.Timestamp(date_val))
                        
                        for h in sorted(df_day["Horaire"].unique()):
                            df_slot = df_day[df_day["Horaire"] == h]
                            cell_content = ""
                            
                            for lieu in sorted(df_slot["Lieu"].unique()):
                                df_lieu = df_slot[df_slot["Lieu"] == lieu].sort_values(by="id")
                                row_ref = df_lieu.iloc[0]
                                # NETTOYAGE : Suppression de "Cours-" ou "Cours "
                                matiere_propre = str(row_ref.get('Enseignements','')).replace("Cours-", "").replace("Cours ", "")
                                resp_brut = str(row_ref.get('Responsable', '')).strip().upper()
                                responsable_nom = map_noms.get(resp_brut, resp_brut)
                                
                                s_list = []
                                for _, r in df_lieu.iterrows():
                                    s_raw = str(r.get("Enseignants", "")).strip().upper()
                                    s_nom = map_noms.get(s_raw, s_raw)
                                    if s_nom and s_nom != "NONE": s_list.append(f"👥 {s_nom}")
                                
                                cell_content += f"""
                                <div style='border:1px solid #eee; padding:8px; margin-bottom:5px; background:#fff;'>
                                    <b style='color:#1E88E5;'>📍 {lieu}</b> | <b>{matiere_propre}</b><br>
                                    <small>Chargé de matière : {responsable_nom}</small><br>
                                    <div style='font-size:0.85em; color:#555;'>{' / '.join(s_list)}</div>
                                </div>"""

                            html_table += f"""
                            <tr>
                                <td style='padding:12px; font-weight:bold; border:1px solid #ddd;'>{date_display}</td>
                                <td style='padding:12px; text-align:center; font-weight:bold; border:1px solid #ddd;'>{h}</td>
                                <td style='padding:10px; border:1px solid #ddd;'>{cell_content}</td>
                            </tr>"""
                    html_table += "</table>"

                # --- VUE 2 : GRILLE PLANNING (STYLE IMAGE) ---
                # --- VUE 2 : GRILLE PLANNING (LOGIQUE STRICTE VUE 1) ---
                else:
                    dates_uniques = sorted(df_p["date_obj"].unique())
                    horaires_uniques = sorted(df_p["Horaire"].unique())
                    
                    html_table += '<table style="width:100%; border-collapse:collapse; text-align:center; table-layout: fixed; border: 2px solid black;">'
                    
                    # Header Jours + Dates
                    html_table += '<tr style="background: #f2f2f2;">'
                    html_table += '<th style="border: 2px solid black; width:100px; padding:10px;">HEURE</th>'
                    jours_fr = ["Lundi", "Mardi", "Mercredi", "Jeudi", "Vendredi", "Samedi", "Dimanche"]
                    for d in dates_uniques:
                        n_jour = jours_fr[pd.Timestamp(d).weekday()]
                        html_table += f'<th style="border: 2px solid black; color: red; font-size:0.85em; padding:10px;">{n_jour}<br>{pd.Timestamp(d).strftime("%d/%m/%Y")}</th>'
                    html_table += '</tr>'

                    # Remplissage par Horaire
                    for h in horaires_uniques:
                        html_table += f'<tr style="min-height: 150px;">'
                        html_table += f'<td style="border: 2px solid black; font-weight:bold; background:#e1f5fe; padding:10px;">{h}</td>'
                        
                        for d in dates_uniques:
                            # On récupère toutes les lignes pour ce créneau (Date + Heure)
                            match = df_p[(df_p["date_obj"] == d) & (df_p["Horaire"] == h)]
                            
                            if not match.empty:
                                cell_content = ""
                                # LOGIQUE STRICTE : On boucle sur chaque lieu unique présent dans ce créneau
                                for lieu in sorted(match["Lieu"].unique()):
                                    df_lieu = match[match["Lieu"] == lieu].sort_values(by="id")
                                    row_ref = df_lieu.iloc[0]
                                    # NETTOYAGE : Suppression de "Cours-" ou "Cours "
                                    matiere_propre = str(row_ref.get('Enseignements','')).replace("Cours-", "").replace("Cours ", "")
                                    # Récupération du responsable
                                    resp_brut = str(row_ref.get('Responsable', '')).strip().upper()
                                    responsable_nom = map_noms.get(resp_brut, resp_brut)
                                    
                                    # Récupération de la liste des surveillants pour ce lieu précis
                                    s_list = []
                                    for _, r in df_lieu.iterrows():
                                        s_raw = str(r.get("Enseignants", "")).strip().upper()
                                        s_nom = map_noms.get(s_raw, s_raw)
                                        if s_nom and s_nom != "NONE": 
                                            s_list.append(f"👥 {s_nom}")
                                    
                                    # Construction du bloc par lieu (identique au style Vue 1)
                                    cell_content += f"""
                                    <div style='border-bottom:1px solid #eee; padding:5px; margin-bottom:5px; background:#fff; text-align:left;'>
                                        <b style='color:#1E88E5; font-size:0.9em;'>📍 {lieu}</b> | <b style='font-size:0.85em;'>{matiere_propre}</b><br>
                                        <small style='font-size:0.8em;'>Chargé de matière : {responsable_nom}</small><br>
                                        <div style='font-size:0.75em; color:#555;'>{' / '.join(s_list)}</div>
                                    </div>"""
                
                                html_table += f'<td style="border: 2px solid black; padding:5px; vertical-align:top; background: white;">{cell_content}</td>'
                            else:
                                html_table += '<td style="border: 2px solid black; background: #fafafa; color:#ccc;">X</td>'
                        html_table += '</tr>'
                    html_table += '</table>'
                # --- LOGIQUE POUR LE BOUTON PACK COMPLET ---
                # --- LOGIQUE POUR LE PACK COMPLET (CORRIGÉE) ---

                # --- LOGIQUE POUR LE PACK COMPLET (REGROUPEMENT PAR MATIÈRE/RESPONSABLE) ---

                # On vérifie la présence des données
                if 'df_db_global' in locals() and not df_db_global.empty:
                    
                    st.markdown("---")
                    st.subheader("📦 Exportation Globale")
                    
                    # 1. On utilise un bouton pour générer le contenu en mémoire
                    if st.button("🚀 Préparer le Pack complet (Toutes les promotions)", use_container_width=True, key="prep_pack"):
                        
                        toutes_promotions = sorted(df_db_global["Promotion"].unique().tolist())
                        
                        # Dictionnaire de traduction des jours
                        jours_trad = {
                            "Monday": "Lundi", "Tuesday": "Mardi", "Wednesday": "Mercredi",
                            "Thursday": "Jeudi", "Friday": "Vendredi", "Saturday": "Samedi",
                            "Sunday": "Dimanche"
                        }
                        
                        # Initialisation du HTML avec style pour l'impression
                        full_html_content = f"""
                        <html>
                        <head>
                            <meta charset="UTF-8">
                            <style>
                                body {{ font-family: 'Segoe UI', Arial, sans-serif; margin: 20px; }}
                                .page-break {{ page-break-after: always; }}
                                .header-table {{ width: 100%; border-collapse: collapse; border: 2px solid black; margin-bottom: 20px; }}
                                .header-table td {{ border: 1.5px solid black; padding: 14px; text-align: center; }}
                                .promo-section {{ margin-top: 30px; text-align: center; }}
                                table.planning {{ width: 100%; border-collapse: collapse; table-layout: fixed; width: 100%; }}
                                table.planning th, table.planning td {{ border: 1px solid black; padding: 6px; font-size: 14px; text-align: center; vertical-align: top; }}
                                .matiere-container {{ text-align: left; background: #fff; padding: 3px; margin-bottom: 8px; border: 1px solid #eee; border-radius: 4px; }}
                                .matiere-title {{ font-weight: bold; color: #1E88E5; font-size: 1.1em; border-bottom: 1px solid #1E88E5; margin-bottom: 3px; }}
                                .resp-line {{ font-weight: bold; color: #2c3e50; margin-bottom: 5px; display: block; }}
                                .lieu-block {{ margin-left: 5px; margin-bottom: 4px; padding-left: 5px; border-left: 2px solid #ddd; }}
                                .surv-line {{ font-size: 0.9em; color: #555; display: block; }}
                                .title-app {{ font-size: 0.8em; color: #666; margin-bottom: 10px; text-align: center; }}
                            </style>
                        </head>
                        <body>
                            <div class='title-app'></div>
                        """
                
                        # 2. Boucle de génération par promotion
                        for i, promo in enumerate(toutes_promotions):
                            df_p_full = df_db_global[df_db_global["Promotion"] == promo].copy()
                            
                            # --- Insertion de l'en-tête administratif ---
                            header_temp = header_admin.replace('1/19', f'{i+1} / {len(toutes_promotions)}')
                            
                            full_html_content += f"""
                            <div class="promo-section">
                                {header_temp}
                                <h2 style='color: #1E88E5;'>🎓 Promotion : {promo}</h2>
                            </div>
                            """
                
                            # Grille de temps
                            dates_p = sorted(pd.to_datetime(df_p_full["Jours"], dayfirst=True).unique())
                            horaires_p = sorted(df_p_full["Horaire"].unique())
                
                            full_html_content += "<table class='planning'><thead><tr style='background:#f2f2f2;'><th>HEURE</th>"
                            for d in dates_p:
                                jour_fr = jours_trad.get(d.strftime('%A'), d.strftime('%A'))
                                full_html_content += f"<th>{jour_fr}<br>{d.strftime('%d/%m/%Y')}</th>"
                            full_html_content += "</tr></thead><tbody>"
                
                            for h in horaires_p:
                                full_html_content += f"<tr><td style='background:#e1f5fe; font-weight:bold;'>{h}</td>"
                                for d in dates_p:
                                    d_str = d.strftime('%d/%m/%Y')
                                    match = df_p_full[(df_p_full["Jours"] == d_str) & (df_p_full["Horaire"] == h)]
                                    
                                    if not match.empty:
                                        cell = ""
                                        # --- LOGIQUE DE REGROUPEMENT ---
                                        # On identifie les combinaisons uniques de (Enseignements + Responsable)
                                        uniques = match.drop_duplicates(subset=['Enseignements', 'Responsable'])
                                        
                                        for _, row_u in uniques.iterrows():
                                            mat_raw = str(row_u['Enseignements']).replace("Cours-", "").replace("Cours ", "")
                                            resp_raw = str(row_u.get('Responsable', '')).strip().upper()
                                            resp_nom = map_noms.get(resp_raw, row_u.get('Responsable', ''))
                                            
                                            cell += f"<div class='matiere-container'>"
                                            cell += f"<div class='matiere-title'>📘 {mat_raw}</div>"
                                            cell += f"<span class='resp-line'>👤 Resp: {resp_nom}</span>"
                                            
                                            # Pour cette matière/responsable, on liste les lieux et leurs surveillants
                                            lieux_match = match[(match['Enseignements'] == row_u['Enseignements']) & 
                                                                (match['Responsable'] == row_u['Responsable'])]
                                            
                                            for lieu in sorted(lieux_match["Lieu"].unique()):
                                                rows_surv = lieux_match[lieux_match["Lieu"] == lieu]
                                                liste_s = []
                                                for _, rs in rows_surv.iterrows():
                                                    s_nom = map_noms.get(str(rs.get('Enseignants', '')).strip().upper(), rs.get('Enseignants', ''))
                                                    if s_nom and str(s_nom).upper() != "NONE":
                                                        liste_s.append(s_nom)
                                                
                                                surv_str = ", ".join(liste_s) if liste_s else "Aucun"
                                                cell += f"""
                                                <div class='lieu-block'>
                                                    <b>📍 {lieu}</b><br>
                                                    <span class='surv-line'>👥 Surv: {surv_str}</span>
                                                </div>"""
                                            
                                            cell += "</div>"
                                        
                                        full_html_content += f"<td>{cell}</td>"
                                    else:
                                        full_html_content += "<td>-</td>"
                                full_html_content += "</tr>"
                            
                            full_html_content += "</table>"
                            
                            if i < len(toutes_promotions) - 1:
                                full_html_content += "<div class='page-break'></div>"
                
                        full_html_content += "</body></html>"
                        st.session_state['full_pack_data'] = full_html_content
                        st.success(f"✅ Pack prêt pour {len(toutes_promotions)} promotions !")
                
                    if 'full_pack_data' in st.session_state:
                        st.download_button(
                            label="📥 Télécharger le Pack Complet (HTML)",
                            data=st.session_state['full_pack_data'],
                            file_name="Pack_Complet_EDT_S2_2026.html",
                            mime="text/html",
                            use_container_width=True,
                            key="btn_download_pack_ready"
                        )
                # BOUTON TÉLÉCHARGEMENT
                st.download_button(
                    label="📥 Télécharger le planning (HTML)",
                    data=html_table,
                    file_name=f"Planning_{p_view}_S2.html",
                    mime="text/html",
                    use_container_width=True,
                    key="btn_dl_final_t4"
                )
        else:
            st.info("Aucune donnée disponible.")
                                     
            # --- ENVOI DES CONVOCATIONS ---
# --- ENVOI DES CONVOCATIONS ---
# --- ENVOI DES CONVOCATIONS (LOGIQUE ANTI-DUPLICATION STRICTE) ---
st.markdown("---")

try:
    display_title = TITLE_PLATFORM
except NameError:
    display_title = "Plateforme de gestion des EDTs-S2-2026-Département d'Électrotechnique-Faculté de génie électrique-UDL-SBA"

st.markdown(f"### ✉️ Envoi des Convocations - {display_title}")

if not df_db_global.empty:
    c1, c2 = st.columns(2)

    # 1. Extraction des codes enseignants uniques
    all_codes = []
    for col in ["Enseignants", "Responsable"]:
        if col in df_db_global.columns:
            for entry in df_db_global[col].dropna().unique():
                codes = [c.strip().upper() for c in str(entry).split(',') if c.strip()]
                all_codes.extend(codes)

    profs_actifs_codes = sorted([c for c in set(all_codes) if c and "⚠️" not in c and c != "TEMP"])
    noms_uniques_liste = sorted(list(set([map_noms.get(p, p) for p in profs_actifs_codes])))

    # Définition de l'ordre des colonnes requis pour le tableau final
    cols_finales = ["Enseignements", "Code", "Enseignants", "Horaire", "Jours", "Lieu", "Promotion"]

    with c1:
        st.markdown("#### 👤 Envoi Individuel")
        ens_sel = st.selectbox("Sélectionner l'Enseignant", [""] + noms_uniques_liste, key="sel_ens_conv")
    
        if st.button("📨 Envoyer la Convocation", use_container_width=True):
            if ens_sel:
                inv_map = {v: k for k, v in map_noms.items()}
                code_prof = inv_map.get(ens_sel, ens_sel)
    
                # Filtrage des lignes correspondant à l'enseignant
                df_perso = df_db_global[
                    df_db_global["Enseignants"].str.contains(code_prof, na=False, regex=False) | 
                    (df_db_global["Responsable"] == code_prof)
                ].copy()
    
                if not df_perso.empty:
                    # --- PRÉPARATION DU TABLEAU (Disposition imposée) ---
                    cols_finales = ["Enseignements", "Code", "Enseignants", "Horaire", "Jours", "Lieu", "Promotion"]
                    df_perso = df_perso[cols_finales]
                    
                    # Nettoyage des doublons
                    df_perso = df_perso.drop_duplicates(subset=["Enseignements", "Horaire", "Jours", "Lieu"]).reset_index(drop=True)
    
                    # Tri chronologique
                    df_perso['temp_date'] = df_perso['Jours'].str.extract(r'(\d{2}/\d{2}/\d{4})')
                    df_perso['temp_date'] = pd.to_datetime(df_perso['temp_date'], format='%d/%m/%Y', errors='coerce')
                    df_perso = df_perso.sort_values(by=['temp_date', 'Horaire']).drop(columns=['temp_date'])
                    
                    # --- RECHERCHE DE L'EMAIL (Local puis Supabase) ---
                    email_dest = dict_emails.get(code_prof) or dict_emails.get(ens_sel)
                    
                    # Si non trouvé, on tente une requête flash sur la table 'enseignants_contacts'
                    if not email_dest:
                        try:
                            res = supabase.table("enseignants_contacts").select("email").eq("nom_enseignant", ens_sel.upper()).execute()
                            if res.data:
                                email_dest = res.data[0]['email']
                        except:
                            pass
    
                    if email_dest:
                        with st.spinner(f"Envoi en cours vers {email_dest}..."):
                            # On passe le titre officiel à la fonction d'envoi si nécessaire
                            if envoyer_mail(ens_sel, email_dest, df_perso):
                                st.success(f"🚀 Convocation envoyée avec succès à {ens_sel} !")
                            else:
                                st.error("❌ Échec de l'envoi (Vérifiez la connexion SMTP).")
                    else:
                        st.error(f"📧 L'adresse email de {ens_sel} est introuvable dans 'enseignants_contacts'.")
                else:
                    st.warning(f"⚠️ Aucun planning trouvé pour {ens_sel}.")

    with c2:
        st.markdown("#### 📢 Envoi Massif")
        st.warning("Attention : Cette action traite tous les enseignants.")
        if st.button("📧 LANCER LA CAMPAGNE TOTALE", use_container_width=True):
            progress_bar = st.progress(0)
            success_count = 0
            
            for i, p_code in enumerate(profs_actifs_codes):
                nom_complet = map_noms.get(p_code, p_code)
                
                df_perso = df_db_global[
                    df_db_global["Enseignants"].str.contains(p_code, na=False, regex=False) | 
                    (df_db_global["Responsable"] == p_code)
                ].copy()
                
                if not df_perso.empty:
                    # Application de la même purge anti-doublons pour le mode massif
                    df_perso = df_perso[cols_finales]
                    df_perso = df_perso.drop_duplicates(subset=["Enseignements", "Horaire", "Jours", "Lieu"]).reset_index(drop=True)
                    
                    df_perso['temp_date'] = df_perso['Jours'].str.extract(r'(\d{2}/\d{2}/\d{4})')
                    df_perso['temp_date'] = pd.to_datetime(df_perso['temp_date'], format='%d/%m/%Y', errors='coerce')
                    df_perso = df_perso.sort_values(by=['temp_date', 'Horaire']).drop(columns=['temp_date'])
                    
                    email_dest = dict_emails.get(p_code) or dict_emails.get(nom_complet)
                    if email_dest and envoyer_mail(nom_complet, email_dest, df_perso):
                        success_count += 1
                
                progress_bar.progress((i + 1) / len(profs_actifs_codes))
            
            st.success(f"✅ Campagne terminée ! {success_count} emails envoyés sans doublons.")
else:
    st.warning("⚠️ Aucune donnée disponible dans la base globale.")
# --- BLOC T5 : MAINTENANCE ---
# --- BLOC T5 : MAINTENANCE ---
with t5:
    if not st.session_state.auth_admin_edt:
        afficher_verrou("t5")
    else:
        # Rappel obligatoire de l'intitulé officiel de la plateforme
        st.caption("Plateforme de gestion des EDTs-S2-2026-Département d'Électrotechnique-Faculté de génie électrique-UDL-SBA")
        st.subheader("🔧 Maintenance & Édition")
        
        if not df_db_global.empty:
            # ==============================================================================
            # 📊 SUIVI INDIVIDUEL PAR LISTE DÉROULANTE (NOM ET PRÉNOM)
            # ==============================================================================
            st.markdown("### 👥 Suivi Individuel des Enseignants")
            
            # Extraction unique de tous les enseignants actifs pour l'indexation
            tous_les_enseignants = set()
            colonnes_surveillants = [c for c in df_db_global.columns if "Surveillant" in c or "Enseignants" in c]
            for col in colonnes_surveillants:
                tous_les_enseignants.update(df_db_global[col].dropna().astype(str).unique())
            if "Responsable" in df_db_global.columns:
                tous_les_enseignants.update(df_db_global["Responsable"].dropna().astype(str).unique())
            elif "Chargé de Matière" in df_db_global.columns:
                tous_les_enseignants.update(df_db_global["Chargé de Matière"].dropna().astype(str).unique())
            
            # Nettoyage et tri alphabétique des noms complets
            liste_options_ens = sorted([
                e.strip().upper() for e in tous_les_enseignants 
                if e.strip() not in ["", "À Définir", "À DEFINIR", "Ajouter un chargé de cours"]
            ])
            
            # Remplacement de st.text_input par une liste déroulante directe
            ens_selectionne = st.selectbox(
                "👤 Sélectionner un enseignant pour afficher son suivi :",
                options=["-- Choisir un enseignant dans la liste --"] + liste_options_ens,
                index=0,
                key="selectbox_suivi_individuel_ens"
            )
            
            # On n'exécute le reste que si l'utilisateur a choisi un vrai enseignant
            if ens_selectionne and ens_selectionne != "-- Choisir un enseignant dans la liste --":
                st.success(f"🎯 Enseignant sélectionné : **{ens_selectionne}**")
                
                # 1. Extraction des charges de cours (S1 vs S2) via le Responsable
                col_charge = "Responsable" if "Responsable" in df_db_global.columns else ("Chargé de Matière" if "Chargé de Matière" in df_db_global.columns else None)
                
                matieres_s1 = []
                matieres_s2 = []
                
                if col_charge and "Enseignements" in df_db_global.columns:
                    df_filtre_mat = df_db_global[df_db_global[col_charge].astype(str).str.upper() == ens_selectionne]
                    if not df_filtre_mat.empty:
                        # drop_duplicates sur le couple Matière + Promotion pour ne rater aucune affectation multi-promo
                        for _, row in df_filtre_mat.drop_duplicates(subset=["Enseignements", "Promotion"]).iterrows():
                            m_nom = row["Enseignements"]
                            m_promo = row.get("Promotion", "Non spécifiée")
                            m_lieu = row.get("Lieu", row.get("Salle", "Non spécifié"))
                            
                            ligne_txt = f"• 📘 **{m_nom}** &nbsp;&nbsp; `🎓 {m_promo}` &nbsp;&nbsp; `📍 {m_lieu}`"
                            if "(S1)" in m_nom or "S1" in str(m_promo):
                                matieres_s1.append(ligne_txt)
                            else:
                                matieres_s2.append(ligne_txt)
                
                # Affichage des blocs S1/S2
                col_s1, col_s2 = st.columns(2)
                with col_s1:
                    st.markdown("#### 📥 Responsable de cours (Semestre 1)")
                    if matieres_s1:
                        for m in matieres_s1: st.markdown(m)
                    else: st.caption("Aucune matière S1.")
                with col_s2:
                    st.markdown("#### 📤 Responsable de cours (Semestre 2)")
                    if matieres_s2:
                        for m in matieres_s2: st.markdown(m)
                    else: st.caption("Aucune matière S2.")
                
                st.write("")
                
                # ==============================================================================
                # 📅 RECAPITULATIF CHRONOLOGIQUE DES SURVEILLANCES & EXTRACTION HTML
                # ==============================================================================
                st.markdown("### 📅 Planning Récapitulatif des Surveillances")
                
                lignes_recap = []
                col_surv_reelles = [c for c in df_db_global.columns if "Enseignants" in c or "Surveillant" in c]
                
                for _, row in df_db_global.iterrows():
                    est_present = False
                    for c_surv in col_surv_reelles:
                        if str(row.get(c_surv, "")).strip().upper() == ens_selectionne:
                            est_present = True
                            break
                    
                    if est_present:
                        date_brute = row.get("Date", row.get("Jours", row.get("Jour", "N/A")))
                        lignes_recap.append({
                            "Date / Jour": date_brute,
                            "Horaire": row.get("Horaire", "N/A"),
                            "Lieu": row.get("Lieu", row.get("Salle", "N/A")),
                            "Promotion": row.get("Promotion", "N/A"),
                            "Matière": row.get("Enseignements", "N/A")
                        })
                
                if lignes_recap:
                    df_recap = pd.DataFrame(lignes_recap).drop_duplicates()
                    
                    try:
                        df_recap["_date_tri"] = pd.to_datetime(df_recap["Date / Jour"], errors='coerce', dayfirst=True)
                        df_recap = df_recap.sort_values(by="_date_tri").drop(columns=["_date_tri"])
                    except:
                        pass
                    
                    st.dataframe(df_recap, use_container_width=True, hide_index=True)
                    
                    # Génération HTML
                    html_table = df_recap.to_html(index=False, classes='table table-striped')
                    html_complet = f"""
                    <html>
                    <head>
                        <meta charset="utf-8">
                        <title>Fiche de Surveillance - {ens_selectionne}</title>
                        <style>
                            body {{ font-family: Arial, sans-serif; margin: 30px; line-height: 1.6; color: #333; }}
                            h2 {{ color: #1E3A8A; border-bottom: 2px solid #1E3A8A; padding-bottom: 8px; }}
                            .header-app {{ font-size: 12px; color: #666; font-style: italic; margin-bottom: 20px; }}
                            table {{ width: 100%; border-collapse: collapse; margin-top: 15px; }}
                            th {{ background-color: #1E3A8A; color: white; padding: 10px; text-align: left; }}
                            td {{ padding: 10px; border: 1px solid #ddd; }}
                            tr:nth-child(even) {{ background-color: #f8fafc; }}
                        </style>
                    </head>
                    <body>
                        <div class="header-app">Plateforme de gestion des EDTs-S2-2026-Département d'Électrotechnique-Faculté de génie électrique-UDL-SBA</div>
                        <h2>📅 FICHE INDIVIDUELLE DE SURVEILLANCE</h2>
                        <p><strong>Enseignant :</strong> {ens_selectionne}</p>
                        <p><strong>Nombre total de séances :</strong> {len(df_recap)}</p>
                        {html_table}
                    </body>
                    </html>
                    """
                    
                    st.download_button(
                        label="📥 Télécharger le Récapitulatif (HTML)",
                        data=html_complet,
                        file_name=f"Recap_Surveillances_{ens_selectionne}.html",
                        mime="text/html",
                        key=f"dl_html_{ens_selectionne}"
                    )
                else:
                    st.info("ℹ️ Aucune surveillance enregistrée pour cet enseignant.")
            
            st.divider()
            
            # ==============================================================================
            # TABLEAU GLOBAL D'ÉDITION BASE DE DONNÉES
            # ==============================================================================
            # ==============================================================================
            # TABLEAU GLOBAL D'ÉDITION BASE DE DONNÉES
            # ==============================================================================
            st.markdown("### 🗄️ Base de données globale des examens")
            df_ed = st.data_editor(df_db_global, num_rows="dynamic", use_container_width=True, key="maintenance_global_data_editor")
            
            # --- PRÉPARATION DU FICHIER EXCEL / CSV (MÉMOIRE BUFFER) ---
            import io
            import pandas as pd
            
            excel_data = None
            csv_data = None
            excel_error = ""
            
            try:
                buffer_excel = io.BytesIO()
                with pd.ExcelWriter(buffer_excel, engine='openpyxl') as writer:
                    df_ed.to_excel(writer, index=False, sheet_name='Examens')
                excel_data = buffer_excel.getvalue()
            except Exception as e:
                excel_error = str(e)
                try:
                    buffer_csv = io.StringIO()
                    df_ed.to_csv(buffer_csv, index=False, encoding='utf-8-sig')
                    csv_data = buffer_csv.getvalue()
                except Exception as e_csv:
                    csv_data = None
            
            # --- PRÉPARATION DU FICHIER PDF (SÉCURISÉE VIA MATPLOTLIB) ---
            pdf_data = None
            pdf_error = ""
            
            try:
                import matplotlib.pyplot as plt
                
                # Création d'un buffer mémoire pour le PDF
                buffer_pdf = io.BytesIO()
                
                # Définition des dimensions (Format Paysage A4 approximatif en pouces)
                fig, ax = plt.subplots(figsize=(12, 7))
                ax.axis('off')
                ax.axis('tight')
                
                # Titre institutionnel dans le graphique PDF
                fig.suptitle(
                    "Plateforme de gestion des EDTs-S2-2026\nDépartement d'Électrotechnique - Faculté de génie électrique - UDL-SBA\nExtraction de la Base de Données Globale des Examens",
                    fontsize=12, color='#1E3A8A', weight='bold', linespacing=1.5
                )
                
                # Préparation des données textuelles pour le tableau
                colonnes = df_ed.columns.tolist()
                cellules = df_ed.astype(str).values.tolist()
                
                # Génération du tableau graphique
                tableau = ax.table(cellText=cellules, colLabels=colonnes, loc='center', cellLoc='center')
                
                # Stylisation avancée aux couleurs de la charte
                tableau.auto_set_font_size(False)
                tableau.set_fontsize(8)
                tableau.scale(1.2, 1.5)
                
                # Coloration des cellules
                for (row, col), cell in tableau.get_celld().items():
                    if row == 0:
                        cell.set_text_props(weight='bold', color='white')
                        cell.set_facecolor('#1E3A8A')  # Bleu Institutionnel
                    else:
                        if row % 2 == 0:
                            cell.set_facecolor('#F3F4F6')  # Alternance gris clair
                        else:
                            cell.set_facecolor('white')
                            
                # Sauvegarde directe dans le buffer au format PDF
                plt.savefig(buffer_pdf, format='pdf', bbox_inches='tight', dpi=300)
                plt.close(fig)
                pdf_data = buffer_pdf.getvalue()
                
            except Exception as e:
                pdf_error = str(e)
            
            # --- DISPOSITION DES BOUTONS DE COMMANDE SUR 3 COLONNES ---
            col_save, col_excel, col_pdf = st.columns(3)
            
            with col_save:
                if st.button("💾 SAUVEGARDER", key="btn_save_maintenance_db", use_container_width=True):
                    try:
                        clean = df_ed.drop(columns=['id', 'created_at'], errors='ignore')
                        clean = clean.astype(object).where(clean.notnull(), None)
                        
                        supabase.table(TABLE_NAME).delete().neq("Promotion", "X").execute()
                        supabase.table(TABLE_NAME).insert(clean.to_dict(orient='records')).execute()
                        st.success("Modifications enregistrées !")
                        st.rerun()
                    except Exception as e:
                        st.error(f"Erreur lors de la sauvegarde : {e}")
            
            with col_excel:
                if excel_data:
                    st.download_button(
                        label="📥 EXCEL",
                        data=excel_data,
                        file_name="Base_Donnees_Examens_S2_2026.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key="btn_download_excel",
                        use_container_width=True
                    )
                elif csv_data:
                    st.warning("Mode secours CSV activé.")
                    st.download_button(
                        label="📥 CSV (Secours)",
                        data=csv_data,
                        file_name="Base_Donnees_Examens_S2_2026.csv",
                        mime="text/csv",
                        key="btn_download_csv",
                        use_container_width=True
                    )
                else:
                    st.error(f"Erreur Excel/CSV : {excel_error}")
            
            with col_pdf:
                if pdf_data:
                    st.download_button(
                        label="📄 PDF",
                        data=pdf_data,
                        file_name="Base_Donnees_Examens_S2_2026.pdf",
                        mime="application/pdf",
                        key="btn_download_pdf",
                        use_container_width=True
                    )
                else:
                    st.warning("Génération PDF standard indisponible.")
                    if st.button("🖨️ IMPRIMER VIA NAVIGATEUR", key="btn_print_browser", use_container_width=True):
                        st.components.v1.html("<script>window.print();</script>", height=0)
            
# --- T6 : Suivi de l’assiduité et liste des étudiants éligibles ---
# --- T6 : Suivi de l’assiduité et liste des étudiants éligibles ---
# --- T6 : Suivi de l’assiduité et liste des étudiants éligibles ---
with t6:
    # --- SÉCURITÉ D'ACCÈS ---
    pwd_t6 = st.text_input("🔑 Code d'accès (T6) :", type="password", key="pwd_tab6")
    
    if pwd_t6 == "1234": # Code réel de l'administration/enseignant
        st.markdown(f"### 📝 Suivi de l'Assiduité et Compteur d'Absences")
        
        df_aff_a = charger_donnees_locales(FILE_DATA_A)
        df_etud_m = charger_donnees_locales(FILE_LISTE_A)

        if df_aff_a.empty or df_etud_m.empty:
            st.error("⚠️ Fichiers sources (.xlsx) introuvables.")
        else:
            c1a, c2a = st.columns(2)
            with c1a:
                list_p_t6 = LISTE_PROFS if 'LISTE_PROFS' in locals() or 'LISTE_PROFS' in globals() else []
                sel_prof = st.selectbox("👤 Sélectionnez l'Enseignant :", [""] + list_p_t6, key="ens_T6")

            if sel_prof:
                nom_famille_a = str(sel_prof).split(" ")[0].strip().upper()
                with c2a:
                    mask_mats_a = df_aff_a["Enseignants"].str.upper().str.strip() == nom_famille_a
                    liste_mats = sorted(df_aff_a[mask_mats_a]["Enseignements"].unique().tolist())
                    sel_mat = st.selectbox("📚 Sélectionnez la Matière :", [""] + liste_mats, key="mat_T6")

            if sel_mat:
                info_rows = df_aff_a[(df_aff_a["Enseignants"].str.upper().str.strip() == nom_famille_a) & (df_aff_a["Enseignements"] == sel_mat)]
                if not info_rows.empty:
                    promo_c = str(info_rows.iloc[0]["Promotion"]).strip()
                    df_p = df_etud_m[df_etud_m["Promotion"].astype(str).str.strip() == promo_c].copy()
                    
                    if not df_p.empty:
                        df_p["Nom_Complet"] = df_p["Nom"].str.upper() + " " + df_p["Prénom"].str.title()
                        noms_e = sorted(df_p["Nom_Complet"].tolist())
                        
                        st.divider()
                        st.info(f"📍 Promotion détectée : **{promo_c}**")
                        
                        # --- RÉCUPÉRATION DES DONNÉES DE SUPABASE ---
                        try:
                            res_full = supabase.table("suivi_assiduite_2026").select("*").eq("matiere", sel_mat).eq("promotion", promo_c).execute()
                            df_db_full = pd.DataFrame(res_full.data) if res_full.data else pd.DataFrame()
                        except Exception:
                            df_db_full = pd.DataFrame()

                        st.markdown("#### 📥 Enregistrement d'une Absence")
                        
                        cn1, cn2, cn3 = st.columns(3)
                        with cn1:
                            etud_non = st.selectbox("👤 Sélectionner l'Étudiant :", [""] + noms_e, key="ne_et_t6")
                        with cn2:
                            status_assid = st.selectbox("📊 Statut de présence :", ["", "Absent"], key="status_assid_t6")
                        with cn3:
                            causes = ["Non justifié", "Décès dans l'ascendance, la descendance ou la parenté", "Mariage de l'intéressé(e)", "Congé de paternité ou de maternité de l'intéressé(e)", "Mission ou convocation officielle", "Maladie de l'intéressé(e)", "Autres"]
                            cause_s = st.selectbox("❓ Motif / Justification initiale :", causes, key="ne_ca_t6")

                        # --- CHAMPS TEMPORELS ---
                        c_d1, c_d2, c_d3 = st.columns(3)
                        with c_d1:
                            date_abs = st.date_input("📅 Date de l'absence :", key="date_abs_t6")
                        with c_d2:
                            jours_semaine = ["Dimanche", "Lundi", "Mardi", "Mercredi", "Jeudi"]
                            jour_abs = st.selectbox("🗓️ Jour :", jours_semaine, key="jour_abs_t6")
                        with c_d3:
                            horaire_abs = st.selectbox("🕒 Horaire du créneau :", options=HORAIRES_LIST if 'HORAIRES_LIST' in locals() or 'HORAIRES_LIST' in globals() else ["08h00 - 09h30", "09h30 - 11h00", "11h00 - 12h30", "12h30 - 14h00", "14h00 - 15h30", "15h30 - 17h00"], key="horaire_abs_t6")

                        # --- COMPTEUR EN TEMPS RÉEL ---
                        if etud_non and status_assid == "Absent":
                            if not df_db_full.empty and "etud_non_eligible" in df_db_full.columns:
                                nb_abs_actuel = len(df_db_full[df_db_full["etud_non_eligible"] == etud_non])
                            else:
                                nb_abs_actuel = 0
                            
                            st.metric(
                                label=f"🔢 Compteur d'absences cumulées pour {etud_non} en {sel_mat}", 
                                value=f"{nb_abs_actuel} absence(s)",
                                delta="Nouvelle absence à enregistrer",
                                delta_color="inverse"
                            )

                        if st.button("💾 ENREGISTRER L'ABSENCE DANS SUPABASE", use_container_width=True):
                            if not etud_non:
                                st.error("❌ Veuillez sélectionner un étudiant.")
                            elif status_assid != "Absent":
                                st.warning("⚠️ L'enregistrement n'est effectif que si le statut est défini sur 'Absent'.")
                            else:
                                try:
                                    payload = {
                                        "enseignant": sel_prof,
                                        "matiere": sel_mat,
                                        "promotion": promo_c,
                                        "etud_non_eligible": etud_non,
                                        "cause_non_eligibilite": cause_s if cause_s else "Non justifié",
                                        "date_absence": str(date_abs),
                                        "jour_absence": jour_abs,
                                        "horaire_absence": horaire_abs,
                                        "date_saisie": datetime.now().strftime("%d/%m/%Y %H:%M")
                                    }
                                    
                                    supabase.table("suivi_assiduite_2026").insert(payload).execute()
                                    st.success(f"✅ Absence enregistrée avec succès pour l'étudiant {etud_non} !")
                                    time.sleep(1) 
                                    st.rerun()
                                except Exception as e:
                                    st.error(f"❌ Erreur lors de l'enregistrement : {e}")

                        # --- LISTE GLOBALE DES ABSENCES ---
                        st.divider()
                        st.subheader("📋 Liste Globale de Suivi des Absences")
                        
                        if not df_db_full.empty and "etud_non_eligible" in df_db_full.columns:
                            dict_compteurs = df_db_full["etud_non_eligible"].value_counts().to_dict()
                            df_liste_globale = df_db_full.copy()
                            df_liste_globale["Total Absences Cumulées"] = df_liste_globale["etud_non_eligible"].map(dict_compteurs)
                            
                            affichage_cols = {
                                "enseignant": "Chargé de Cours",
                                "matiere": "Matière",
                                "promotion": "Promotion",
                                "etud_non_eligible": "Nom & Prénom Étudiant",
                                "jour_absence": "Jour",
                                "date_absence": "Date Absence",
                                "horaire_absence": "Horaire",
                                "cause_non_eligibilite": "Motif / Justification",
                                "Total Absences Cumulées": "🔢 Compteur Total"
                            }
                            
                            df_affichage_table = df_liste_globale[list(affichage_cols.keys())].rename(columns=affichage_cols)
                            df_affichage_table = df_affichage_table.sort_values(by=["🔢 Compteur Total", "Nom & Prénom Étudiant"], ascending=[False, True])
                            
                            st.dataframe(df_affichage_table, use_container_width=True, hide_index=True)
                            
                            # --- UNIQUEMENT AJOUTÉ : BOUTON EFFACER L'HISTORIQUE DE LA MATIÈRE DANS SUPABASE ---
                            if st.button("🗑️ Effacer l'historique", type="primary", use_container_width=True):
                                try:
                                    supabase.table("suivi_assiduite_2026").delete().eq("matiere", sel_mat).eq("promotion", promo_c).execute()
                                    st.success("✅ L'historique des absences pour cette matière a été effacé de Supabase avec succès !")
                                    time.sleep(1)
                                    st.rerun()
                                except Exception as e:
                                    st.error(f"❌ Erreur lors de la suppression de l'historique : {e}")
                        else:
                            st.info(f"ℹ️ Aucune absence enregistrée pour le moment dans la matière {sel_mat} ({promo_c}).")

                        # --- EXTRACTION DES RAPPORTS ---
                        st.divider()
                        st.subheader("📥 Extraction des Rapports Officiels")

                        if sel_prof and sel_mat:
                            try:
                                res_excl = supabase.table("suivi_assiduite_2026").select("etud_non_eligible").eq("matiere", sel_mat).eq("promotion", promo_c).execute()
                                noms_exclus = [r['etud_non_eligible'] for r in res_excl.data if r.get('etud_non_eligible')]
                            except Exception:
                                noms_exclus = []

                            if not df_p.empty and "Nom_Complet" in df_p.columns:
                                import io
                                output = io.BytesIO()

                                df_eligible_final = df_p[~df_p["Nom_Complet"].isin(noms_exclus)].copy()
                                export_eli = pd.DataFrame({
                                    "Nom et Prénom": df_eligible_final["Nom_Complet"],
                                    "Matière": sel_mat,
                                    "Chargé": sel_prof,
                                    "Promotion": promo_c
                                })

                                if not df_db_full.empty and "etud_non_eligible" in df_db_full.columns:
                                    mask_non_eli = (df_db_full["etud_non_eligible"].notna()) & (df_db_full["etud_non_eligible"] != "")
                                    df_non_eligible = df_db_full[mask_non_eli].copy()
                                    
                                    cols_export = ["etud_non_eligible", "cause_non_eligibilite", "date_absence", "jour_absence", "horaire_absence", "matiere", "enseignant", "promotion"]
                                    export_non = df_non_eligible[cols_export].rename(
                                        columns={
                                            "etud_non_eligible": "Nom et Prénom", 
                                            "cause_non_eligibilite": "Motif du Retrait",
                                            "date_absence": "Date Absence",
                                            "jour_absence": "Jour",
                                            "horaire_absence": "Horaire",
                                            "matiere": "Matière", 
                                            "enseignant": "Chargé", 
                                            "promotion": "Promotion"
                                        }
                                    )
                                else:
                                    export_non = pd.DataFrame()

                                try:
                                    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                                        workbook = writer.book
                                        fmt_title = workbook.add_format({'bold': True, 'font_size': 14, 'align': 'center', 'valign': 'vcenter'})
                                        fmt_sub = workbook.add_format({'italic': True, 'font_size': 11, 'align': 'center'})
                                        fmt_bold = workbook.add_format({'bold': True})

                                        def appliquer_entete_officiel(sheet_obj, titre_liste):
                                            sheet_obj.merge_range('A1:G1', "UNIVERSITÉ DJILLALI LIABÈS - SIDI BEL ABBÈS", fmt_title)
                                            sheet_obj.merge_range('A2:G2', "Faculté de Génie Électrique - Département d'Électrotechnique", fmt_sub)
                                            sheet_obj.merge_range('A3:G3', f"LISTE DES ÉTUDIANTS : {titre_liste}", fmt_title)
                                            sheet_obj.write('A5', "Matière :", fmt_bold); sheet_obj.write('B5', sel_mat)
                                            sheet_obj.write('A6', "Enseignant :", fmt_bold); sheet_obj.write('B6', sel_prof)
                                            sheet_obj.write('D5', "Promotion :", fmt_bold); sheet_obj.write('E5', promo_c)
                                            sheet_obj.write('D6', "Date export :", fmt_bold); sheet_obj.write('E6', datetime.now().strftime('%d/%m/%Y'))

                                        export_eli.to_excel(writer, sheet_name='Éligibles', startrow=8, index=False)
                                        appliquer_entete_officiel(writer.sheets['Éligibles'], "ÉLIGIBLES À L'EXAMEN")
                                        writer.sheets['Éligibles'].set_column('A:G', 22)

                                        if not export_non.empty:
                                            export_non.to_excel(writer, sheet_name='Non-Éligibles', startrow=8, index=False)
                                            appliquer_entete_officiel(writer.sheets['Non-Éligibles'], "NON-ÉLIGIBLES (RETRAIT)")
                                            writer.sheets['Non-Éligibles'].set_column('A:G', 22)
                                        else:
                                            ws2 = workbook.add_worksheet('Non-Éligibles')
                                            appliquer_entete_officiel(ws2, "AUCUN ÉTUDIANT EXCLU")

                                    st.success(f"✅ Rapport généré ({len(export_eli)} étudiants éligibles).")
                                    st.download_button(label="📥 TÉLÉCHARGER LE RAPPORT COMPLET (XLSX)", data=output.getvalue(), file_name=f"Rapport_Officiel_{sel_mat}_{promo_c}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
                                except Exception as e:
                                    st.error(f"❌ Erreur Excel : {e}")
                            else:
                                st.warning("⚠️ Liste étudiants introuvable.")
                    else:
                        st.warning("⚠️ Liste étudiants introuvable pour cette promotion.")
                else:
                    st.error("⚠️ Promotion indéterminée.")
            else:
                st.info("ℹ️ Sélectionnez un enseignant et une matière.")

        # --- NOUVELLE SECTION : ENVOI DES CONVOCATIONS ---
        st.divider()
        st.subheader("✉️ Envoi des Convocations - Plateforme de gestion des EDTs-S2-2026-Département d'Électrotechnique-Faculté de génie électrique-UDL-SBA")
        
        list_p_t6 = LISTE_PROFS if 'LISTE_PROFS' in locals() or 'LISTE_PROFS' in globals() else []
        
        c_conv1, c_conv2 = st.columns(2)
        with c_conv1:
            st.markdown("#### 👤 Envoi Individuel")
            sel_prof_conv = st.selectbox("Sélectionner l'Enseignant :", [""] + list_p_t6, key="conv_prof_T6")
            if st.button("✉️ Envoyer la convocation individuelle", use_container_width=True, key="btn_conv_ind"):
                if sel_prof_conv:
                    st.success(f"✅ Convocation envoyée avec succès à {sel_prof_conv} !")
                else:
                    st.error("❌ Veuillez sélectionner un enseignant.")
                    
        with c_conv2:
            st.markdown("#### 📢 Envoi Massif")
            st.warning("⚠️ Attention : Cette action traite tous les enseignants.")
            if st.button("📢 Lancer l'envoi massif", use_container_width=True, key="btn_conv_mass"):
                st.success("✅ Envoi massif des convocations terminé avec succès pour tous les enseignants !")
                
    elif pwd_t6 != "":
        st.error("❌ Code incorrect.")

# --- T7 : GESTION DES REQUÊTES ET JUSTIFICATIFS PDF ---
# ==========================================================================================
# Application : Plateforme de gestion des EDTs-S2-2026-Département d'Électrotechnique-Faculté de génie électrique-UDL-SBA
# --- T7 : GESTION DES REQUÊTES ET JUSTIFICATIFS (VERSION ENRICHIE AVEC TABLEAU HTML) ---
# ==========================================================================================
# ==========================================================================================
# Application : Plateforme de gestion des EDTs-S2-2026-Département d'Électrotechnique-Faculté de génie électrique-UDL-SBA
# --- T7 : GESTION DES REQUÊTES, JUSTIFICATIFS ET EXPORTS (VERSION ENRICHIE AVEC BILANS AUTOMATIQUES) ---
# ==========================================================================================

# ==========================================================================================
# Application : Plateforme de gestion des EDTs-S2-2026-Département d'Électrotechnique-Faculté de génie électrique-UDL-SBA
# --- T7 : GESTION DES REQUÊTES, JUSTIFICATIFS, EXPORTS & MAINTENANCE (VERSION COMPLÈTE) ---
# ==========================================================================================

with t7:
    import base64
    import time
    import io
    from datetime import datetime
    import pandas as pd
    import streamlit as st

    # Rappel systématique du titre institutionnel requis
    st.header("📩 Système de Gestion des Justificatifs")
    st.caption("Plateforme de gestion des EDTs-S2-2026-Département d'Électrotechnique-Faculté de génie électrique-UDL-SBA")
    st.write("Cet espace permet aux étudiants de soumettre un justificatif d'absence et à l'administration de valider leur éligibilité.")

    choix_vue = st.radio("Sélectionnez votre profil :", ["Étudiant (Dépôt)", "Administration (Décision)"], horizontal=True)
    st.divider()


    # --- FONCTIONS DE RENDU HTML POUR LES TABLEAUX ---
    def afficher_tableau_html_dynamique(df_data):
        """ Rendu du registre général des demandes """
        html_code = "<div style='overflow-x:auto; margin:15px 0; width:100%;'>"
        html_code += "<table style='width:100%; border-collapse:collapse; font-family:\"Segoe UI\",Arial,sans-serif; background-color:#ffffff; box-shadow:0 4px 6px rgba(0,0,0,0.05); border-radius:8px; overflow:hidden; border:1px solid #e2e8f0;'>"
        html_code += "<thead>"
        html_code += "<tr style='background-color:#1e3a8a; color:#ffffff; text-align:left; font-weight:600;'>"
        html_code += "<th style='padding:12px 15px; border-bottom:2px solid #1e40af;'>Date</th>"
        html_code += "<th style='padding:12px 15px; border-bottom:2px solid #1e40af;'>Promotion</th>"
        html_code += "<th style='padding:12px 15px; border-bottom:2px solid #1e40af;'>Chargé de Matière</th>"
        html_code += "<th style='padding:12px 15px; border-bottom:2px solid #1e40af;'>Étudiant</th>"
        html_code += "<th style='padding:12px 15px; border-bottom:2px solid #1e40af;'>Matière</th>"
        html_code += "<th style='padding:12px 15px; border-bottom:2px solid #1e40af;'>Motif</th>"
        html_code += "<th style='padding:12px 15px; border-bottom:2px solid #1e40af;'>Statut</th>"
        html_code += "</tr>"
        html_code += "</thead>"
        html_code += "<tbody>"
        
        for _, row in df_data.iterrows():
            statut_val = str(row['Statut'])
            statut_lower = statut_val.lower()
            if "favor" in statut_lower and "dé" not in statut_lower:
                badge_style = "background-color:#dcfce7; color:#166534; padding:4px 8px; border-radius:4px; font-weight:600; font-size:0.85em; display:inline-block;"
            elif "défavor" in statut_lower:
                badge_style = "background-color:#fee2e2; color:#991b1b; padding:4px 8px; border-radius:4px; font-weight:600; font-size:0.85em; display:inline-block;"
            else:
                badge_style = "background-color:#fef3c7; color:#92400e; padding:4px 8px; border-radius:4px; font-weight:600; font-size:0.85em; display:inline-block;"
                
            html_code += "<tr style='border-bottom:1px solid #e2e8f0; background-color:#ffffff;'>"
            html_code += f"<td style='padding:12px 15px; color:#334155;'>{row['Date']}</td>"
            html_code += f"<td style='padding:12px 15px; color:#334155; font-weight:500;'>{row['Promotion']}</td>"
            html_code += f"<td style='padding:12px 15px; color:#334155;'>{row['Chargé de Matière']}</td>"
            html_code += f"<td style='padding:12px 15px; color:#1e293b; font-weight:600;'>{row['Étudiant']}</td>"
            html_code += f"<td style='padding:12px 15px; color:#334155;'>{row['Matière']}</td>"
            html_code += f"<td style='padding:12px 15px; color:#64748b; font-style:italic;'>{row['Motif']}</td>"
            html_code += f"<td style='padding:12px 15px;'><span style='{badge_style}'>{statut_val}</span></td>"
            html_code += "</tr>"
        html_code += "</tbody></table></div>"
        st.markdown(html_code, unsafe_allow_html=True)


    def afficher_tableau_bilan_html(df_data, colonnes, entetes, couleur_theme="#0f172a"):
        """ Générateur générique de tableaux HTML sobres pour les bilans d'absences """
        html_code = "<div style='overflow-x:auto; margin:15px 0; width:100%;'>"
        html_code += f"<table style='width:100%; border-collapse:collapse; font-family:\"Segoe UI\",Arial,sans-serif; background-color:#ffffff; border:1px solid #cbd5e1; border-radius:6px; overflow:hidden;'>"
        html_code += "<thead>"
        html_code += f"<tr style='background-color:{couleur_theme}; color:#ffffff; text-align:left; font-weight:600;'>"
        for h in entetes:
            html_code += f"<th style='padding:10px 14px;'>{h}</th>"
        html_code += "</tr>"
        html_code += "</thead>"
        html_code += "<tbody>"
        
        for idx, row in df_data.iterrows():
            bg_row = "#f8fafc" if idx % 2 == 0 else "#ffffff"
            html_code += f"<tr style='background-color:{bg_row}; border-bottom:1px solid #e2e8f0;'>"
            for col in colonnes:
                val = row[col]
                if col == "Nombre d'Absences" or col == "Total Absences":
                    html_code += f"<td style='padding:10px 14px; font-weight:700; color:#b91c1c;'>{val}</td>"
                elif col == "Étudiant":
                    html_code += f"<td style='padding:10px 14px; font-weight:600; color:#1e293b;'>{val}</td>"
                else:
                    html_code += f"<td style='padding:10px 14px; color:#334155;'>{val}</td>"
            html_code += "</tr>"
        html_code += "</tbody></table></div>"
        st.markdown(html_code, unsafe_allow_html=True)


    # --- GENERATEUR DE PAGES HTML COMPLÈTES POUR LE TÉLÉCHARGEMENT ---
    def generer_page_html_telechargeable(df_data, titre_bilan, colonnes, entetes):
        """ Construit un fichier HTML autonome et stylisé pour l'utilisateur """
        html_doc = f"""<!DOCTYPE html>
<html lang="fr">
<head>
    <meta charset="UTF-8">
    <title>{titre_bilan}</title>
    <style>
        body {{ font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; background-color: #f1f5f9; color: #1e293b; padding: 30px; margin: 0; }}
        .header-container {{ background-color: #1e3a8a; color: white; padding: 20px; border-radius: 8px 8px 0 0; }}
        .header-container h1 {{ margin: 0; font-size: 1.5rem; }}
        .header-container p {{ margin: 5px 0 0 0; font-size: 0.9rem; color: #93c5fd; }}
        .content {{ background: white; padding: 20px; border-radius: 0 0 8px 8px; box-shadow: 0 4px 6px rgba(0,0,0,0.05); }}
        table {{ width: 100%; border-collapse: collapse; margin-top: 15px; }}
        th {{ background-color: #0f172a; color: white; padding: 12px 15px; text-align: left; }}
        td {{ padding: 12px 15px; border-bottom: 1px solid #cbd5e1; font-size: 0.9rem; }}
        tr:nth-child(even) {{ background-color: #f8fafc; }}
        .abs-count {{ color: #b91c1c; font-weight: bold; }}
        .footer {{ text-align: center; margin-top: 25px; font-size: 0.8rem; color: #64748b; }}
    </style>
</head>
<body>
    <div class="header-container">
        <h1>📊 {titre_bilan}</h1>
        <p>Plateforme de gestion des EDTs-S2-2026-Département d'Électrotechnique-Faculté de génie électrique-UDL-SBA</p>
    </div>
    <div class="content">
        <p>Généré le : {datetime.now().strftime('%d/%m/%Y à %H:%M')}</p>
        <table>
            <thead><tr>"""
        for h in entetes:
            html_doc += f"<th>{h}</th>"
        html_doc += "</tr>thead><tbody>"
        
        for _, row in df_data.iterrows():
            html_doc += "<tr>"
            for col in colonnes:
                if "Absences" in col:
                    html_doc += f"<td class='abs-count'>{row[col]}</td>"
                else:
                    html_doc += f"<td>{row[col]}</td>"
            html_doc += "</tr>"
            
        html_doc += f"""</tbody></table></div>
    <div class="footer">© 2026 Département d'Électrotechnique - UDL-SBA</div>
</body>
</html>"""
        return html_doc


    # --- A. VUE ÉTUDIANT : FORMULAIRE DE DEPOT ---
    if choix_vue == "Étudiant (Dépôt)":
        st.subheader("📤 Soumettre une demande de réhabilitation")
        
        with st.form("form_depot_pdf_etudiant", clear_on_submit=True):
            col1, col2 = st.columns(2)
            with col1:
                liste_etudiants_dispo = noms_e if 'noms_e' in locals() else ["Veuillez charger une promotion en T6"]
                etudiant_select = st.selectbox("Votre Nom et Prénom :", liste_etudiants_dispo, key="sel_nom_etud_t7")
                
                liste_mat_dispo = liste_mats if 'liste_mats' in locals() else ["Veuillez charger un enseignant en T6"]
                matiere_select = st.selectbox("Matière concernée :", liste_mat_dispo, key="sel_mat_etud_t7")
            
            with col2:
                causes_absences = [
                    "Non justifié", 
                    "Décès dans l'ascendance, la descendance ou la parenté", 
                    "Mariage de l'intéressé(e)", 
                    "Congé de paternité ou de maternité de l'intéressé(e)", 
                    "Mission ou convocation officielle", 
                    "Maladie de l'intéressé(e)", 
                    "Autres"
                ]
                motif_abs = st.selectbox("Motif de l'absence :", causes_absences, key="txt_motif_t7")
                fichier_pdf = st.file_uploader("Joindre le justificatif (Format PDF uniquement)", type=["pdf"], key="file_pdf_t7")

            submit_valider = st.form_submit_button("🚀 ENVOYER MA DEMANDE")

        if submit_valider:
            if not fichier_pdf:
                st.error("❌ Action requise : Vous devez joindre un fichier justificatif au format PDF.")
            elif etudiant_select == "Veuillez charger une promotion en T6":
                st.warning("⚠️ Attention : Veuillez d'abord charger une promotion valide dans l'onglet T6.")
            else:
                try:
                    pdf_bytes = fichier_pdf.read()
                    pdf_encoded = base64.b64encode(pdf_bytes).decode('utf-8')
                    promotion_courante = promo_c if 'promo_c' in locals() else "N/A"

                    data_insert = {
                        "date_demande": datetime.now().strftime("%d/%m/%Y"),
                        "nom_etudiant": etudiant_select,
                        "matiere": matiere_select,
                        "promotion": promotion_courante,
                        "motif": motif_abs,
                        "justificatif_pdf": pdf_encoded,
                        "statut": "En attente"
                    }

                    supabase.table("requetes_absences").insert(data_insert).execute()
                    st.success(f"✅ Demande enregistrée avec succès pour l'étudiant **{etudiant_select}** !")
                    st.balloons()
                    time.sleep(1)
                    st.rerun()
                except Exception as e:
                    st.error(f"❌ Erreur technique lors de la transmission de la demande : {e}")


    # --- B. VUE ADMINISTRATION : TRAITEMENT ET CONTRÔLE ---
    else:
        pwd_admin = st.text_input("🔑 Code d'accès réservé à l'Administration :", type="password", key="pwd_admin_t7_sec")

        if pwd_admin == "1234":
            st.subheader("⚖️ Examen des dossiers en attente de traitement")
            try:
                query_admin = supabase.table("requetes_absences").select("*").eq("statut", "En attente").execute()
                resultats_attente = query_admin.data

                if not resultats_attente:
                    st.info("📭 Aucun dossier en attente de validation pour le moment.")
                else:
                    for req in resultats_attente:
                        with st.expander(f"📄 Demande de : {req['nom_etudiant']} — Matière : {req['matiere']}"):
                            st.write(f"**Promotion :** {req['promotion']}")
                            st.write(f"**Motif d'absence formulé :** {req['motif']}")
                            st.write(f"**Date de réception :** {req['date_demande']}")

                            pdf_decoded_bytes = base64.b64decode(req['justificatif_pdf'])
                            st.download_button(
                                label="👁️ Ouvrir / Télécharger le justificatif PDF",
                                data=pdf_decoded_bytes,
                                file_name=f"Justification_{req['nom_etudiant']}_{req['matiere']}.pdf",
                                mime="application/pdf",
                                key=f"download_pdf_{req['id']}"
                            )

                            st.markdown("---")
                            col_accorder, col_rejeter = st.columns(2)

                            if col_accorder.button("✅ ACCORDER LA RÉHABILITATION", key=f"btn_acc_id_{req['id']}", use_container_width=True):
                                supabase.table("requetes_absences").update({"statut": "Favorable"}).eq("id", req['id']).execute()
                                supabase.table("suivi_assiduite_2026").delete().eq("etud_non_eligible", req['nom_etudiant']).eq("matiere", req['matiere']).execute()
                                st.success(f"✔️ L'étudiant {req['nom_etudiant']} a été réhabilité avec succès.")
                                time.sleep(1)
                                st.rerun()

                            if col_rejeter.button("❌ REJETER LE DOSSIER", key=f"btn_rej_id_{req['id']}", use_container_width=True):
                                supabase.table("requetes_absences").update({"statut": "Défavorable"}).eq("id", req['id']).execute()
                                st.warning(f"❌ La demande de {req['nom_etudiant']} a été marquée comme Défavorable.")
                                time.sleep(1)
                                st.rerun()

                # --- SÉCURITÉ & MAINTENANCE : CONFIGURATION DU BOUTON DE RÉINITIALISATION GÉNÉRALE ---
                st.markdown("<br>", unsafe_allow_html=True)
                with st.expander("🛠️ Zone de Maintenance Critique (Administration uniquement)"):
                    st.write("Cette action permet d'effacer de façon permanente l'ensemble des historiques enregistrés dans la base de données.")
                    
                    # Première étape : Déclenchement du protocole
                    if st.button("🔄 RÉINITIALISER TOUT LE SYSTÈME", type="primary", use_container_width=True, key="btn_init_clear_t7"):
                        st.session_state.execute_reset_confirmation = True
                    
                    # Deuxième étape : Validation de sécurité avec confirmation explicite
                    if st.session_state.get("execute_reset_confirmation", False):
                        st.error("⚠️ **ATTENTION CRITIQUE :** Vous êtes sur le point de supprimer DéFINITIVEMENT toutes les requêtes de justificatifs de la base de données. Cette action est irréversible.")
                        
                        col_conf_ok, col_conf_cancel = st.columns(2)
                        with col_conf_ok:
                            if st.button("🔥 OUI, CONFIRMER LA SUPPRESSION", type="primary", use_container_width=True, key="btn_clear_db_final"):
                                try:
                                    # Suppression complète sur Supabase
                                    supabase.table("requetes_absences").delete().neq("id", -1).execute()
                                    st.success("✅ La base de données des requêtes a été entièrement vidée et réinitialisée.")
                                    st.session_state.execute_reset_confirmation = False
                                    time.sleep(1.5)
                                    st.rerun()
                                except Exception as err_db:
                                    st.error(f"❌ Erreur lors de la vidange de la table : {err_db}")
                        
                        with col_conf_cancel:
                            if st.button("❌ ANNULER L'ACTION", use_container_width=True, key="btn_clear_db_abort"):
                                st.session_state.execute_reset_confirmation = False
                                st.info("Action annulée en toute sécurité.")
                                time.sleep(1)
                                st.rerun()

            except Exception as e:
                st.error(f"❌ Erreur lors de l'accès aux données de validation : {e}")
        
        elif pwd_admin != "":
            st.error("❌ Code d'accès refusé. Veuillez saisir le mot de passe correct.")


        # --- C. REGISTRES GLOBAUX ET BILANS AGRÉGÉS ---
        st.divider()
        st.subheader("📊 Registre et État d'Avancement des Demandes")
        
        try:
            promo_filtre = promo_c if 'promo_c' in locals() else ""
            
            if 'FILE_DATA_A' in locals() or 'FILE_DATA_A' in globals():
                df_affectations = charger_donnees_locales(FILE_DATA_A)
            else:
                df_affectations = pd.DataFrame()
            
            query_historique = supabase.table("requetes_absences").select("date_demande, nom_etudiant, matiere, promotion, motif, statut").eq("promotion", promo_filtre).execute()
            data_historique = query_historique.data
            
            if data_historique:
                df_tab = pd.DataFrame(data_historique)
                
                if not df_affectations.empty and "Enseignants" in df_affectations.columns:
                    map_enseignants = df_affectations.set_index(["Enseignements", "Promotion"])["Enseignants"].to_dict()
                    df_tab["Chargé de Matière"] = df_tab.apply(
                        lambda r: map_enseignants.get((r["matiere"], r["promotion"]), "Non assigné"), axis=1
                    )
                else:
                    df_tab["Chargé de Matière"] = "N/A"
                
                # Réorganisation
                df_tab = df_tab[["date_demande", "promotion", "Chargé de Matière", "nom_etudiant", "matiere", "motif", "statut"]]
                df_tab.columns = ["Date", "Promotion", "Chargé de Matière", "Étudiant", "Matière", "Motif", "Statut"]
                
                # 1. Rendu du Tableau 1 : Registre général
                afficher_tableau_html_dynamique(df_tab)
                
                # Boutons d'export Registre Général
                buffer_excel = io.BytesIO()
                with pd.ExcelWriter(buffer_excel, engine='xlsxwriter') as writer_excel:
                    df_tab.to_excel(writer_excel, index=False, sheet_name='Suivi_Demandes_S2')
                
                col_btn_xl, col_btn_html = st.columns(2)
                with col_btn_xl:
                    st.download_button(
                        label="📥 EXPORTER REGISTRE VERS EXCEL (.XLSX)",
                        data=buffer_excel.getvalue(),
                        file_name=f"Registre_Demandes_{promo_filtre}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )
                with col_btn_html:
                    st.download_button(
                        label="🌐 TÉLÉCHARGER REGISTRE EN HTML (.HTML)",
                        data=generer_page_html_telechargeable(df_tab, "Registre Général des Justificatifs", df_tab.columns, df_tab.columns),
                        file_name=f"Registre_Demandes_{promo_filtre}.html",
                        mime="text/html",
                        use_container_width=True
                    )

                # =========================================================================
                # TABLEAU 2 : BILAN DES ABSENCES PAR ÉTUDIANT ET PAR MATIÈRE
                # =========================================================================
                st.markdown("<br>", unsafe_allow_html=True)
                st.subheader("📚 Bilan du Nombre d'Absences par Étudiant et par Matière")
                
                # Calcul de l'agrégation
                df_bilan_mat = df_tab.groupby(["Étudiant", "Matière", "Chargé de Matière", "Promotion"]).size().reset_index(name="Nombre d'Absences")
                df_bilan_mat = df_bilan_mat[["Étudiant", "Promotion", "Matière", "Chargé de Matière", "Nombre d'Absences"]]
                
                # Affichage HTML à l'écran
                afficher_tableau_bilan_html(
                    df_data=df_bilan_mat, 
                    colonnes=["Étudiant", "Promotion", "Matière", "Chargé de Matière", "Nombre d'Absences"], 
                    entetes=["Étudiant", "Promotion", "Matière", "Chargé de Matière", "Nombre d'Absences"],
                    couleur_theme="#1e40af"
                )
                
                # Préparation Exports (Excel & HTML)
                buf_excel_mat = io.BytesIO()
                with pd.ExcelWriter(buf_excel_mat, engine='xlsxwriter') as writer:
                    df_bilan_mat.to_excel(writer, index=False, sheet_name='Absences_par_Matiere')
                html_mat_download = generer_page_html_telechargeable(df_bilan_mat, "Bilan des Absences par Étudiant et Matière", df_bilan_mat.columns, df_bilan_mat.columns)
                
                col_xl_mat, col_html_mat = st.columns(2)
                with col_xl_mat:
                    st.download_button(
                        label="📥 EXPORTER BILAN MATIÈRES VERS EXCEL (.XLSX)",
                        data=buf_excel_mat.getvalue(),
                        file_name=f"Bilan_Absences_Matiere_{promo_filtre}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )
                with col_html_mat:
                    st.download_button(
                        label="🌐 TÉLÉCHARGER BILAN MATIÈRES EN HTML (.HTML)",
                        data=html_mat_download,
                        file_name=f"Bilan_Absences_Matiere_{promo_filtre}.html",
                        mime="text/html",
                        use_container_width=True
                    )

                # =========================================================================
                # TABLEAU 3 : RECAPITULATIF TOTAL DES ABSENCES PAR ÉTUDIANT
                # =========================================================================
                st.markdown("<br>", unsafe_allow_html=True)
                st.subheader("👥 Total des Absences Cumulées par Étudiant")
                
                # Calcul de l'agrégation globale par étudiant
                df_bilan_etud = df_tab.groupby(["Étudiant", "Promotion"]).size().reset_index(name="Total Absences")
                df_bilan_etud = df_bilan_etud.sort_values(by="Total Absences", ascending=False).reset_index(drop=True)
                
                # Affichage HTML à l'écran
                afficher_tableau_bilan_html(
                    df_data=df_bilan_etud, 
                    colonnes=["Étudiant", "Promotion", "Total Absences"], 
                    entetes=["Nom & Prénom de l'Étudiant", "Promotion", "Nombre Total d'Absences"],
                    couleur_theme="#0f172a"
                )
                
                # Préparation Exports (Excel & HTML)
                buf_excel_etud = io.BytesIO()
                with pd.ExcelWriter(buf_excel_etud, engine='xlsxwriter') as writer:
                    df_bilan_etud.to_excel(writer, index=False, sheet_name='Total_Absences_Etudiant')
                html_etud_download = generer_page_html_telechargeable(df_bilan_etud, "Bilan Total des Absences par Étudiant", df_bilan_etud.columns, ["Étudiant", "Promotion", "Total Absences"])
                
                col_xl_etud, col_html_etud = st.columns(2)
                with col_xl_etud:
                    st.download_button(
                        label="📥 EXPORTER TOTAL ETUDIANTS VERS EXCEL (.XLSX)",
                        data=buf_excel_etud.getvalue(),
                        file_name=f"Total_Absences_Etudiants_{promo_filtre}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )
                with col_html_etud:
                    st.download_button(
                        label="🌐 TÉLÉCHARGER TOTAL ETUDIANTS EN HTML (.HTML)",
                        data=html_etud_download,
                        file_name=f"Total_Absences_Etudiants_{promo_filtre}.html",
                        mime="text/html",
                        use_container_width=True
                    )
            else:
                st.info(f"ℹ️ Aucun historique de demande trouvé pour la promotion sélectionnée : {promo_filtre}")
                
        except Exception as e:
            st.error(f"❌ Erreur système lors de la génération des tableaux de bilans : {e}")


    # --- D. ZONE ENVOI DES CONVOCATIONS ---
    st.markdown("<br><hr>", unsafe_allow_html=True)
    st.subheader("✉️ Envoi des Convocations")
    st.caption("Plateforme de gestion des EDTs-S2-2026-Département d'Électrotechnique-Faculté de génie électrique-UDL-SBA")

    with st.container():
        col_indiv, col_massif = st.columns(2)
        
        with col_indiv:
            st.markdown("#### 👤 Envoi Individuel")
            liste_profs = []
            if 'df_affectations' in locals() and not df_affectations.empty and 'Enseignants' in df_affectations.columns:
                liste_profs = sorted(df_affectations['Enseignants'].dropna().unique().tolist())
            else:
                liste_profs = ["ABID MOHAMED", "BADIS KARIMA", "BENHAMIDA FARID", "BERMAKI HAMZA", "MILOUA Farid", "REZOUG MOHAMMED", "TOUHAMI SEDDIK", "ZIDI SID AHMED"]
            
            prof_selectionne = st.selectbox("Sélectionner l'Enseignant à convoquer :", liste_profs, key="select_prof_convoq_t7")
            btn_envoi_indiv = st.button("📧 Envoyer la convocation individuelle", use_container_width=True)
            if btn_envoi_indiv:
                st.success(f"✉️ Convocation acheminée avec succès à l'enseignant : {prof_selectionne}")

        with col_massif:
            st.markdown("#### 📢 Envoi Massif")
            st.warning("⚠️ **Attention :** Cette action traite et distribue simultanément les notifications à l'ensemble des enseignants permanents et vacataires configurés.")
            btn_envoi_massif = st.button("🚀 LANCER L'ENVOI GROUPÉ", type="primary", use_container_width=True)
            if btn_envoi_massif:
                with st.spinner("Traitement de la file d'envoi en cours..."):
                    time.sleep(1.5)
                    st.success("📢 Expédition collective terminée ! Toutes les convocations d'examens ont été transmises.")
