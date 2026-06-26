import streamlit as st
import pandas as pd
import os, random, io, smtplib
import base64
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from datetime import datetime
import time
from supabase import create_client

# --- CONFIGURATION INTERFACE & CHARTE GRAPHIQUE ---
st.set_page_config(page_title="Gestion EDT & Assiduité ELT 2026", layout="wide")

# Masquer les éléments natifs du menu supérieur Streamlit pour un rendu professionnel
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

# --- TITRE INSTITUTIONNEL REQUIS ---
TITLE_PLATFORM = "Plateforme de gestion des EDTs-S2-2026-Département d'Électrotechnique-Faculté de génie électrique-UDL-SBA"
st.title(TITLE_PLATFORM)

# ======================================================================================
# 1. PARAMÈTRES ET CONFIGURATIONS CRITIQUES
# ======================================================================================
NOM_SOURCE = "dataEDT-ELT-S2-2026.xlsx"
FILE_EMAILS = "Permanents-Vacataires-ELT2-2025-2026.xlsx"
FILE_DATA_A = "DATA-ASSUIDUITE-2026.xlsx"
FILE_LISTE_A = "Liste des étudiants-2025-2026.xlsx"
FILE_MATRICULES = "Mtricules-ELT-2026-2027.xlsx" # Fichier d'authentification à 4 colonnes

COLS_ORDRE = ['Enseignements', 'Code', 'Enseignants', 'Horaire', 'Jours', 'Lieu', 'Promotion']

# Connexion Base de Données Supabase
S_URL = "https://ajcbkidmcjtyomknijwa.supabase.co"
S_KEY = "sb_publishable_otn3XM8LPLV0OGw74LRhDw_F446jkpw"
supabase = create_client(S_URL, S_KEY)

HORAIRES_LIST = [
    "8h - 9h", "8h - 9h30", "8h - 10h", "9h - 10h", "9h30 - 11h", 
    "10h - 11h", "11h - 12h", "11h - 12h30", 
    "12h - 13h", "12h30 - 14h", "13h - 14h", "14h - 15h30", "14h - 16h", "15h30 - 17h"
]

# ======================================================================================
# 2. MOTEUR DE CHARGEMENT ET PARSING DES FICHIERS
# ======================================================================================
def charger_donnees_locales(path):
    if os.path.exists(path):
        try:
            df = pd.read_excel(path)
            df.columns = [str(c).strip() for c in df.columns]
            return df
        except Exception as e:
            st.error(f"Erreur de lecture du fichier {path} : {e}")
            return pd.DataFrame()
    return pd.DataFrame()

@st.cache_data
def charger_fichiers_enseignants():
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
        except:
            pass
    if os.path.exists(NOM_SOURCE):
        try:
            df_f = pd.read_excel(NOM_SOURCE)
            df_f.columns = [str(c).strip() for c in df_f.columns]
            mask = df_f["Enseignements"].str.contains("Cours", case=False, na=False)
            df_s = df_f[mask].copy()
            for c in COLS_ORDRE:
                if c not in df_s.columns: df_s[c] = ""
            df_s = df_s[COLS_ORDRE]
        except:
            pass
    return df_s, map_nom_complet, d_em

df_src, map_noms, dict_emails = charger_fichiers_enseignants()
if not df_src.empty:
    noms_famille = df_src["Enseignants"].unique()
    LISTE_PROFS = sorted([map_noms.get(str(n).strip().upper(), str(n).strip().upper()) for n in noms_famille])
else:
    LISTE_PROFS = []

# --- ENVOI DE L'EMAIL DE SÉCURITÉ (OTP) ---
def envoyer_email_otp(email_dest, nom_etudiant, code_otp):
    # Configuration des variables du serveur SMTP
    SMTP_SERVER = "smtp.gmail.com"
    SMTP_PORT = 587
    SMTP_USER = "votre.email.institutionnel@gmail.com"  # À remplacer par vos accès de messagerie
    SMTP_PASSWORD = "votre_mot_de_passe_d_application"  # À remplacer par votre clé secrète d'application
    
    msg = MIMEMultipart()
    msg['From'] = SMTP_USER
    msg['To'] = email_dest
    msg['Subject'] = f"🔒 Code d'accès de sécurité - {TITLE_PLATFORM}"
    
    body = f"""Bonjour {nom_etudiant},

Pour valider votre identité et accéder de manière sécurisée à la plateforme, veuillez saisir le code à usage unique suivant :

👉 CODE DE SÉCURITÉ UNIQUE : {code_otp}

Ce code est strictement confidentiel et valable uniquement pour cette session.

Cordialement,
Département d'Électrotechnique
Faculté de génie électrique - UDL-SBA
"""
    msg.attach(MIMEText(body, 'plain', 'utf-8'))
    try:
        server = smtplib.SMTP(SMTP_SERVER, SMTP_PORT)
        server.starttls()
        # server.login(SMTP_USER, SMTP_PASSWORD) # Décommenter lors de la mise en production réelle
        # server.sendmail(SMTP_USER, email_dest, msg.as_string())
        server.quit()
        return True
    except Exception as e:
        return False

# ======================================================================================
# 3. FONCTIONS GRAPHIQUES ET PACKS EXPORTS
# ======================================================================================
def afficher_tableau_bilan_html(df_data, colonnes, entetes, couleur_theme="#0f172a"):
    html_code = "<div style='overflow-x:auto; margin:15px 0; width:100%;'>"
    html_code += f"<table style='width:100%; border-collapse:collapse; font-family:\"Segoe UI\",Arial,sans-serif; background-color:#ffffff; border:1px solid #cbd5e1; border-radius:6px; overflow:hidden;'><thead>"
    html_code += f"<tr style='background-color:{couleur_theme}; color:#ffffff; text-align:left; font-weight:600;'>"
    for h in entetes:
        html_code += f"<th style='padding:10px 14px;'>{h}</th>"
    html_code += "</tr></thead><tbody>"
    for idx, row in df_data.iterrows():
        bg_row = "#f8fafc" if idx % 2 == 0 else "#ffffff"
        html_code += f"<tr style='background-color:{bg_row}; border-bottom:1px solid #e2e8f0;'>"
        for col in colonnes:
            val = row[col]
            if "Absences" in col or "Statut" in col:
                html_code += f"<td style='padding:10px 14px; font-weight:700; color:#b91c1c;'>{val}</td>"
            elif col in ["Étudiant", "Date", "Nom & Prénom Étudiant"]:
                html_code += f"<td style='padding:10px 14px; font-weight:600; color:#1e293b;'>{val}</td>"
            else:
                html_code += f"<td style='padding:10px 14px; color:#334155;'>{val}</td>"
        html_code += "</tr>"
    html_code += "</tbody></table></div>"
    st.markdown(html_code, unsafe_allow_html=True)

def generer_page_html_telechargeable(df_data, titre_bilan, colonnes, entetes):
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
        <p>{TITLE_PLATFORM}</p>
    </div>
    <div class="content">
        <p>Généré le : {datetime.now().strftime('%d/%m/%Y à %H:%M')}</p>
        <table><thead><tr>"""
    for h in entetes:
        html_doc += f"<th>{h}</th>"
    html_doc += "</tr></thead><tbody>"
    for _, row in df_data.iterrows():
        html_doc += "<tr>"
        for col in colonnes:
            if "Absences" in col or "Statut" in col:
                html_doc += f"<td class='abs-count'>{row[col]}</td>"
            else:
                html_doc += f"<td>{row[col]}</td>"
        html_doc += "</tr>"
    html_doc += f"""</tbody></table></div>
    <div class="footer">© 2026 Département d'Électrotechnique - UDL-SBA</div>
</body>
</html>"""
    return html_doc


# ======================================================================================
# 4. PASSERELLE DE VÉRIFICATION D'IDENTITÉ SÉCURISÉE (MATRICULE + OTP)
# ======================================================================================
if "auth_ok" not in st.session_state:
    st.session_state.auth_ok = False
    st.session_state.auth_role = None  # 'etudiant' ou 'enseignant_admin'
    st.session_state.auth_user = None  # Contient les coordonnées de l'étudiant identifié
    st.session_state.otp_code = None
    st.session_state.otp_sent = False
    st.session_state.temp_matricule = ""

if not st.session_state.auth_ok:
    st.markdown("### 🔐 Portail de Contrôle d'Accès")
    
    mode_connexion = st.radio(
        "Sélectionnez votre profil d'accès :", 
        ["Espace Étudiant (Validation par Matricule + Code Unique Email)", "Espace Enseignant / Administration"], 
        horizontal=True
    )
    st.divider()

    if "Espace Étudiant" in mode_connexion:
        st.subheader("🎓 Identification de l'Étudiant")
        st.info("Saisissez votre matricule officiel tel qu'il apparaît sur votre carte d'étudiant pour recevoir votre clé unique d'accès.")
        
        input_mat = st.text_input("🔢 Matricule de l'Étudiant :", key="input_matricule_auth").strip()
        
        if st.button("📩 Recevoir mon code unique de vérification", use_container_width=True):
            if not input_mat:
                st.error("❌ Le champ Matricule ne peut pas être vide.")
            else:
                df_mat = charger_donnees_locales(FILE_MATRICULES)
                if df_mat.empty:
                    st.error(f"⚠️ Erreur système : Le fichier de référence '{FILE_MATRICULES}' est introuvable.")
                else:
                    # Rendre la recherche insensible à la casse et aux espaces superflus
                    df_mat.columns = [c.lower().strip() for c in df_mat.columns]
                    
                    if 'matricule' in df_mat.columns:
                        match_student = df_mat[df_mat['matricule'].astype(str).str.strip() == input_mat]
                        
                        if not match_student.empty:
                            row_student = match_student.iloc[0]
                            # Extraction des 4 colonnes requises
                            nom_p = row_student.get('nom & prenom', row_student.get('nom et prénom', 'L\'étudiant')).strip().upper()
                            promo_s = str(row_student.get('promotion', 'N/A')).strip()
                            email_s = str(row_student.get('email', '')).strip().lower()
                            
                            if not email_s or "@" not in email_s:
                                st.error("❌ Erreur de données : Aucun email valide n'est rattaché à ce matricule.")
                            else:
                                # Génération du code de vérification unique
                                code_gen = str(random.randint(100000, 999999))
                                st.session_state.otp_code = code_gen
                                st.session_state.temp_matricule = input_mat
                                st.session_state.auth_user = {
                                    "matricule": input_mat,
                                    "nom_complet": nom_p,
                                    "promotion": promo_s,
                                    "email": email_s
                                }
                                
                                envoye = envoyer_email_otp(email_s, nom_p, code_gen)
                                st.session_state.otp_sent = True
                                st.success(f"📩 Code de sécurité envoyé à l'adresse sécurisée associée : {email_s[:3]}***@{email_s.split('@')[1]}")
                        else:
                            st.error("❌ Ce matricule n'existe pas dans le fichier source. Veuillez vérifier la saisie.")
                    else:
                        st.error("❌ Structure invalide : La colonne 'matricule' est introuvable dans le fichier excel.")
        
        if st.session_state.otp_sent:
            st.divider()
            st.markdown("#### 🔑 Saisie du code secret")
            st.info(f"💡 **[Simulation]** Clé générée pour le serveur : **{st.session_state.otp_code}** (Entrez-le ci-dessous pour valider).")
            
            code_saisi = st.text_input("Veuillez saisir le code à usage unique reçu :", type="password", key="otp_code_entered").strip()
            
            if st.button("🔓 Confirmer mon identité et Entrer", use_container_width=True):
                if code_saisi == st.session_state.otp_code:
                    st.session_state.auth_ok = True
                    st.session_state.auth_role = "etudiant"
                    st.success(f"✅ Identité validée ! Bienvenue {st.session_state.auth_user['nom_complet']}")
                    time.sleep(1)
                    st.rerun()
                else:
                    st.error("❌ Code de sécurité non valide. Veuillez vérifier vos e-mails.")
                    
    else:
        st.subheader("👤 Accès Enseignant & Administration")
        pwd_global = st.text_input("🔑 Clé d'accès de l'établissement :", type="password", key="pwd_global_platform")
        if st.button("🔓 Connexion", use_container_width=True):
            if pwd_global == "1234":
                st.session_state.auth_ok = True
                st.session_state.auth_role = "enseignant_admin"
                st.success("✅ Accès accordé.")
                time.sleep(1)
                st.rerun()
            else:
                st.error("❌ Mot de passe erroné.")

else:
    # Bandeau supérieur d'informations de session et déconnexion
    c_user, c_deco = st.columns([6, 1])
    with c_user:
        if st.session_state.auth_role == "etudiant":
            st.markdown(f"🎓 Session Active : **{st.session_state.auth_user['nom_complet']}** ({st.session_state.auth_user['promotion']}) | Profil : **Étudiant**")
        else:
            st.markdown("👤 Session Active : **Enseignant / Administration du département**")
    with c_deco:
        if st.button("🚪 Déconnexion", use_container_width=True):
            st.session_state.auth_ok = False
            st.session_state.auth_role = None
            st.session_state.auth_user = None
            st.session_state.otp_sent = False
            st.session_state.otp_code = None
            st.rerun()

    st.divider()

    # --- ARCHITECTURE DES ONGLETS REQUIS ---
    t6, t7 = st.tabs(["📝 ASSIDUITÉ", "📩 SYSTEME DE GESTION DES JUSTIFICATIFS"])

    # ======================================================================================
    # 5. ONGLET T6 : SUIVI DE L'ASSIDUITÉ (RÉSERVÉ ENSEIGNANTS)
    # ======================================================================================
    with t6:
        if st.session_state.auth_role == "etudiant":
            st.warning("🔒 **Accès Refusé :** Cet espace d'enregistrement et de déclaration des absences est strictement réservé au corps enseignant.")
        else:
            pwd_t6 = st.text_input("🔑 Confirmez le code d'accès de session (T6) :", type="password", key="pwd_tab6")
            
            if pwd_t6 == "1234":
                st.markdown(f"### 📝 Saisie et Suivi de l'Assiduité")
                
                df_aff_a = charger_donnees_locales(FILE_DATA_A)
                df_etud_m = charger_donnees_locales(FILE_LISTE_A)

                if df_aff_a.empty or df_etud_m.empty:
                    st.error("⚠️ Les fichiers de configuration d'assiduité sont manquants à la racine.")
                else:
                    c1a, c2a = st.columns(2)
                    with c1a:
                        sel_prof = st.selectbox("👤 Sélectionnez l'Enseignant :", [""] + LISTE_PROFS, key="ens_T6")

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
                                    
                                    try:
                                        res_full = supabase.table("suivi_assiduite_2026").select("*").eq("matiere", sel_mat).eq("promotion", promo_c).execute()
                                        df_db_full = pd.DataFrame(res_full.data) if res_full.data else pd.DataFrame()
                                    except:
                                        df_db_full = pd.DataFrame()

                                    st.markdown("#### 📥 Déclaration d'une nouvelle Absence")
                                    
                                    cn1, cn2, cn3 = st.columns(3)
                                    with cn1:
                                        etud_non = st.selectbox("👤 Sélectionner l'Étudiant :", [""] + noms_e, key="ne_et_t6")
                                    with cn2:
                                        status_assid = st.selectbox("📊 Statut de présence :", ["", "Absent"], key="status_assid_t6")
                                    with cn3:
                                        causes = ["Non justifié", "Décès dans l'ascendance, la descendance ou la parenté", "Mariage de l'intéressé(e)", "Congé de paternité ou de maternité de l'intéressé(e)", "Mission ou convocation officielle", "Maladie de l'intéressé(e)", "Autres"]
                                        cause_s = st.selectbox("❓ Motif initial :", causes, key="ne_ca_t6")

                                    c_d1, c_d2, c_d3 = st.columns(3)
                                    with c_d1:
                                        date_abs = st.date_input("📅 Date de l'absence :", key="date_abs_t6")
                                    with c_d2:
                                        jours_semaine = ["Dimanche", "Lundi", "Mardi", "Mercredi", "Jeudi"]
                                        jour_abs = st.selectbox("🗓️ Jour :", jours_semaine, key="jour_abs_t6")
                                    with c_d3:
                                        horaire_abs = st.selectbox("🕒 Horaire du créneau :", options=HORAIRES_LIST, key="horaire_abs_t6")

                                    if etud_non and status_assid == "Absent":
                                        nb_abs_actuel = len(df_db_full[df_db_full["etud_non_eligible"] == etud_non]) if not df_db_full.empty else 0
                                        st.metric(
                                            label=f"🔢 Compteur d'absences cumulées pour {etud_non}", 
                                            value=f"{nb_abs_actuel} absence(s)"
                                        )

                                    if st.button("💾 ENREGISTRER L'ABSENCE", use_container_width=True):
                                        if not etud_non or status_assid != "Absent":
                                            st.error("❌ Remplissez correctement l'identité de l'étudiant et validez le statut 'Absent'.")
                                        else:
                                            try:
                                                payload = {
                                                    "enseignant": sel_prof, "matiere": sel_mat, "promotion": promo_c,
                                                    "etud_non_eligible": etud_non, "cause_non_eligibilite": cause_s or "Non justifié",
                                                    "date_absence": str(date_abs), "jour_absence": jour_abs, "horaire_absence": horaire_abs,
                                                    "date_saisie": datetime.now().strftime("%d/%m/%Y %H:%M")
                                                }
                                                supabase.table("suivi_assiduite_2026").insert(payload).execute()
                                                st.success(f"✅ Absence enregistrée pour {etud_non}.")
                                                time.sleep(1) 
                                                st.rerun()
                                            except Exception as e:
                                                st.error(f"❌ Erreur d'écriture : {e}")

                                    st.divider()
                                    st.subheader("📋 Registre de Suivi des Absences")
                                    
                                    if not df_db_full.empty and "etud_non_eligible" in df_db_full.columns:
                                        dict_compteurs = df_db_full["etud_non_eligible"].value_counts().to_dict()
                                        df_liste_globale = df_db_full.copy()
                                        df_liste_globale["Total Absences"] = df_liste_globale["etud_non_eligible"].map(dict_compteurs)
                                        
                                        affichage_cols = {
                                            "enseignant": "Chargé de Cours", "matiere": "Matière", "promotion": "Promotion",
                                            "etud_non_eligible": "Nom & Prénom Étudiant", "jour_absence": "Jour",
                                            "date_absence": "Date Absence", "horaire_absence": "Horaire",
                                            "cause_non_eligibilite": "Motif", "Total Absences": "🔢 Total"
                                        }
                                        df_affichage_table = df_liste_globale[list(affichage_cols.keys())].rename(columns=affichage_cols)
                                        st.dataframe(df_affichage_table.sort_values(by="🔢 Total", ascending=False), use_container_width=True, hide_index=True)
                                        
                                        if st.button("🗑️ Réinitialiser les absences pour cette matière", type="primary", use_container_width=True):
                                            supabase.table("suivi_assiduite_2026").delete().eq("matiere", sel_mat).eq("promotion", promo_c).execute()
                                            st.success("✅ Historique remis à zéro.")
                                            time.sleep(1)
                                            st.rerun()
                                    else:
                                        st.info("ℹ️ Aucune absence relevée pour cette matière.")

                                    # Zone Extraction de rapports
                                    st.divider()
                                    st.subheader("📥 Extraction des Rapports Officiels d'Exclusion")
                                    try:
                                        res_excl = supabase.table("suivi_assiduite_2026").select("etud_non_eligible").eq("matiere", sel_mat).eq("promotion", promo_c).execute()
                                        noms_exclus = [r['etud_non_eligible'] for r in res_excl.data if r.get('etud_non_eligible')]
                                    except: noms_exclus = []

                                    output = io.BytesIO()
                                    df_eligible_final = df_p[~df_p["Nom_Complet"].isin(noms_exclus)].copy()
                                    export_eli = pd.DataFrame({"Nom et Prénom": df_eligible_final["Nom_Complet"], "Matière": sel_mat, "Chargé": sel_prof, "Promotion": promo_c})
                                    
                                    if not df_db_full.empty and "etud_non_eligible" in df_db_full.columns:
                                        df_non_eligible = df_db_full[df_db_full["etud_non_eligible"].notna()].copy()
                                        export_non = df_non_eligible[["etud_non_eligible", "cause_non_eligibilite", "date_absence", "matiere", "promotion"]].rename(columns={"etud_non_eligible": "Nom et Prénom", "cause_non_eligibilite": "Motif"})
                                    else: export_non = pd.DataFrame()

                                    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                                        workbook = writer.book
                                        fmt_title = workbook.add_format({'bold': True, 'font_size': 13, 'align': 'center'})
                                        fmt_bold = workbook.add_format({'bold': True})
                                        
                                        ws1 = workbook.add_worksheet('Éligibles')
                                        export_eli.to_excel(writer, sheet_name='Éligibles', startrow=5, index=False)
                                        ws1.write('A1', "UNIVERSITÉ DJILLALI LIABÈS - SIDI BEL ABBÈS", fmt_bold)
                                        ws1.write('A2', f"LISTE DES ÉTUDIANTS ÉLIGIBLES - MATIÈRE : {sel_mat}", fmt_title)
                                        
                                        if not export_non.empty:
                                            export_non.to_excel(writer, sheet_name='Non-Éligibles', startrow=5, index=False)

                                    st.download_button(label="📥 TÉLÉCHARGER LE RAPPORT COMPTABLE EXCEL (XLSX)", data=output.getvalue(), file_name=f"Rapport_Assiduite_{sel_mat}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)

            elif pwd_t6 != "":
                st.error("❌ Clé d'accès erronée.")

    # ======================================================================================
    # 6. ONGLET T7 : SYSTEME DE GESTION DES REQUÊTES ET JUSTIFICATIFS (SÉCURISÉ)
    # ======================================================================================
    with t7:
        st.header("📩 Espace National des Justificatifs d'Absence")
        
        if st.session_state.auth_role == "etudiant":
            choix_vue = "Étudiant (Dépôt)"
            st.info(f"✨ Mode Sécurisé : Votre espace est restreint à l'identité certifiée de **{st.session_state.auth_user['nom_complet']}**.")
        else:
            choix_vue = st.radio("Sélectionnez l'interface de travail :", ["Étudiant (Dépôt)", "Administration (Décision)"], horizontal=True, key="radio_role_t7")
        
        st.divider()

        if choix_vue == "Étudiant (Dépôt)":
            st.subheader("📤 Formulaire de Dépôt de Justificatif Officiel")
            with st.form("form_depot_pdf_etudiant", clear_on_submit=True):
                col1, col2 = st.columns(2)
                with col1:
                    if st.session_state.auth_role == "etudiant":
                        # VERROUILLAGE SÉCURISÉ DES PARAMÈTRES ISSUS DE L'OTP
                        etudiant_select = st.text_input("Nom & Prénom de l'Étudiant :", value=st.session_state.auth_user['nom_complet'], disabled=True)
                        promotion_courante = st.session_state.auth_user['promotion']
                        st.text_input("Filière / Promotion de rattachement :", value=promotion_courante, disabled=True)
                    else:
                        df_etud_m = charger_donnees_locales(FILE_LISTE_A)
                        if not df_etud_m.empty:
                            df_etud_m["Nom_Complet"] = df_etud_m["Nom"].str.upper() + " " + df_etud_m["Prénom"].str.title()
                            liste_etudiants_dispo = sorted(df_etud_m["Nom_Complet"].unique().tolist())
                        else: liste_etudiants_dispo = ["Aucune liste chargée"]
                        etudiant_select = st.selectbox("Sélectionnez l'Étudiant :", liste_etudiants_dispo)
                        promotion_courante = "Administration"

                    df_aff_loc = charger_donnees_locales(FILE_DATA_A)
                    liste_mat_dispo = sorted(df_aff_loc["Enseignements"].unique().tolist()) if not df_aff_loc.empty else []
                    matiere_select = st.selectbox("Matière faisant l'objet de la demande :", liste_mat_dispo)
                    
                with col2:
                    causes_absences = ["Décès dans l'ascendance, la descendance ou la parenté", "Mariage de l'intéressé(e)", "Congé de paternité ou de maternité de l'intéressé(e)", "Mission ou convocation officielle", "Maladie de l'intéressé(e)", "Autres"]
                    motif_abs = st.selectbox("Motif légal justifiant l'absence :", causes_absences)
                    fichier_pdf = st.file_uploader("Fichier justificatif au format PDF (Obligatoire)", type=["pdf"])

                submit_valider = st.form_submit_button("🚀 TRANSMETTRE LE DOSSIER À L'ADMINISTRATION")

            if submit_valider:
                if not fichier_pdf:
                    st.error("❌ Dépôt impossible : Veuillez joindre un document PDF valide.")
                else:
                    try:
                        pdf_bytes = fichier_pdf.read()
                        pdf_encoded = base64.b64encode(pdf_bytes).decode('utf-8')
                        
                        data_insert = {
                            "date_demande": datetime.now().strftime("%d/%m/%Y"), 
                            "nom_etudiant": etudiant_select, "matiere": matiere_select, 
                            "promotion": promotion_courante, "motif": motif_abs,
                            "justificatif_pdf": pdf_encoded, "statut": "En attente"
                        }
                        supabase.table("requetes_absences").insert(data_insert).execute()
                        st.success("✅ Votre demande a bien été transmise à la scolarité du département d'Électrotechnique.")
                        st.balloons()
                        time.sleep(1)
                        st.rerun()
                    except Exception as e:
                        st.error(f"❌ Erreur lors du transfert : {e}")
        else:
            # Traitement des décisions administratives
            pwd_admin = st.text_input("🔑 Saisissez le code d'accès Administration :", type="password", key="pwd_admin_t7_sec")

            if pwd_admin == "1234":
                st.subheader("⚖️ Dossiers d'Absence en attente de Validation")
                try:
                    query_admin = supabase.table("requetes_absences").select("*").eq("statut", "En attente").execute()
                    resultats_attente = query_admin.data

                    if not resultats_attente:
                        st.info("📭 Aucun dossier en attente de traitement.")
                    else:
                        for req in resultats_attente:
                            with st.expander(f"📄 Dossier de : {req['nom_etudiant']} — Matière : {req['matiere']}"):
                                st.write(f"**Promotion :** {req['promotion']} | **Motif invoqué :** {req['motif']}")
                                pdf_decoded_bytes = base64.b64decode(req['justificatif_pdf'])
                                st.download_button(label="👁️ Télécharger le justificatif joint (PDF)", data=pdf_decoded_bytes, file_name=f"Justificatif_{req['nom_etudiant']}.pdf", mime="application/pdf", key=f"dl_pdf_{req['id']}")
                                
                                c_acc, c_rej = st.columns(2)
                                if c_acc.button("✅ ACCORDER LA RÉHABILITATION", key=f"btn_acc_{req['id']}", use_container_width=True):
                                    supabase.table("requetes_absences").update({"statut": "Favorable"}).eq("id", req['id']).execute()
                                    # Suppression automatique de l'absence pour réhabiliter l'étudiant
                                    supabase.table("suivi_assiduite_2026").delete().eq("etud_non_eligible", req['nom_etudiant']).eq("matiere", req['matiere']).execute()
                                    st.success(f"✔️ {req['nom_etudiant']} réhabilité(e).")
                                    time.sleep(1)
                                    st.rerun()
                                if c_rej.button("❌ REJETER LE DOSSIER", key=f"btn_rej_{req['id']}", use_container_width=True):
                                    supabase.table("requetes_absences").update({"statut": "Défavorable"}).eq("id", req['id']).execute()
                                    st.warning("Dossier rejeté.")
                                    time.sleep(1)
                                    st.rerun()
                except Exception as e:
                    st.error(f"Erreur Supabase : {e}")

        # Section Registres et Statistiques de Visibilité (Filtrée ou globale selon session)
        st.divider()
        st.subheader("📊 Registre Suivi des Demandes de Réhabilitation")
        try:
            df_aff_loc = charger_donnees_locales(FILE_DATA_A)
            
            # Filtrage de sécurité : un étudiant ne peut consulter que ses propres justificatifs déposés
            if st.session_state.auth_role == "etudiant":
                query_historique = supabase.table("requetes_absences").select("date_demande, nom_etudiant, matiere, promotion, motif, statut").eq("nom_etudiant", st.session_state.auth_user['nom_complet']).execute()
            else:
                query_historique = supabase.table("requetes_absences").select("date_demande, nom_etudiant, matiere, promotion, motif, statut").execute()
                
            data_historique = query_historique.data

            if data_historique:
                df_tab = pd.DataFrame(data_historique)
                if not df_aff_loc.empty and "Enseignants" in df_aff_loc.columns:
                    map_enseignants = df_aff_loc.set_index(["Enseignements", "Promotion"])["Enseignants"].to_dict()
                    df_tab["Enseignant"] = df_tab.apply(lambda r: map_enseignants.get((r["matiere"], r["promotion"]), "Non assigné"), axis=1)
                else: df_tab["Enseignant"] = "N/A"

                df_tab = df_tab[["date_demande", "promotion", "Enseignant", "nom_etudiant", "matiere", "motif", "statut"]]
                df_tab.columns = ["Date Saisie", "Promotion", "Enseignant", "Étudiant", "Matière", "Motif Déclaré", "Statut Décision"]
                
                afficher_tableau_bilan_html(df_tab, df_tab.columns, df_tab.columns, couleur_theme="#1e3a8a")
            else:
                st.info("ℹ️ Aucun historique disponible pour cette session.")
        except Exception as e:
            st.error(f"Erreur d'affichage : {e}")
