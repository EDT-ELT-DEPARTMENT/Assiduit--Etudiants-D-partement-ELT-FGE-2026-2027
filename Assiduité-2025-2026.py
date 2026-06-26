import streamlit as st
import pandas as pd
import os, random, io, smtplib
import base64
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from datetime import datetime, timedelta
from supabase import create_client
import time

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
TITLE_PLATFORM = "Plateforme de gestion d"assiduité-S2-2026-2027-Département d'Électrotechnique-Faculté de génie électrique-UDL-SBA"
st.title(TITLE_PLATFORM)

# ======================================================================================
# 1. PARAMÈTRES ET CONFIGURATIONS CRITIQUES
# ======================================================================================
NOM_SOURCE = "dataEDT-ELT-S2-2026.xlsx"
FILE_EMAILS = "Permanents-Vacataires-ELT2-2025-2026.xlsx"
FILE_DATA_A = "DATA-ASSUIDUITE-2026.xlsx"
FILE_LISTE_A = "Liste des étudiants-2025-2026.xlsx"

TABLE_NAME = "surveillances_2026"
COLS_ORDRE = ['Enseignements', 'Code', 'Enseignants', 'Horaire', 'Jours', 'Lieu', 'Promotion']

# Connexion Base de Données Supabase
S_URL = "https://ajcbkidmcjtyomknijwa.supabase.co"
S_KEY = "sb_publishable_otn3XM8LPLV0OGw74LRhDw_F446jkpw"
supabase = create_client(S_URL, S_KEY)

# Liste des 14 créneaux réglementaires
HORAIRES_LIST = [
    "8h - 9h", "8h - 9h30", "8h - 10h", "9h - 10h", "9h30 - 11h", 
    "10h - 11h", "11h - 12h", "11h - 12h30", 
    "12h - 13h", "12h30 - 14h", "13h - 14h", "14h - 15h30", "14h - 16h", "15h30 - 17h"
]

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
# 2. MOTEUR DE CHARGEMENT ET PARSING DES FICHIERS LOCAUX
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

# ======================================================================================
# 3. CONFECTION DES PACKS EXPORTS (FONCTIONS GRAPHIQUES)
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
            if "Absences" in col:
                html_code += f"<td style='padding:10px 14px; font-weight:700; color:#b91c1c;'>{val}</td>"
            elif col == "Étudiant" or col == "Date":
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

# --- ROUTAGE DES TABS UNIQUES ---
t6, t7 = st.tabs(["📝 ASSIDUITÉ", "📩 SYSTEME DE GESTION DES JUSTIFICATIFS"])

# ======================================================================================
# 4. ONGLET T6 : SUIVI DE L'ASSIDUITÉ ET LOGIQUE D'ENREGISTREMENT
# ======================================================================================
with t6:
    pwd_t6 = st.text_input("🔑 Code d'accès (T6) :", type="password", key="pwd_tab6")
    
    if pwd_t6 == "1234":
        st.markdown(f"### 📝 Suivi de l'Assiduité et Compteur d'Absences")
        
        df_aff_a = charger_donnees_locales(FILE_DATA_A)
        df_etud_m = charger_donnees_locales(FILE_LISTE_A)

        if df_aff_a.empty or df_etud_m.empty:
            st.error("⚠️ Fichiers sources (.xlsx) introuvables ou vides à la racine du serveur.")
        else:
            c1a, c2a = st.columns(2)
            with c1a:
                list_p_t6 = LISTE_PROFS
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
                            
                            # Récupération en temps réel depuis la base Supabase
                            try:
                                res_full = supabase.table("suivi_assiduite_2026").select("*").eq("matiere", sel_mat).eq("promotion", promo_c).execute()
                                df_db_full = pd.DataFrame(res_full.data) if res_full.data else pd.DataFrame()
                            except:
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

                            c_d1, c_d2, c_d3 = st.columns(3)
                            with c_d1:
                                date_abs = st.date_input("📅 Date de l'absence :", key="date_abs_t6")
                            with c_d2:
                                jours_semaine = ["Dimanche", "Lundi", "Mardi", "Mercredi", "Jeudi"]
                                jour_abs = st.selectbox("🗓️ Jour :", jours_semaine, key="jour_abs_t6")
                            with c_d3:
                                horaire_abs = st.selectbox("🕒 Horaire du créneau :", options=HORAIRES_LIST, key="horaire_abs_t6")

                            # Compteur dynamique en fonction des filtres actifs
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

                            # Table d'affichage globale
                            st.divider()
                            st.subheader("📋 Liste Globale de Suivi des Absences")
                            
                            if not df_db_full.empty and "etud_non_eligible" in df_db_full.columns:
                                dict_compteurs = df_db_full["etud_non_eligible"].value_counts().to_dict()
                                df_liste_globale = df_db_full.copy()
                                df_liste_globale["Total Absences Cumulées"] = df_liste_globale["etud_non_eligible"].map(dict_compteurs)
                                
                                affichage_cols = {
                                    "enseignant": "Chargé de Cours", "matiere": "Matière", "promotion": "Promotion",
                                    "etud_non_eligible": "Nom & Prénom Étudiant", "jour_absence": "Jour",
                                    "date_absence": "Date Absence", "horaire_absence": "Horaire",
                                    "cause_non_eligibilite": "Motif / Justification", "Total Absences Cumulées": "🔢 Compteur Total"
                                }
                                df_affichage_table = df_liste_globale[list(affichage_cols.keys())].rename(columns=affichage_cols)
                                df_affichage_table = df_affichage_table.sort_values(by=["🔢 Compteur Total", "Nom & Prénom Étudiant"], ascending=[False, True])
                                
                                st.dataframe(df_affichage_table, use_container_width=True, hide_index=True)
                                
                                # Bouton d'effacement spécifique pour la matière active
                                if st.button("🗑️ Effacer l'historique", type="primary", use_container_width=True):
                                    try:
                                        supabase.table("suivi_assiduite_2026").delete().eq("matiere", sel_mat).eq("promotion", promo_c).execute()
                                        st.success("✅ L'historique des absences pour cette matière a été effacé avec succès !")
                                        time.sleep(1)
                                        st.rerun()
                                    except Exception as e:
                                        st.error(f"❌ Erreur lors de la suppression de l'historique : {e}")
                            else:
                                st.info(f"ℹ️ Aucune absence enregistrée pour le moment dans la matière {sel_mat} ({promo_c}).")

                            # Section Extraction des rapports Officiels
                            st.divider()
                            st.subheader("📥 Extraction des Rapports Officiels")
                            if sel_prof and sel_mat:
                                try:
                                    res_excl = supabase.table("suivi_assiduite_2026").select("etud_non_eligible").eq("matiere", sel_mat).eq("promotion", promo_c).execute()
                                    noms_exclus = [r['etud_non_eligible'] for r in res_excl.data if r.get('etud_non_eligible')]
                                except:
                                    noms_exclus = []

                                if not df_p.empty and "Nom_Complet" in df_p.columns:
                                    output = io.BytesIO()
                                    df_eligible_final = df_p[~df_p["Nom_Complet"].isin(noms_exclus)].copy()
                                    export_eli = pd.DataFrame({
                                        "Nom et Prénom": df_eligible_final["Nom_Complet"], "Matière": sel_mat, "Chargé": sel_prof, "Promotion": promo_c
                                    })
                                    if not df_db_full.empty and "etud_non_eligible" in df_db_full.columns:
                                        mask_non_eli = (df_db_full["etud_non_eligible"].notna()) & (df_db_full["etud_non_eligible"] != "")
                                        df_non_eligible = df_db_full[mask_non_eli].copy()
                                        cols_export = ["etud_non_eligible", "cause_non_eligibilite", "date_absence", "jour_absence", "horaire_absence", "matiere", "enseignant", "promotion"]
                                        export_non = df_non_eligible[cols_export].rename(columns={
                                            "etud_non_eligible": "Nom et Prénom", "cause_non_eligibilite": "Motif du Retrait",
                                            "date_absence": "Date Absence", "jour_absence": "Jour", "horaire_absence": "Horaire",
                                            "matiere": "Matière", "enseignant": "Chargé", "promotion": "Promotion"
                                        })
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

            # Zone d'envoi de convocations
            st.divider()
            st.subheader("✉️ Envoi des Convocations")
            c_conv1, c_conv2 = st.columns(2)
            with c_conv1:
                st.markdown("#### 👤 Envoi Individuel")
                sel_prof_conv = st.selectbox("Sélectionner l'Enseignant :", [""] + LISTE_PROFS, key="conv_prof_T6")
                if st.button("✉️ Envoyer la convocation individuelle", use_container_width=True, key="btn_conv_ind"):
                    if sel_prof_conv:
                        st.success(f"✅ Convocation envoyée avec succès à {sel_prof_conv} !")
                    else:
                        st.error("❌ Veuillez sélectionner un enseignant.")
            with c_conv2:
                st.markdown("#### 📢 Envoi Massif")
                st.warning("⚠️ Attention : Cette action traite tous les enseignants configurés.")
                if st.button("📢 Lancer l'envoi massif", use_container_width=True, key="btn_conv_mass"):
                    st.success("✅ Envoi massif des convocations terminé avec succès pour tous les enseignants !")
    elif pwd_t6 != "":
        st.error("❌ Code incorrect.")

# ======================================================================================
# 5. ONGLET T7 : SYSTEME COMPLET DE GESTION DES REQUÊTES ET JUSTIFICATIFS
# ======================================================================================
with t7:
    st.header("📩 Système de Gestion des Justificatifs")
    st.caption(TITLE_PLATFORM)
    st.write("Espace permettant la soumission des justificatifs d'absence étudiants et validation administrative.")

    choix_vue = st.radio("Sélectionnez votre profil :", ["Étudiant (Dépôt)", "Administration (Décision)"], horizontal=True)
    st.divider()

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
                causes_absences = ["Non justifié", "Décès dans l'ascendance, la descendance ou la parenté", "Mariage de l'intéressé(e)", "Congé de paternité ou de maternité de l'intéressé(e)", "Mission ou convocation officielle", "Maladie de l'intéressé(e)", "Autres"]
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
                        "date_demande": datetime.now().strftime("%d/%m/%Y"), "nom_etudiant": etudiant_select,
                        "matiere": matiere_select, "promotion": promotion_courante, "motif": motif_abs,
                        "justificatif_pdf": pdf_encoded, "statut": "En attente"
                    }
                    supabase.table("requetes_absences").insert(data_insert).execute()
                    st.success(f"✅ Demande enregistrée avec succès pour l'étudiant **{etudiant_select}** !")
                    st.balloons()
                    time.sleep(1)
                    st.rerun()
                except Exception as e:
                    st.error(f"❌ Erreur technique lors de la transmission : {e}")
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
                            pdf_decoded_bytes = base64.b64decode(req['justificatif_pdf'])
                            st.download_button(label="👁️ Ouvrir / Télécharger le justificatif PDF", data=pdf_decoded_bytes, file_name=f"Justification_{req['nom_etudiant']}.pdf", mime="application/pdf", key=f"dl_pdf_{req['id']}")
                            
                            st.markdown("---")
                            col_accorder, col_rejeter = st.columns(2)
                            if col_accorder.button("✅ ACCORDER LA RÉHABILITATION", key=f"btn_acc_id_{req['id']}", use_container_width=True):
                                supabase.table("requetes_absences").update({"statut": "Favorable"}).eq("id", req['id']).execute()
                                supabase.table("suivi_assiduite_2026").delete().eq("etud_non_eligible", req['nom_etudiant']).eq("matiere", req['matiere']).execute()
                                st.success("✔️ L'étudiant a été réhabilité.")
                                time.sleep(1)
                                st.rerun()
                            if col_rejeter.button("❌ REJETER LE DOSSIER", key=f"btn_rej_id_{req['id']}", use_container_width=True):
                                supabase.table("requetes_absences").update({"statut": "Défavorable"}).eq("id", req['id']).execute()
                                st.warning("❌ Demande rejetée.")
                                time.sleep(1)
                                st.rerun()

                # Zone Maintenance Critique T7
                st.markdown("<br>", unsafe_allow_html=True)
                with st.expander("🛠️ Zone de Maintenance Critique (Administration uniquement)"):
                    if st.button("🔄 RÉINITIALISER TOUT LE SYSTÈME", type="primary", use_container_width=True, key="btn_clear_t7_db"):
                        st.session_state.execute_reset_confirmation = True
                    if st.session_state.get("execute_reset_confirmation", False):
                        st.error("⚠️ **ATTENTION CRITIQUE :** Effacement définitif de toutes les données justificatives.")
                        col_c1, col_c2 = st.columns(2)
                        with col_c1:
                            if st.button("🔥 OUI, CONFIRMER LA SUPPRESSION", type="primary", use_container_width=True):
                                supabase.table("requetes_absences").delete().neq("id", -1).execute()
                                st.session_state.execute_reset_confirmation = False
                                st.success("Base réinitialisée.")
                                time.sleep(1)
                                st.rerun()
                        with col_c2:
                            if st.button("❌ ANNULER L'ACTION", use_container_width=True):
                                st.session_state.execute_reset_confirmation = False
                                st.rerun()
            except Exception as e:
                st.error(f"Erreur d'accès : {e}")

        # Section Registre Général et Double Exports (Excel et HTML)
        st.divider()
        st.subheader("📊 Registre et État d'Avancement des Demandes")
        try:
            promo_filtre = promo_c if 'promo_c' in locals() else ""
            df_aff_loc = charger_donnees_locales(FILE_DATA_A)
            query_historique = supabase.table("requetes_absences").select("date_demande, nom_etudiant, matiere, promotion, motif, statut").eq("promotion", promo_filtre).execute()
            data_historique = query_historique.data

            if data_historique:
                df_tab = pd.DataFrame(data_historique)
                if not df_aff_loc.empty and "Enseignants" in df_aff_loc.columns:
                    map_enseignants = df_aff_loc.set_index(["Enseignements", "Promotion"])["Enseignants"].to_dict()
                    df_tab["Chargé de Matière"] = df_tab.apply(lambda r: map_enseignants.get((r["matiere"], r["promotion"]), "Non assigné"), axis=1)
                else:
                    df_tab["Chargé de Matière"] = "N/A"

                df_tab = df_tab[["date_demande", "promotion", "Chargé de Matière", "nom_etudiant", "matiere", "motif", "statut"]]
                df_tab.columns = ["Date", "Promotion", "Chargé de Matière", "Étudiant", "Matière", "Motif", "Statut"]
                
                # Rendu du Tableau principal
                afficher_tableau_bilan_html(df_tab, df_tab.columns, df_tab.columns, couleur_theme="#1e3a8a")

                # Exports Registre Général
                buf_reg = io.BytesIO()
                with pd.ExcelWriter(buf_reg, engine='xlsxwriter') as w: df_tab.to_excel(w, index=False)
                html_reg = generer_page_html_telechargeable(df_tab, "Registre Général des Demandes", df_tab.columns, df_tab.columns)
                
                col_x, col_h = st.columns(2)
                col_x.download_button("📥 REGISTRE VERS EXCEL", data=buf_reg.getvalue(), file_name="Registre.xlsx", use_container_width=True)
                col_h.download_button("🌐 REGISTRE EN HTML", data=html_reg, file_name="Registre.html", use_container_width=True)

                # --- TABLEAU 2 : BILAN PAR ÉTUDIANT ET MATIÈRE ---
                st.markdown("<br>", unsafe_allow_html=True)
                st.subheader("📚 Bilan du Nombre d'Absences par Étudiant et par Matière")
                df_bilan_mat = df_tab.groupby(["Étudiant", "Matière", "Chargé de Matière", "Promotion"]).size().reset_index(name="Nombre d'Absences")
                df_bilan_mat = df_bilan_mat[["Étudiant", "Promotion", "Matière", "Chargé de Matière", "Nombre d'Absences"]]
                
                afficher_tableau_bilan_html(df_bilan_mat, df_bilan_mat.columns, df_bilan_mat.columns, couleur_theme="#1e40af")
                
                buf_mat = io.BytesIO()
                with pd.ExcelWriter(buf_mat, engine='xlsxwriter') as w: df_bilan_mat.to_excel(w, index=False)
                html_mat = generer_page_html_telechargeable(df_bilan_mat, "Bilan Absences par Matière", df_bilan_mat.columns, df_bilan_mat.columns)
                
                cx1, cx2 = st.columns(2)
                cx1.download_button("📥 BILAN MATIÈRES VERS EXCEL", data=buf_mat.getvalue(), file_name="Bilan_Matieres.xlsx", use_container_width=True)
                cx2.download_button("🌐 BILAN MATIÈRES EN HTML", data=html_mat, file_name="Bilan_Matieres.html", use_container_width=True)

                # --- TABLEAU 3 : TOTAL CUMULÉ ---
                st.markdown("<br>", unsafe_allow_html=True)
                st.subheader("👥 Total des Absences Cumulées par Étudiant")
                df_bilan_etud = df_tab.groupby(["Étudiant", "Promotion"]).size().reset_index(name="Total Absences").sort_values(by="Total Absences", ascending=False)
                
                afficher_tableau_bilan_html(df_bilan_etud, df_bilan_etud.columns, ["Nom de l'Étudiant", "Promotion", "Total Absences"], couleur_theme="#0f172a")
                
                buf_etud = io.BytesIO()
                with pd.ExcelWriter(buf_etud, engine='xlsxwriter') as w: df_bilan_etud.to_excel(w, index=False)
                html_etud = generer_page_html_telechargeable(df_bilan_etud, "Total Absences par Étudiant", df_bilan_etud.columns, df_bilan_etud.columns)
                
                cc1, cc2 = st.columns(2)
                cc1.download_button("📥 TOTAL ÉTUDIANTS VERS EXCEL", data=buf_etud.getvalue(), file_name="Total_Absences.xlsx", use_container_width=True)
                cc2.download_button("🌐 TOTAL ÉTUDIANTS EN HTML", data=html_etud, file_name="Total_Absences.html", use_container_width=True)
            else:
                st.info("ℹ️ Aucun historique de demande disponible pour le moment.")
        except Exception as e:
            st.error(f"Erreur Génération rapports T7 : {e}")
