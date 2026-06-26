import streamlit as st
import smtplib
import random
import pandas as pd
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

# 1. Rappel obligatoire du Titre de la plateforme
TITRE_PLATEFORME = "Plateforme de gestion des EDTs-S2-2026-Département d'Électrotechnique-Faculté de génie électrique-UDL-SBA"
st.set_page_config(page_title=TITRE_PLATEFORME, layout="wide")
st.title(TITRE_PLATEFORME)

# Fonction réelle d'envoi d'e-mail par SMTP
def envoyer_code_par_email(email_destinataire, code_unique):
    # Configuration du serveur SMTP de Gmail
    smtp_server = "chef.department.elt.fge@gmail.com"
    smtp_port = 587
    
    # Récupération sécurisée des identifiants depuis st.secrets
    try:
        smtp_user = st.secrets["chef.department.elt.fge@gmail.com"]
        smtp_password = st.secrets["gkzs pdza yodb icvd"]
    except KeyError:
        st.error("⚠️ Configuration SMTP manquante dans `.streamlit/secrets.toml`")
        return False

    # Création du message
    sujet = "🔑 Votre code d'activation unique - Gestion des EDTs S2 2026"
    corps_message = f"""
    Bonjour,
    
    Vous avez demandé un accès à la Plateforme de gestion des EDTs (S2-2026).
    Voici votre clé unique d'accès à usage unique :
    
    👉 {code_unique}
    
    Veuillez saisir ce code sur l'interface pour valider votre identité.
    
    Cordialement,
    Département d'Électrotechnique
    Faculté de génie électrique - UDL-SBA
    """
    
    msg = MIMEMultipart()
    msg['From'] = smtp_user
    msg['To'] = email_destinataire
    msg['Subject'] = sujet
    msg.attach(MIMEText(corps_message, 'plain', 'utf-8'))
    
    try:
        server = smtplib.SMTP(smtp_server, smtp_port)
        server.starttls()  # Sécurisation de la connexion
        server.login(smtp_user, smtp_password)
        server.sendmail(smtp_user, email_destinataire, msg.as_string())
        server.quit()
        return True
    except Exception as e:
        st.error(f"Erreur technique lors de l'envoi de l'e-mail : {e}")
        return False

# Initialisation du State Streamlit si non existant
if "code_serveur" not in st.session_state:
    st.session_state.code_serveur = None
if "email_envoye" not in st.session_state:
    st.session_state.email_envoye = False
if "authentifie" not in st.session_state:
    st.session_state.authentifie = False

# 🔐 Portail de Contrôle d'Accès
st.header("🔐 Portail de Contrôle d'Accès")
profil = st.radio(
    "Sélectionnez votre profil d'accès :",
    [
        "Espace Étudiant (Validation par Identité Complète + Code Unique Email)",
        "Espace Enseignant / Administration"
    ]
)

if "Espace Étudiant" in profil:
    if not st.session_state.authentifie:
        st.subheader("🎓 Identification de l'Étudiant")
        st.write("Saisissez votre matricule, nom, prénom et e-mail pour recevoir votre clé unique d'accès.")
        
        # Formulaire de saisie
        matricule = st.text_input("🔢 Matricule de l'Étudiant (Mat. Etudiant) :")
        nom = st.text_input("👤 Nom de famille (en Majuscules) :").upper()
        prenom = st.text_input("👤 Prénom :")
        email_etudiant = st.text_input("📧 Votre Adresse E-mail (pour l'envoi du code) :")
        
        if st.button("📩 Demander mon code d'activation"):
            if matricule and nom and prenom and email_etudiant:
                # Génération d'un code robuste à 6 chiffres
                code_genere = str(random.randint(100000, 999999))
                
                with st.spinner("Envoi du code de sécurité en cours..."):
                    # Appel de la fonction d'envoi réel
                    succes = envoyer_code_par_email(email_etudiant, code_genere)
                    
                    if succes:
                        st.session_state.code_serveur = code_genere
                        st.session_state.email_envoye = True
                        st.success(f"✅ Code de sécurité envoyé avec succès à l'adresse : {email_etudiant}")
            else:
                st.warning("⚠️ Veuillez remplir tous les champs avant de demander le code.")

        # 🔑 Saisie du code secret (S'affiche uniquement après envoi de l'e-mail)
        if st.session_state.email_envoye:
            st.write("---")
            st.subheader("🔑 Saisie du code secret")
            code_saisi = st.text_input("Veuillez saisir le code à usage unique reçu par e-mail :", type="default")
            
            if st.button("Valider et Entrer"):
                if code_saisi == st.session_state.code_serveur:
                    st.session_state.authentifie = True
                    st.success("🎉 Authentification réussie !")
                    st.rerun()
                else:
                    st.error("❌ Code incorrect. Veuillez vérifier votre boîte de réception ou générer un nouveau code.")
                    
    else:
        # Zone connectée : Affichage de l'EDT selon la disposition demandée
        st.success(f"Bienvenue dans votre espace étudiant, {prenom} {nom} !")
        st.subheader("🗓️ Votre Emploi du Temps - Semestre 2 (2026)")
        
        # Données de l'EDT mémorisées avec la disposition demandée :
        # Enseignements, Code, Enseignants, Horaire, Jours, Lieu, Promotion
        data_edt = {
            "Enseignements": [
                "Stabilité et dynamique des réseaux électriques", 
                "Éclairage LED: Principes et applications",
                "Techniques d'intelligence artificielle",
                "Intégration des ressources renouvelables aux réseaux électriques",
                "Dimensionnement des Réseaux électriques industriels",
                "Technique de la haute tension",
                "Conduite des réseaux électriques",
                "Réseaux électriques intelligents"
            ],
            "Code": ["Cours-SDRE-RE", "Cours-LEDPA-RE", "Cours-TIA-RE", "Cours-IRRRE-RE", "Cours-DREI-RE", "Cours-THT-RE", "Cours-CdRE-RE", "Cours-REI-RE"],
            "Enseignants": ["Zidi", "Bermaki", "Touhami", "BENHAMIDA", "Rezoug", "Bellebna", "Benhamida", "Maamar"],
            "Horaire": ["8h - 9h30", "9h30 - 11h", "9h30 - 11h", "14h00 - 15h30", "9h30 - 11h", "12h30 - 14h00", "8h - 9h30", "9h30 - 11h"],
            "Jours": ["Dimanche", "Dimanche", "Lundi", "Lundi", "Mardi", "Mardi", "Mercredi", "Mercredi"],
            "Lieu": ["S06", "S06", "S06", "S06", "S06", "S06", "S06", "S06"],
            "Promotion": ["M2RE", "M2RE", "M2RE", "M2RE", "M2RE", "M2RE", "M2RE", "M2RE"]
        }
        
        df_edt = pd.DataFrame(data_edt)
        st.dataframe(df_edt, use_container_width=True)

elif "Espace Enseignant" in profil:
    st.subheader("👤 Espace Enseignant / Administration")
    st.info("Interface de gestion administrative en cours d'initialisation.")
