import streamlit as st
import fitz  # PyMuPDF
import re
import pandas as pd
import json
import requests
import datetime
from io import BytesIO
import io
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
import pytesseract
from PIL import Image

st.set_page_config(page_title="METRE-PRO System", page_icon="🏗️", layout="wide")

# ==========================================
# INJECTION CSS POUR DESIGN PRO
# ==========================================
st.markdown("""
<style>
/* Masquer le menu et le footer de Streamlit pour faire plus "Logiciel" */
#MainMenu {visibility: hidden;}
footer {visibility: hidden;}
header {visibility: hidden;}

/* Design du Header / Logo */
.header-container {
    background-color: #1a2c42;
    padding: 20px;
    border-radius: 10px;
    margin-bottom: 30px;
    display: flex;
    align-items: center;
    border-left: 8px solid #f39c12;
}
.logo-icon {
    font-size: 45px;
    margin-right: 20px;
}
.app-title {
    color: white;
    font-size: 30px;
    font-weight: 800;
    margin: 0;
    font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
}
.app-subtitle {
    color: #bdc3c7;
    font-size: 15px;
    margin: 0;
    margin-top: 5px;
}
</style>
""", unsafe_allow_html=True)

# ==========================================
# 0. API KEY ET BASE DE DONNÉES
# ==========================================
GROQ_TOKEN = "gsk_" + "rL7iv9leEcHqfbdKrnlvWGdyb3FYSrEtimGJWVH6NruYOP3pFqsG"
API_URL = "https://api.groq.com/openai/v1/chat/completions"

BASE_DONNEES = {
    "IPE400": {"desc": "Profilé IPE 400 - Acier S275", "unite": "U", "prix_u": 2500.0},
    "HEA120": {"desc": "Profilé HEA 120 - Acier S275", "unite": "U", "prix_u": 850.0},
    "HEA300": {"desc": "Profilé HEA 300 - Acier S275", "unite": "U", "prix_u": 2100.0},
    "UPN80": {"desc": "Profilé UPN 80", "unite": "U", "prix_u": 400.0},
    "UPN200": {"desc": "Profilé UPN 200", "unite": "U", "prix_u": 1200.0},
    "L70*7": {"desc": "Cornière à ailes égales 70x7", "unite": "U", "prix_u": 150.0},
    "BOULON M16": {"desc": "Boulon d'assemblage M16 HR", "unite": "U", "prix_u": 15.0},
    "PL 300*300*20": {"desc": "Platine d'ancrage 300x300 Ep:20mm", "unite": "U", "prix_u": 350.0},
    "SIKAGROUT": {"desc": "Mortier de scellement Sikagrout", "unite": "Sac", "prix_u": 150.0},
    "POTEAU BETON": {"desc": "Poteau en Béton Armé", "unite": "U", "prix_u": 1200.0},
    "TUBE EN PVC": {"desc": "Tube PVC Évacuation", "unite": "ml", "prix_u": 35.0},
    "BARDAGE": {"desc": "Revêtement / Bardage", "unite": "m²", "prix_u": 95.0},
}

def get_item_info(item_name):
    item_upper = item_name.upper()
    for key in BASE_DONNEES.keys():
        if key in item_upper:
            return BASE_DONNEES[key]
            
    if "BÉTON" in item_upper or "BETON" in item_upper: return {"desc": "Ouvrage en Béton", "unite": "m3", "prix_u": 800.0}
    if "TUBE" in item_upper or "PVC" in item_upper: return {"desc": f"Tube {item_name}", "unite": "ml", "prix_u": 40.0}
    if "TOLE" in item_upper or "TÔLE" in item_upper or "BARDAGE" in item_upper: return {"desc": f"Tôle / Bardage", "unite": "m²", "prix_u": 100.0}
    if "SIKA" in item_upper: return {"desc": "Produit d'étanchéité/scellement", "unite": "Sac", "prix_u": 150.0}
    if "BOULON" in item_upper or "BLS" in item_upper or "TIGE" in item_upper: return {"desc": f"Fixation {item_name}", "unite": "U", "prix_u": 20.0}
    if "IPE" in item_upper or "HEA" in item_upper or "UPN" in item_upper or "HEB" in item_upper: return {"desc": f"Profilé métallique {item_name}", "unite": "U", "prix_u": 1000.0}
    if "PL " in item_upper or "PLATINE" in item_upper or "GOUSSET" in item_upper: return {"desc": f"Platine / Gousset", "unite": "U", "prix_u": 150.0}
    
    return {"desc": f"Élément divers", "unite": "Ens", "prix_u": 250.0}

# ==========================================
# 1. ExtractionEngine (Hybride : Vectoriel + OCR)
# ==========================================
class ExtractionEngine:
    @staticmethod
    def extract_all(doc):
        st.info("📖 **Lecture hybride intelligente :** Le système lit le texte et scanne les images page par page... ⏳")
        full_text = ""
        
        progress_bar = st.progress(0)
        total_pages = len(doc)
        
        for i, page in enumerate(doc):
            # 1. Tentative d'extraction vectorielle
            page_text = page.get_text("text").strip()
            
            # 2. Si le texte est faible ou si la page contient des images (plan mixte), on force l'OCR
            images = page.get_images(full=True)
            if len(page_text) < 1000 or len(images) > 0:
                try:
                    pix = page.get_pixmap(dpi=150) # 150 DPI pour équilibre Vitesse/Qualité
                    img = Image.open(io.BytesIO(pix.tobytes("jpeg")))
                    ocr_text = pytesseract.image_to_string(img, lang="fra").strip()
                    page_text += "\n" + ocr_text
                except Exception:
                    pass
            
            full_text += f"\n--- PAGE {i+1} ---\n" + page_text
            progress_bar.progress((i + 1) / total_pages)
            
        progress_bar.empty()
        return full_text

# ==========================================
# 2. Parser Métier (Intégration Llama 3.1 - JSON PRO)
# ==========================================
class ParserMetier:
    @staticmethod
    def parse_with_ai(text):
        headers = {
            "Authorization": f"Bearer {GROQ_TOKEN}",
            "Content-Type": "application/json"
        }
        # Nettoyage du texte (suppression des espaces multiples) pour économiser des Tokens
        clean_text = re.sub(r'\s+', ' ', text)
        
        prompt = f"""Tu es un expert en BTP et Métré. Analyse le texte suivant extrait d'un plan d'architecture/charpente.
Ta mission est d'extraire 1) Les informations du projet (Cartouche) et 2) TOUS les matériaux.

RÈGLE ABSOLUE : Tu DOIS extraire absolument TOUT ce qui ressemble à un matériau ou élément de construction, même si ce n'est pas standard (ex: Lierne, Lisse, Panne, Poutre, IPE, HEA, Tube, Béton, Acier, Armature, Ø12, Boulon, Platine, Cornière, etc.). Ne laisse RIEN de côté. Si un élément est mentionné plusieurs fois, additionne les quantités.

Tu dois répondre UNIQUEMENT avec un objet JSON valide ayant cette structure exacte :
{{
    "metadata": {{
        "projet": "Nom du projet ou titre du plan (laisse vide si introuvable)",
        "societe": "Nom de l'entreprise, maitre d'ouvrage, client ou bureau d'étude (laisse vide si introuvable)",
        "date_plan": "Date trouvée sur le plan (laisse vide si introuvable)",
        "description": "Un bref résumé (1-2 phrases) de ce que représente ce plan (ex: Construction métallique d'un auvent...). Laisse vide si introuvable."
    }},
    "materiaux": [
        {{"element": "TUBE EN PVC", "infos": "DN125", "unite": "ml", "quantite": 5}},
        {{"element": "IPE 400", "infos": "Long = 200mm", "unite": "U", "quantite": 12}}
    ]
}}

Texte à analyser :
{clean_text[:12000]}
"""
        
        payload = {
            "model": "llama-3.1-8b-instant",
            "messages": [{"role": "user", "content": prompt}],
            "temperature": 0.1
        }
        
        try:
            response = requests.post(API_URL, headers=headers, json=payload)
            if response.status_code == 200:
                result = response.json()['choices'][0]['message']['content']
                match = re.search(r'(\{.*\}|\[.*\])', result, re.DOTALL)
                if match:
                    try:
                        data_json = json.loads(match.group(0))
                    except:
                        data_json = []
                        
                    if isinstance(data_json, list):
                        items = data_json
                        metadata = {"projet": "", "societe": "", "date_plan": "", "description": ""}
                    elif isinstance(data_json, dict):
                        items = data_json.get("materiaux", data_json.get("matériaux", data_json.get("materials", [])))
                        metadata = data_json.get("metadata", {"projet": "", "societe": "", "date_plan": "", "description": ""})
                    else:
                        items = []
                        metadata = {"projet": "", "societe": "", "date_plan": "", "description": ""}
                    
                    merged = {}
                    for item in items:
                        elem = str(item.get("element", "")).strip().upper()
                        infos = str(item.get("infos", "")).strip()
                        unite = str(item.get("unite", "U")).strip()
                        try: qty = float(item.get("quantite", 1))
                        except: qty = 1.0
                        
                        if elem:
                            key = f"{elem}___{infos}___{unite}"
                            if key in merged:
                                merged[key]["quantite"] += qty
                            else:
                                merged[key] = {"element": elem, "infos": infos, "unite": unite, "quantite": qty}
                                
                    st.success("✔️ **Analyse terminée :** Le plan a été traité et les détails ont été structurés avec succès.")
                    return {"metadata": metadata, "materiaux": list(merged.values()), "raw_response": result}
                else:
                    st.warning("⚠️ Impossible de formater les données. (Passage au mode dégradé).")
                    return ParserMetier.parse_regex(text)
            else:
                st.error(f"Erreur Serveur ({response.status_code}): {response.text} - Passage au mode dégradé.")
                return ParserMetier.parse_regex(text)
        except Exception as e:
            st.error(f"Erreur de connexion : {e}. Passage au mode dégradé.")
            return ParserMetier.parse_regex(text)

    @staticmethod
    def parse_regex(text):
        elements = []
        profiles = re.findall(r'\b(IPE|HEA|HEB|UPN)\s*(\d+)\b', text, re.IGNORECASE)
        for p in profiles: elements.append({"element": f"{p[0].upper()}{p[1]}", "infos": "", "unite": "U", "quantite": 1})
            
        cornieres = re.findall(r'\bL\s*(\d+\*\d+)\b', text, re.IGNORECASE)
        for c in cornieres: elements.append({"element": f"Cornière L{c}", "infos": "", "unite": "U", "quantite": 1})
            
        boulons = re.findall(r'(\d+)\s*Bls\s*(M\d+)', text, re.IGNORECASE)
        for b in boulons:
            elements.append({"element": f"Boulon {b[1].upper()}", "infos": "", "unite": "U", "quantite": int(b[0])})
            
        return {"metadata": {"projet": "Inconnu", "societe": "Inconnu", "date_plan": "", "description": "Extraction manuelle Regex."}, "materiaux": elements}

# ==========================================
# 3. Exporter
# ==========================================
class Exporter:
    @staticmethod
    def to_excel(df, total_general, metadata, logo_bytes=None):
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name='Métré_Détaillé', startrow=8)
            worksheet = writer.sheets['Métré_Détaillé']
            
            header_font = Font(bold=True, color="FFFFFF")
            header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
            border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
            
            # En-tête PRO avec Métadonnées
            worksheet['A1'] = "DEVIS ESTIMATIF DÉTAILLÉ (Généré par METRE-PRO)"
            worksheet['A1'].font = Font(bold=True, size=14, color="1F4E78")
            
            worksheet['A3'] = "Projet :"
            worksheet['B3'] = metadata.get('projet', 'Non spécifié')
            worksheet['A4'] = "Société / Client :"
            worksheet['B4'] = metadata.get('societe', 'Non spécifié')
            worksheet['A5'] = "Date du plan :"
            date_plan = metadata.get('date_plan', '')
            worksheet['B5'] = date_plan if date_plan else "Non spécifiée"
            
            worksheet['A6'] = "Généré le :"
            date_export = datetime.datetime.now().strftime("%d/%m/%Y à %H:%M")
            worksheet['B6'] = date_export
            
            worksheet['A3'].font = Font(bold=True, color="1F4E78")
            worksheet['A4'].font = Font(bold=True, color="1F4E78")
            worksheet['A5'].font = Font(bold=True, color="1F4E78")
            worksheet['A6'].font = Font(bold=True, color="1F4E78")
            
            if logo_bytes:
                try:
                    from openpyxl.drawing.image import Image as OpenpyxlImage
                    img = OpenpyxlImage(BytesIO(logo_bytes))
                    # Redimensionner l'image pour qu'elle tienne dans l'en-tête (Hauteur ~80px)
                    ratio = 80 / img.height
                    img.height = 80
                    img.width = int(img.width * ratio)
                    worksheet.add_image(img, 'F2')
                except Exception:
                    pass
            
            col_widths = {'A': 25, 'B': 35, 'C': 30, 'D': 10, 'E': 15, 'F': 18, 'G': 18}
            for col_letter, width in col_widths.items():
                worksheet.column_dimensions[col_letter].width = width
                
            for col_num in range(len(df.columns)):
                cell = worksheet.cell(row=9, column=col_num+1)
                cell.font = header_font
                cell.fill = header_fill
                cell.border = border
                cell.alignment = Alignment(horizontal="center", vertical="center")
            
            for row_idx, row in enumerate(df.values, 10):
                for col_idx, value in enumerate(row, 1):
                    cell = worksheet.cell(row=row_idx, column=col_idx, value=value)
                    cell.border = border
                    if col_idx in [6, 7]: cell.number_format = '#,##0.00'
                        
            total_row = len(df) + 10
            worksheet.merge_cells(start_row=total_row, start_column=1, end_row=total_row, end_column=6)
            worksheet.cell(row=total_row, column=1, value="TOTAL GÉNÉRAL").font = Font(bold=True)
            worksheet.cell(row=total_row, column=1).alignment = Alignment(horizontal="right")
            
            cell_total = worksheet.cell(row=total_row, column=7, value=total_general)
            cell_total.font = Font(bold=True, color="9C0006")
            cell_total.fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
            cell_total.number_format = '#,##0.00'
            cell_total.border = border

        return output.getvalue()

    @staticmethod
    def to_csv(df):
        return df.to_csv(index=False, sep=';').encode('utf-8-sig')


# ==========================================
# INTERFACE LOGICIEL
# ==========================================

# Affichage du Header PRO
st.markdown("""
<div class="header-container">
    <div class="logo-icon">🏗️</div>
    <div>
        <p class="app-title">METRE-PRO SYSTEM</p>
        <p class="app-subtitle">Solution Automatisée d'Extraction de Quantités et Métré BTP</p>
    </div>
</div>
""", unsafe_allow_html=True)

# Initialisation de la mémoire du navigateur (Session State) pour sauvegarder les résultats
if "df" not in st.session_state: st.session_state.df = None
if "total_general" not in st.session_state: st.session_state.total_general = 0.0
if "last_file_name" not in st.session_state: st.session_state.last_file_name = None

if "metadata" not in st.session_state: st.session_state.metadata = {}
if "logo_bytes" not in st.session_state: st.session_state.logo_bytes = None
if "plan_preview" not in st.session_state: st.session_state.plan_preview = None

col1, col2 = st.columns([1, 2])
with col1:
    uploaded_file = st.file_uploader("Étape 1 : Importer le Plan (Format PDF)", type=["pdf"])

if uploaded_file is not None:
    # Si l'utilisateur importe un nouveau fichier, on réinitialise les résultats
    if st.session_state.last_file_name != uploaded_file.name:
        st.session_state.df = None
        st.session_state.total_general = 0.0
        st.session_state.metadata = {}
        st.session_state.logo_bytes = None
        st.session_state.plan_preview = None
        st.session_state.last_file_name = uploaded_file.name
        
    with col1:
        # Ajout d'un bouton pour lancer l'analyse manuellement
        start_btn = st.button("🚀 Étape 2 : Lancer l'Analyse Automatique", type="primary", use_container_width=True)
        
    if start_btn:
        doc = fitz.open(stream=uploaded_file.read(), filetype="pdf")
        
        # Génération de l'aperçu du plan (Première page)
        try:
            pix = doc[0].get_pixmap(dpi=150)
            st.session_state.plan_preview = Image.open(io.BytesIO(pix.tobytes("jpeg")))
        except Exception:
            pass
        
        # Extraction du logo (la plus grande image trouvée dans le document)
        logo_bytes = None
        for page in doc:
            images = page.get_images(full=True)
            if images:
                max_size = 0
                for img in images:
                    try:
                        base_image = doc.extract_image(img[0])
                        w, h = base_image["width"], base_image["height"]
                        if w * h > max_size:
                            max_size = w * h
                            best_img = base_image["image"]
                    except: pass
                if max_size > 5000: # Ignorer les petites icônes
                    logo_bytes = best_img
                    break
        st.session_state.logo_bytes = logo_bytes
        
        # Extraction hybride du texte complet
        text = ExtractionEngine.extract_all(doc)
        
        if text and len(text.strip()) > 10:
            with st.spinner("⏳ Analyse intelligente en cours... Veuillez patienter (Cela peut prendre 10 à 30 secondes)."):
                
                # Exécution silencieuse et automatique de l'analyse intelligente
                resultats_dict = ParserMetier.parse_with_ai(text)
                metadata = resultats_dict.get("metadata", {})
                resultats = resultats_dict.get("materiaux", [])
                raw_response = resultats_dict.get("raw_response", "")
                
                if len(resultats) > 0:
                    grouped_data = {}
                    tot = 0.0
                    
                    for item in resultats:
                        ref = item.get("element", "Inconnu")
                        info_db = get_item_info(ref)
                        
                        unite = item.get("unite", "")
                        if unite == "" or unite == "U": 
                            unite = info_db["unite"]
                            
                        try: qty = float(item.get("quantite", 1))
                        except: qty = 1.0
                        
                        infos_supp = item.get("infos", "")
                        prix_u = info_db["prix_u"]
                        total_ligne = qty * prix_u
                        
                        key = f"{ref}____{infos_supp}____{unite}____{prix_u}"
                        
                        if key in grouped_data:
                            grouped_data[key]["Quantité"] += qty
                            grouped_data[key]["Total Ligne"] += total_ligne
                        else:
                            grouped_data[key] = {
                                "Référence": ref,
                                "Désignation": info_db["desc"],
                                "Infos / Dimensions": infos_supp,
                                "Unité": unite,
                                "Quantité": qty,
                                "Prix Unitaire": prix_u,
                                "Total Ligne": total_ligne
                            }
                        
                        tot += total_ligne
                        
                    data = list(grouped_data.values())
                        
                    # Sauvegarde dans la session (Mémoire)
                    st.session_state.df = pd.DataFrame(data).sort_values(by="Total Ligne", ascending=False)
                    st.session_state.total_general = tot
                    st.session_state.metadata = metadata
                else:
                    st.warning("⚠️ Aucun élément trouvé dans ce plan.")
                    with st.expander("🔍 Mode Débogage (Voir pourquoi l'IA n'a rien trouvé)"):
                        st.write("Ceci arrive souvent si le plan est une image (Scanné) ou s'il ne contient pas de vrai texte.")
                        st.text_area("Texte extrait du PDF (Ce que l'IA a vu) :", text[:2000], height=200)
                        st.text_area("Réponse brute de l'IA :", raw_response, height=200)
        else:
            st.error("❌ Ce PDF est une image complète (Scanné) sans texte vectoriel. L'OCR est obligatoire pour l'analyser.")

    # Affichage des résultats s'ils sont dans la mémoire
    if st.session_state.df is not None:
        df = st.session_state.df
        total_general = st.session_state.total_general
        metadata = st.session_state.metadata
        
        with col2:
            st.metric(label="TOTAL GÉNÉRAL", value=f"{total_general:,.2f} DH")
            
            st.write("---")
            st.write("### 📤 Exporter le Métré")
            file_date = datetime.datetime.now().strftime("%Y-%m-%d_%H-%M")
            
            st.download_button("📊 Télécharger Fichier EXCEL (PRO)", Exporter.to_excel(df, total_general, metadata, st.session_state.logo_bytes), f"METRE_{file_date}.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
            st.download_button("📑 Télécharger Données (CSV)", Exporter.to_csv(df), f"METRE_{file_date}.csv", "text/csv", use_container_width=True)
        
        st.write("### 📋 Étape 3 : Résultat du Métré")
        
        # Affichage des métadonnées
        md_col1, md_col2, md_col3 = st.columns(3)
        md_col1.info(f"**🏢 Projet :** {metadata.get('projet', 'Non spécifié')}")
        md_col2.info(f"**💼 Client/Bureau :** {metadata.get('societe', 'Non spécifié')}")
        
        date_plan = metadata.get('date_plan', '')
        date_export = datetime.datetime.now().strftime("%d/%m/%Y à %H:%M")
        md_col3.info(f"**📅 Date / Heure :** {date_export}")
        
        # Création d'une copie du dataframe pour l'affichage avec la ligne TOTAL
        df_display = df.copy()
        total_row_df = pd.DataFrame([{
            "Référence": "TOTAL GÉNÉRAL", "Désignation": "", "Infos / Dimensions": "", 
            "Unité": "", "Quantité": None, "Prix Unitaire": None, "Total Ligne": total_general
        }])
        df_display = pd.concat([df_display, total_row_df], ignore_index=True)
        
        # Application d'un style spécifique pour la ligne Total
        def highlight_total(s):
            if s.name == len(df_display) - 1: return ['background-color: #f39c12; color: white; font-weight: bold'] * len(s)
            return [''] * len(s)
            
        st.dataframe(df_display.style.apply(highlight_total, axis=1).format({"Prix Unitaire": "{:,.2f}", "Total Ligne": "{:,.2f}"}, na_rep=""), use_container_width=True)
        
        st.write("### 📝 Synthèse du Plan")
        st.info(metadata.get('description', "Aucune description trouvée dans ce plan."))
        
        if st.session_state.plan_preview:
            st.write("### 🖼️ Aperçu du Plan")
            st.image(st.session_state.plan_preview, use_container_width=True)
