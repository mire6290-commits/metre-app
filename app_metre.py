import streamlit as st
import fitz  # PyMuPDF
import re
import pandas as pd
import json
import requests
from io import BytesIO
import io
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
import pytesseract
from PIL import Image

st.set_page_config(page_title="Architecture IA Métré", page_icon="🏗️", layout="wide")

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
# 1. Classifier
# ==========================================
class PDFClassifier:
    @staticmethod
    def classify(doc):
        text_length = sum([len(page.get_text("text").strip()) for page in doc])
        return "VECTORIEL" if text_length > 50 else "SCANNE"

# ==========================================
# 2. ExtractionEngine
# ==========================================
class ExtractionEngine:
    @staticmethod
    def extract_vectoriel(doc):
        return "\n".join([page.get_text("text") for page in doc])
    
    @staticmethod
    def extract_scanne(doc):
        st.info("📷 **Lecture des images (OCR) en cours :** Ce plan est scanné. L'application convertit les images en textes... ⏳")
        text = ""
        for page in doc:
            pix = page.get_pixmap(dpi=200)
            img = Image.open(io.BytesIO(pix.tobytes("jpeg")))
            try:
                text += pytesseract.image_to_string(img, lang="fra") + "\n"
            except Exception as e:
                st.error("⚠️ Moteur Tesseract OCR introuvable sur le serveur. Veuillez vérifier packages.txt.")
                return ""
        return text

# ==========================================
# 3. Parser Métier (Intégration Llama 3 - JSON PRO)
# ==========================================
class ParserMetier:
    @staticmethod
    def parse_with_ai(text):
        headers = {
            "Authorization": f"Bearer {GROQ_TOKEN}",
            "Content-Type": "application/json"
        }
        
        prompt = f"""Tu es un expert en BTP et Métré. Analyse le texte suivant extrait d'un plan d'architecture/charpente.
Ta mission est d'extraire TOUS les matériaux, profilés, bétons, accessoires, et leurs détails.

Tu dois répondre UNIQUEMENT avec un tableau JSON valide. Chaque objet doit avoir :
- "element": Le nom du matériau (ex: IPE 400, SIKAGROUT, TUBE PVC)
- "infos": Les détails supplémentaires trouvés comme les dimensions, épaisseur, classe (ex: "Ep:08mm", "DN125", "CL6.8", "L=200mm"). Laisse vide si introuvable.
- "unite": L'unité logique de mesure (ex: "U", "ml", "m²", "kg", "Ens").
- "quantite": La quantité trouvée (nombre entier ou décimal). S'il n'y a pas de quantité claire, mets 1.

Exemple :
[
    {{"element": "TUBE EN PVC", "infos": "DN125", "unite": "ml", "quantite": 5}},
    {{"element": "IPE 400", "infos": "Long = 200mm", "unite": "U", "quantite": 12}}
]

Texte à analyser :
{text[:5000]}
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
                match = re.search(r'\[\s*\{.*?\}\s*\]', result, re.DOTALL)
                if match:
                    items = json.loads(match.group(0))
                    
                    # Regrouper les éléments identiques pour additionner leurs quantités
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
                                
                    st.success("✨ **Succès :** L'IA a lu le PDF et a structuré tous les détails (Dimensions, Quantités) comme un vrai Métreur ! 🧠")
                    return list(merged.values())
                else:
                    st.warning("⚠️ L'IA n'a pas pu formater le JSON correctement. (Passage au Regex).")
                    return ParserMetier.parse_regex(text)
            else:
                st.error(f"Erreur API ({response.status_code}): {response.text} - Passage au Regex.")
                return ParserMetier.parse_regex(text)
        except Exception as e:
            st.error(f"Erreur de connexion : {e}. Passage au Regex.")
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
            
        return elements

# ==========================================
# 4. Exporter
# ==========================================
class Exporter:
    @staticmethod
    def to_excel(df, total_general):
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name='Métré_Détaillé', startrow=2)
            worksheet = writer.sheets['Métré_Détaillé']
            
            header_font = Font(bold=True, color="FFFFFF")
            header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
            border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
            
            worksheet['A1'] = "DEVIS ESTIMATIF DÉTAILLÉ (Généré par IA)"
            worksheet['A1'].font = Font(bold=True, size=14, color="1F4E78")
            
            # Ajustement des largeurs des colonnes PRO
            col_widths = {'A': 20, 'B': 35, 'C': 30, 'D': 10, 'E': 15, 'F': 18, 'G': 18}
            for col_letter, width in col_widths.items():
                worksheet.column_dimensions[col_letter].width = width
                
            for col_num in range(len(df.columns)):
                cell = worksheet.cell(row=3, column=col_num+1)
                cell.font = header_font
                cell.fill = header_fill
                cell.border = border
                cell.alignment = Alignment(horizontal="center", vertical="center")
            
            for row_idx, row in enumerate(df.values, 4):
                for col_idx, value in enumerate(row, 1):
                    cell = worksheet.cell(row=row_idx, column=col_idx, value=value)
                    cell.border = border
                    if col_idx in [6, 7]: cell.number_format = '#,##0.00'
                        
            total_row = len(df) + 4
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
# INTERFACE STREAMLIT
# ==========================================
st.title("🧠 Pipeline IA Avancé : Lecture Totale du Plan (PRO)")

use_ai = st.toggle("Activer l'Intelligence Artificielle (Groq Llama 3) pour lire tout le plan", value=True)

col1, col2 = st.columns([1, 2])
with col1:
    uploaded_file = st.file_uploader("📥 Importer le Plan (PDF Input)", type=["pdf"])

if uploaded_file is not None:
    doc = fitz.open(stream=uploaded_file.read(), filetype="pdf")
    pdf_type = PDFClassifier.classify(doc)
    
    if pdf_type == "VECTORIEL":
        text = ExtractionEngine.extract_vectoriel(doc)
    else:
        text = ExtractionEngine.extract_scanne(doc)
    
    if text and len(text.strip()) > 10:
        with st.spinner("🤖 L'IA est en train de lire le plan ligne par ligne..."):
            
            if use_ai:
                resultats = ParserMetier.parse_with_ai(text)
            else:
                resultats = ParserMetier.parse_regex(text)
            
            if len(resultats) > 0:
                data = []
                total_general = 0.0
                
                for item in resultats:
                    ref = item["element"]
                    info_db = get_item_info(ref)
                    
                    unite = item.get("unite")
                    if unite == "" or unite == "U": 
                        unite = info_db["unite"] # Utilise la DB si l'IA n'a pas trouvé mieux que "U"
                        
                    qty = item["quantite"]
                    infos_supp = item.get("infos", "")
                    
                    prix_u = info_db["prix_u"]
                    total_ligne = qty * prix_u
                    total_general += total_ligne
                    
                    data.append({
                        "Référence": ref,
                        "Désignation": info_db["desc"],
                        "Infos / Dimensions": infos_supp,
                        "Unité": unite,
                        "Quantité": qty,
                        "Prix Unitaire": prix_u,
                        "Total Ligne": total_ligne
                    })
                    
                df = pd.DataFrame(data).sort_values(by="Total Ligne", ascending=False)
                
                with col2:
                    st.metric(label="TOTAL GÉNÉRAL", value=f"{total_general:,.2f} DH")
                
                st.write("### ⚙️ Résultat de l'Extraction (Matériaux découverts)")
                st.dataframe(df.style.format({"Prix Unitaire": "{:,.2f}", "Total Ligne": "{:,.2f}"}), use_container_width=True)
                
                st.write("### 📤 Exports (Outputs)")
                exp_col1, exp_col2 = st.columns(2)
                
                with exp_col1:
                    st.download_button("📊 Télécharger EXCEL (PRO)", Exporter.to_excel(df, total_general), "Metre_PRO.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
                with exp_col2:
                    st.download_button("📑 Télécharger CSV (ERP)", Exporter.to_csv(df), "ERP_Import.csv", "text/csv", use_container_width=True)
            else:
                st.warning("⚠️ Aucun élément trouvé dans ce plan.")
    else:
        st.error("❌ Ce PDF est une image complète (Scanné) sans texte vectoriel. L'OCR est obligatoire pour l'analyser.")
