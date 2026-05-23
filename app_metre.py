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

def calculate_weights(designation, longueur_mm):
    desig_upper = str(designation).upper().replace(" ", "")
    
    poids_kg_ml = None
    poids_unit = 0.0
    is_linear = True
    
    # Base de données simple des poids linéiques (kg/m)
    db = {
        "IPE120": 10.40, "IPE140": 12.90, "IPE180": 18.80, "IPE200": 22.40, "IPE240": 30.70, 
        "IPE270": 36.10, "IPE300": 42.20, "IPE400": 66.30, "IPE450": 77.60, "IPE500": 90.70,
        "HEA120": 19.90, "HEA140": 24.70, "HEA180": 35.50, "HEA200": 42.30, "HEA240": 60.30,
        "HEA300": 88.30, "HEA400": 125.0,
        "L50*5": 3.77, "L60*6": 5.42, "L70*7": 7.38, "L80*8": 9.66,
        "UPN80": 8.70, "UPN100": 10.60, "UPN120": 13.40, "UPN200": 25.30,
        "D14": 1.20, "D16": 1.58, "D20": 2.47, "D24": 3.55,
        "TUBEC40*40*2": 2.31, "TUBE-C40*40*2": 2.31,
    }
    
    # 1. Chercher dans la BDD
    for key, val in db.items():
        if key in desig_upper:
            poids_kg_ml = val
            break
            
    # 2. Vérifier si c'est une platine / tôle (TN, TH, PL, PLATINE) -> ex: TN300*300*20
    plate_match = re.search(r'(?:TN|PL|PLATINE|TH|EP|TÔLE|TOLE).*?(\d+)(?:\*|X)(\d+)(?:\*|X)(\d+)', desig_upper)
    if plate_match:
        is_linear = False
        l = float(plate_match.group(1))
        w = float(plate_match.group(2))
        t = float(plate_match.group(3))
        # Densité acier = 8000 kg/m3 (comme dans le tableau modèle) ou 7850
        poids_unit = (l / 1000.0) * (w / 1000.0) * (t / 1000.0) * 8000.0
        poids_kg_ml = None
    
    # 3. Calcul unitaire si linéaire
    if is_linear and poids_kg_ml is not None:
        try:
            long_val = float(longueur_mm)
            if long_val > 0:
                poids_unit = poids_kg_ml * (long_val / 1000.0)
            else:
                poids_unit = poids_kg_ml # si pas de longueur on suppose 1 mètre
        except:
            poids_unit = poids_kg_ml
            
    return poids_kg_ml, round(poids_unit, 2)


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
            page_text = page.get_text("text").strip()
            images = page.get_images(full=True)
            if len(page_text) < 1000 or len(images) > 0:
                try:
                    pix = page.get_pixmap(dpi=150)
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
        clean_text = re.sub(r'\s+', ' ', text)
        
        prompt = f"""Tu es un expert en BTP et Métré de Charpente Métallique. Analyse le texte suivant extrait d'un plan.
Ta mission est d'extraire les éléments structuraux pour créer un tableau de nomenclature exact.

RÈGLE ABSOLUE : Tu DOIS extraire tous les éléments (Poteaux, Potelets, Traverses, Pannes, Liernes, Bracons, Platines, Goussets, Tiges, etc.).
Structure chaque ligne avec : pos (numéro de position si présent), nomenclature (ex: POTEAU, PANNE), designation (profilé ou plaque, ex: IPE400, TN2000*1000*20), quantite (nombre), longueur_mm (longueur en mm, null si non applicable).

Tu dois répondre UNIQUEMENT avec un objet JSON valide ayant cette structure exacte :
{{
    "metadata": {{
        "projet": "Nom du projet",
        "societe": "Nom du client",
        "date_plan": "Date"
    }},
    "materiaux": [
        {{"pos": "1", "nomenclature": "POTEAU", "designation": "IPE400", "quantite": 14, "longueur_mm": 4000}},
        {{"pos": "57", "nomenclature": "PLATINE PIED DE POTEAU", "designation": "TN300*300*20", "quantite": 14, "longueur_mm": null}}
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
                    try: data_json = json.loads(match.group(0))
                    except: data_json = {}
                        
                    items = data_json.get("materiaux", [])
                    metadata = data_json.get("metadata", {})
                    
                    st.success("✔️ **Analyse terminée :** La nomenclature a été structurée avec succès.")
                    return {"metadata": metadata, "materiaux": items, "raw_response": result}
                else:
                    st.warning("⚠️ Impossible de formater les données JSON.")
                    return {"metadata": {}, "materiaux": []}
            else:
                st.error(f"Erreur Serveur ({response.status_code}).")
                return {"metadata": {}, "materiaux": []}
        except Exception as e:
            st.error(f"Erreur de connexion : {e}.")
            return {"metadata": {}, "materiaux": []}

# ==========================================
# 3. Exporter
# ==========================================
class Exporter:
    @staticmethod
    def to_excel(df, total_brut, total_net, metadata, logo_bytes=None):
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name='Nomenclature', startrow=5)
            worksheet = writer.sheets['Nomenclature']
            
            # Styles
            header_font = Font(bold=True, color="000000")
            header_fill = PatternFill(start_color="33CCFF", end_color="33CCFF", fill_type="solid")
            border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
            
            # En-tête
            worksheet['B2'] = "Client:"
            worksheet['C2'] = metadata.get('societe', 'Non spécifié')
            worksheet['B3'] = "Projet:"
            worksheet['C3'] = metadata.get('projet', 'Non spécifié')
            worksheet['E2'] = "Date:"
            worksheet['F2'] = metadata.get('date_plan', datetime.datetime.now().strftime("%d/%m/%Y"))
            
            for row in range(2, 4):
                for col in ['B', 'C', 'E', 'F']:
                    worksheet[f"{col}{row}"].font = Font(bold=True)
            
            worksheet.merge_cells('B4:H4')
            worksheet['B4'] = "2 - Ossature métallique :"
            worksheet['B4'].font = Font(bold=True, italic=True)
            worksheet['B4'].fill = PatternFill(start_color="00FFFF", end_color="00FFFF", fill_type="solid")
            worksheet['B4'].alignment = Alignment(horizontal="center")
            worksheet['B4'].border = border
            
            # Colonnes
            col_widths = {'A': 5, 'B': 30, 'C': 25, 'D': 10, 'E': 15, 'F': 15, 'G': 15, 'H': 15}
            for col_letter, width in col_widths.items():
                worksheet.column_dimensions[col_letter].width = width
                
            for col_num in range(len(df.columns)):
                cell = worksheet.cell(row=6, column=col_num+1)
                cell.font = header_font
                cell.fill = header_fill
                cell.border = border
                cell.alignment = Alignment(horizontal="center", vertical="center")
            
            # Lignes de données
            for row_idx, row in enumerate(df.values, 7):
                for col_idx, value in enumerate(row, 1):
                    cell = worksheet.cell(row=row_idx, column=col_idx, value=value)
                    cell.border = border
                    cell.alignment = Alignment(horizontal="center")
                    if col_idx in [6, 7, 8] and isinstance(value, (int, float)): 
                        cell.number_format = '#,##0.00'
                    if col_idx == 4: cell.font = Font(color="00B050") # Vert pour quantité
                    if col_idx == 5: cell.fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid") # Jaune pour longueur
                        
            # Ligne Boulonnerie
            row_boulonnerie = len(df) + 7
            worksheet.merge_cells(start_row=row_boulonnerie, start_column=2, end_row=row_boulonnerie, end_column=3)
            worksheet.cell(row=row_boulonnerie, column=2, value="BOULONNERIE + SOUDAGE").alignment = Alignment(horizontal="center")
            worksheet.cell(row=row_boulonnerie, column=4, value="5%")
            boulonnerie_val = worksheet.cell(row=row_boulonnerie, column=8, value=total_net - total_brut)
            boulonnerie_val.number_format = '#,##0.00'
            
            for c in range(1, 9): worksheet.cell(row=row_boulonnerie, column=c).border = border
            
            # Ligne Total Net
            row_total = row_boulonnerie + 2
            worksheet.merge_cells(start_row=row_total, start_column=6, end_row=row_total, end_column=7)
            worksheet.cell(row=row_total, column=6, value="Poids Tot Net en Kg").font = Font(bold=True)
            worksheet.cell(row=row_total, column=6).fill = PatternFill(start_color="8EA9DB", end_color="8EA9DB", fill_type="solid")
            worksheet.cell(row=row_total, column=6).border = border
            worksheet.cell(row=row_total, column=7).border = border
            
            cell_total = worksheet.cell(row=row_total, column=8, value=total_net)
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

st.markdown("""
<div class="header-container">
    <div class="logo-icon">🏗️</div>
    <div>
        <p class="app-title">METRE-PRO SYSTEM</p>
        <p class="app-subtitle">Génération Automatique de Nomenclatures - Ossature Métallique</p>
    </div>
</div>
""", unsafe_allow_html=True)

if "df" not in st.session_state: st.session_state.df = None
if "total_brut" not in st.session_state: st.session_state.total_brut = 0.0
if "total_net" not in st.session_state: st.session_state.total_net = 0.0
if "metadata" not in st.session_state: st.session_state.metadata = {}
if "last_file_name" not in st.session_state: st.session_state.last_file_name = None

col1, col2 = st.columns([1, 2])
with col1:
    uploaded_file = st.file_uploader("Étape 1 : Importer le Plan (Format PDF)", type=["pdf"])

if uploaded_file is not None:
    if st.session_state.last_file_name != uploaded_file.name:
        st.session_state.df = None
        st.session_state.last_file_name = uploaded_file.name
        
    with col1:
        start_btn = st.button("🚀 Étape 2 : Extraire la Nomenclature", type="primary", use_container_width=True)
        
    if start_btn:
        doc = fitz.open(stream=uploaded_file.read(), filetype="pdf")
        text = ExtractionEngine.extract_all(doc)
        
        if text and len(text.strip()) > 10:
            with st.spinner("⏳ Analyse intelligente en cours..."):
                resultats_dict = ParserMetier.parse_with_ai(text)
                metadata = resultats_dict.get("metadata", {})
                materiaux = resultats_dict.get("materiaux", [])
                
                if len(materiaux) > 0:
                    data = []
                    for item in materiaux:
                        pos = item.get("pos", "")
                        nom = item.get("nomenclature", "")
                        desig = item.get("designation", "")
                        try: qty = float(item.get("quantite", 1))
                        except: qty = 1.0
                        try: long_mm = float(item.get("longueur_mm", 0))
                        except: long_mm = 0.0
                        
                        poids_ml, poids_u = calculate_weights(desig, long_mm)
                        poids_tot = round(qty * poids_u, 2)
                        
                        data.append({
                            "Pos": pos,
                            "Nomenclature": nom,
                            "Désignation": desig,
                            "Quantité": qty,
                            "Long (mm)": long_mm if long_mm > 0 else "----",
                            "Poids (kg/ml)": poids_ml if poids_ml is not None else "----",
                            "Poids kg/Unit": poids_u,
                            "Poids Tot Kg": poids_tot
                        })
                        
                    df = pd.DataFrame(data)
                    
                    # Totaux
                    total_brut = df["Poids Tot Kg"].sum()
                    total_net = total_brut * 1.05  # + 5% Boulonnerie et Soudage
                    
                    st.session_state.df = df
                    st.session_state.total_brut = total_brut
                    st.session_state.total_net = total_net
                    st.session_state.metadata = metadata
                else:
                    st.warning("⚠️ Aucun élément trouvé.")
        else:
            st.error("❌ PDF illisible ou scanné sans OCR.")

    if st.session_state.df is not None:
        df = st.session_state.df
        total_brut = st.session_state.total_brut
        total_net = st.session_state.total_net
        metadata = st.session_state.metadata
        
        with col2:
            st.metric(label="POIDS TOTAL NET (Avec 5% Assemblages)", value=f"{total_net:,.2f} Kg")
            
            st.write("---")
            st.write("### 📤 Exporter la Nomenclature")
            file_date = datetime.datetime.now().strftime("%Y-%m-%d")
            
            st.download_button("📊 Télécharger Fichier EXCEL", Exporter.to_excel(df, total_brut, total_net, metadata), f"Nomenclature_{file_date}.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
            st.download_button("📑 Télécharger Données (CSV)", Exporter.to_csv(df), f"Nomenclature_{file_date}.csv", "text/csv", use_container_width=True)
            
        st.write("### 📋 Étape 3 : Résultat de la Nomenclature")
        
        # Tableau
        df_display = df.copy()
        
        # Format d'affichage pour Streamlit
        st.dataframe(
            df_display.style.format({
                "Quantité": "{:,.0f}",
                "Poids kg/Unit": "{:,.2f}", 
                "Poids Tot Kg": "{:,.2f}"
            }), 
            use_container_width=True,
            height=500
        )
        
        st.info(f"**Boulonnerie + Soudage (5%)** : {(total_net - total_brut):,.2f} Kg")
        st.success(f"**Poids Total Net** : {total_net:,.2f} Kg")
