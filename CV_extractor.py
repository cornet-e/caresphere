import re
import io
import pdfplumber
import pandas as pd
import streamlit as st

def extract_header_info(file_bytes: bytes) -> dict:
    """Extrait les informations d'en-tête de la page 1 du rapport PDF."""
    
    # Ordre voulu
    header_fields = {
        "Model": None,
        "Serial No.": None,
        "Nickname": None,
        "Instrument Code": None,
        "Control Material": None,
        "Lot No. Level 1": None,
        "Lot No. Level 2": None,
        "Lot No. Level 3": None,
        "Report Period": None
    }

    # Liste des champs pour le regex
    field_names = list(header_fields.keys())

    # Regex : capture "Field: valeur" jusqu'au prochain champ connu ou fin de texte
    pattern = re.compile(
        r'(' + '|'.join(re.escape(f) for f in field_names) + r')\s*:\s*(.*?)\s*(?=(?:' + '|'.join(re.escape(f) for f in field_names) + r')\s*:|$)',
        re.DOTALL
    )

    with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
        if not pdf.pages:
            return header_fields

        text = pdf.pages[0].extract_text() or ""

        # Stopper le texte après 'Report Period' si 'Accredited' suit
        if "Accredited" in text:
            text = text.split("Accredited")[0]

        for match in pattern.finditer(text):
            key = match.group(1)
            value = match.group(2).strip().replace("_", " ")  # nettoyer les underscores
            header_fields[key] = value

    # Réordonner explicitement les Lot No.
    ordered_lots = ["Lot No. Level 1", "Lot No. Level 2", "Lot No. Level 3"]
    ordered_lots_values = [header_fields[k] for k in ordered_lots]
    for k, v in zip(ordered_lots, ordered_lots_values):
        header_fields[k] = v

    return header_fields



st.set_page_config(page_title="Extraction CV% PDF", page_icon="📊", layout="wide")

st.title("1️⃣ 📊 Extraction automatique des CV% depuis un PDF (XN-CHECK)")

# --- Téléversement du fichier ---
uploaded_file = st.file_uploader("Choisissez un fichier PDF", type=["pdf"])

# --- Sélection des pages ---
col1, col2 = st.columns(2)
start_page = col1.number_input("Page de début (numérotation humaine)", min_value=1, value=2)
end_page = col2.number_input("Page de fin (numérotation humaine)", min_value=1, value=7)

# --- Regex principaux ---
param_L1_regex = re.compile(r'^([A-Z0-9][A-Z0-9\-\+#/%\.\(\)]{0,30}?)\s+L1\b(.*)$')
level_regex = re.compile(r'^(L[23])\b(.*)$')

def find_cv_in_tokens(tokens, raw_tokens):
    for i in range(len(tokens)-1):
        if ('.' in raw_tokens[i] and re.match(r'^\d+$', raw_tokens[i+1])):
            return float(tokens[i])
        if re.match(r'^\d+(\.\d+)?$', raw_tokens[i]) and re.match(r'^\d+$', raw_tokens[i+1]):
            val = float(tokens[i])
            if val < 50:
                return val
    return None

def extract_cv_from_pdf(file_bytes: bytes, start_page: int, end_page: int) -> pd.DataFrame:
    results = []
    with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
        total_pages = len(pdf.pages)  # ✅ CORRECTION
        lines = []
        for p_idx in range(start_page - 1, min(end_page, total_pages)):
            page = pdf.pages[p_idx]
            text = page.extract_text() or ""
            for li, l in enumerate(text.splitlines()):
                lines.append(l.strip())

    current_param = None
    for idx, line in enumerate(lines):
        # PARAM L1
        m1 = param_L1_regex.match(line)
        if m1:
            current_param = m1.group(1)
            rest = m1.group(2)
            raw_nums = re.findall(r'[-+]?\d+\.\d+|[-+]?\d+', rest)
            nums = [n.replace('+','') for n in raw_nums]
            cv = find_cv_in_tokens(nums, raw_nums)
            if cv is not None:
                results.append({"Parameter": current_param, "Level": "L1", "CV%": cv})
            continue

        # L2/L3
        m2 = level_regex.match(line)
        if m2 and current_param:
            level = m2.group(1)
            rest = m2.group(2)
            raw_nums = re.findall(r'[-+]?\d+\.\d+|[-+]?\d+', rest)
            nums = [n.replace('+','') for n in raw_nums]
            cv = find_cv_in_tokens(nums, raw_nums)
            if cv is not None:
                results.append({"Parameter": current_param, "Level": level, "CV%": cv})
            continue

        # Param sur une ligne + '%' sur la suivante + L1 plus loin
        if re.match(r'^[A-Z0-9\-\+#]{1,20}$', line):
            if idx+1 < len(lines) and lines[idx+1] == '%':
                for j in range(idx+2, min(len(lines), idx+10)):
                    if lines[j].startswith("L1"):
                        current_param = line + '%'
                        raw_nums = re.findall(r'[-+]?\d+\.\d+|[-+]?\d+', lines[j])
                        nums = [n.replace('+','') for n in raw_nums]
                        cv = find_cv_in_tokens(nums, raw_nums)
                        if cv is not None:
                            results.append({"Parameter": current_param, "Level": "L1", "CV%": cv})
                        break

    df = pd.DataFrame(results).drop_duplicates().reset_index(drop=True)
    return df



# --- Traitement ---
if uploaded_file is not None:
    file_bytes = uploaded_file.read()  # Lire une seule fois
    st.info("📑 Traitement en cours...")

    # --- Extraction CV ---
    df_cv = extract_cv_from_pdf(file_bytes, start_page, end_page)

    # --- Extraction header ---
    header_info = extract_header_info(file_bytes)
    df_header = pd.DataFrame(list(header_info.items()), columns=["Field", "Value"])

    if not df_cv.empty:
        st.success(f"{len(df_cv)} enregistrements extraits ✅")
        st.dataframe(df_cv, use_container_width=True)

        # --- Création Excel avec 2 feuilles ---
        xlsx_bytes = io.BytesIO()
        with pd.ExcelWriter(xlsx_bytes, engine="openpyxl") as writer:
            df_cv.to_excel(writer, sheet_name="CV%", index=False)
            df_header.to_excel(writer, sheet_name="Header Info", index=False)
        xlsx_bytes.seek(0)

        # --- Boutons téléchargement ---
        csv_bytes = df_cv.to_csv(index=False, sep=';').encode('utf-8')

        col1, col2 = st.columns(2)
        col1.download_button("💾 Télécharger CSV", csv_bytes, "extraction_cv.csv", "text/csv")
        col2.download_button("📊 Télécharger Excel", xlsx_bytes, "extraction_cv.xlsx", 
                             "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    else:
        st.warning("⚠️ Aucune donnée extraite sur la plage de pages spécifiée.")

    # --- Extraction des métadonnées en-tête ---
    header_info = extract_header_info(file_bytes)
    st.subheader("🧾 Informations du rapport")
    cols = st.columns(1)
    items = list(header_info.items())
    for i, (key, value) in enumerate(items):
        col = cols[i % 1]
        col.markdown(f"**{key}**: {value if value else '—'}")

# --- Étape 2 optionnelle : comparaison avec CV référence ---
    st.subheader("2️⃣ 📊 Comparaison optionnelle avec CV de référence")
    ref_file = st.file_uploader("📁 Choisissez le fichier Excel des CV de référence", type=["xlsx"], key="ref")
    
    if ref_file is not None and not df_cv.empty:
        df_ref = pd.read_excel(ref_file)
        expected_cols = ["Parameter", "Level", "CV%"]
        if not all(col in df_ref.columns for col in expected_cols):
            st.error("Le fichier de référence doit contenir les colonnes : Parameter, Level, CV%")
        else:
            # Merge sur Parameter + Level
            df_merged = df_cv.merge(df_ref, on=["Parameter", "Level"], suffixes=("", "_ref"))
            
            # Nouvelle colonne "Conformité"
            df_merged["Conformité"] = df_merged.apply(lambda row: "Conforme" if row["CV%"] <= row["CV%_ref"] else "Non Conforme", axis=1)

            # Surlignage rouge si Non Conforme
            def highlight_exceed(s):
                if s["Conformité"] == "Non Conforme":
                    return ['background-color: red']*len(s)
                else:
                    return ['']*len(s)

            st.dataframe(df_merged.style.apply(highlight_exceed, axis=1), use_container_width=True)

            # Export Excel comparatif
            xlsx_bytes2 = io.BytesIO()
            with pd.ExcelWriter(xlsx_bytes2, engine="openpyxl") as writer:
                df_merged.to_excel(writer, sheet_name="CV%", index=False)
                df_header.to_excel(writer, sheet_name="Header Info", index=False)
            xlsx_bytes2.seek(0)
            st.download_button("📊 Télécharger Excel comparatif", xlsx_bytes2, "extraction_cv_comparatif.xlsx",
                            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib import colors
import datetime

# --- Bouton génération PDF ---
if uploaded_file is not None and not df_cv.empty:
    st.subheader("📄 Génération d'un rapport PDF")

    if st.button("📝 Générer le rapport PDF"):
        buffer = io.BytesIO()
        doc = SimpleDocTemplate(buffer, pagesize=A4)
        styles = getSampleStyleSheet()
        story = []

        # --- Titre ---
        story.append(Paragraph("<b>Rapport d'extraction des CV%</b>", styles['Title']))
        story.append(Spacer(1, 12))
        story.append(Paragraph(f"Date de génération : {datetime.datetime.now().strftime('%d/%m/%Y %H:%M')}", styles['Normal']))
        story.append(Spacer(1, 12))

        # --- Informations en-tête ---
        story.append(Paragraph("<b>Informations du rapport</b>", styles['Heading2']))
        header_data = [["Champ", "Valeur"]] + [[k, v if v else "—"] for k, v in header_info.items()]
        t = Table(header_data, hAlign='LEFT')
        t.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
        ]))
        story.append(t)
        story.append(Spacer(1, 12))

        # --- Tableau CV extraits ---
        story.append(Paragraph("<b>Tableau des CV extraits</b>", styles['Heading2']))
        cv_data = [df_cv.columns.tolist()] + df_cv.values.tolist()
        t2 = Table(cv_data, hAlign='LEFT')
        t2.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor("#004080")),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
        ]))
        story.append(t2)
        story.append(Spacer(1, 12))

        # --- Tableau comparatif si dispo ---
        if 'df_merged' in locals() and not df_merged.empty:
            story.append(Paragraph("<b>Comparaison avec CV de référence</b>", styles['Heading2']))

            merged_columns = df_merged.columns.tolist()
            merged_rows = df_merged.values.tolist()
            merged_data = [merged_columns] + merged_rows

            t3 = Table(merged_data, hAlign='LEFT')

            # Styles de base
            table_style = TableStyle([
                ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor("#800000")),
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
                ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
            ])

            # Repérage de la colonne "Conformité"
            if "Conformité" in df_merged.columns:
                conf_idx = df_merged.columns.get_loc("Conformité")
                # Appliquer un fond rouge aux lignes Non Conforme
                for i, row in enumerate(merged_rows, start=1):  # +1 car ligne 0 = en-tête
                    if str(row[conf_idx]).strip().lower() == "non conforme":
                        table_style.add('BACKGROUND', (0, i), (-1, i), colors.Color(1, 0.8, 0.8))  # rouge clair

            t3.setStyle(table_style)
            story.append(t3)
            story.append(Spacer(1, 12))


        # --- Construction PDF ---
        doc.build(story)
        buffer.seek(0)

        # --- Téléchargement ---
        st.download_button(
            label="📥 Télécharger le rapport PDF",
            data=buffer,
            file_name="rapport_cv.pdf",
            mime="application/pdf"
        )
