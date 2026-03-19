# -*- coding: utf-8 -*-
"""
Created on Thu Mar 19 11:49:20 2026

@author: Ha Ho
"""

import streamlit as st
import pandas as pd
import io
import random
import re
from docx2python import docx2python

# --- STEP 1: EXTRACTION LOGIC ---
def extract_docx_to_df(docx_file):
    with docx2python(docx_file) as docx_content:
        all_text = docx_content.text
        lines = [line.strip() for line in all_text.split('\n') if line.strip()]

    data = []
    q_pattern = re.compile(r'^(\d+[\.\)]|Question\s+\d+)', re.IGNORECASE)
    opt_pattern = re.compile(r'^([A-Z])[\.\)]', re.IGNORECASE)
    current_row = {"Question Text": None}
    
    for text in lines:
        if q_pattern.match(text):
            if current_row.get("Question Text"):
                data.append(current_row.copy())
                current_row = {"Question Text": None}
            current_row["Question Text"] = re.sub(q_pattern, '', text).strip()
        elif opt_pattern.match(text):
            match = opt_pattern.match(text)
            letter = match.group(1).upper()
            current_row[f"Option {letter}"] = re.sub(opt_pattern, '', text).strip()
        elif current_row.get("Question Text") and not any(key.startswith("Option") for key in current_row):
             current_row["Question Text"] += " " + text

    if current_row.get("Question Text"):
        data.append(current_row)
    
    df = pd.DataFrame(data)
    
    # Smart Shuffle Logic
    def analyze(row):
        opts = [str(row[c]) for c in df.columns if c.startswith("Option") and pd.notna(row[c]) and str(row[c]).strip() != ""]
        all_opts_text = " ".join(opts).lower()
        if len(opts) <= 2: return "No", "True/False Detected"
        fixed = ["above", "below", "both", "all of", "none of", "neither"]
        if any(w in all_opts_text for w in fixed) or re.search(r'\b[A-F]\s*(&|and|or)\s*[A-F]\b', all_opts_text, re.IGNORECASE):
            return "No", "Fixed keywords detected"
        return "Yes", ""

    analysis = df.apply(analyze, axis=1)
    df["Shuffle? (Yes/No)"] = [res[0] for res in analysis]
    df["Teacher Notes"] = [res[1] for res in analysis]
    df["Correct Answer"] = ""
    
    cols = ["Question Text"] + sorted([c for c in df.columns if c.startswith("Option")]) + ["Correct Answer", "Shuffle? (Yes/No)", "Teacher Notes"]
    return df.reindex(columns=cols).fillna("")

# --- STEP 2: SHUFFLE LOGIC ---
def shuffle_to_excel(input_df, num_versions):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        for v in range(1, num_versions + 1):
            rows = input_df.to_dict(orient='records')
            random.shuffle(rows)
            version_data = []
            for i, q in enumerate(rows, 1):
                opt_cols = sorted([c for c in input_df.columns if c.startswith("Option")])
                current_opts = []
                correct_txt = None
                ans_key = str(q.get('Correct Answer', '')).strip().upper()

                for col in opt_cols:
                    val = str(q[col]).strip()
                    if val and val != "nan" and val != "":
                        current_opts.append(val)
                        if col.endswith(ans_key): correct_txt = val

                if str(q.get('Shuffle? (Yes/No)', 'Yes')).strip().lower() == 'yes':
                    random.shuffle(current_opts)

                new_row = {"No.": i, "Question Text": q['Question Text']}
                for j in range(len(current_opts)):
                    new_row[f"Option {chr(65+j)}"] = current_opts[j]
                
                new_key = chr(65 + current_opts.index(correct_txt)) if correct_txt in current_opts else ""
                new_row["Correct Key"] = new_key
                version_data.append(new_row)

            df_v = pd.DataFrame(version_data).fillna("")
            df_v.to_excel(writer, sheet_name=f"Version {v}", index=False)
    return output.getvalue()

# --- STREAMLIT UI ---
st.set_page_config(page_title="OB Exam Automator", layout="wide")
st.title("📝 OB Exam Automator")
st.markdown("### Securely extract, review, and shuffle exam versions.")

col1, col2 = st.columns(2)

with col1:
    st.header("Step 1: Extract")
    uploaded_docx = st.file_uploader("Upload Original Exam (.docx)", type="docx")
    if uploaded_docx:
        df_extracted = extract_docx_to_df(uploaded_docx)
        st.success(f"Extracted {len(df_extracted)} questions!")
        
        towrite = io.BytesIO()
        df_extracted.to_excel(towrite, index=False, engine='xlsxwriter')
        st.download_button("📥 Download Excel Template", data=towrite.getvalue(), file_name="Exam_Template_to_Fill.xlsx")

with col2:
    st.header("Step 2: Shuffle")
    uploaded_excel = st.file_uploader("Upload FILLED Excel Template", type="xlsx")
    num_v = st.slider("Number of versions", 1, 10, 3)
    
    if uploaded_excel:
        if st.button("🚀 Generate Shuffled Versions"):
            df_to_shuffle = pd.read_excel(uploaded_excel)
            final_file = shuffle_to_excel(df_to_shuffle, num_v)
            st.success(f"Generated {num_v} versions!")
            st.download_button("📥 Download Final Shuffled Exam", data=final_file, file_name="Shuffled_Exam_Versions.xlsx")

st.divider()
st.info("🔒 **Privacy Note:** No files are stored on this server. All processing happens in temporary memory.")