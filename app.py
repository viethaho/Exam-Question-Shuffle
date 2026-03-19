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
def shuffle_to_excel(input_template, output_excel, num_versions=3):
    # 1. Load the original teacher-filled data
    df_original = pd.read_excel(input_template, sheet_name="Exam Data")
    
    # --- CHANGE 1: Detect the maximum width of the exam ---
    # This finds all 'Option' columns (A, B, C, D, E, etc.)
    opt_cols = sorted([c for c in df_original.columns if c.startswith("Option")])
    max_opt_count = len(opt_cols)
    
    # We will use this to track answers for the Master Key
    master_key_data = []

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        for v in range(1, num_versions + 1):
            # Convert to list of dicts for shuffling
            rows = df_original.to_dict(orient='records')
            
            # We want to keep track of original question IDs for the master key
            # If your Excel doesn't have a unique ID, we use the original row index
            for idx, row in enumerate(rows):
                row['_original_id'] = idx + 1
            
            random.shuffle(rows)
            
            version_data = []
            version_answers = {} # Stores {Original_Q_Num: New_Letter}

            for i, q in enumerate(rows, 1):
                current_opts = []
                correct_txt = None
                ans_key = str(q.get('Correct Answer', '')).strip().upper()

                # Collect valid option text
                for col in opt_cols:
                    val = str(q[col]).strip()
                    if val and val != "nan" and val != "":
                        current_opts.append(val)
                        if col.endswith(ans_key): 
                            correct_txt = val

                # Shuffle options if the teacher allowed it
                should_shuffle = str(q.get('Shuffle? (Yes/No)', 'Yes')).strip().lower() == 'yes'
                if should_shuffle:
                    random.shuffle(current_opts)

                # --- CHANGE 2: Build the row with 'Elastic' padding ---
                new_row = {"No.": i, "Question Text": q['Question Text']}
                
                # We always loop through the MAX count to keep columns aligned
                for index in range(max_opt_count):
                    letter = chr(65 + index)
                    if index < len(current_opts):
                        new_row[f"Option {letter}"] = current_opts[index]
                    else:
                        new_row[f"Option {letter}"] = "" # Blank spacer

                # Re-map the correct letter
                new_key = ""
                if correct_txt in current_opts:
                    new_key = chr(65 + current_opts.index(correct_txt))
                
                new_row["Correct Key"] = new_key
                version_data.append(new_row)
                
                # Track for Master Key
                version_answers[q['_original_id']] = new_key

            # --- CHANGE 3: Standardize Column Order before saving ---
            standard_columns = ["No.", "Question Text"] + opt_cols + ["Correct Key"]
            df_v = pd.DataFrame(version_data)
            df_v = df_v.reindex(columns=standard_columns).fillna("")
            
            sheet_name = f"Version {v}"
            df_v.to_excel(writer, sheet_name=sheet_name, index=False)
            
            # Formatting the sheet
            worksheet = writer.sheets[sheet_name]
            worksheet.set_column('B:B', 55) # Wide Question Text
            worksheet.set_column('C:H', 22) # Option columns
            worksheet.set_column('I:I', 12) # Correct Key column
            
            # Add this version's answers to the master list
            version_answers['Version'] = f"Version {v}"
            master_key_data.append(version_answers)

        # --- STEP 4: Create the Master Key Summary Sheet ---
        df_master = pd.DataFrame(master_key_data)
        # Move 'Version' column to the front
        cols = ['Version'] + [c for c in df_master.columns if c != 'Version']
        df_master = df_master[cols]
        
        df_master.to_excel(writer, sheet_name="MASTER KEY", index=False)
        
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
