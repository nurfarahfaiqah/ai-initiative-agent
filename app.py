import streamlit as st
import pandas as pd
import duckdb
import numpy as np
import json

st.title("AI Initiative Discovery Agent")

st.write("Upload datasets to analyze customer pain points and generate initiative ideas.")

uploaded_files = st.file_uploader(
    "Upload datasets",
    accept_multiple_files=True
)

datasets = {}

if uploaded_files:

    for file in uploaded_files:

        if file.name.endswith(".csv"):
            df = pd.read_csv(file)

        elif file.name.endswith(".xlsx"):
            df = pd.read_excel(file)

        else:
            continue

        datasets[file.name] = df

    st.success(f"{len(datasets)} datasets loaded")

    for name, df in datasets.items():
        st.subheader(name)
        st.dataframe(df.head())

if st.button("Run Analysis"):

    profiles = {}

    for name, df in datasets.items():

        profiles[name] = {
            "rows": df.shape[0],
            "columns": df.shape[1],
            "column_names": df.columns.tolist()
        }

    st.subheader("Dataset Profiles")
    st.json(profiles)

    st.subheader("Generated Prompt for AI")

    prompt = f"""
You are a senior strategy consultant.

Analyze the datasets and identify:
- Problem statement
- Key insights
- Initiative opportunities
- KPI recommendations
- 30-60-90 day plan

Dataset profiles:

{json.dumps(profiles, indent=2)}
"""

    st.code(prompt)

    st.info("Copy the prompt above into Google AI Studio (Gemini 2.5 Pro) to generate the executive insights.")