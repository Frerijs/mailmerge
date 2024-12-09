# streamlit_app.py

import streamlit as st
import pandas as pd
from mail_merge import perform_mail_merge
import matplotlib.pyplot as plt
import os

st.title("Mail Merge Lietotne")

st.sidebar.header("Iestatījumi")

# Augšupielādējam CSV failu
uploaded_file = st.file_uploader("Augšupielādējiet CSV failu", type=["csv"])

if uploaded_file is not None:
    data = pd.read_csv(uploaded_file)
    st.write("CSV Saturs:")
    st.dataframe(data)
    
    # Parādām dažas vizualizācijas
    st.header("Datu Vizualizācijas")
    
    numeric_columns = data.select_dtypes(include=['int64', 'float64']).columns
    if not numeric_columns.empty:
        selected_column = st.selectbox("Izvēlieties kolonnas vizualizācijai", numeric_columns)
        fig, ax = plt.subplots()
        data[selected_column].hist(ax=ax)
        st.pyplot(fig)
    else:
        st.write("Nav pieejamu skaitlisku kolonnu vizualizācijai.")
    
    # Veicam mail merge
    if st.button("Veikt Mail Merge"):
        template_path = "template.docx"  # Pārliecinieties, ka template.docx ir pieejams
        output_dir = "output_documents"
        perform_mail_merge(template_path, uploaded_file, output_dir)
        st.success(f"Mail merge veiksmīgi pabeigts! Dokumenti saglabāti mapē: {output_dir}")
        
        # Parādām lejupielādes saites
        for file in os.listdir(output_dir):
            file_path = os.path.join(output_dir, file)
            with open(file_path, "rb") as f:
                st.download_button(label=f"Lejupielādēt {file}", data=f, file_name=file, mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
