# streamlit_app.py

import streamlit as st
import pandas as pd
from docx import Document
import matplotlib.pyplot as plt
import os
import io

def perform_mail_merge(template_path, csv_data, output_dir):
    """
    Veic mail merge, izmantojot Word šablonu un CSV datus.

    Args:
        template_path (str): Ceļš uz Word šablonu (`template.docx`).
        csv_data (pd.DataFrame): Pandas DataFrame ar CSV datiem.
        output_dir (str): Mape, kurā saglabāt izveidotos dokumentus.
    
    Returns:
        list: Saraksts ar izveidoto dokumentu ceļiem.
    """
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
    
    created_files = []
    
    for index, row in csv_data.iterrows():
        doc = Document(template_path)
        
        # Aizvietojam placeholderus ar CSV datiem
        for paragraph in doc.paragraphs:
            for key, value in row.items():
                placeholder = f'{{{{{key}}}}}'
                if placeholder in paragraph.text:
                    paragraph.text = paragraph.text.replace(placeholder, str(value))
        
        # Saglabājam jauno dokumentu
        name = row.get('Name', f'Document_{index}')
        output_path = os.path.join(output_dir, f"{name}_{index}.docx")
        doc.save(output_path)
        created_files.append(output_path)
    
    return created_files

def main():
    st.title("Mail Merge Lietotne")
    
    st.sidebar.header("Iestatījumi")
    
    # Augšupielādējam CSV failu
    uploaded_file = st.file_uploader("Augšupielādējiet CSV failu", type=["csv"])
    
    if uploaded_file is not None:
        try:
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
                created_files = perform_mail_merge(template_path, data, output_dir)
                
                st.success(f"Mail merge veiksmīgi pabeigts! Dokumenti saglabāti mapē: {output_dir}")
                
                # Parādām lejupielādes saites
                for file_path in created_files:
                    file_name = os.path.basename(file_path)
                    with open(file_path, "rb") as f:
                        file_bytes = f.read()
                        st.download_button(
                            label=f"Lejupielādēt {file_name}",
                            data=file_bytes,
                            file_name=file_name,
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                        )
        except Exception as e:
            st.error(f"Kļūda apstrādājot CSV failu: {e}")

if __name__ == "__main__":
    main()
