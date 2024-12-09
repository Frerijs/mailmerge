# streamlit_app.py

import streamlit as st
import pandas as pd
from docx import Document
import matplotlib.pyplot as plt
import os
import io

def remove_all_paragraphs(doc):
    """
    Noņem visus paragrafus no dokumenta, lai novērstu tukšas lapas.
    """
    for p in doc.paragraphs:
        p._element.getparent().remove(p._element)

def perform_mail_merge_single_doc(template_path, csv_data, output_path):
    """
    Veic mail merge, izmantojot Word šablonu un CSV datus, un saglabā visus rezultātus vienā .docx failā.

    Args:
        template_path (str): Ceļš uz Word šablonu (`template.docx`).
        csv_data (pd.DataFrame): Pandas DataFrame ar CSV datiem.
        output_path (str): Ceļš uz izvadītāja .docx failu.
    
    Returns:
        str: Izvades faila ceļš.
    """
    # Inicializē izvadītāja dokumentu
    output_doc = Document()
    remove_all_paragraphs(output_doc)  # Noņem tukšo paragrafu

    for index, row in csv_data.iterrows():
        # Nolasām šablonu
        doc = Document(template_path)
        
        # Aizvietojam placeholderus ar CSV datiem
        for paragraph in doc.paragraphs:
            for key, value in row.items():
                placeholder = f'{{{{{key}}}}}'
                if placeholder in paragraph.text:
                    paragraph.text = paragraph.text.replace(placeholder, str(value))
        
        # Pievienojam saturu izvadītāja dokumentam
        for element in doc.element.body:
            output_doc.element.body.append(element)
        
        # Pievienojam lappuses pārtraukumu starp ierakstiem, ja nav pēdējais
        if index < len(csv_data) - 1:
            output_doc.add_page_break()
    
    # Saglabājam izvadītāja dokumentu
    output_doc.save(output_path)
    return output_path

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
                if not os.path.exists(output_dir):
                    os.makedirs(output_dir)
                
                output_path = os.path.join(output_dir, "merged_documents.docx")
                perform_mail_merge_single_doc(template_path, data, output_path)
                
                st.success(f"Mail merge veiksmīgi pabeigts! Dokumenti saglabāti failā: {output_path}")
                
                # Parādām lejupielādes saiti
                with open(output_path, "rb") as f:
                    file_bytes = f.read()
                    st.download_button(
                        label="Lejupielādēt Merged Dokumentu",
                        data=file_bytes,
                        file_name="merged_documents.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    )
        except Exception as e:
            st.error(f"Kļūda apstrādājot CSV failu: {e}")

if __name__ == "__main__":
    main()
