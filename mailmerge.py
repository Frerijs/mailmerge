import streamlit as st
import pandas as pd
from docx import Document
import matplotlib.pyplot as plt
import os
import io
import csv
import re

def clear_document(doc):
    """
    Noņem visus elementus no dokumenta, lai sagatavotu to satura pievienošanai.
    """
    for element in doc.element.body[:]:
        doc.element.body.remove(element)

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
    # Inicializē izvadītāja dokumentu un noņem visus saturus
    output_doc = Document()
    clear_document(output_doc)

    first_record = True

    for index, row in csv_data.iterrows():
        # Nolasām šablonu
        doc = Document(template_path)
        
        # Aizvietojam placeholderus ar CSV datiem
        for paragraph in doc.paragraphs:
            for key, value in row.items():
                # Definējam gan {{key}}, gan {[key]} formātus
                placeholders = [f'{{{{{key}}}}}', f'{{[{key}]}}']
                for placeholder in placeholders:
                    if placeholder in paragraph.text:
                        # Pārbaudām, vai vērtība nav NaN, ja tā ir, aizvietojam ar tukšu stringu
                        if pd.isna(value):
                            replacement = ""
                        else:
                            replacement = str(value)
                        paragraph.text = paragraph.text.replace(placeholder, replacement)
                        # Pievienojam diagnostikas ziņojumu
                        st.write(f"Aizvietots `{placeholder}` ar `{replacement}`")
                    else:
                        # Pievienojam diagnostikas ziņojumu, ja vietturs netiek atrasts
                        st.write(f"Vietturs `{placeholder}` netika atrasts paragrafā.")

        # Pievienojam lappuses pārtraukumu, ja nav pirmais ieraksts
        if not first_record:
            output_doc.add_page_break()
        else:
            first_record = False

        # Pievienojam saturu manuāli
        for para in doc.paragraphs:
            # Izveidojam jaunu paragrafu ar tādu pašu stilu un tekstu
            p = output_doc.add_paragraph()
            p.style = para.style
            for run in para.runs:
                r = p.add_run(run.text)
                r.bold = run.bold
                r.italic = run.italic
                r.underline = run.underline

        for table in doc.tables:
            # Izveidojam jaunu tabulu ar tādu pašu kolonnu skaitu
            table_copy = output_doc.add_table(rows=0, cols=len(table.columns))
            for row_table in table.rows:
                cells = table_copy.add_row().cells
                for i, cell in enumerate(row_table.cells):
                    cells[i].text = cell.text

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
            # Nolasām CSV ar pareizu kodējumu un Python engine, lai labāk apstrādātu multi-line fields
            data = pd.read_csv(uploaded_file, encoding='utf-8', engine='python', quoting=csv.QUOTE_ALL)
            st.write("### CSV Saturs:")
            st.dataframe(data)
            
            # Pārbaudām CSV kolonnas nosaukumus pirms pārveides
            st.write("### CSV Kolonnas Pirms Pārveides:", data.columns.tolist())

            # Automātiska kolonnu nosaukumu pārveide ar regex: aizvieto jebkuru neatbilstīgu rakstzīmi ar zemessvītri
            # `[^\w]+` - meklē vienu vai vairākus rakstzīmes, kas nav vārdu rakstzīmes (a-zA-Z0-9_)
            data.columns = data.columns.str.replace(r'[^\w]+', '_', regex=True)
            st.write("### Atjauninātās Kolonnas Pēc Pārveides:", data.columns.tolist())

            # Definējam kolonnu nosaukumu karti
            csv_column_to_placeholder = {
                "Vārds_uzvārds_nosaukums": "Vārds_uzvārds_nosaukums",
                "Adrese": "Adrese",
                "kadapz": "kadapz",
                "Nekustamā_īpašuma_nosaukums": "Nekustamā_īpašuma_nosaukums",
                "uzruna": "uzruna",
                "Atrasts_Zemes_Vienības_Kadastra_Apzīmējums_lapā_1": "Atrasts_Zemes_Vienības_Kadastra_Apzīmēju",
                "Uzņēmums": "Uzņēmums",
                "Vieta": "Vieta",
                "Pagasts_un_Novads": "Pagasts_un_Novads",
                "Tikšanās_vieta_un_laiks": "Tikšanās_vieta_un_laiks",
                "Tikšanās_datums": "Tikšanās_datums",
                "Mērnieks_Vārds_Uzvārds": "Mērnieks_Vārds_Uzvārds",
                "Mērnieks_Telefons": "Mērnieks_Telefons",
                "Sagatavotājs_Vārds_Uzvārds_Telefons": "Sagatavotājs_Vārds_Uzvārds_Telefons",
                "Sagatavotājs_e_pasts": "Sagatavotājs_e_pasts"
            }

            # Veicam kolonnu nosaukumu pārveidi ar manuālu kartēšanu
            data.rename(columns=csv_column_to_placeholder, inplace=True)
            st.write("### Kolonnu Nosaukumi Pēc Manuālās Pārveides:", data.columns.tolist())

            # Definējam nepieciešamās kolonnu nosaukumus pēc pārveides
            required_columns = list(csv_column_to_placeholder.values())

            # Pārbaudām, vai visi nepieciešamie kolonnu nosaukumi ir klāt
            missing_columns = set(required_columns) - set(data.columns)
            if missing_columns:
                st.error(f"Trūkst kolonnas: {missing_columns}")
            else:
                st.success("Visas nepieciešamās kolonnas ir klāt pēc pārveides.")

                # Turpināsim ar mail merge procesu
                st.write("### Pārbaudām vietturu aizvietošanu:")
                # Parādām dažas CSV datu rindas
                st.write("#### Piemērs no CSV datiem:")
                st.write(data.head())

                # Parādām direktorijas saturu (diagnostika)
                st.write("### Current working directory:", os.getcwd())
                st.write("### Files in directory:", os.listdir('.'))

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
