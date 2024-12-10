# streamlit_app.py

import streamlit as st
import pandas as pd
from docxtpl import DocxTemplate
from docx import Document
import matplotlib.pyplot as plt
import os
import csv
import re

def perform_mail_merge(template_path, csv_data, output_path):
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

    first_record = True

    for index, row in csv_data.iterrows():
        # Nolasām šablonu
        doc = DocxTemplate(template_path)

        # Sagatavo kontekstu, aizvietojot NaN ar 'nav'
        context = {key: (str(value) if pd.notna(value) else "nav") for key, value in row.items()}

        st.write(f"Aizvieto ar datiem: {context}")

        # Renderē šablonu ar kontekstu
        doc.render(context)

        # Saglabā renderēto dokumentu uz pagaidu failu
        temp_output = f"temp_{index}.docx"
        doc.save(temp_output)

        # Nolasām renderēto pagaidu dokumentu
        sub_doc = Document(temp_output)

        # Pievienojam lappuses pārtraukumu, ja nav pirmais ieraksts
        if not first_record:
            output_doc.add_page_break()
        else:
            first_record = False

        # Pievienojam saturu no renderētā dokumenta uz izvadītāja dokumentu
        for element in sub_doc.element.body:
            output_doc.element.body.append(element)

        # Dzēšam pagaidu failu
        os.remove(temp_output)

    # Saglabājam izvadītāja dokumentu
    output_doc.save(output_path)
    return output_path

def main():
    st.title("Mail Merge Lietotne ar DocxTemplate")

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
               # "Vārds_Uzvārds_nosaukums": "Vārds_Uzvārds_nosaukums",
                "Adrese": "Adrese",
                "kadapz": "kadapz",
                "Nekustamā_īpašuma_nosaukums": "Nekustamā_īpašuma_nosaukums",
                "uzruna": "uzruna",
                "Atrasts_Zemes_Vienības_Kadastra_Apzīmējums_lapā_1": "Atrasts_Zemes_Vienības_Kadastra_Apzīmējums_lapā_1",
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

            # Aizvietojam visus NaN ar "nav"
            data.fillna("nav", inplace=True)
            st.write("### CSV Saturs Pēc NaN Aizvietošanas:", data.head())

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
                    perform_mail_merge(template_path, data, output_path)
                    
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
