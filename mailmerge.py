# streamlit_app.py

import streamlit as st
import pandas as pd
from docxtpl import DocxTemplate
import matplotlib.pyplot as plt  # Importējam matplotlib
import os
import io
import csv

def perform_mail_merge(template_path, csv_data, output_path):
    """
    Veic mail merge, izmantojot docxtpl, un saglabā visus rezultātus vienā .docx failā.

    Args:
        template_path (str): Ceļš uz Word šablonu (`template.docx`).
        csv_data (pd.DataFrame): Pandas DataFrame ar CSV datiem.
        output_path (str): Ceļš uz izvadītāja .docx failu.

    Returns:
        str: Izvades faila ceļš.
    """
    try:
        template = DocxTemplate(template_path)
    except Exception as e:
        st.error(f"Neizdevās ielādēt šablonu: {e}")
        return None

    # Pārveidojam CSV datus par sarakstu vārdnīcām
    records = csv_data.to_dict(orient='records')

    context = {
        'records': records
    }

    try:
        template.render(context)
    except Exception as e:
        st.error(f"Neizdevās renderēt šablonu: {e}")
        return None

    try:
        template.save(output_path)
    except Exception as e:
        st.error(f"Neizdevās saglabāt izvadīgo dokumentu: {e}")
        return None

    return output_path

def main():
    st.title("Mail Merge Lietotne ar docxtpl")

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
                "Vieta": "Vieta",
                "Uzņēmums": "Uzņēmums",
                "Nekustamā_īpašuma_nosaukums": "Nekustamā_īpašuma_nosaukums",
                "Atrasts_Zemes_Vienības_Kadastra_Apzīmējums_lapā_1": "Atrasts_Zemes_Vienības_Kadastra_Apzīmējums_lapā_1",
                "Pagasts_un_Novads": "Pagasts_un_Novads",
                "kadapz": "kadapz",
                "uzruna": "uzruna",
                "Tikšanās_datums": "Tikšanās_datums",
                "Tikšanās_vieta_un_laiks": "Tikšanās_vieta_un_laiks",
                "Mērnieks_Telefons": "Mērnieks_Telefons",
                "Mērnieks_Vārds_Uzvārds": "Mērnieks_Vārds_Uzvārds",
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
                    result = perform_mail_merge(template_path, data, output_path)
                    
                    if result:
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
