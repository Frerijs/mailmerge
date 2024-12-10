# streamlit_app.py

import streamlit as st
import pandas as pd
from docxtpl import DocxTemplate
import os
import csv  # Pārliecinieties, ka csv modulis ir importēts
from io import BytesIO
import zipfile

def clean_address_field(address):
    """
    Tīra 'Adrese' lauku, aizvietojot rindiņu pārtraukumus ar komatiem.
    
    Args:
        address (str): Oriģinālais adrese kā virkne.
    
    Returns:
        str: Izlabotais adrese bez rindiņu pārtraukumiem.
    """
    if isinstance(address, str):
        return address.replace('\n', ', ').replace('\r', ', ').strip()
    return address

def perform_mail_merge(template_path, records, output_path):
    """
    Veic mail merge visiem ierakstiem vienā dokumentā.
    
    Args:
        template_path (str): Ceļš uz Word šablonu (`template.docx`).
        records (list of dict): Saraksts ar vārdnīcām katram ierakstam.
        output_path (str): Ceļš uz izvadīgo dokumentu.
    
    Returns:
        bool: Vai mail merge veiksmīgi pabeigts.
    """
    try:
        template = DocxTemplate(template_path)
    except Exception as e:
        st.error(f"Neizdevās ielādēt šablonu: {e}")
        return False

    try:
        context = {'records': records}
        # Tīram 'Adrese' laukus katram ierakstam
        for record in context['records']:
            record['Adrese'] = clean_address_field(record['Adrese'])
        
        # Renderējam šablonu ar kontekstu
        template.render(context)
        
        # Saglabājam izvadīgo dokumentu
        template.save(output_path)
        return True
    except Exception as e:
        st.error(f"Kļūda renderējot dokumentu: {e}")
        return False

def main():
    st.title("Mail Merge Lietotne ar docxtpl")
    
    st.sidebar.header("Iestatījumi")
    
    # Augšupielādējam CSV failu
    uploaded_file = st.file_uploader("Augšupielādējiet CSV failu", type=["csv"])
    
    if uploaded_file is not None:
        try:
            # Nolasām CSV ar pandas, izmantojot pareizās opcijas
            data = pd.read_csv(
                uploaded_file,
                encoding='utf-8',
                engine='python',
                quoting=csv.QUOTE_ALL,
                skip_blank_lines=False
            )
            st.write("### CSV Saturs:")
            st.dataframe(data)
    
            # Pārbaudām CSV kolonnas nosaukumus pirms pārveides
            st.write("### CSV Kolonnas Pirms Pārveides:", data.columns.tolist())
    
            # Automātiska kolonnu nosaukumu pārveide ar regex: aizvieto neatbilstošas rakstzīmes ar zemessvītri
            data.columns = data.columns.str.replace(r'[^\w]+', '_', regex=True)
            st.write("### Atjauninātās Kolonnas Pēc Pārveides:", data.columns.tolist())
    
            # Definējam kolonnu nosaukumu karti
            csv_column_to_placeholder = {
                "Vārds_uzvārds_nosaukums": "Vārds_uzvārds_nosaukums",
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
    
                # Pārveidojam CSV datus par sarakstu vārdnīcām
                records = data.to_dict(orient='records')
                st.write("### Ieraksti:", records)
    
                # Veicam mail merge procesu, izveidojot vienu dokumentu ar visiem ierakstiem
                if st.button("Veikt Mail Merge"):
                    template_path = "template.docx"  # Pārliecinieties, ka template.docx ir pieejams
                    output_path = "merged_documents.docx"
    
                    success = perform_mail_merge(template_path, records, output_path)
    
                    if success:
                        st.success(f"Mail merge veiksmīgi pabeigts! Izveidotais dokuments: {output_path}")
                        
                        # Saglabājam izveidoto dokumentu kā lejupielāde
                        with open(output_path, "rb") as f:
                            doc_bytes = f.read()
                            st.download_button(
                                label="Lejupielādēt Merged Dokumentu",
                                data=doc_bytes,
                                file_name="merged_documents.docx",
                                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                            )
        except pd.errors.ParserError as e:
            st.error(f"CSV Parsing Kļūda: {e}")
        except Exception as e:
            st.error(f"Kļūda apstrādājot CSV failu: {e}")

if __name__ == "__main__":
    main()
