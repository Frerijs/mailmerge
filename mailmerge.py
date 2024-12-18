# streamlit_app.py

import streamlit as st
import pandas as pd
from docxtpl import DocxTemplate
from docx import Document
import os
import csv  # Šis imports ir nepieciešams, lai izmantotu csv moduli
from io import BytesIO
from docxcompose.composer import Composer

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

def perform_mail_merge(template_path, records, output_dir):
    """
    Veic mail merge katram ierakstam atsevišķā dokumentā.
    
    Args:
        template_path (str): Ceļš uz Word šablonu (`template.docx`).
        records (list of dict): Saraksts ar vārdnīcām katram ierakstam.
        output_dir (str): Ceļš uz izvadīgo direktoriju.
    
    Returns:
        list: Izvadīgo dokumentu ceļu saraksts.
    """
    output_paths = []
    try:
        template = DocxTemplate(template_path)
    except Exception as e:
        st.error(f"Neizdevās ielādēt šablonu: {e}")
        return output_paths

    for idx, record in enumerate(records):
        try:
            context = record.copy()
            # Tīram 'Adrese' lauku
            context['Adrese'] = clean_address_field(context['Adrese'])
            
            # Renderējam šablonu ar kontekstu
            template.render(context)
            
            # Saglabājam izvadīgo dokumentu
            output_path = os.path.join(output_dir, f"merged_document_{idx+1}.docx")
            template.save(output_path)
            output_paths.append(output_path)
        except Exception as e:
            st.error(f"Kļūda renderējot ierakstu {idx+1}: {e}")
            continue

    return output_paths

def merge_word_documents(file_paths, merged_output_path):
    """
    Apvieno vairākus Word dokumentus vienā dokumentā ar lapu pārtraukumiem.

    Args:
        file_paths (list): Saraksts ar Word dokumentu ceļiem, kas jāapvieno.
        merged_output_path (str): Ceļš, kur saglabāt apvienoto dokumentu.
    """
    if not file_paths:
        st.error("Nav dokumentu, kas varētu tikt apvienoti.")
        return

    try:
        master = Document(file_paths[0])
        composer = Composer(master)

        for file_path in file_paths[1:]:
            doc = Document(file_path)
            composer.append(doc)

        composer.save(merged_output_path)
        st.success(f"Apvienotais dokuments saglabāts kā: {merged_output_path}")
    except Exception as e:
        st.error(f"Kļūda apvienojot dokumentus: {e}")

def main():
    st.title("Mail Merge Lietotne ar docxtpl un docxcompose")

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
                quoting=csv.QUOTE_ALL,  # Nodrošina, ka visas pēdiņas tiek pareizi apstrādātas
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

                # Veicam mail merge procesu, izveidojot atsevišķus dokumentus katram ierakstam
                if st.button("Veikt Mail Merge"):
                    template_path = "template.docx"  # Pārliecinieties, ka template.docx ir pieejams
                    output_dir = "output_documents"
                    if not os.path.exists(output_dir):
                        os.makedirs(output_dir)

                    output_paths = perform_mail_merge(template_path, records, output_dir)

                    if output_paths:
                        st.success(f"Mail merge veiksmīgi pabeigts! Izveidotie dokumenti: {len(output_paths)}")
                        
                        # Izveidojam apvienoto dokumentu ar docxcompose
                        merged_document_path = os.path.join(output_dir, "apvienotais_dokuments.docx")
                        merge_word_documents(output_paths, merged_document_path)

                        # Sagatavojam apvienotā dokumenta saturu lejupielādei
                        if os.path.exists(merged_document_path):
                            with open(merged_document_path, "rb") as f:
                                merged_file = BytesIO(f.read())
                            
                            st.download_button(
                                label="Lejupielādēt Apvienoto Dokumentu",
                                data=merged_file,
                                file_name="apvienotais_dokuments.docx",
                                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                            )
                            
                            st.info("Lejupielādējiet apvienoto dokumentu `apvienotais_dokuments.docx`.")

                            # Dzēšam atsevišķos dokumentus, lai paliktu tikai ar apvienoto dokumentu
                            for file in output_paths:
                                try:
                                    os.remove(file)
                                except Exception as e:
                                    st.warning(f"Neizdevās dzēst failu {file}: {e}")
                            
                            st.success("Atsevišķie dokumenti tika veiksmīgi dzēsti.")
                    else:
                        st.error("Nav izveidoti nekādi dokumenti pēc mail merge procesa.")
        except pd.errors.ParserError as e:
            st.error(f"CSV Parsing Kļūda: {e}")
        except Exception as e:
            st.error(f"Kļūda apstrādājot CSV failu: {e}")

if __name__ == "__main__":
    main()
