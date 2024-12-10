# streamlit_app.py

import streamlit as st
import pandas as pd
from docxtpl import DocxTemplate
from docx import Document
from docx.enum.text import WD_BREAK
import os
from io import BytesIO

def clean_address_field(address):
    if isinstance(address, str):
        return address.replace('\n', ', ').replace('\r', ', ').strip()
    return address

def perform_mail_merge_single_document(template_path, records, output_path):
    try:
        template = DocxTemplate(template_path)
    except Exception as e:
        st.error(f"Neizdevās ielādēt šablonu: {e}")
        return False

    for idx, record in enumerate(records):
        try:
            context = record.copy()
            context['Adrese'] = clean_address_field(context['Adrese'])
            template.render(context)
            if idx == 0:
                template.save(output_path)
            else:
                # Pievienojam lappuses pārtraukumu un ierakstu
                doc = Document(output_path)
                doc.add_page_break()
                new_doc = DocxTemplate(template_path)
                new_doc.render(context)
                temp_path = os.path.join("temp.docx")
                new_doc.save(temp_path)
                temp_doc = Document(temp_path)
                for element in temp_doc.element.body:
                    doc.element.body.append(element)
                doc.save(output_path)
                os.remove(temp_path)
        except Exception as e:
            st.error(f"Kļūda renderējot ierakstu {idx+1}: {e}")
            continue

    return True

def main():
    st.title("Mail Merge Lietotne ar docxtpl")

    st.sidebar.header("Iestatījumi")

    uploaded_file = st.file_uploader("Augšupielādējiet CSV failu", type=["csv"])

    if uploaded_file is not None:
        try:
            data = pd.read_csv(
                uploaded_file,
                encoding='utf-8',
                engine='python',
                quoting=csv.QUOTE_ALL,
                skip_blank_lines=False
            )
            st.write("### CSV Saturs:")
            st.dataframe(data)

            st.write("### CSV Kolonnas Pirms Pārveides:", data.columns.tolist())

            data.columns = data.columns.str.replace(r'[^\w]+', '_', regex=True)
            st.write("### Atjauninātās Kolonnas Pēc Pārveides:", data.columns.tolist())

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

            data.rename(columns=csv_column_to_placeholder, inplace=True)
            st.write("### Kolonnu Nosaukumi Pēc Manuālās Pārveides:", data.columns.tolist())

            data.fillna("nav", inplace=True)
            st.write("### CSV Saturs Pēc NaN Aizvietošanas:", data.head())

            required_columns = list(csv_column_to_placeholder.values())

            missing_columns = set(required_columns) - set(data.columns)
            if missing_columns:
                st.error(f"Trūkst kolonnas: {missing_columns}")
            else:
                st.success("Visas nepieciešamās kolonnas ir klāt pēc pārveides.")

                records = data.to_dict(orient='records')
                st.write("### Ieraksti:", records)

                if st.button("Veikt Mail Merge"):
                    template_path = "template.docx"
                    output_path = "merged_documents.docx"

                    success = perform_mail_merge_single_document(template_path, records, output_path)

                    if success:
                        st.success(f"Mail merge veiksmīgi pabeigts! Izveidots dokuments: {output_path}")
                        
                        with open(output_path, "rb") as f:
                            file_bytes = f.read()
                            st.download_button(
                                label="Lejupielādēt Merged Dokumentu",
                                data=file_bytes,
                                file_name="merged_documents.docx",
                                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                            )
        except pd.errors.ParserError as e:
            st.error(f"CSV Parsing Kļūda: {e}")
        except Exception as e:
            st.error(f"Kļūda apstrādājot CSV failu: {e}")

if __name__ == "__main__":
    main()
