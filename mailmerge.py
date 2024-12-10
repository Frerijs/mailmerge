# streamlit_app.py

import streamlit as st
import pandas as pd
from docx import Document
import matplotlib.pyplot as plt
import os
import io
import csv
import re

def replace_placeholders(doc, row):
    """
    Aizvieto vietturu dokumentā, saglabājot noformējumu.
    
    Args:
        doc (Document): Word dokumenta objekts.
        row (pd.Series): Datu rinda no CSV.
    """
    for paragraph in doc.paragraphs:
        for key, value in row.items():
            placeholders = [f'{{{{{key}}}}}', f'{{[{key}]}}']
            for placeholder in placeholders:
                if placeholder in paragraph.text:
                    inline = paragraph.runs
                    for run in inline:
                        if placeholder in run.text:
                            run.text = run.text.replace(placeholder, str(value))
                            st.write(f"Aizvietots `{placeholder}` ar `{value}`")
    for table in doc.tables:
        for row_table in table.rows:
            for cell in row_table.cells:
                for paragraph in cell.paragraphs:
                    for key, value in row.items():
                        placeholders = [f'{{{{{key}}}}}', f'{{[{key}]}}']
                        for placeholder in placeholders:
                            if placeholder in paragraph.text:
                                inline = paragraph.runs
                                for run in inline:
                                    if placeholder in run.text:
                                        run.text = run.text.replace(placeholder, str(value))
                                        st.write(f"Aizvietots `{placeholder}` ar `{value}` tabulā")

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
    # Inicializē izvadītāja dokumentu un saglabā šablona stilu un struktūru
    template_doc = Document(template_path)
    output_doc = Document()

    # Kopē šablona stilu no šablona uz izvadītāja dokumentu
    output_doc.styles = template_doc.styles

    first_record = True

    for index, row in csv_data.iterrows():
        # Nolasām šablonu
        doc = Document(template_path)
        
        # Aizvietojam placeholderus ar CSV datiem, saglabājot formatējumu
        replace_placeholders(doc, row)

        # Pievienojam lappuses pārtraukumu, ja nav pirmais ieraksts
        if not first_record:
            output_doc.add_page_break()
        else:
            first_record = False

        # Pievienojam saturu no šablona pēc aizvietošanas
        for element in doc.element.body:
            output_doc.element.body.append(element)

    # Saglabājam izvadītāja dokumentu
    output_doc.save(output_path)
    return output_path
