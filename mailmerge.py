# mail_merge.py

import pandas as pd
from docx import Document
from docx.shared import Pt
import os

def perform_mail_merge(template_path, csv_path, output_dir):
    # Nolasām CSV datus
    data = pd.read_csv(csv_path)
    
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
    
    for index, row in data.iterrows():
        doc = Document(template_path)
        
        # Aizvietojam placeholderus ar CSV datiem
        for paragraph in doc.paragraphs:
            for key, value in row.items():
                placeholder = f'{{{{{key}}}}}'
                if placeholder in paragraph.text:
                    paragraph.text = paragraph.text.replace(placeholder, str(value))
        
        # Saglabājam jauno dokumentu
        output_path = os.path.join(output_dir, f"{row.get('Name', 'Document')}_{index}.docx")
        doc.save(output_path)
    
    return output_dir
