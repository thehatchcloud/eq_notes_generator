import argparse
import pandas as pd
import numpy as np

from pathlib import Path
from jinja2 import Environment, FileSystemLoader
from weasyprint import HTML


parser = argparse.ArgumentParser()
parser.add_argument('--report_name', help="The file path and name of the report")
parser.add_argument('--template_name', help="The name of the report template file.", default="notes_template.html")


args = parser.parse_args()

env = Environment(loader=FileSystemLoader('.'))
template = env.get_template(args.template_name)

df = pd.read_excel(args.report_name, sheet_name="Sheet0")

for note in df.iloc:
    template_vars = {
        'First_Name': note['First Name'],
        'Last_or_Org_Name': note['Last or Org Name'],
        'Lesson_Date': note['Lesson Date'],
        'Lesson_Notes': note['Lesson Notes'],
        'Term': note['Term'],
        'Instructor': note['Instructor'],
        'Billing_Code': note['Billing Code'],
        'Rider_Level': note['Rider Level'],
        'Horse': note['Horse'],
        'Leader': note['Leader'],
        'Saddle': note['Saddle'],
        'Side_Walker_Hold': note['Side Walker Hold'],
        'BridleReins': note['Bridle/Reins'],
        'Mount': note['Mount'],
        'Dismount': note['Dismount'],
        'Warm_Up': note['Warm Up'],
        'Warm_Up_Note': note['Warm Up Note'],
        'Constituent_Number': note['Constituent Number'],
        'Objective': note['Objective'],
        'Progression': note['Progression '],
        'Progression_Note': note['Progression Note'],
        'Cool_Down': note['Cool Down'],
        'Cool_Down_Note': note['Cool Down Note'],
        'Long_Term_Goal': note['Long Term Goal'],
        'Long_Term_Goal_Note': note['Long Term Goal Note'],
        'Short_Term_Goal_1': note['Short Term Goal 1'],
        'Short_Term_Goal_1_Note': note['Short Term Goal 1 Note'],
        'Short_Term_Goal_2': note['Short Term Goal 2'],
        'Short_Term_Goal_2_Note': note['Short Term Goal 2 Note'],
        'Comments': note['Comments'],
    }
    
    html_out = template.render(template_vars)
    output_path = Path(f"C:\\Users\\kahat\\OneDrive\\Documents\\Desktop\\Lesson_Notes\\{note['Lesson Date']}_{note['First Name']}_{note['Last or Org Name']}_lesson_notes.pdf")
    HTML(string=html_out).write_pdf(output_path)
