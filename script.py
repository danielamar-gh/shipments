import pandas as pd
from docx import Document
from datetime import datetime
import os
from docx2pdf import convert

# 1. קריאת הנתונים מקובץ CSV עם תמיכה בקידודים שונים ודילוג על שורות מיותרות
def load_data(csv_file):
    encodings = ["utf-8", "ISO-8859-8", "latin1", "windows-1255"]
    for enc in encodings:
        try:
            df = pd.read_csv(csv_file, encoding=enc, skiprows=2)
            df.columns = df.columns.str.strip()
            print(f"Successfully loaded file with encoding: {enc}")
            print("Columns in CSV:", df.columns.tolist())
            return df
        except UnicodeDecodeError:
            continue
    raise ValueError("Failed to decode CSV file with supported encodings")

# 2. סיכום נתוני המשלוחים בניכוי החזרות
def summarize_shipments(df):
    column_mapping = {
        'סוג מסמך': 'Document Type',
        'שם חשבון במסמך': 'Customer Name',
        'שם פריט': 'Item Name',
        'תאריך': 'Date',
        'סה\"כ בתנועה': 'Total Movement',
        'כמות': 'Quantity',
        'מחיר נטו': 'Net Price'
    }
    
    df.rename(columns=column_mapping, inplace=True)
    required_columns = list(column_mapping.values())
    for col in required_columns:
        if col not in df.columns:
            raise ValueError(f"Missing required column: {col}. Found columns: {df.columns.tolist()}")
    
    df['Date'] = pd.to_datetime(df['Date'], format='%d/%m/%Y', errors='coerce').dt.strftime('%d/%m/%Y')
    today_date = datetime.today().strftime('%d/%m/%Y')
    df = df[(df['Date'] == today_date) & ((df['Document Type'] == 'תעודת משלוח') | (df['Document Type'] == 'החזרה'))]
    df.loc[:, 'Total Movement'] = pd.to_numeric(df['Total Movement'], errors='coerce')
    
    summary = df.groupby(['Date', 'Document Type', 'Customer Name', 'Item Name', 'Net Price', 'Quantity']).agg({'Total Movement': 'sum'}).reset_index()
    summary.rename(columns={'Total Movement': 'Total Quantity'}, inplace=True)
    summary.sort_values(by=['Document Type'], ascending=False, inplace=True)
    
    print("Summary Data:")
    print(summary)
    
    if summary.empty:
        print("⚠️ Warning: The summary DataFrame is empty. No data to insert into the table.")
    
    return summary

# 3. פונקציה להחלפת טקסט תוך שמירה על העיצוב
def replace_text_across_runs(paragraph, placeholder, replacement_text):
    full_text = "".join(run.text for run in paragraph.runs)
    if placeholder in full_text:
        full_text = full_text.replace(placeholder, replacement_text)
        for run in paragraph.runs:
            run.text = ""
        paragraph.runs[0].text = full_text  # כתיבה מחדש תוך שמירה על העיצוב

# 4. יצירת קובץ PDF לכל לקוח
def fill_word_template_per_customer(summary, template_path, output_dir):
    os.makedirs(output_dir, exist_ok=True)
    unique_customers = summary['Customer Name'].unique()
    current_date = datetime.today().strftime('%d%m%Y')
    
    for customer in unique_customers:
        customer_data = summary[summary['Customer Name'] == customer]
        doc = Document(template_path)
        
        # מילוי שדות דינמיים תוך שמירה על העיצוב
        for paragraph in doc.paragraphs:
            replace_text_across_runs(paragraph, "{שם חשבון במסמך}", customer)
            replace_text_across_runs(paragraph, "{תאריך}", datetime.today().strftime('%d-%m-%Y'))
        
        table = doc.tables[0]
        headers = [cell.text.strip().replace("\u200e", "").replace("\ufeff", "").replace("{", "").replace("}", "") for cell in table.rows[0].cells]
        print(f"Processing customer: {customer}")
        print("Table Headers:", headers)
        
        column_map = {
            "תאריך": "Date",
            "סוג מסמך": "Document Type",
            "שם חשבון במסמך": "Customer Name",
            "שם פריט": "Item Name",
            "מחיר נטו": "Net Price",
            "כמות": "Quantity",
            'סה\"כ בתנועה': "Total Quantity"
        }
        
        column_indices = {column_map[key]: headers.index(key) for key in headers if key in column_map}
        print("Column Indices Mapping:", column_indices)
        
        for _, row in customer_data.iterrows():
            new_row = table.add_row().cells
            for col_name, df_col in column_indices.items():
                new_row[df_col].text = "{:,}".format(int(row[col_name])) if col_name == "Total Quantity" else str(row[col_name])
        
        word_file = os.path.join(output_dir, f"{current_date}_{customer}.docx")
        pdf_file = os.path.join(output_dir, f"{current_date}_{customer}.pdf")
        doc.save(word_file)
        convert(word_file)
        os.remove(word_file)  # מחיקת קובץ ה-Word לאחר יצירת ה-PDF
        print(f"מסמך PDF נשמר: {pdf_file}")

# דוגמה לשימוש
csv_file = 'shipments.csv'
template_path = 'Template.docx'
output_dir = os.path.join(os.path.dirname(__file__), 'customer_reports')

df = load_data(csv_file)
summary = summarize_shipments(df)
fill_word_template_per_customer(summary, template_path, output_dir)