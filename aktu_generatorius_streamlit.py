import os
import re
import shutil
from openpyxl import load_workbook
import win32com.client
from datetime import datetime
import calendar
import pythoncom

streamlit_code = '''
import streamlit as st
import os
import re
from openpyxl import load_workbook
import win32com.client
from datetime import datetime
import calendar
import pythoncom

def get_current_month_end_and_name():
    today = datetime.today()
    current_month_end = today.replace(day=calendar.monthrange(today.year, today.month)[1])
    month_names = [
        "sausio", "vasario", "kovo", "balandžio", "gegužės", "birželio",
        "liepos", "rugpjūčio", "rugsėjo", "spalio", "lapkričio", "gruodžio"
    ]
    current_month_name = month_names[today.month - 1]
    return current_month_end.strftime("%Y-%m-%d"), current_month_name

def process_excels_in_folder(folder_path):
    if not os.path.exists(folder_path):
        st.error(f"❌ Aplankas nerastas: {folder_path}")
        return

    current_month_end, current_month_name = get_current_month_end_and_name()
    months = [
        "sausio", "vasario", "kovo", "balandžio", "gegužės", "birželio",
        "liepos", "rugpjūčio", "rugsėjo", "spalio", "lapkričio", "gruodžio"
    ]
    all_files_processed = True

    pythoncom.CoInitialize()
    excel_app = win32com.client.Dispatch("Excel.Application")
    excel_app.Visible = False

    for root, _, files in os.walk(folder_path):
        st.write(f"📁 Tikrinamas aplankas: {root}")
        for filename in files:
            if filename.endswith(".xlsx"):
                file_path = os.path.join(root, filename)
                st.write(f"🔄 Apdorojamas failas: {filename}")
                try:
                    workbook = load_workbook(file_path)
                    sheet = workbook.active

                    if sheet["C5"].value:
                        sheet["C5"].value = current_month_end
                        st.success(f"✅ Data įrašyta į C5: {current_month_end}")
                    else:
                        st.warning("⚠️ Langelyje C5 nėra reikšmės.")

                    if sheet["A9"].value:
                        cell_value = sheet["A9"].value.strip().lower()
                        for month in months:
                            if month in cell_value:
                                sheet["A9"].value = re.sub(month, current_month_name, cell_value, flags=re.IGNORECASE)
                                st.success(f"✅ Langelyje A9 pakeista į: {sheet['A9'].value}")
                                break
                    else:
                        st.warning("⚠️ Langelyje A9 nėra reikšmės.")

                    workbook.save(file_path)
                    workbook.close()
                    st.info(f"💾 Pakeitimai išsaugoti: {filename}")

                    excel_workbook = excel_app.Workbooks.Open(file_path)
                    pdf_file_path = file_path.replace(".xlsx", ".pdf")

                    counter = 1
                    while os.path.exists(pdf_file_path):
                        base, ext = os.path.splitext(pdf_file_path)
                        pdf_file_path = f"{base}_v{counter}{ext}"
                        counter += 1

                    excel_workbook.ExportAsFixedFormat(0, pdf_file_path)
                    excel_workbook.Close()
                    st.success(f"📄 PDF sugeneruotas: {pdf_file_path}")

                except Exception as e:
                    st.error(f"❌ Klaida faile '{filename}': {e}")
                    all_files_processed = False

    excel_app.Quit()
    if all_files_processed:
        st.success("✅ Visi failai sėkmingai apdoroti ir konvertuoti į PDF!")
    else:
        st.warning("⚠️ Įvyko klaidų – kai kurie failai nebuvo apdoroti.")

st.title("📄 Aktų generavimo įrankis")
base_folder_path = st.text_input("Įveskite Excel failų aplanko kelią:", "C:\\\\Users\\\\sigitaabasoviene\\\\OneDrive - Corpus A, UAB\\\\Desktop\\\\Aktai")

if st.button("▶️ Pradėti apdorojimą"):
    process_excels_in_folder(base_folder_path)
'''

file_path = "/mnt/data/aktu_generatorius_streamlit.py"
with open(file_path, "w", encoding="utf-8") as f:
    f.write(streamlit_code)

file_path
