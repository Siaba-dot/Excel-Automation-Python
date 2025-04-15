import streamlit as st
import os
import re
import shutil
from openpyxl import load_workbook
from datetime import datetime
import calendar
import zipfile

# 📆 Funkcija apskaičiuoti einamojo mėnesio pabaigą ir pavadinimą
def get_current_month_end_and_name():
    today = datetime.today()
    current_month_end = today.replace(day=calendar.monthrange(today.year, today.month)[1])
    month_names = [
        "sausio", "vasario", "kovo", "balandžio", "gegužės", "birželio",
        "liepos", "rugpjūtčio", "rugsėjo", "spalio", "lapkričio", "gruodžio"
    ]
    current_month_name = month_names[today.month - 1]
    return current_month_end.strftime("%Y-%m-%d"), current_month_name

# 🗄 Funkcija Excel failams apdoroti
def process_excels_in_folder(folder_path):
    current_month_end, current_month_name = get_current_month_end_and_name()
    months = [
        "sausio", "vasario", "kovo", "balandžio", "gegužės", "birželio",
        "liepos", "rugpjūtčio", "rugsėjo", "spalio", "lapkričio", "gruodžio"
    ]
    all_files_processed = True

    for root, _, files in os.walk(folder_path):
        for filename in files:
            if filename.endswith(".xlsx"):
                file_path = os.path.join(root, filename)
                try:
                    workbook = load_workbook(file_path)
                    sheet = workbook.active

                    if sheet["C5"].value:
                        sheet["C5"].value = current_month_end

                    if sheet["A9"].value:
                        cell_value = sheet["A9"].value.strip().lower()
                        for month in months:
                            if month in cell_value:
                                sheet["A9"].value = re.sub(month, current_month_name, cell_value, flags=re.IGNORECASE)
                                break

                    workbook.save(file_path)
                    workbook.close()

                except Exception as e:
                    st.error(f"❌ Klaida apdorojant failą {filename}: {e}")
                    all_files_processed = False

    return all_files_processed

# 🗂 Funkcija ZIP failo sukūrimui
def create_zip_from_folder(folder_path, zip_path):
    with zipfile.ZipFile(zip_path, 'w') as zipf:
        for root, _, files in os.walk(folder_path):
            for file in files:
                file_path = os.path.join(root, file)
                arcname = os.path.relpath(file_path, folder_path)
                zipf.write(file_path, arcname)

# 📄 Streamlit UI
st.title("📄 Aktų generavimo aplikacija")

base_folder_path = st.text_input("Įveskite aplanko su Excel failais kelią:", "C:\\Users\\sigitaabasoviene\\OneDrive - Corpus A, UAB\\Desktop\\Aktai")

if st.button("Generuoti aktus"):
    if not os.path.exists(base_folder_path):
        st.error("❌ Nurodytas aplankas neegzistuoja.")
    else:
        st.info("⏳ Pradedamas aktų generavimo procesas...")
        success = process_excels_in_folder(base_folder_path)

        if success:
            zip_path = "aktai_sugeneruoti.zip"
            create_zip_from_folder(base_folder_path, zip_path)

            with open(zip_path, "rb") as f:
                st.download_button(
                    label="📦 Atsisiųsti visus Excel failus (.zip)",
                    data=f.read(),
                    file_name=zip_path,
                    mime="application/zip"
                )
            os.remove(zip_path)
            st.success("✅ Visi failai apdoroti sėkmingai!")
        else:
            st.warning("⚠️ Kai kurių failų apdoroti nepavyko. Patikrinkite klaidas.")
