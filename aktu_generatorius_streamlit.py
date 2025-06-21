from zipfile import ZipFile
import os

# Sukuriame struktūrą, kuri būtų naudinga GitHub projektui
project_files = {
    "aktai_app_streamlit/main.py": """
import streamlit as st
import os
import re
from openpyxl import load_workbook
from datetime import datetime
import calendar
import zipfile
import tempfile

def get_current_month_end_and_name():
    today = datetime.today()
    current_month_end = today.replace(day=calendar.monthrange(today.year, today.month)[1])
    month_names = [
        "sausio", "vasario", "kovo", "balandžio", "gegužės", "birželio",
        "liepos", "rugpjūtčio", "rugsėjo", "spalio", "lapkričio", "gruodžio"
    ]
    current_month_name = month_names[today.month - 1]
    return current_month_end.strftime("%Y-%m-%d"), current_month_name

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

def create_zip_from_folder(folder_path, zip_path):
    with zipfile.ZipFile(zip_path, 'w') as zipf:
        for root, _, files in os.walk(folder_path):
            for file in files:
                file_path = os.path.join(root, file)
                arcname = os.path.relpath(file_path, folder_path)
                zipf.write(file_path, arcname)

st.title("📄 Aktų generavimo aplikacija")

uploaded_zip = st.file_uploader("Įkelkite ZIP failą su Excel dokumentais:", type="zip")

if uploaded_zip is not None:
    with tempfile.TemporaryDirectory() as tmp_dir:
        zip_path = os.path.join(tmp_dir, "input.zip")
        with open(zip_path, "wb") as f:
            f.write(uploaded_zip.read())

        extract_dir = os.path.join(tmp_dir, "extracted")
        os.makedirs(extract_dir, exist_ok=True)
        with zipfile.ZipFile(zip_path, 'r') as zip_ref:
            zip_ref.extractall(extract_dir)

        if st.button("Generuoti aktus"):
            st.info("⏳ Pradedamas aktų generavimo procesas...")
            success = process_excels_in_folder(extract_dir)

            if success:
                output_zip_path = os.path.join(tmp_dir, "aktai_sugeneruoti.zip")
                create_zip_from_folder(extract_dir, output_zip_path)

                with open(output_zip_path, "rb") as f:
                    st.download_button(
                        label="📦 Atsisiųsti visus Excel failus (.zip)",
                        data=f.read(),
                        file_name="aktai_sugeneruoti.zip",
                        mime="application/zip"
                    )
                st.success("✅ Visi failai apdoroti sėkmingai!")
            else:
                st.warning("⚠️ Kai kurių failų apdoroti nepavyko. Patikrinkite klaidas.")
""",
    "aktai_app_streamlit/requirements.txt": "streamlit\nopenpyxl\n",
    "aktai_app_streamlit/README.md": """
# 📄 Aktų generavimo aplikacija

Ši Streamlit aplikacija leidžia automatiškai atnaujinti Excel failų datas bei mėnesių pavadinimus pagal einamąjį mėnesį.

## 🔧 Funkcionalumas

- Įkelkite ZIP archyvą su Excel failais
- Aplikacija apdoroja kiekvieną dokumentą:
  - Atnaujina datą langelyje `C5`
  - Keičia mėnesio pavadinimą langelyje `A9`
- Sugeneruoja naują ZIP atsisiuntimui

## 🧰 Naudojamos bibliotekos

- `streamlit`
- `openpyxl`
- `zipfile`

## 🚀 Paleidimas

```bash
streamlit run main.py
