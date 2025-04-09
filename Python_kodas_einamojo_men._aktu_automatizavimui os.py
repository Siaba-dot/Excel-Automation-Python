import streamlit as st
import os
import re
import shutil
from openpyxl import load_workbook
import win32com.client
from twilio.rest import Client
from datetime import datetime
import calendar

# 📂 Twilio konfigūracija
account_sid = 'AC66dc5106a5a3a0737fad9e2b40e49475'
auth_token = 'b08d14e922ecc00cd0e35be3a253ffb7'
twilio_number = '+13252210350'
recipient_number = '+37067017827'

# 📆 Funkcija apskaičiuoti einamojo mėnesio pabaigą ir pavadinimą
def get_current_month_end_and_name():
    today = datetime.today()
    current_month_end = today.replace(day=calendar.monthrange(today.year, today.month)[1])
    month_names = [
        "sausio", "vasario", "kovo", "balandžio", "gegužės", "birželio",
        "liepos", "rugpjūčio", "rugsėjo", "spalio", "lapkričio", "gruodžio"
    ]
    current_month_name = month_names[today.month - 1]
    return current_month_end.strftime("%Y-%m-%d"), current_month_name

# 📂 Funkcija patikrinti, ar aplankas egzistuoja
def check_folder_exists(folder_path):
    if not os.path.exists(folder_path):
        st.error(f"Klaida: Aplankas '{folder_path}' nerastas.")
        return False
    return True

# 📄 Funkcija Excel failams apdoroti ir konvertuoti į PDF
def process_excels_in_all_subfolders(base_folder_path):
    if not check_folder_exists(base_folder_path):
        return

    current_month_end, current_month_name = get_current_month_end_and_name()
    months = [
        "sausio", "vasario", "kovo", "balandžio", "gegužės", "birželio",
        "liepos", "rugpjūčio", "rugsėjo", "spalio", "lapkričio", "gruodžio"
    ]
    all_files_processed = True

    import pythoncom
    pythoncom.CoInitialize()

    excel_app = win32com.client.Dispatch("Excel.Application")
    excel_app.Visible = False

    for root, _, files in os.walk(base_folder_path):
        st.write(f"Apdorojamas aplankas: {root}")
        for filename in files:
            if filename.endswith(".xlsx"):
                file_path = os.path.join(root, filename)
                try:
                    st.write(f"🔄 Apdorojamas failas: {filename}")

                    workbook = load_workbook(file_path)
                    sheet = workbook.active

                    if sheet["C5"].value:
                        sheet["C5"].value = current_month_end
                        st.success(f"✅ Langelyje C5 data pakeista į: {current_month_end}")

                    if sheet["A9"].value:
                        cell_value = sheet["A9"].value.strip().lower()
                        for month in months:
                            if month in cell_value:
                                sheet["A9"].value = re.sub(month, current_month_name, cell_value, flags=re.IGNORECASE)
                                st.success(f"✅ Langelyje A9 tekstas pakeistas į: {sheet['A9'].value}")
                                break

                    workbook.save(file_path)
                    workbook.close()

                    new_filename = re.sub(r"(\d{4})_(\d{2})", f"{current_month_end[:4]}_{current_month_end[5:7]}", filename)
                    if new_filename != filename:
                        new_file_path = os.path.join(root, new_filename)
                        if not os.path.exists(new_file_path):
                            os.rename(file_path, new_file_path)
                            file_path = new_file_path
                            st.success(f"✅ Failas pervadintas į: {new_filename}")

                    excel_workbook = excel_app.Workbooks.Open(file_path)
                    pdf_file_path = file_path.replace(".xlsx", ".pdf")

                    counter = 1
                    while os.path.exists(pdf_file_path):
                        base_name, ext = os.path.splitext(pdf_file_path)
                        pdf_file_path = f"{base_name}_v{counter}{ext}"
                        counter += 1

                    excel_workbook.ExportAsFixedFormat(0, pdf_file_path)
                    excel_workbook.Close()
                    st.success(f"✅ PDF sugeneruotas: {pdf_file_path}")

                except Exception as e:
                    st.error(f"❌ Klaida apdorojant failą {filename}: {e}")
                    all_files_processed = False

    excel_app.Quit()

    if all_files_processed:
        send_sms("Visi failai apdoroti ir sugeneruoti PDF failai.", recipient_number)
        st.success("Visi failai apdoroti ir sugeneruoti PDF failai.")
    else:
        send_sms("Kai kurių failų apdoroti nepavyko. Patikrinkite klaidas.", recipient_number)
        st.warning("Kai kurių failų apdoroti nepavyko. Patikrinkite klaidas.")

    # Sukuriame ZIP failą iš kopijos ir siūlome atsisiųsti
    temp_zip_path = "temp_aktai.zip"
    shutil.make_archive('temp_aktai', 'zip', base_folder_path)

    with open(temp_zip_path, "rb") as f:
        zip_bytes = f.read()

    st.download_button(
        label="📦 Atsisiųsti visus aktus (.zip)",
        data=zip_bytes,
        file_name=datetime.now().strftime("Sugeneruoti_aktai_%Y%m%d_%H%M.zip"),
        mime="application/zip"
    )

    if os.path.exists(temp_zip_path):
        os.remove(temp_zip_path)

# 📢 Funkcija siųsti SMS
def send_sms(message, recipient):
    try:
        client = Client(account_sid, auth_token)
        safe_message = message.encode('ascii', 'ignore').decode('ascii')
        client.messages.create(body=safe_message, from_=twilio_number, to=recipient)
        st.success(f"📲 SMS išsiųsta į {recipient}: {safe_message}")
    except Exception as e:
        st.error(f"❌ Klaida siunčiant SMS: {e}")

# Streamlit UI
st.title("📄 Aktų generavimo aplikacija")

base_folder_path = st.text_input("Įveskite aplanko su Excel failais kelią", "C:\\Users\\sigitaabasoviene\\OneDrive - Corpus A, UAB\\Desktop\\Aktai")

if st.button("Generuoti aktus"):
    # Sukuriame darbinę kopiją
    backup_folder = base_folder_path + "_kopija"
    if os.path.exists(backup_folder):
        shutil.rmtree(backup_folder)  # Ištriname jei jau buvo sena kopija
    shutil.copytree(base_folder_path, backup_folder)
    st.info(f"✅ Sukurta aplanko kopija: {backup_folder}")

    # Dirbame su kopija, ne su originalu
    process_excels_in_all_subfolders(backup_folder)
