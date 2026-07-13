import streamlit as st
from fpdf import FPDF
import os

st.title("Installer Form")

# --- Personal Information ---
st.header("Personal Information")
date = st.date_input("Date")
installer_name = st.text_input("Installer Name")
rfc = st.text_input("RFC")
bank_account = st.text_input("Bank Account")
bank_name = st.text_input("Bank Name")
cel_number = st.text_input("Cell Number")

# --- About Installation ---
st.header("About Installation")
installation_place = st.text_input("Installation Place")
customer_code = st.text_input("Customer Code")
customer_name = st.text_input("Customer Name")

# --- Installation Table ---
st.subheader("Installation Details")
num_rows = st.number_input("Number of items", min_value=1, max_value=20, value=1)
items = []
for i in range(int(num_rows)):
    st.write(f"Item {i+1}")
    model_name = st.text_input(f"Model Name {i+1}")
    serie_id = st.text_input(f"Serie ID {i+1}")
    amount = st.number_input(f"Amount {i+1}", min_value=0, value=1)
    items.append((model_name, serie_id, amount))

# --- Attachments (MULTIPLE FILES) ---
st.header("Attachments")
bank_info_files = st.file_uploader("Bank Information (Screenshots)", type=["png", "jpg", "jpeg"], accept_multiple_files=True)
id_picture_files = st.file_uploader("ID Pictures", type=["png", "jpg", "jpeg"], accept_multiple_files=True)
evidence_files = st.file_uploader("Evidence of Completion", type=["png", "jpg", "jpeg"], accept_multiple_files=True)

# Save uploaded files
upload_dir = "uploads"
os.makedirs(upload_dir, exist_ok=True)

def save_files(file_list):
    saved_paths = []
    if file_list:
        for file in file_list:
            path = os.path.join(upload_dir, file.name)
            with open(path, "wb") as f:
                f.write(file.getbuffer())
            saved_paths.append(path)
    return saved_paths

bank_paths = save_files(bank_info_files)
id_paths = save_files(id_picture_files)
evidence_paths = save_files(evidence_files)

# --- Generate PDF ---
if st.button("Generate PDF"):

    pdf = FPDF()
    pdf.set_auto_page_break(auto=True, margin=15)

    # ---------------- PAGE 1 ----------------
    pdf.add_page()

    # GREE LOGO
    pdf.image("gree_logo.png", x=10, y=8, w=40)

    # Title
    pdf.set_font("Arial", "B", 18)
    pdf.set_text_color(0, 80, 160)
    pdf.cell(200, 15, txt="INSTALLER FORM", ln=True, align="C")

    pdf.set_text_color(0, 0, 0)
    pdf.set_font("Arial", size=12)

    # Draw section boxes
    def section_header(title):
        pdf.set_fill_color(220, 230, 255)
        pdf.set_font("Arial", "B", 13)
        pdf.cell(190, 10, txt=title, ln=True, fill=True)
        pdf.set_font("Arial", size=12)

    # PERSONAL INFORMATION
    section_header("PERSONAL INFORMATION")
    pdf.cell(190, 8, txt=f"Date: {date}", ln=True)
    pdf.cell(190, 8, txt=f"Installer Name: {installer_name}", ln=True)
    pdf.cell(190, 8, txt=f"RFC: {rfc}", ln=True)
    pdf.cell(190, 8, txt=f"Bank Account: {bank_account}", ln=True)
    pdf.cell(190, 8, txt=f"Bank Name: {bank_name}", ln=True)
    pdf.cell(190, 8, txt=f"Cell Number: {cel_number}", ln=True)

    pdf.ln(5)

    # ABOUT INSTALLATION
    section_header("ABOUT INSTALLATION")
    pdf.cell(190, 8, txt=f"Installation Place: {installation_place}", ln=True)
    pdf.cell(190, 8, txt=f"Customer Code: {customer_code}", ln=True)
    pdf.cell(190, 8, txt=f"Customer Name: {customer_name}", ln=True)

    pdf.ln(5)

    # INSTALLATION TABLE
    section_header("INSTALLATION DETAILS")

    # Table header
    pdf.set_fill_color(200, 200, 200)
    pdf.cell(10, 8, "No", border=1, fill=True)
    pdf.cell(60, 8, "Model Name", border=1, fill=True)
    pdf.cell(60, 8, "Serie ID", border=1, fill=True)
    pdf.cell(20, 8, "Amount", border=1, fill=True)
    pdf.ln()

    # Table rows
    for i, (model_name, serie_id, amount) in enumerate(items, start=1):
        pdf.cell(10, 8, str(i), border=1)
        pdf.cell(60, 8, model_name, border=1)
        pdf.cell(60, 8, serie_id, border=1)
        pdf.cell(20, 8, str(amount), border=1)
        pdf.ln()

    # ---------------- PAGE 2 ----------------
    pdf.add_page()
    pdf.set_font("Arial", "B", 16)
    pdf.set_text_color(0, 80, 160)
    pdf.cell(200, 10, txt="ATTACHMENTS", ln=True, align="C")
    pdf.set_text_color(0, 0, 0)
    pdf.ln(5)

    def add_images(title, paths):
    section_header(title)
    if not paths:
        pdf.cell(190, 8, txt="No files uploaded.", ln=True)
        pdf.ln(5)
        return

    max_width = 90   # ancho máximo de cada imagen
    max_height = 120 # alto máximo de cada imagen
    images_per_page = 2
    count = 0

    for path in paths:
        # Mostrar nombre del archivo
        pdf.cell(190, 8, txt=os.path.basename(path), ln=True)

        # Insertar imagen con tamaño limitado
        pdf.image(path, w=max_width, h=max_height)
        pdf.ln(10)

        count += 1
        # Cada 2 imágenes, agregar nueva página
        if count % images_per_page == 0 and count < len(paths):
            pdf.add_page()
            section_header(title)


    add_images("Bank Information", bank_paths)
    add_images("ID Pictures", id_paths)
    add_images("Evidence of Completion", evidence_paths)

    pdf.output("installer_form.pdf")

    with open("installer_form.pdf", "rb") as f:
        st.download_button("Download PDF", f, file_name="installer_form.pdf")

    st.success("PDF generated successfully!")
