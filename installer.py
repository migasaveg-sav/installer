"""
Installer Form — multi-installer payment request app.

Flow:
  1. Fill shared job info (date, site, customer) once per submission.
  2. Add one or more installers. Each installer has personal/banking data,
     an items table (model catalog with fixed price -> commission), and
     their own attachments (bank info, ID, evidence photos).
  3. Generate: one PDF per installer (for the individual payment file) plus
     one Excel workbook summarizing the whole submission in a readable,
     tabular form (no need to open every PDF to check totals/data).

Model catalog is persisted to models_catalog.json next to this script, so
prices stay consistent across submissions and new models can be added
in place, either from the catalog manager or inline on an item row.
"""

import io
import json
import os
import shutil
import uuid
import zipfile
from datetime import date as date_cls

import pandas as pd
import streamlit as st
from fpdf import FPDF
from PIL import Image, ImageOps

# --------------------------------------------------------------------------
# Paths & constants
# --------------------------------------------------------------------------
APP_DIR = os.path.dirname(os.path.abspath(__file__))
CATALOG_PATH = os.path.join(APP_DIR, "models_catalog.json")
WORK_DIR = os.path.join(APP_DIR, "_work")  # processed/compressed images live here
LOGO_PATH = os.path.join(APP_DIR, "gree_logo.png")

os.makedirs(WORK_DIR, exist_ok=True)

ATTACHMENT_CATEGORIES = [
    ("bank", "Bank Information (Screenshots)"),
    ("id", "ID Pictures"),
    ("evidence", "Evidence of Completion"),
]

# Bank info and ID pictures carry text that has to stay readable (account
# numbers, CLABE, the ID itself), so they get a much bigger, one-per-row
# layout instead of the compact gallery used for evidence photos — and,
# coming first, they naturally land in the upper half of the Attachments
# pages.
LARGE_ATTACHMENT_CATEGORIES = {"bank", "id"}

# Image handling
MAX_IMAGE_DIM = 1600      # px, longest side after compression
JPEG_QUALITY = 85
LOW_RES_WARNING_PX = 800  # warn if longest side below this (text may be unreadable)

# PDF image grid: images are packed left-to-right at a fixed target row
# height (like a photo-gallery "justified" layout) instead of a rigid 2-column
# grid, so narrow portrait phone photos don't each waste half their cell.
IMG_ROW_TARGET_H = 62   # mm, height every image in a row is scaled to (evidence)
IMG_GAP = 4              # mm, gap between images and between rows
IMG_CAPTION_H = 4        # mm, space reserved for the "#N" caption under a row

# Bank info / ID pictures: one large image per row instead of a packed grid.
LARGE_IMG_MAX_W = 160    # mm
LARGE_IMG_MAX_H = 140    # mm, roughly half the usable page height


# --------------------------------------------------------------------------
# Model catalog helpers
# --------------------------------------------------------------------------
# Each model in the catalog is {"item_code": "...", "price": 850.0}. This is
# the validation list for both "Model Name" and "Unit Price" on an item row:
# picking a model from here locks in its code and price instead of retyping
# them (and risking a typo that miscalculates the commission).
def load_catalog() -> dict:
    if not os.path.exists(CATALOG_PATH):
        return {}
    try:
        with open(CATALOG_PATH, "r", encoding="utf-8") as f:
            data = json.load(f)
        catalog = {}
        for k, v in data.items():
            if isinstance(v, dict):
                catalog[str(k)] = {"item_code": str(v.get("item_code", "")), "price": float(v.get("price", 0) or 0)}
            else:  # backward compatibility with the old {name: price} format
                catalog[str(k)] = {"item_code": "", "price": float(v or 0)}
        return catalog
    except Exception:
        return {}


def save_catalog(catalog: dict) -> None:
    clean = {}
    for k, v in catalog.items():
        name = str(k).strip()
        if not name:
            continue
        clean[name] = {"item_code": str(v.get("item_code", "")).strip(), "price": float(v.get("price", 0) or 0)}
    with open(CATALOG_PATH, "w", encoding="utf-8") as f:
        json.dump(clean, f, ensure_ascii=False, indent=2, sort_keys=True)


# --------------------------------------------------------------------------
# Text helper (fpdf core fonts are Latin-1 only, which covers Spanish accents)
# --------------------------------------------------------------------------
def t(s) -> str:
    return str(s).encode("latin-1", "replace").decode("latin-1")


def money(v) -> str:
    try:
        return f"${float(v):,.2f}"
    except Exception:
        return "$0.00"


def summarize_by_model(items: list) -> list:
    """Group item rows by model and sum quantity/total — mirrors the
    'N EQUIPO <model> <total>' breakdown used when filing the incentive
    payment request, so it can be read straight off the top of the table
    instead of added up by hand."""
    order, agg = [], {}
    for it in items:
        key = it.get("model_name", "")
        if key not in agg:
            agg[key] = {"model_name": key, "qty": 0, "total": 0.0}
            order.append(key)
        agg[key]["qty"] += it.get("quantity", 0) or 0
        agg[key]["total"] += (it.get("quantity", 0) or 0) * (it.get("incentive", 0) or 0)
    return [agg[k] for k in order]


# --------------------------------------------------------------------------
# Image processing: fix orientation, compress, and report low-resolution files
# --------------------------------------------------------------------------
def process_image(uploaded_file, dest_dir) -> dict:
    """Save a normalized/compressed copy of an uploaded image; return metadata."""
    img = Image.open(uploaded_file)
    img = ImageOps.exif_transpose(img)  # fix phone-camera rotation
    if img.mode not in ("RGB", "L"):
        img = img.convert("RGB")
    elif img.mode == "L":
        img = img.convert("RGB")

    orig_w, orig_h = img.size
    low_res = max(orig_w, orig_h) < LOW_RES_WARNING_PX

    scale = min(1.0, MAX_IMAGE_DIM / max(orig_w, orig_h))
    if scale < 1.0:
        img = img.resize((max(1, int(orig_w * scale)), max(1, int(orig_h * scale))), Image.LANCZOS)

    out_name = f"{uuid.uuid4().hex}.jpg"
    out_path = os.path.join(dest_dir, out_name)
    img.save(out_path, "JPEG", quality=JPEG_QUALITY, optimize=True)

    return {
        "path": out_path,
        "orig_name": getattr(uploaded_file, "name", out_name),
        "width": img.size[0],
        "height": img.size[1],
        "low_res": low_res,
        "orig_size_kb": getattr(uploaded_file, "size", None) and uploaded_file.size / 1024,
        "compressed_size_kb": os.path.getsize(out_path) / 1024,
    }


def process_uploaded_files(files, dest_dir) -> list:
    return [process_image(f, dest_dir) for f in (files or [])]


# --------------------------------------------------------------------------
# PDF building
# --------------------------------------------------------------------------
def section_header(pdf: FPDF, title: str):
    pdf.set_fill_color(220, 230, 255)
    pdf.set_font("Arial", "B", 13)
    pdf.cell(190, 10, txt=t(title), ln=True, fill=True)
    pdf.set_font("Arial", size=12)


def pack_rows(images: list, usable_w: float, row_h: float, gap: float):
    """Shelf-pack images into rows: each image is scaled to a fixed row
    height (preserving aspect ratio), then placed left-to-right until the
    next one would overflow the page width, starting a new row. This is
    what photo galleries do, and it fits far more portrait phone photos per
    page than a rigid fixed-column grid, since a tall narrow photo only
    takes the width it actually needs instead of a whole half-page cell."""
    rows, current, x_used = [], [], 0.0
    for entry in images:
        ratio = entry["width"] / entry["height"] if entry["height"] else 1
        w = min(row_h * ratio, usable_w)  # never wider than the page itself
        needed = w if not current else w + gap
        if current and x_used + needed > usable_w:
            rows.append(current)
            current, x_used = [], 0.0
            needed = w
        current.append((entry, w))
        x_used += needed
    if current:
        rows.append(current)
    return rows


def add_images_grid(pdf: FPDF, title: str, images: list):
    """Lay out images as a justified photo gallery (see pack_rows) so
    portrait/landscape/odd-ratio photos never get stretched, and pages
    aren't left half-empty the way a fixed 2-column grid would for narrow
    portrait photos. Each image gets a small "#N" caption (low-resolution
    ones get a warning marker); the full original filename stays out of the
    PDF and lives in the Excel "Attachments" sheet, matched by that number."""
    section_header(pdf, title)
    if not images:
        pdf.set_font("Arial", "I", 11)
        pdf.cell(190, 8, txt="No files uploaded.", ln=True)
        pdf.set_font("Arial", size=12)
        pdf.ln(3)
        return

    left_margin = pdf.l_margin
    usable_w = 190
    row_h = IMG_ROW_TARGET_H
    row_total_h = row_h + IMG_CAPTION_H + IMG_GAP

    rows = pack_rows(images, usable_w, row_h, IMG_GAP)

    y = pdf.get_y()
    if y + row_total_h > 270:
        pdf.add_page()
        section_header(pdf, title)
        y = pdf.get_y()

    idx = 0
    any_low_res = False
    for row in rows:
        if y + row_total_h > 270:
            pdf.add_page()
            section_header(pdf, title)
            y = pdf.get_y()

        x = left_margin
        for entry, w in row:
            idx += 1
            pdf.image(entry["path"], x=x, y=y, w=w, h=row_h)

            low_res = bool(entry.get("low_res"))
            any_low_res = any_low_res or low_res
            caption = f"#{idx}*" if low_res else f"#{idx}"
            pdf.set_xy(x, y + row_h)
            pdf.set_font("Arial", "BI" if low_res else "I", 7)
            pdf.cell(w, IMG_CAPTION_H, txt=t(caption), align="C")
            pdf.set_font("Arial", size=12)

            x += w + IMG_GAP

        y += row_total_h

    if any_low_res:
        pdf.set_font("Arial", "I", 8)
        pdf.set_xy(left_margin, y)
        pdf.cell(190, 5, txt=t("* baja resolucion: el texto podria no ser legible"), ln=True)
        pdf.set_font("Arial", size=12)
        y = pdf.get_y()

    pdf.set_xy(left_margin, y)


def contain_fit(px_w, px_h, max_w, max_h):
    """Scale (px_w, px_h) to fit inside (max_w, max_h), preserving aspect
    ratio, without ever stretching the image."""
    ratio = px_w / px_h if px_h else 1
    w, h = max_w, max_w / ratio
    if h > max_h:
        h = max_h
        w = h * ratio
    return w, h


def add_images_stack(pdf: FPDF, title: str, images: list, max_w=LARGE_IMG_MAX_W, max_h=LARGE_IMG_MAX_H):
    """Lay out images one per row, large, stacked vertically — used for bank
    info and ID pictures, where the text inside the photo has to stay
    legible, unlike the many small evidence photos that just need to show
    'it's installed'."""
    section_header(pdf, title)
    if not images:
        pdf.set_font("Arial", "I", 11)
        pdf.cell(190, 8, txt="No files uploaded.", ln=True)
        pdf.set_font("Arial", size=12)
        pdf.ln(3)
        return

    left_margin = pdf.l_margin
    y = pdf.get_y()
    any_low_res = False

    for idx, entry in enumerate(images, start=1):
        w, h = contain_fit(entry["width"], entry["height"], max_w, max_h)
        row_total_h = h + IMG_CAPTION_H + IMG_GAP

        if y + row_total_h > 270:
            pdf.add_page()
            section_header(pdf, title)
            y = pdf.get_y()

        pdf.image(entry["path"], x=left_margin, y=y, w=w, h=h)

        low_res = bool(entry.get("low_res"))
        any_low_res = any_low_res or low_res
        caption = f"#{idx}*" if low_res else f"#{idx}"
        pdf.set_xy(left_margin, y + h)
        pdf.set_font("Arial", "BI" if low_res else "I", 8)
        pdf.cell(w, IMG_CAPTION_H, txt=t(caption), align="L")
        pdf.set_font("Arial", size=12)

        y += row_total_h

    if any_low_res:
        pdf.set_font("Arial", "I", 8)
        pdf.set_xy(left_margin, y)
        pdf.cell(190, 5, txt=t("* baja resolucion: el texto podria no ser legible"), ln=True)
        pdf.set_font("Arial", size=12)
        y = pdf.get_y()

    pdf.set_xy(left_margin, y)


def _pdf_output_bytes(pdf: FPDF) -> bytes:
    """Normalize FPDF.output() to bytes regardless of fpdf2 version/quirks.

    fpdf2's .output() normally returns a bytearray, and bytes(...) on that
    works fine. But if the wrong package ever ends up installed (the legacy
    PyPI "fpdf" package, which is a different, unmaintained library that
    also exposes an `FPDF` class under the same import name), .output()
    returns a plain str instead — and bytes(some_str) with no encoding
    raises "TypeError: string argument without an encoding" at exactly this
    line. Handle both shapes defensively so a bad environment fails loudly
    with a clear message instead of this cryptic TypeError.
    """
    out = pdf.output()
    if isinstance(out, (bytes, bytearray)):
        return bytes(out)
    if isinstance(out, str):
        return out.encode("latin-1", "replace")
    raise TypeError(
        f"Unexpected type from FPDF.output(): {type(out)!r}. This usually means "
        "the wrong 'fpdf' package is installed (legacy 'fpdf' instead of 'fpdf2'). "
        "Check requirements.txt pins 'fpdf2', not 'fpdf'."
    )


def build_installer_pdf(job: dict, installer: dict) -> bytes:
    pdf = FPDF()
    pdf.set_auto_page_break(auto=True, margin=15)

    # ---------------- PAGE 1: data ----------------
    pdf.add_page()

    if os.path.exists(LOGO_PATH):
        pdf.image(LOGO_PATH, x=10, y=8, w=40)

    pdf.set_font("Arial", "B", 18)
    pdf.set_text_color(0, 80, 160)
    pdf.cell(200, 15, txt="INSTALLER PAYMENT FORM", ln=True, align="C")
    pdf.set_text_color(0, 0, 0)
    pdf.set_font("Arial", size=12)

    section_header(pdf, "JOB INFORMATION")
    pdf.cell(190, 8, txt=t(f"Date: {job['date']}"), ln=True)
    pdf.cell(190, 8, txt=t(f"Installation Place: {job['installation_place']}"), ln=True)
    pdf.cell(190, 8, txt=t(f"Customer Code: {job['customer_code']}"), ln=True)
    pdf.cell(190, 8, txt=t(f"Customer Name: {job['customer_name']}"), ln=True)
    pdf.ln(5)

    section_header(pdf, "PERSONAL INFORMATION")
    pdf.cell(190, 8, txt=t(f"Installer Name: {installer['name']}"), ln=True)
    pdf.cell(190, 8, txt=t(f"RFC / CURP: {installer['rfc']}"), ln=True)
    pdf.cell(190, 8, txt=t(f"Bank Account: {installer['bank_account']}"), ln=True)
    pdf.cell(190, 8, txt=t(f"Bank Name: {installer['bank_name']}"), ln=True)
    pdf.cell(190, 8, txt=t(f"Cell Number: {installer['cell_number']}"), ln=True)
    pdf.ln(5)

    # --- Summary by model (totals first, like the incentive breakdown) ---
    section_header(pdf, "SUMMARY BY MODEL")
    pdf.set_fill_color(200, 200, 200)
    pdf.set_font("Arial", "B", 11)
    pdf.cell(100, 8, "Item Model", border=1, fill=True)
    pdf.cell(30, 8, "Quantity", border=1, fill=True, align="C")
    pdf.cell(60, 8, "Total", border=1, fill=True, align="C")
    pdf.ln()
    pdf.set_font("Arial", size=11)

    grand_total = 0.0
    for row in summarize_by_model(installer["items"]):
        grand_total += row["total"]
        pdf.cell(100, 8, t(row["model_name"]), border=1)
        pdf.cell(30, 8, str(row["qty"]), border=1, align="C")
        pdf.cell(60, 8, money(row["total"]), border=1, align="R")
        pdf.ln()

    pdf.set_font("Arial", "B", 12)
    pdf.cell(130, 10, "TOTAL TO PAY", border=1)
    pdf.cell(60, 10, money(grand_total), border=1, align="R")
    pdf.ln(15)

    # --- Line-item detail: No / Item Model / Serie ID 1 / Serie ID 2 / Quantity / Incentive / Total ---
    pdf.set_font("Arial", "B", 10)
    pdf.set_fill_color(235, 235, 235)
    pdf.cell(8, 8, "No", border=1, fill=True)
    pdf.cell(45, 8, "Item Model", border=1, fill=True)
    pdf.cell(35, 8, "Serie ID 1", border=1, fill=True)
    pdf.cell(35, 8, "Serie ID 2", border=1, fill=True)
    pdf.cell(15, 8, "Qty", border=1, fill=True, align="C")
    pdf.cell(25, 8, "Incentive", border=1, fill=True, align="C")
    pdf.cell(27, 8, "Total", border=1, fill=True, align="C")
    pdf.ln()
    pdf.set_font("Arial", size=10)

    for i, item in enumerate(installer["items"], start=1):
        line_total = item["quantity"] * item["incentive"]
        pdf.cell(8, 7, str(i), border=1)
        pdf.cell(45, 7, t(item["model_name"]), border=1)
        pdf.cell(35, 7, t(item.get("serie_id_1", "")), border=1)
        pdf.cell(35, 7, t(item.get("serie_id_2", "")), border=1)
        pdf.cell(15, 7, str(item["quantity"]), border=1, align="C")
        pdf.cell(25, 7, money(item["incentive"]), border=1, align="R")
        pdf.cell(27, 7, money(line_total), border=1, align="R")
        pdf.ln()

    pdf.ln(10)

    # ---------------- PAGE 2+: attachments ----------------
    pdf.add_page()
    pdf.set_font("Arial", "B", 16)
    pdf.set_text_color(0, 80, 160)
    pdf.cell(200, 10, txt="ATTACHMENTS", ln=True, align="C")
    pdf.set_text_color(0, 0, 0)
    pdf.ln(5)

    for key, label in ATTACHMENT_CATEGORIES:
        images = installer["attachments"].get(key, [])
        if key in LARGE_ATTACHMENT_CATEGORIES:
            add_images_stack(pdf, label, images)
        else:
            add_images_grid(pdf, label, images)

    return _pdf_output_bytes(pdf)


# --------------------------------------------------------------------------
# Excel (readable summary) building
# --------------------------------------------------------------------------
def build_summary_excel(job: dict, installers: list) -> bytes:
    detail_rows = []
    summary_rows = []
    attachments_rows = []

    for inst in installers:
        inst_total = sum(it["quantity"] * it["incentive"] for it in inst["items"])
        n_equip = sum(it["quantity"] for it in inst["items"])

        summary_rows.append({
            "Date": job["date"],
            "Installation Place": job["installation_place"],
            "Customer Code": job["customer_code"],
            "Customer Name": job["customer_name"],
            "Installer Name": inst["name"],
            "RFC / CURP": inst["rfc"],
            "Bank Name": inst["bank_name"],
            "Bank Account": inst["bank_account"],
            "Cell Number": inst["cell_number"],
            "Units Installed": n_equip,
            "Total To Pay": inst_total,
        })

        for i, item in enumerate(inst["items"], start=1):
            detail_rows.append({
                "Installer Name": inst["name"],
                "RFC / CURP": inst["rfc"],
                "No": i,
                "Item Model": item["model_name"],
                "Serie ID 1": item.get("serie_id_1", ""),
                "Serie ID 2": item.get("serie_id_2", ""),
                "Quantity": item["quantity"],
                "Incentive": item["incentive"],
                "Total": item["quantity"] * item["incentive"],
            })

        for key, label in ATTACHMENT_CATEGORIES:
            for pos, entry in enumerate(inst["attachments"].get(key, []), start=1):
                attachments_rows.append({
                    "Installer Name": inst["name"],
                    "Category": label,
                    "PDF #": pos,  # matches the "#N" caption under the photo in that installer's PDF
                    "File Name": entry["orig_name"],
                    "Resolution": f"{entry['width']}x{entry['height']}",
                    "Low Resolution Warning": "YES" if entry["low_res"] else "",
                })

    df_summary = pd.DataFrame(summary_rows)
    df_detail = pd.DataFrame(detail_rows)
    df_attach = pd.DataFrame(attachments_rows)

    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df_summary.to_excel(writer, sheet_name="Summary", index=False)
        df_detail.to_excel(writer, sheet_name="Detail", index=False)
        if not df_attach.empty:
            df_attach.to_excel(writer, sheet_name="Attachments", index=False)

        money_cols_by_sheet = {"Summary": ["Total To Pay"], "Detail": ["Incentive", "Total"]}
        for sheet_name, df in (("Summary", df_summary), ("Detail", df_detail)):
            if df.empty:
                continue
            ws = writer.sheets[sheet_name]

            # Bold header + autosize columns
            for col_idx, col_name in enumerate(df.columns, start=1):
                header_cell = ws.cell(row=1, column=col_idx)
                header_cell.font = header_cell.font.copy(bold=True)
                max_len = max([len(str(col_name))] + [len(str(v)) for v in df[col_name]] + [8])
                ws.column_dimensions[header_cell.column_letter].width = min(max_len + 2, 40)

            # Currency formatting for money columns
            for money_col in money_cols_by_sheet.get(sheet_name, []):
                if money_col in df.columns:
                    c_idx = list(df.columns).index(money_col) + 1
                    for r in range(2, len(df) + 2):
                        ws.cell(row=r, column=c_idx).number_format = "$#,##0.00"

            ws.freeze_panes = "A2"

    return buf.getvalue()


# ==========================================================================
# STREAMLIT UI
# ==========================================================================
def main():
    st.set_page_config(page_title="Installer Form", layout="wide")
    st.title("Installer Form")

    # --- Model catalog manager: the validation list for Model Name / Item Code / Unit Price ---
    with st.expander("Manage model catalog (name + item code + fixed price / commission)", expanded=False):
        catalog = load_catalog()
        cat_df = pd.DataFrame([
            {"Model": k, "Item Code": v["item_code"], "Price": v["price"]}
            for k, v in sorted(catalog.items())
        ]) if catalog else pd.DataFrame(columns=["Model", "Item Code", "Price"])

        edited_cat = st.data_editor(
            cat_df, num_rows="dynamic", width="stretch", key="catalog_editor",
            column_config={"Price": st.column_config.NumberColumn(format="$%.2f", min_value=0.0)},
        )
        if st.button("Save catalog"):
            new_catalog = {
                str(row["Model"]).strip(): {"item_code": str(row["Item Code"] or "").strip(), "price": float(row["Price"] or 0)}
                for _, row in edited_cat.iterrows() if str(row["Model"]).strip()
            }
            save_catalog(new_catalog)
            st.success(f"Catalog saved ({len(new_catalog)} models).")
            st.rerun()

    catalog = load_catalog()
    model_options = sorted(catalog.keys())
    NEW_MODEL_OPTION = "+ Model not in list (enter manually)"

    # --- Job (shared) information ---
    st.header("Job Information")
    c1, c2 = st.columns(2)
    with c1:
        job_date = st.date_input("Date", value=date_cls.today())
        installation_place = st.text_input("Installation Place")
    with c2:
        customer_code = st.text_input("Customer Code")
        customer_name = st.text_input("Customer Name")

    st.divider()

    # --- Installers (dynamic list) ---
    if "installer_ids" not in st.session_state:
        st.session_state.installer_ids = [str(uuid.uuid4())]

    st.header("Installers")
    st.caption("Add one entry per installer that needs to be paid for this job.")

    col_add, col_info = st.columns([1, 4])
    with col_add:
        if st.button("➕ Add installer"):
            st.session_state.installer_ids.append(str(uuid.uuid4()))
            st.rerun()

    installers_payload = []

    for idx, iid in enumerate(st.session_state.installer_ids, start=1):
        with st.expander(f"Installer {idx}", expanded=True):
            top = st.columns([5, 1])
            with top[1]:
                if len(st.session_state.installer_ids) > 1 and st.button("🗑️ Remove", key=f"remove_{iid}"):
                    st.session_state.installer_ids.remove(iid)
                    st.rerun()

            c1, c2 = st.columns(2)
            with c1:
                name = st.text_input("Installer Name", key=f"name_{iid}")
                rfc = st.text_input("RFC / CURP", key=f"rfc_{iid}")
                cel_number = st.text_input("Cell Number", key=f"cel_{iid}")
            with c2:
                bank_name = st.text_input("Bank Name", key=f"bankname_{iid}")
                bank_account = st.text_input("Bank Account", key=f"bankacc_{iid}")

            num_items_key = f"numitems_{iid}"

            # --- Live "summary by model" preview, shown ABOVE the item rows ---
            # Reads each row's current value out of session_state before that
            # row's own widgets are (re)created further down, so the totals
            # at the top always reflect what's currently in the table below.
            def _row_preview(i):
                choice = st.session_state.get(f"model_{iid}_{i}")
                if choice is None:
                    return None
                if choice == NEW_MODEL_OPTION:
                    name = st.session_state.get(f"newmodel_{iid}_{i}", "")
                    price = st.session_state.get(f"newincentive_{iid}_{i}", 0.0)
                else:
                    name = choice
                    price = catalog.get(choice, {}).get("price", 0.0)
                quantity = st.session_state.get(f"qty_{iid}_{i}", 0) or 0
                return {"model_name": name, "incentive": float(price or 0), "quantity": quantity}

            preview_items = [p for p in (
                _row_preview(i) for i in range(int(st.session_state.get(num_items_key, 1)))
            ) if p and p["model_name"]]

            st.markdown("**Summary by model**")
            if preview_items:
                summary_rows_preview = summarize_by_model(preview_items)
                st.dataframe(
                    pd.DataFrame([
                        {"Item Model": r["model_name"], "Quantity": r["qty"], "Total": money(r["total"])}
                        for r in summary_rows_preview
                    ]),
                    hide_index=True, width="stretch",
                )
                grand = sum(r["total"] for r in summary_rows_preview)
            else:
                st.caption("No items yet.")
                grand = 0.0
            st.markdown(f"**Total to pay: {money(grand)}**")

            num_items = st.number_input(
                "Number of items", min_value=1, max_value=20, value=1, key=num_items_key
            )
            items = []
            for i in range(int(num_items)):
                # Build the expander's label from whatever was saved on the
                # previous run (same trick as the live preview above), so a
                # collapsed row still shows which model/qty it holds instead
                # of a bare "Item N" — that's what makes collapsing useful
                # once several items have been added.
                row_preview = _row_preview(i)
                if row_preview and row_preview["model_name"]:
                    item_label = (
                        f"Item {i + 1}: {row_preview['model_name']} "
                        f"(x{int(row_preview['quantity'] or 0)}) — {money(row_preview['quantity'] * row_preview['incentive'])}"
                    )
                else:
                    item_label = f"Item {i + 1}"

                with st.expander(item_label, expanded=True, key=f"item_expander_{iid}_{i}"):
                    ic1, ic2, ic3, ic4, ic5 = st.columns([3, 2, 2, 1, 1.5])
                    with ic1:
                        choice = st.selectbox(
                            "Item Model",
                            options=model_options + [NEW_MODEL_OPTION],
                            key=f"model_{iid}_{i}",
                        )
                        if choice == NEW_MODEL_OPTION:
                            # Manual entry for every field when the model isn't in the catalog
                            new_name = st.text_input("Model name", key=f"newmodel_{iid}_{i}", placeholder="Model name")
                            new_incentive = st.number_input(
                                "Incentive", min_value=0.0, step=0.01, key=f"newincentive_{iid}_{i}"
                            )
                            model_name, incentive = new_name, new_incentive
                        else:
                            model_name = choice
                            incentive = catalog.get(choice, {}).get("price", 0.0)
                            st.caption(f"Incentive: {money(incentive)}")
                    with ic2:
                        serie_id_1 = st.text_input("Serie ID 1", key=f"serie1_{iid}_{i}")
                    with ic3:
                        serie_id_2 = st.text_input("Serie ID 2", key=f"serie2_{iid}_{i}")
                    with ic4:
                        quantity = st.number_input(
                            "Quantity", min_value=0, value=1, key=f"qty_{iid}_{i}",
                        )
                    with ic5:
                        st.caption(f"Total: {money(quantity * incentive)}")

                items.append({
                    "model_name": model_name, "serie_id_1": serie_id_1, "serie_id_2": serie_id_2,
                    "quantity": quantity, "incentive": incentive,
                })

            item_total = sum(it["quantity"] * it["incentive"] for it in items)
            st.markdown(f"**Subtotal for this installer: {money(item_total)}**")

            st.subheader("Attachments")
            st.caption("Photos are auto-compressed, rotated upright, and fit to the page without distortion.")
            attachments_raw = {}
            for key, label in ATTACHMENT_CATEGORIES:
                attachments_raw[key] = st.file_uploader(
                    label, type=["png", "jpg", "jpeg"], accept_multiple_files=True, key=f"files_{key}_{iid}"
                )

            installers_payload.append({
                "id": iid, "name": name, "rfc": rfc, "bank_account": bank_account,
                "bank_name": bank_name, "cell_number": cel_number,
                "items": items, "attachments_raw": attachments_raw,
            })

    st.divider()

    grand_total = sum(
        sum(it["quantity"] * it["incentive"] for it in inst["items"]) for inst in installers_payload
    )
    st.subheader(f"Grand total for this submission: {money(grand_total)}")

    # --------------------------------------------------------------------------
    # Generate
    # --------------------------------------------------------------------------
    if st.button("Generate documents", type="primary"):
        job = {
            "date": job_date, "installation_place": installation_place,
            "customer_code": customer_code, "customer_name": customer_name,
        }

        # Persist any inline "model not in list" entries so future submissions reuse them
        fresh_catalog = load_catalog()
        for inst in installers_payload:
            for item in inst["items"]:
                if item["model_name"] and item["model_name"] not in fresh_catalog:
                    fresh_catalog[item["model_name"]] = {"item_code": "", "price": float(item["incentive"] or 0)}
        save_catalog(fresh_catalog)

        low_res_alerts = []
        final_installers = []
        with st.spinner("Processing images and building documents..."):
            for inst in installers_payload:
                dest_dir = os.path.join(WORK_DIR, inst["id"])
                os.makedirs(dest_dir, exist_ok=True)
                attachments = {}
                for key, _ in ATTACHMENT_CATEGORIES:
                    processed = process_uploaded_files(inst["attachments_raw"].get(key), dest_dir)
                    attachments[key] = processed
                    for p in processed:
                        if p["low_res"]:
                            low_res_alerts.append(f"{inst['name'] or 'Installer'}: {p['orig_name']}")

                final_installers.append({**inst, "attachments": attachments})

            pdfs = {}
            for inst in final_installers:
                pdf_bytes = build_installer_pdf(job, inst)
                safe_name = (inst["name"] or inst["id"]).strip().replace(" ", "_") or inst["id"]
                pdfs[f"{safe_name}.pdf"] = pdf_bytes

            excel_bytes = build_summary_excel(job, final_installers)

            zip_buf = io.BytesIO()
            with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
                for fname, data in pdfs.items():
                    zf.writestr(fname, data)
                zf.writestr("submission_summary.xlsx", excel_bytes)
            zip_buf.seek(0)

            # Clean up temp compressed images now that they're embedded in the PDFs
            for inst in final_installers:
                shutil.rmtree(os.path.join(WORK_DIR, inst["id"]), ignore_errors=True)

        if low_res_alerts:
            st.warning(
                "These photos have low resolution and text may not be legible:\n\n"
                + "\n".join(f"- {a}" for a in low_res_alerts)
            )

        st.success("Documents generated successfully!")

        st.download_button(
            "⬇️ Download everything (ZIP: PDFs + Excel summary)",
            data=zip_buf,
            file_name=f"installer_submission_{job_date}.zip",
            mime="application/zip",
        )

        st.download_button(
            "⬇️ Download Excel summary only",
            data=excel_bytes,
            file_name=f"submission_summary_{job_date}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

        with st.expander("Download individual installer PDFs"):
            for fname, data in pdfs.items():
                st.download_button(f"⬇️ {fname}", data=data, file_name=fname, mime="application/pdf", key=f"dl_{fname}")


if __name__ == "__main__":
    main()
