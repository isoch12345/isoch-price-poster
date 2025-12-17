import io, os, zipfile, tempfile
from datetime import datetime
import pandas as pd
from PIL import Image, ImageDraw, ImageFont
import streamlit as st

CANVAS_W, CANVAS_H = 1080, 1920
COMPANY_NAME = "INNOVATIVE SOCH"
EMAILS = [
    "info@innovativesoch.com",
    "sapnarani@innovativesoch.com",
    "sandeep@innovativesoch.com",
    "sales@innovativesoch.com",
    "purchase@innovativesoch.com",
]
DISCLAIMER = "Prices are indicative. Final order confirmation by call."

def try_font(paths, size):
    for p in paths:
        try: return ImageFont.truetype(p, size)
        except: pass
    return ImageFont.load_default()

def money(v):
    try: return f"{float(v):.2f}"
    except: return ""

def wrap_text(draw, text, font, max_width):
    words = text.split()
    lines, line = [], ""
    for w in words:
        test = (line + " " + w).strip()
        if draw.textlength(test, font=font) <= max_width:
            line = test
        else:
            if line: lines.append(line)
            line = w
    if line: lines.append(line)
    return lines

def build_image(template_img, destination, rows, out_path, date_str):
    img = template_img.copy().convert("RGB")
    draw = ImageDraw.Draw(img)

    font_paths = [
        "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf",
        "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
    ]
    H1 = try_font(font_paths, 58)
    H2 = try_font(font_paths, 42)
    H3 = try_font(font_paths, 34)
    BODY = try_font(font_paths, 34)
    SMALL = try_font(font_paths, 26)

    DARK = (40, 30, 25)
    LINE = (160, 130, 110)
    pad = 70

    y = 90
    draw.text((pad, y), f"TODAY PRICES — {destination.upper()}", fill=DARK, font=H2)
    y += 60

    w = draw.textlength(COMPANY_NAME, font=H1)
    draw.text(((CANVAS_W - w) / 2, y), COMPANY_NAME, fill=DARK, font=H1)
    y += 80

    w = draw.textlength(date_str, font=H3)
    draw.text(((CANVAS_W - w) / 2, y), date_str, fill=DARK, font=H3)
    y += 40

    draw.line((pad, y, CANVAS_W - pad, y), fill=LINE, width=3)
    y += 30

    table_left, table_right = pad, CANVAS_W - pad
    col_item_w = int((table_right - table_left) * 0.68)
    header_h = 70

    draw.rectangle((table_left, y, table_right, y + header_h), outline=LINE, width=3)
    draw.text((table_left + 20, y + 18), "ITEM", fill=DARK, font=H3)
    draw.text((table_left + col_item_w + 20, y + 18), "PRICE (₹/KG)", fill=DARK, font=H3)
    y += header_h

    row_h = 78
    rows = rows.head(14).copy()

    for _, r in rows.iterrows():
        draw.line((table_left, y + row_h, table_right, y + row_h), fill=LINE, width=2)
        draw.line((table_left + col_item_w, y, table_left + col_item_w, y + row_h), fill=LINE, width=2)
        draw.text((table_left + 20, y + 20), str(r["Product"]), fill=DARK, font=BODY)
        draw.text((table_left + col_item_w + 20, y + 20), f"₹ {money(r['DeliveredPrice'])}", fill=DARK, font=BODY)
        y += row_h

    footer_y = CANVAS_H - 260
    draw.line((pad, footer_y, CANVAS_W - pad, footer_y), fill=LINE, width=3)
    footer_y += 18

    # Disclaimer
    for line in wrap_text(draw, DISCLAIMER, SMALL, CANVAS_W - 2*pad)[:2]:
        draw.text((pad, footer_y), line, fill=DARK, font=SMALL)
        footer_y += 32

    footer_y += 8

    # Emails
    email_text = "Email: " + " | ".join(EMAILS)
    for line in wrap_text(draw, email_text, SMALL, CANVAS_W - 2*pad)[:3]:
        draw.text((pad, footer_y), line, fill=DARK, font=SMALL)
        footer_y += 34

    img.save(out_path, "PNG")

def run(excel_bytes, template_bytes):
    df = pd.read_excel(io.BytesIO(excel_bytes))
    df.columns = [c.strip() for c in df.columns]
    if "Product " in df.columns and "Product" not in df.columns:
        df = df.rename(columns={"Product ": "Product"})

    if "For" in df.columns:
        df["DeliveredPrice"] = df["For"]
    else:
        for col in ["Ex Price", "Freight", "Margin", "GST(5%)"]:
            if col not in df.columns:
                df[col] = 0
        df["DeliveredPrice"] = df["Ex Price"] + df["Freight"] + df["Margin"] + df["GST(5%)"]

    # Required columns check
    if "Destination" not in df.columns or "Product" not in df.columns:
        raise ValueError("Excel must contain columns: Destination, Product (and either For or Ex Price+Freight+Margin+GST).")

    df["Destination"] = df["Destination"].astype(str).str.strip()
    df["Product"] = df["Product"].astype(str).str.strip()
    df["DeliveredPrice"] = pd.to_numeric(df["DeliveredPrice"], errors="coerce")
    df = df.dropna(subset=["DeliveredPrice"])

    today = datetime.now().strftime("%d-%m-%Y")

    template_img = Image.open(io.BytesIO(template_bytes)).convert("RGB")
    if template_img.size != (CANVAS_W, CANVAS_H):
        template_img = template_img.resize((CANVAS_W, CANVAS_H))

    with tempfile.TemporaryDirectory() as tmp:
        out_files = []
        for dest, g in df.groupby("Destination"):
            g2 = g[["Product", "DeliveredPrice"]].groupby("Product", as_index=False)["DeliveredPrice"].min()
            out_path = os.path.join(tmp, f"{dest}_{today}.png".replace("/", "-"))
            build_image(template_img, dest, g2.sort_values("Product"), out_path, today)
            out_files.append(out_path)

        zip_buf = io.BytesIO()
        with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as z:
            for f in out_files:
                z.write(f, arcname=os.path.basename(f))
        zip_buf.seek(0)
        return zip_buf

st.set_page_config(page_title="ISOCH Price Poster Generator", layout="centered")
st.title("ISOCH Price Poster Generator")
st.write("Upload Excel + Template PNG → Download ZIP (Destination-wise Posters)")

excel = st.file_uploader("Upload Excel (.xlsx)", type=["xlsx"])
template = st.file_uploader("Upload Template PNG (1080×1920)", type=["png"])

if excel and template:
    if st.button("Generate Posters"):
        try:
            zip_file = run(excel.read(), template.read())
            st.success("Done! Download your ZIP below.")
            st.download_button("Download ZIP", data=zip_file, file_name="destination_posters.zip", mime="application/zip")
        except Exception as e:
            st.error(f"Error: {e}")
