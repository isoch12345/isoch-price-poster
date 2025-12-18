import io
from datetime import datetime
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders

import pandas as pd
from PIL import Image, ImageDraw, ImageFont
import streamlit as st


# ----------------- SETTINGS -----------------
TEMPLATE_PATH = "isoch_template_final.png"

CANVAS_W, CANVAS_H = 1080, 1600
PAD = 80

HEADER_H = 200
STRIP_H = 60
FOOTER_H = 140

TABLE_LEFT = PAD
TABLE_RIGHT = CANVAS_W - PAD
TABLE_TOP = HEADER_H + STRIP_H + 20
TABLE_BOTTOM = CANVAS_H - 170  # just above footer divider

COL_PRODUCT_W = int((TABLE_RIGHT - TABLE_LEFT) * 0.65)
COL_PRICE_W = (TABLE_RIGHT - TABLE_LEFT) - COL_PRODUCT_W

ROW_H = 62

# Colors
BLACK = (28, 28, 28)
GRAY = (110, 120, 125)
RED = (155, 42, 42)
GOLD = (201, 169, 92)
ROW_BG = (252, 249, 244)
WHITE = (255, 255, 255)

DISCLAIMER = "Prices are indicative. Final order confirmation by call."
EMAIL_LINE = "📧 info@innovativesoch.com"

DEFAULT_SALES_TEAM = [
    "sapnarani@innovativesoch.com",
    "sandeep@innovativesoch.com",
    "sales@innovativesoch.com",
    "purchase@innovativesoch.com",
]
# -------------------------------------------


def load_font(size, bold=False):
    path = "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf" if bold else "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"
    return ImageFont.truetype(path, size)


H1 = load_font(44, True)
H2 = load_font(30, True)
H3 = load_font(26, True)
BODY = load_font(28, True)
BODY_REG = load_font(28, False)
SMALL = load_font(22, False)


def load_template():
    img = Image.open(TEMPLATE_PATH).convert("RGB")
    if img.size != (CANVAS_W, CANVAS_H):
        img = img.resize((CANVAS_W, CANVAS_H))
    return img


def parse_excel(excel_bytes: bytes) -> pd.DataFrame:
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

    if "Destination" not in df.columns or "Product" not in df.columns:
        raise ValueError("Excel must contain Destination, Product, and For (or pricing columns).")

    df["Destination"] = df["Destination"].astype(str).str.strip()
    df["Product"] = df["Product"].astype(str).str.strip()
    df["DeliveredPrice"] = pd.to_numeric(df["DeliveredPrice"], errors="coerce")
    df = df.dropna(subset=["DeliveredPrice"])

    # keep cheapest if duplicates
    df = df.groupby(["Destination", "Product"], as_index=False)["DeliveredPrice"].min()
    return df


def ellipsize(draw: ImageDraw.ImageDraw, text: str, font: ImageFont.ImageFont, max_width: int) -> str:
    """Trim text to fit max_width with …"""
    if draw.textlength(text, font=font) <= max_width:
        return text
    if max_width <= 20:
        return "…"
    s = text
    while s and draw.textlength(s + "…", font=font) > max_width:
        s = s[:-1]
    return (s + "…") if s else "…"


def build_poster(destination: str, date_str: str, rows: pd.DataFrame) -> bytes:
    img = load_template()
    draw = ImageDraw.Draw(img)

    # Title line (single line: destination left, date right)
    title_text = f"TODAY PRICES — {destination.upper()}"
    draw.text((TABLE_LEFT, 160), title_text, font=H2, fill=BLACK)

    dw = draw.textlength(date_str, font=H3)
    draw.text((TABLE_RIGHT - dw, 160), date_str, font=H3, fill=BLACK)

    # Destination label inside content (optional small)
    draw.text((TABLE_LEFT, TABLE_TOP - 55), "📍 DESTINATION", font=H3, fill=BLACK)

    # Table header row
    y = TABLE_TOP
    draw.rectangle((TABLE_LEFT, y, TABLE_RIGHT, y + ROW_H), fill=WHITE)

    draw.text((TABLE_LEFT + 12, y + 18), "PRODUCT", font=H3, fill=BLACK)

    price_label = "PRICE (₹/KG)"
    plw = draw.textlength(price_label, font=H3)
    draw.text((TABLE_RIGHT - 12 - plw, y + 18), price_label, font=H3, fill=BLACK)

    draw.line((TABLE_LEFT, y + ROW_H, TABLE_RIGHT, y + ROW_H), fill=GOLD, width=2)
    y += ROW_H + 6

    # Rows
    max_rows_area_bottom = CANVAS_H - FOOTER_H - 30
    max_product_width = COL_PRODUCT_W - 24
    for i, r in enumerate(rows.itertuples(index=False)):
        if y + ROW_H > max_rows_area_bottom:
            break

        if i % 2 == 0:
            draw.rectangle((TABLE_LEFT, y, TABLE_RIGHT, y + ROW_H), fill=ROW_BG)

        product = str(r.Product).strip()
        product = ellipsize(draw, product, BODY, max_product_width)
        draw.text((TABLE_LEFT + 12, y + 18), product, font=BODY, fill=BLACK)

        price = f"₹ {float(r.DeliveredPrice):.2f}"
        pw = draw.textlength(price, font=BODY_REG)
        draw.text((TABLE_RIGHT - 12 - pw, y + 18), price, font=BODY_REG, fill=BLACK)

        y += ROW_H

    # Footer (fixed, never overlaps)
    fy = CANVAS_H - FOOTER_H + 5
    draw.text(
        (CANVAS_W / 2, fy + 15),
        "✔ Consistent Quality   ✔ Reliable Supply   ✔ Transparent Pricing",
        font=SMALL,
        fill=RED,
        anchor="mm",
    )
    draw.text((CANVAS_W / 2, fy + 50), DISCLAIMER, font=SMALL, fill=GRAY, anchor="mm")
    draw.text((CANVAS_W / 2, fy + 82), EMAIL_LINE, font=SMALL, fill=BLACK, anchor="mm")

    buf = io.BytesIO()
    img.save(buf, format="PNG")
    buf.seek(0)
    return buf.getvalue()


def send_email(to_list, subject, body, attachments: dict):
    host = st.secrets.get("SMTP_HOST", "")
    port = int(st.secrets.get("SMTP_PORT", 465))
    user = st.secrets.get("SMTP_USER", "")
    pwd = st.secrets.get("SMTP_PASS", "")

    if not (host and user and pwd):
        raise ValueError("SMTP secrets missing in Streamlit Secrets (SMTP_HOST/SMTP_PORT/SMTP_USER/SMTP_PASS).")

    msg = MIMEMultipart()
    msg["From"] = user
    msg["To"] = ", ".join(to_list)
    msg["Subject"] = subject
    msg.attach(MIMEText(body, "plain"))

    for filename, file_bytes in attachments.items():
        part = MIMEBase("application", "octet-stream")
        part.set_payload(file_bytes)
        encoders.encode_base64(part)
        part.add_header("Content-Disposition", f'attachment; filename="{filename}"')
        msg.attach(part)

    with smtplib.SMTP_SSL(host, port) as server:
        server.login(user, pwd)
        server.sendmail(user, to_list, msg.as_string())


# ----------------- STREAMLIT UI -----------------
st.set_page_config(page_title="ISOCH Price Poster Generator", layout="centered")
st.title("ISOCH Price Poster Generator")

uploaded = st.file_uploader("Upload Excel (.xlsx)", type=["xlsx"])

if uploaded:
    try:
        df = parse_excel(uploaded.read())
        destinations = sorted(df["Destination"].unique().tolist())
    except Exception as e:
        st.error(f"Excel error: {e}")
        st.stop()

    date_str = datetime.now().strftime("%d-%m-%Y")

    sel = st.multiselect("Select destinations", destinations, default=destinations[:1] if destinations else [])
    st.caption("Clean output guaranteed (no overlaps).")

    if st.button("Generate Posters", type="primary"):
        if not sel:
            st.warning("Select at least one destination.")
            st.stop()

        posters = {}
        for d in sel:
            rows = df[df["Destination"] == d][["Product", "DeliveredPrice"]].sort_values("Product")
            posters[f"ISOCH_{d}_{date_str}.png"] = build_poster(d, date_str, rows)

        st.success("Generated.")
        for fname, data in posters.items():
            st.download_button(f"⬇️ Download {fname}", data=data, file_name=fname, mime="image/png", key=fname)

        st.markdown("---")
        st.subheader("Send to Sales Team (optional)")
        to_list = st.text_input("Recipients (comma separated)", value=", ".join(DEFAULT_SALES_TEAM))
        if st.button("Email Posters Now"):
            recipients = [x.strip() for x in to_list.split(",") if x.strip()]
            subject = f"ISOCH | Delivered Prices | {date_str}"
            body = f"Dear Team,\n\nPlease find attached today's delivered price updates.\n\n{DISCLAIMER}\n\nRegards,\nInnovative Soch\n"
            send_email(recipients, subject, body, posters)
            st.success("Email sent.")
else:
    st.info("Upload the Excel to start.")
