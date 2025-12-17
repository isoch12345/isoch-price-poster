import io, os, json, zipfile, tempfile, smtplib
from datetime import datetime
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders

import pandas as pd
from PIL import Image, ImageDraw, ImageFont
import streamlit as st

# ---------------- CONFIG ----------------
CANVAS_W, CANVAS_H = 1080, 1920
COMPANY_NAME = "INNOVATIVE SOCH"
TEMPLATE_PATH = "template_light.png"
MAPPING_PATH = "destination_mapping.json"

DEFAULT_SALES_TEAM = [
    "info@innovativesoch.com",
    "sapnarani@innovativesoch.com",
    "sandeep@innovativesoch.com",
    "sales@innovativesoch.com",
    "purchase@innovativesoch.com",
]

DISCLAIMER = "Prices are indicative. Final order confirmation by call."
# ----------------------------------------


def try_font(paths, size):
    for p in paths:
        try:
            return ImageFont.truetype(p, size)
        except Exception:
            pass
    return ImageFont.load_default()


def money(v):
    try:
        return f"{float(v):.2f}"
    except Exception:
        return ""


def wrap_text(draw, text, font, max_width):
    words = text.split()
    lines, line = [], ""
    for w in words:
        test = (line + " " + w).strip()
        if draw.textlength(test, font=font) <= max_width:
            line = test
        else:
            if line:
                lines.append(line)
            line = w
    if line:
        lines.append(line)
    return lines


def load_template():
    if not os.path.exists(TEMPLATE_PATH):
        raise FileNotFoundError(f"Missing {TEMPLATE_PATH}. Upload it to repo root.")
    img = Image.open(TEMPLATE_PATH).convert("RGB")
    if img.size != (CANVAS_W, CANVAS_H):
        img = img.resize((CANVAS_W, CANVAS_H))
    return img


def load_mapping():
    if not os.path.exists(MAPPING_PATH):
        return {}
    try:
        with open(MAPPING_PATH, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}


def save_mapping(mapping: dict):
    # NOTE: On Streamlit Cloud, this usually persists on the running instance,
    # but may reset after redeploy. If you need guaranteed persistence later,
    # we can store mapping in Google Sheet/DB.
    with open(MAPPING_PATH, "w", encoding="utf-8") as f:
        json.dump(mapping, f, indent=2, ensure_ascii=False)


def build_poster(template_img, destination, rows, out_path, date_str, email_line):
    img = template_img.copy()
    draw = ImageDraw.Draw(img)

    font_paths = [
        "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf",
        "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
    ]
    TITLE = try_font(font_paths, 48)
    H2 = try_font(font_paths, 36)
    H3 = try_font(font_paths, 30)
    BODY_B = try_font(font_paths, 32)
    BODY = try_font(font_paths, 30)
    SMALL = try_font(font_paths, 24)

    DARK = (35, 28, 24)
    LINE = (150, 135, 125)

    pad = 90

    # Header text zones (fits the template)
    y = 210
    draw.text((pad, y), f"TODAY PRICES — {destination.upper()}", fill=DARK, font=TITLE)
    y += 62

    w = draw.textlength(date_str, font=H3)
    draw.text((CANVAS_W - pad - w, y), date_str, fill=DARK, font=H3)

    # Table zone
    table_left, table_right = pad, CANVAS_W - pad
    table_top = 410
    col_item_w = int((table_right - table_left) * 0.68)

    # Table header
    y = table_top
    draw.line((table_left, y, table_right, y), fill=LINE, width=3)
    y += 16
    draw.text((table_left + 10, y), "ITEM", fill=DARK, font=H2)
    draw.text((table_left + col_item_w + 10, y), "PRICE (₹/KG)", fill=DARK, font=H2)
    y += 54
    draw.line((table_left, y, table_right, y), fill=LINE, width=2)

    # Rows
    row_h = 72
    max_rows = 14
    rows = rows.head(max_rows).copy()

    for i, r in rows.iterrows():
        y += 14
        product = str(r["Product"]).strip()
        price = money(r["DeliveredPrice"])

        # zebra row background (very subtle)
        if (i % 2) == 0:
            draw.rectangle((table_left, y-8, table_right, y + row_h - 8), fill=(255, 255, 255))

        draw.text((table_left + 10, y + 14), product, fill=DARK, font=BODY_B)

        # right aligned price
        price_text = f"₹ {price}"
        pw = draw.textlength(price_text, font=BODY)
        draw.text((table_right - 10 - pw, y + 16), price_text, fill=DARK, font=BODY)

        # row divider
        draw.line((table_left, y + row_h, table_right, y + row_h), fill=LINE, width=1)
        y += row_h - 8

    # Footer text
    fy = CANVAS_H - 235
    # disclaimer
    for line in wrap_text(draw, DISCLAIMER, SMALL, CANVAS_W - 2*pad)[:2]:
        draw.text((pad, fy), line, fill=DARK, font=SMALL)
        fy += 30
    fy += 6

    # email line
    for line in wrap_text(draw, email_line, SMALL, CANVAS_W - 2*pad)[:3]:
        draw.text((pad, fy), line, fill=DARK, font=SMALL)
        fy += 30

    img.save(out_path, "PNG")


def make_zip(files: list[str]) -> io.BytesIO:
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        for f in files:
            z.write(f, arcname=os.path.basename(f))
    buf.seek(0)
    return buf


def send_email_with_zip(to_list, subject, body, zip_bytes: bytes, filename="destination_posters.zip"):
    host = st.secrets.get("SMTP_HOST", "")
    port = int(st.secrets.get("SMTP_PORT", 465))
    user = st.secrets.get("SMTP_USER", "")
    pwd  = st.secrets.get("SMTP_PASS", "")

    if not (host and user and pwd):
        raise ValueError("SMTP secrets missing. Add SMTP_HOST/SMTP_PORT/SMTP_USER/SMTP_PASS in Streamlit Secrets.")

    msg = MIMEMultipart()
    msg["From"] = user
    msg["To"] = ", ".join(to_list)
    msg["Subject"] = subject

    msg.attach(MIMEText(body, "plain"))

    part = MIMEBase("application", "zip")
    part.set_payload(zip_bytes)
    encoders.encode_base64(part)
    part.add_header("Content-Disposition", f'attachment; filename="{filename}"')
    msg.attach(part)

    with smtplib.SMTP_SSL(host, port) as server:
        server.login(user, pwd)
        server.sendmail(user, to_list, msg.as_string())


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
        raise ValueError("Excel must have: Destination, Product, and either For OR (Ex Price, Freight, Margin, GST(5%)).")

    df["Destination"] = df["Destination"].astype(str).str.strip()
    df["Product"] = df["Product"].astype(str).str.strip()
    df["DeliveredPrice"] = pd.to_numeric(df["DeliveredPrice"], errors="coerce")
    df = df.dropna(subset=["DeliveredPrice"])
    return df


# ---------------- UI ----------------
st.set_page_config(page_title="ISOCH Price Poster Generator", layout="centered")

st.markdown("## ISOCH Price Poster Generator (Sales)")
st.caption("Upload Excel → Select destinations → Generate posters + WhatsApp text + Email ZIP")

excel = st.file_uploader("Upload Excel (.xlsx)", type=["xlsx"])

mapping = load_mapping()
all_destinations = sorted(list(mapping.keys()))

# If Excel uploaded, show destinations from Excel as well
df = None
excel_destinations = []
if excel:
    try:
        df = parse_excel(excel.read())
        excel_destinations = sorted(df["Destination"].unique().tolist())
    except Exception as e:
        st.error(f"Excel error: {e}")

dest_pool = sorted(set(all_destinations + excel_destinations))
selected_destinations = st.multiselect("Select Destinations (multi)", dest_pool, default=excel_destinations[:1] if excel_destinations else [])

with st.expander("⚙️ Destination → Email Mapping (Edit)"):
    st.write("Set who should receive which destination price update.")
    st.info("Tip: Add multiple emails separated by comma.")

    # allow adding/updating mapping
    dest_to_edit = st.selectbox("Destination", options=(dest_pool if dest_pool else ["(add destinations via Excel)"]))
    emails_str = ", ".join(mapping.get(dest_to_edit, DEFAULT_SALES_TEAM)) if dest_to_edit and dest_to_edit != "(add destinations via Excel)" else ", ".join(DEFAULT_SALES_TEAM)
    new_emails_str = st.text_area("Emails for this destination", value=emails_str, height=90)

    c1, c2 = st.columns(2)
    with c1:
        if st.button("Save Mapping"):
            emails_list = [e.strip() for e in new_emails_str.split(",") if e.strip()]
            mapping[dest_to_edit] = emails_list
            save_mapping(mapping)
            st.success(f"Saved mapping for {dest_to_edit}.")
    with c2:
        if st.button("Set All Destinations to Default Team"):
            for d in dest_pool:
                mapping[d] = DEFAULT_SALES_TEAM
            save_mapping(mapping)
            st.success("Applied default team to all destinations.")

# WhatsApp message preview
today = datetime.now().strftime("%d-%m-%Y")
wa_dest_text = " | ".join(selected_destinations) if selected_destinations else "(select destinations)"
whatsapp_text = (
    f"📢 *ISOCH – Today’s Delivered Prices* ({today})\n"
    f"📍 {wa_dest_text}\n\n"
    f"{DISCLAIMER}\n"
    f"📧 info@innovativesoch.com"
)

st.markdown("### Quick Message for WhatsApp")
st.code(whatsapp_text)

if st.button("📋 Copy WhatsApp Message"):
    st.toast("Copy from the box above (WhatsApp message).", icon="✅")

# Generate & Send
if st.button("🛠️ Generate & Send", type="primary"):
    if not df is None and selected_destinations:
        try:
            template_img = load_template()
            date_str = today

            posters = []
            per_dest_recipients = {}
            with tempfile.TemporaryDirectory() as tmp:
                for dest in selected_destinations:
                    g = df[df["Destination"] == dest].copy()
                    if g.empty:
                        continue

                    # lowest price per product
                    g2 = g[["Product", "DeliveredPrice"]].groupby("Product", as_index=False)["DeliveredPrice"].min().sort_values("Product")

                    # recipient list for this destination
                    recips = mapping.get(dest, DEFAULT_SALES_TEAM)
                    per_dest_recipients[dest] = recips
                    email_line = "Email: " + " | ".join(sorted(set(recips)))

                    out_path = os.path.join(tmp, f"ISOCH_{dest}_{date_str}.png".replace("/", "-"))
                    build_poster(template_img, dest, g2, out_path, date_str, email_line)
                    posters.append(out_path)

                if not posters:
                    st.warning("No posters generated (check destinations / Excel).")
                    st.stop()

                zip_buf = make_zip(posters)
                zip_bytes = zip_buf.getvalue()

            # Email sending: send ONE ZIP to the combined recipients (unique)
            all_to = sorted(set([e for dest, lst in per_dest_recipients.items() for e in lst]))
            subject = f"ISOCH | Delivered Prices | {date_str} | " + ", ".join(selected_destinations[:4]) + ("..." if len(selected_destinations) > 4 else "")
            body = (
                "Dear Team,\n\n"
                f"Please find attached today's delivered price posters for:\n- " + "\n- ".join(selected_destinations) + "\n\n"
                f"{DISCLAIMER}\n\n"
                "Regards,\n"
                "Innovative Soch\n"
            )

            send_email_with_zip(all_to, subject, body, zip_bytes, filename=f"ISOCH_Price_Posters_{date_str}.zip")

            st.success("✅ Generated posters + emailed ZIP to sales team mapping.")
            st.download_button("Download ZIP (optional)", data=zip_bytes, file_name=f"ISOCH_Price_Posters_{date_str}.zip", mime="application/zip")

        except Exception as e:
            st.error(f"Error: {e}")
    else:
        st.warning("Upload Excel and select at least one destination.")
