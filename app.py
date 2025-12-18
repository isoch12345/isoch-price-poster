import io, os, json, base64, smtplib
from datetime import datetime, date
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders

import pandas as pd
import requests
from PIL import Image, ImageDraw, ImageFont
import streamlit as st


# ----------------- TEMPLATE + LAYOUT -----------------
TEMPLATE_PATH = "isoch_template_final.png"
CANVAS_W, CANVAS_H = 1080, 1600
PAD = 80

HEADER_H = 200
STRIP_H = 60
FOOTER_H = 140

TABLE_LEFT = PAD
TABLE_RIGHT = CANVAS_W - PAD
TABLE_TOP = HEADER_H + STRIP_H + 20

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

MAPPING_PATH = "config/destination_mapping.json"
LOG_PATH = "logs/price_sent_log.jsonl"
# ----------------------------------------------------


# ----------------- FONTS -----------------
def load_font(size, bold=False):
    path = "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf" if bold else "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"
    return ImageFont.truetype(path, size)

H2 = load_font(30, True)
H3 = load_font(26, True)
BODY_B = load_font(28, True)
BODY = load_font(28, False)
SMALL = load_font(22, False)
# ----------------------------------------


# ----------------- GITHUB HELPERS -----------------
def _gh_cfg():
    owner = st.secrets.get("GITHUB_OWNER", "")
    repo = st.secrets.get("GITHUB_REPO", "")
    branch = st.secrets.get("GITHUB_BRANCH", "master")
    token = st.secrets.get("GITHUB_TOKEN", "")
    if not (owner and repo and token):
        raise ValueError("Missing GitHub secrets: GITHUB_OWNER/GITHUB_REPO/GITHUB_BRANCH/GITHUB_TOKEN")
    return owner, repo, branch, token

def gh_read_text(path: str) -> tuple[str, str | None]:
    owner, repo, branch, token = _gh_cfg()
    api = f"https://api.github.com/repos/{owner}/{repo}/contents/{path}"
    headers = {"Authorization": f"token {token}", "Accept": "application/vnd.github+json"}
    r = requests.get(api, headers=headers, params={"ref": branch}, timeout=30)
    if r.status_code == 200:
        sha = r.json()["sha"]
        content_b64 = r.json().get("content", "")
        content = base64.b64decode(content_b64).decode("utf-8", errors="ignore")
        return content, sha
    if r.status_code == 404:
        return "", None
    raise ValueError(f"GitHub read failed: {r.status_code} {r.text}")

def gh_write_text(path: str, new_content: str, message: str, sha: str | None):
    owner, repo, branch, token = _gh_cfg()
    api = f"https://api.github.com/repos/{owner}/{repo}/contents/{path}"
    headers = {"Authorization": f"token {token}", "Accept": "application/vnd.github+json"}
    payload = {
        "message": message,
        "content": base64.b64encode(new_content.encode("utf-8")).decode("utf-8"),
        "branch": branch,
    }
    if sha:
        payload["sha"] = sha
    u = requests.put(api, headers=headers, json=payload, timeout=30)
    if u.status_code not in (200, 201):
        raise ValueError(f"GitHub write failed: {u.status_code} {u.text}")

def load_mapping_from_github() -> dict:
    content, _ = gh_read_text(MAPPING_PATH)
    return json.loads(content) if content.strip() else {}

def save_mapping_to_github(mapping: dict):
    content, sha = gh_read_text(MAPPING_PATH)
    new_content = json.dumps(mapping, indent=2, ensure_ascii=False) + "\n"
    gh_write_text(MAPPING_PATH, new_content, "update destination mapping", sha)

def append_log_to_github(line_obj: dict):
    content, sha = gh_read_text(LOG_PATH)
    new_line = json.dumps(line_obj, ensure_ascii=False)
    new_content = (content.rstrip("\n") + "\n" + new_line + "\n") if content.strip() else (new_line + "\n")
    gh_write_text(LOG_PATH, new_content, f"log price sent {line_obj.get('date')}", sha)

def tail_logs_from_github(n=20) -> list[dict]:
    content, _ = gh_read_text(LOG_PATH)
    lines = [x.strip() for x in content.splitlines() if x.strip()]
    out = []
    for line in lines[-n:]:
        try:
            out.append(json.loads(line))
        except:
            pass
    return out
# ---------------------------------------------------


# ----------------- CORE DATA -----------------
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
        raise ValueError("Excel must contain: Destination, Product, and For (or Ex Price/Freight/Margin/GST).")

    df["Destination"] = df["Destination"].astype(str).str.strip()
    df["Product"] = df["Product"].astype(str).str.strip()
    df["DeliveredPrice"] = pd.to_numeric(df["DeliveredPrice"], errors="coerce")
    df = df.dropna(subset=["DeliveredPrice"])

    # de-dupe: cheapest per destination+product
    df = df.groupby(["Destination", "Product"], as_index=False)["DeliveredPrice"].min()
    return df
# ---------------------------------------------------


# ----------------- IMAGE HELPERS -----------------
def load_template():
    if not os.path.exists(TEMPLATE_PATH):
        raise FileNotFoundError(f"Template not found: {TEMPLATE_PATH}")
    img = Image.open(TEMPLATE_PATH).convert("RGB")
    if img.size != (CANVAS_W, CANVAS_H):
        img = img.resize((CANVAS_W, CANVAS_H))
    return img

def ellipsize(draw, text, font, max_width):
    if draw.textlength(text, font=font) <= max_width:
        return text
    s = text
    while s and draw.textlength(s + "…", font=font) > max_width:
        s = s[:-1]
    return (s + "…") if s else "…"

def build_poster(destination: str, date_str: str, rows: pd.DataFrame) -> bytes:
    img = load_template()
    draw = ImageDraw.Draw(img)

    # Title line (dest left, date right) on safe zone
    title = f"TODAY PRICES — {destination.upper()}"
    draw.text((TABLE_LEFT, 160), title, font=H2, fill=BLACK)
    dw = draw.textlength(date_str, font=H3)
    draw.text((TABLE_RIGHT - dw, 160), date_str, font=H3, fill=BLACK)

    # Table header
    y = TABLE_TOP
    draw.rectangle((TABLE_LEFT, y, TABLE_RIGHT, y + ROW_H), fill=WHITE)
    draw.text((TABLE_LEFT + 12, y + 18), "PRODUCT", font=H3, fill=BLACK)
    price_label = "PRICE (₹/KG)"
    plw = draw.textlength(price_label, font=H3)
    draw.text((TABLE_RIGHT - 12 - plw, y + 18), price_label, font=H3, fill=BLACK)
    draw.line((TABLE_LEFT, y + ROW_H, TABLE_RIGHT, y + ROW_H), fill=GOLD, width=2)
    y += ROW_H + 6

    # Rows (no overlap guaranteed)
    max_rows_area_bottom = CANVAS_H - FOOTER_H - 30
    max_product_width = COL_PRODUCT_W - 24
    for i, r in enumerate(rows.itertuples(index=False)):
        if y + ROW_H > max_rows_area_bottom:
            break

        if i % 2 == 0:
            draw.rectangle((TABLE_LEFT, y, TABLE_RIGHT, y + ROW_H), fill=ROW_BG)

        product = ellipsize(draw, str(r.Product).strip(), BODY_B, max_product_width)
        draw.text((TABLE_LEFT + 12, y + 18), product, font=BODY_B, fill=BLACK)

        price = f"₹ {float(r.DeliveredPrice):.2f}"
        pw = draw.textlength(price, font=BODY)
        draw.text((TABLE_RIGHT - 12 - pw, y + 18), price, font=BODY, fill=BLACK)

        y += ROW_H

    # Footer (fixed)
    fy = CANVAS_H - FOOTER_H + 5
    draw.text((CANVAS_W / 2, fy + 15),
              "✔ Consistent Quality   ✔ Reliable Supply   ✔ Transparent Pricing",
              font=SMALL, fill=RED, anchor="mm")
    draw.text((CANVAS_W / 2, fy + 50), DISCLAIMER, font=SMALL, fill=GRAY, anchor="mm")
    draw.text((CANVAS_W / 2, fy + 82), EMAIL_LINE, font=SMALL, fill=BLACK, anchor="mm")

    buf = io.BytesIO()
    img.save(buf, format="PNG")
    buf.seek(0)
    return buf.getvalue()

def build_master_table(date_str: str, df_sel: pd.DataFrame) -> bytes:
    """One master image: destination blocks + rows."""
    img = load_template()
    draw = ImageDraw.Draw(img)

    title = "MASTER PRICE UPDATE"
    draw.text((TABLE_LEFT, 160), title, font=H2, fill=BLACK)
    dw = draw.textlength(date_str, font=H3)
    draw.text((TABLE_RIGHT - dw, 160), date_str, font=H3, fill=BLACK)

    y = TABLE_TOP
    draw.rectangle((TABLE_LEFT, y, TABLE_RIGHT, y + ROW_H), fill=WHITE)
    draw.text((TABLE_LEFT + 12, y + 18), "DESTINATION / PRODUCT", font=H3, fill=BLACK)
    price_label = "PRICE (₹/KG)"
    plw = draw.textlength(price_label, font=H3)
    draw.text((TABLE_RIGHT - 12 - plw, y + 18), price_label, font=H3, fill=BLACK)
    draw.line((TABLE_LEFT, y + ROW_H, TABLE_RIGHT, y + ROW_H), fill=GOLD, width=2)
    y += ROW_H + 6

    max_rows_area_bottom = CANVAS_H - FOOTER_H - 30
    max_text_width = (TABLE_RIGHT - TABLE_LEFT) - 220

    for dest in sorted(df_sel["Destination"].unique().tolist()):
        if y + ROW_H > max_rows_area_bottom:
            break

        # Destination header row
        draw.rectangle((TABLE_LEFT, y, TABLE_RIGHT, y + ROW_H), fill=(255, 255, 255))
        draw.text((TABLE_LEFT + 12, y + 18), f"📍 {dest}", font=BODY_B, fill=BLACK)
        y += ROW_H

        rows = df_sel[df_sel["Destination"] == dest][["Product", "DeliveredPrice"]].sort_values("Product")
        for i, r in enumerate(rows.itertuples(index=False)):
            if y + ROW_H > max_rows_area_bottom:
                break
            if i % 2 == 0:
                draw.rectangle((TABLE_LEFT, y, TABLE_RIGHT, y + ROW_H), fill=ROW_BG)

            text = ellipsize(draw, f"   {r.Product}", BODY, max_text_width)
            draw.text((TABLE_LEFT + 12, y + 18), text, font=BODY, fill=BLACK)

            price = f"₹ {float(r.DeliveredPrice):.2f}"
            pw = draw.textlength(price, font=BODY)
            draw.text((TABLE_RIGHT - 12 - pw, y + 18), price, font=BODY, fill=BLACK)
            y += ROW_H

        y += 6

    # Footer
    fy = CANVAS_H - FOOTER_H + 5
    draw.text((CANVAS_W / 2, fy + 15),
              "✔ Consistent Quality   ✔ Reliable Supply   ✔ Transparent Pricing",
              font=SMALL, fill=RED, anchor="mm")
    draw.text((CANVAS_W / 2, fy + 50), DISCLAIMER, font=SMALL, fill=GRAY, anchor="mm")
    draw.text((CANVAS_W / 2, fy + 82), EMAIL_LINE, font=SMALL, fill=BLACK, anchor="mm")

    buf = io.BytesIO()
    img.save(buf, format="PNG")
    buf.seek(0)
    return buf.getvalue()
# ---------------------------------------------------


# ----------------- EMAIL -----------------
def send_email_with_attachments(to_list, subject, body, attachments: dict):
    host = st.secrets.get("SMTP_HOST", "")
    port = int(st.secrets.get("SMTP_PORT", 465))
    user = st.secrets.get("SMTP_USER", "")
    pwd = st.secrets.get("SMTP_PASS", "")
    if not (host and user and pwd):
        raise ValueError("SMTP secrets missing in Streamlit Secrets.")

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
# ---------------------------------------------------


# ----------------- STREAMLIT UI -----------------
st.set_page_config(page_title="ISOCH Price Poster Generator", layout="centered")
st.title("ISOCH Price Poster Generator")

uploaded = st.file_uploader("Upload Excel (.xlsx)", type=["xlsx"])

# load mapping + logs (GitHub)
mapping = {}
mapping_ok = False
log_ok = False
mapping_error = None
log_error = None

try:
    mapping = load_mapping_from_github()
    mapping_ok = True
except Exception as e:
    mapping_error = str(e)

try:
    _ = tail_logs_from_github(1)
    log_ok = True
except Exception as e:
    log_error = str(e)

if uploaded:
    df = parse_excel(uploaded.read())
    destinations = sorted(df["Destination"].unique().tolist())
    date_str = datetime.now().strftime("%d-%m-%Y")

    st.markdown("### Select destinations")
    sel = st.multiselect("Destinations", destinations, default=destinations[:1] if destinations else [])

    st.markdown("### Mapping: Destination → Sales Recipients")
    if mapping_ok:
        st.success("Mapping loaded from GitHub.")
    else:
        st.warning(f"Mapping not writable/available (email+images still work): {mapping_error}")

    with st.expander("Edit mapping"):
        if destinations:
            d = st.selectbox("Destination to edit", destinations)
            current = mapping.get(d, DEFAULT_SALES_TEAM)
            new_val = st.text_area("Recipients (comma-separated)", value=", ".join(current), height=90)
            if st.button("Save mapping"):
                mapping[d] = [x.strip() for x in new_val.split(",") if x.strip()]
                try:
                    save_mapping_to_github(mapping)
                    st.success("Saved mapping to GitHub.")
                except Exception as e:
                    st.error(f"Could not save mapping: {e}")

    st.markdown("### WhatsApp Caption")
    caption = (
        f"📢 *ISOCH – Today’s Delivered Prices* ({date_str})\n"
        f"📍 " + (" | ".join(sel) if sel else "(select destinations)") + "\n\n"
        f"{DISCLAIMER}\n"
        f"{EMAIL_LINE}"
    )
    st.code(caption)

    with st.expander("Last logs"):
        if log_ok:
            try:
                st.json(tail_logs_from_github(15))
            except Exception as e:
                st.write(f"Could not load logs: {e}")
        else:
            st.warning(f"Logs not available: {log_error}")

    st.markdown("---")
    send_mode = st.radio(
        "Email sending mode",
        ["Send ONE email to everyone (combined recipients)", "Send destination-wise to mapped recipients"],
        index=1
    )

    if st.button("Generate + Email + Downloads", type="primary"):
        if not sel:
            st.warning("Select at least one destination.")
            st.stop()

        df_sel = df[df["Destination"].isin(sel)].copy()

        # Build outputs
        outputs = {}
        master_png = build_master_table(date_str, df_sel)
        outputs[f"ISOCH_MASTER_{date_str}.png"] = master_png

        for d in sel:
            rows = df_sel[df_sel["Destination"] == d][["Product", "DeliveredPrice"]].sort_values("Product")
            outputs[f"ISOCH_{d}_{date_str}.png"] = build_poster(d, date_str, rows)

        # Email sending
        subject = f"ISOCH | Delivered Prices | {date_str}"
        body = (
            "Dear Team,\n\n"
            "Please find attached today's delivered price update.\n\n"
            f"{DISCLAIMER}\n\n"
            "Regards,\nInnovative Soch\n"
        )

        if send_mode.startswith("Send ONE"):
            # union recipients across selected destinations
            recips = set()
            for d in sel:
                for r in mapping.get(d, DEFAULT_SALES_TEAM):
                    recips.add(r)
            to_list = sorted(recips) if recips else DEFAULT_SALES_TEAM
            send_email_with_attachments(to_list, subject, body, outputs)

            # log
            try:
                append_log_to_github({
                    "date": date.today().isoformat(),
                    "time": datetime.now().strftime("%H:%M"),
                    "mode": "one_email_union",
                    "destinations": sel,
                    "recipients": to_list,
                    "files": list(outputs.keys()),
                })
            except Exception as e:
                st.warning(f"Email sent, but log failed: {e}")

            st.success(f"✅ Email sent to {len(to_list)} recipients.")

        else:
            # destination wise
            total_sent = 0
            for d in sel:
                to_list = mapping.get(d, DEFAULT_SALES_TEAM)
                attach = {
                    f"ISOCH_{d}_{date_str}.png": outputs[f"ISOCH_{d}_{date_str}.png"],
                    f"ISOCH_MASTER_{date_str}.png": outputs[f"ISOCH_MASTER_{date_str}.png"],
                }
                send_email_with_attachments(to_list, subject, body, attach)
                total_sent += 1

            try:
                append_log_to_github({
                    "date": date.today().isoformat(),
                    "time": datetime.now().strftime("%H:%M"),
                    "mode": "destination_wise",
                    "destinations": sel,
                    "mapping_used": {d: mapping.get(d, DEFAULT_SALES_TEAM) for d in sel},
                    "files": list(outputs.keys()),
                })
            except Exception as e:
                st.warning(f"Emails sent, but log failed: {e}")

            st.success(f"✅ Sent {total_sent} destination-wise emails.")

        # Downloads (NO ZIP)
        st.markdown("## Downloads")
        st.download_button(
            "⬇️ Download MASTER Table",
            data=master_png,
            file_name=f"ISOCH_MASTER_{date_str}.png",
            mime="image/png"
        )
        for d in sel:
            fname = f"ISOCH_{d}_{date_str}.png"
            st.download_button(
                f"⬇️ Download {d} Poster",
                data=outputs[fname],
                file_name=fname,
                mime="image/png",
                key=f"dl_{fname}",
            )

else:
    st.info("Upload your Excel to start.")
# ---------------------------------------------------
