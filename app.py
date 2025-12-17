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

# ---------------- CONFIG ----------------
CANVAS_W, CANVAS_H = 1080, 1920
COMPANY_NAME = "INNOVATIVE SOCH"
TEMPLATE_PATH = "template_light.png"

MAPPING_PATH = "config/destination_mapping.json"
LOG_PATH = "logs/price_sent_log.jsonl"

DEFAULT_TEAM = [
    "info@innovativesoch.com",
    "sapnarani@innovativesoch.com",
    "sandeep@innovativesoch.com",
    "sales@innovativesoch.com",
    "purchase@innovativesoch.com",
]

DISCLAIMER = "Prices are indicative. Final order confirmation by call."
# ----------------------------------------


# ========== helpers ==========
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


# ========== GitHub read/write ==========
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
    if not content.strip():
        return {}
    return json.loads(content)

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


# ========== Image generation ==========
def build_destination_poster(template_img, destination, rows, date_str, email_line) -> bytes:
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

    y = 210
    draw.text((pad, y), f"TODAY PRICES — {destination.upper()}", fill=DARK, font=TITLE)
    y += 62
    w = draw.textlength(date_str, font=H3)
    draw.text((CANVAS_W - pad - w, y), date_str, fill=DARK, font=H3)

    table_left, table_right = pad, CANVAS_W - pad
    table_top = 410

    y = table_top
    draw.line((table_left, y, table_right, y), fill=LINE, width=3)
    y += 16
    draw.text((table_left + 10, y), "ITEM", fill=DARK, font=H2)
    draw.text((table_right - 270, y), "PRICE (₹/KG)", fill=DARK, font=H2)
    y += 54
    draw.line((table_left, y, table_right, y), fill=LINE, width=2)

    row_h = 72
    rows = rows.head(14).copy()

    for i, r in rows.iterrows():
        y += 14
        product = str(r["Product"]).strip()
        price_text = f"₹ {money(r['DeliveredPrice'])}"

        if (i % 2) == 0:
            draw.rectangle((table_left, y-8, table_right, y + row_h - 8), fill=(255, 255, 255))

        draw.text((table_left + 10, y + 14), product, fill=DARK, font=BODY_B)
        pw = draw.textlength(price_text, font=BODY)
        draw.text((table_right - 10 - pw, y + 16), price_text, fill=DARK, font=BODY)

        draw.line((table_left, y + row_h, table_right, y + row_h), fill=LINE, width=1)
        y += row_h - 8

    fy = CANVAS_H - 235
    for line in wrap_text(draw, DISCLAIMER, SMALL, CANVAS_W - 2*pad)[:2]:
        draw.text((pad, fy), line, fill=DARK, font=SMALL)
        fy += 30
    fy += 6
    for line in wrap_text(draw, email_line, SMALL, CANVAS_W - 2*pad)[:3]:
        draw.text((pad, fy), line, fill=DARK, font=SMALL)
        fy += 30

    buf = io.BytesIO()
    img.save(buf, format="PNG")
    buf.seek(0)
    return buf.getvalue()

def build_master_table_image(template_img, df_selected: pd.DataFrame, date_str: str) -> bytes:
    img = template_img.copy()
    draw = ImageDraw.Draw(img)

    font_paths = [
        "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf",
        "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
    ]
    TITLE = try_font(font_paths, 46)
    H2 = try_font(font_paths, 32)
    H3 = try_font(font_paths, 28)
    BODY = try_font(font_paths, 26)
    BODY_B = try_font(font_paths, 26)
    SMALL = try_font(font_paths, 22)

    DARK = (35, 28, 24)
    LINE = (150, 135, 125)
    pad = 90

    y = 200
    draw.text((pad, y), "MASTER PRICE UPDATE", fill=DARK, font=TITLE)
    w = draw.textlength(date_str, font=H3)
    draw.text((CANVAS_W - pad - w, y + 10), date_str, fill=DARK, font=H3)

    y += 70
    draw.line((pad, y, CANVAS_W - pad, y), fill=LINE, width=3)
    y += 22

    left, right = pad, CANVAS_W - pad
    max_y = CANVAS_H - 320
    truncated = False

    for dest in sorted(df_selected["Destination"].unique().tolist()):
        if y + 60 > max_y:
            truncated = True
            break

        draw.text((left, y), dest.upper(), fill=DARK, font=H2)
        y += 42
        draw.line((left, y, right, y), fill=LINE, width=2)
        y += 12

        g = df_selected[df_selected["Destination"] == dest][["Product", "DeliveredPrice"]].copy()
        g = g.groupby("Product", as_index=False)["DeliveredPrice"].min().sort_values("Product")

        for _, r in g.iterrows():
            if y + 40 > max_y:
                truncated = True
                break
            product = str(r["Product"]).strip()
            price = f"₹ {money(r['DeliveredPrice'])}"

            prod_lines = wrap_text(draw, product, BODY_B, (right - left - 260))
            prod_line = prod_lines[0] if prod_lines else product

            draw.text((left, y), prod_line, fill=DARK, font=BODY_B)
            pw = draw.textlength(price, font=BODY)
            draw.text((right - pw, y), price, fill=DARK, font=BODY)
            y += 34

        y += 14
        if truncated:
            break

    if truncated:
        draw.text((left, y), "... (More rows truncated)", fill=DARK, font=SMALL)

    fy = CANVAS_H - 240
    draw.line((pad, fy - 20, CANVAS_W - pad, fy - 20), fill=LINE, width=2)
    for line in wrap_text(draw, DISCLAIMER, SMALL, CANVAS_W - 2*pad)[:2]:
        draw.text((pad, fy), line, fill=DARK, font=SMALL)
        fy += 28

    buf = io.BytesIO()
    img.save(buf, format="PNG")
    buf.seek(0)
    return buf.getvalue()


# ========== Email ==========
def send_email_with_attachments(to_list, subject, body, attachments: dict):
    host = st.secrets.get("SMTP_HOST", "")
    port = int(st.secrets.get("SMTP_PORT", 465))
    user = st.secrets.get("SMTP_USER", "")
    pwd = st.secrets.get("SMTP_PASS", "")

    if not (host and user and pwd):
        raise ValueError("SMTP secrets missing in Streamlit Settings → Secrets.")

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


# ---------------- UI ----------------
st.set_page_config(page_title="ISOCH Price Poster Generator", layout="centered")
st.markdown("## ISOCH Price Poster Generator (Purchase → Sales)")
st.caption("Upload Excel → Map destinations → Generate Master Table + Posters → Email + GitHub Log")

# Load mapping (GitHub)
mapping = {}
try:
    mapping = load_mapping_from_github()
except Exception as e:
    st.warning(f"Mapping load warning: {e} (You can still generate/send; mapping edits need GitHub write access.)")

excel = st.file_uploader("Upload Excel (.xlsx)", type=["xlsx"])
df = None
destinations = []
if excel:
    try:
        df = parse_excel(excel.read())
        destinations = sorted(df["Destination"].unique().tolist())
    except Exception as e:
        st.error(f"Excel error: {e}")

selected_destinations = st.multiselect("Select Destinations (multi)", destinations, default=destinations[:1] if destinations else [])

# Mapping UI
with st.expander("⚙️ Destination → Sales Recipients Mapping"):
    st.info("Enter comma-separated emails for each destination. Saved in GitHub (solid persistence).")
    if destinations:
        d = st.selectbox("Destination", destinations)
        current = mapping.get(d, DEFAULT_TEAM)
        new_val = st.text_area("Recipients (comma-separated)", value=", ".join(current), height=90)
        if st.button("Save Mapping to GitHub"):
            try:
                mapping[d] = [x.strip() for x in new_val.split(",") if x.strip()]
                save_mapping_to_github(mapping)
                st.success("Saved mapping to GitHub.")
            except Exception as e:
                st.error(f"Could not save mapping: {e}")
    else:
        st.write("Upload Excel to load destinations first.")

# WhatsApp caption
today_str = datetime.now().strftime("%d-%m-%Y")
whatsapp_text = (
    f"📢 *ISOCH – Today’s Delivered Prices* ({today_str})\n"
    f"📍 " + (" | ".join(selected_destinations) if selected_destinations else "(select destinations)") + "\n\n"
    f"{DISCLAIMER}\n"
    f"📧 info@innovativesoch.com"
)
st.markdown("### WhatsApp Caption")
st.code(whatsapp_text)

# Show last logs
with st.expander("🧾 Last Sent Logs (from GitHub)"):
    try:
        logs = tail_logs_from_github(20)
        if not logs:
            st.write("No logs yet.")
        else:
            st.json(logs)
    except Exception as e:
        st.write(f"Could not load logs: {e}")

if st.button("🛠️ Generate & Send", type="primary"):
    if df is None or not selected_destinations:
        st.warning("Upload Excel and select at least one destination.")
        st.stop()

    try:
        template = load_template()
        df_sel = df[df["Destination"].isin(selected_destinations)].copy()

        # Build Master + posters
        master_png = build_master_table_image(template, df_sel, today_str)

        attachments = {f"ISOCH_MASTER_{today_str}.png": master_png}

        # Collect recipients based on mapping
        all_recipients = set()
        for dest in selected_destinations:
            recips = mapping.get(dest, DEFAULT_TEAM)
            for r in recips:
                all_recipients.add(r)

            g = df_sel[df_sel["Destination"] == dest][["Product", "DeliveredPrice"]].copy()
            g = g.groupby("Product", as_index=False)["DeliveredPrice"].min().sort_values("Product")
            email_line = "Email: " + " | ".join(sorted(set(recips)))
            attachments[f"ISOCH_{dest}_{today_str}.png"] = build_destination_poster(template, dest, g, today_str, email_line)

        to_list = sorted(all_recipients) if all_recipients else DEFAULT_TEAM

        subject = f"ISOCH | Delivered Prices | {today_str}"
        body = (
            "Dear Team,\n\n"
            "Please find attached today's price update:\n"
            "- MASTER table (all selected destinations)\n"
            "- Destination posters\n\n"
            f"{DISCLAIMER}\n\n"
            "Regards,\nInnovative Soch\n"
        )

        send_email_with_attachments(to_list, subject, body, attachments)

        # GitHub log (after successful send)
        try:
            append_log_to_github({
                "date": date.today().isoformat(),
                "time": datetime.now().strftime("%H:%M"),
                "by": "purchase",
                "destinations": selected_destinations,
                "recipients": to_list,
            })
        except Exception as e:
            st.warning(f"Email sent, but logging failed: {e}")

        st.success("✅ Sent successfully.")

        st.download_button(
            "⬇️ Download MASTER Table (PNG)",
            data=master_png,
            file_name=f"ISOCH_MASTER_{today_str}.png",
            mime="image/png",
        )

        st.markdown("### Destination Posters (Download)")
        for dest in selected_destinations:
            fname = f"ISOCH_{dest}_{today_str}.png"
            st.download_button(
                f"⬇️ Download {dest} Poster",
                data=attachments[fname],
                file_name=fname,
                mime="image/png",
                key=f"dl_{dest}",
            )

    except Exception as e:
        st.error(f"Error: {e}")
