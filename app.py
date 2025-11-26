import os
import csv
import time
import threading
import imaplib
import smtplib
import email
from email.header import decode_header
from email.mime.text import MIMEText
import ssl
import traceback
from datetime import datetime

import customtkinter as ctk
from tkinter import messagebox
import keyring
from fpdf import FPDF

# ----------------- CONSTANTS & PATHS -----------------

APP_NAME = "EmailAssistantPro"
LOG_DIR = "logs"
ATTACH_DIR = "attachments"
os.makedirs(LOG_DIR, exist_ok=True)
os.makedirs(ATTACH_DIR, exist_ok=True)

LOG_CSV_PATH = os.path.join(LOG_DIR, "email_log.csv")

CATEGORIES = [
    "Billing / Payment",
    "Order / Purchase",
    "Support Request",
    "Client Lead",
    "Other",
]


# ----------------- CLASSIFIER -----------------

def classify_email(subject: str, body: str):
    text = (subject or "") + " " + (body or "")
    text = text.lower()

    billing = ["invoice", "payment", "bill", "billing", "due", "overdue"]
    order = ["order", "shipment", "tracking", "delivery", "purchase"]
    support = ["issue", "error", "bug", "problem", "support", "help", "not working", "failed"]
    lead = ["quote", "pricing", "project", "proposal", "hire", "collaboration", "work with you", "service"]
    urgent = ["urgent", "asap", "immediately", "critical", "important"]

    if any(k in text for k in billing):
        category = "Billing / Payment"
    elif any(k in text for k in order):
        category = "Order / Purchase"
    elif any(k in text for k in support):
        category = "Support Request"
    elif any(k in text for k in lead):
        category = "Client Lead"
    else:
        category = "Other"

    is_urgent = any(k in text for k in urgent)
    return category, is_urgent


def build_reply(category: str, sender_name: str | None = None):
    sender_name = sender_name or "there"

    templates = {
        "Billing / Payment": f"""Hi {sender_name},

Thank you for your message about billing/payment. We have received your request and will review the details shortly.

If this is about a specific invoice, please include the invoice number or date.

Best regards,
[Your Name]
""",
        "Order / Purchase": f"""Hi {sender_name},

Thank you for contacting us about your order.

We will check the order status and get back to you. If you have an order ID or tracking number, please include it.

Best regards,
[Your Name]
""",
        "Support Request": f"""Hi {sender_name},

Thank you for reaching out! We have received your support request.

We will review the issue and respond with an update as soon as possible.

Best regards,
[Your Name]
""",
        "Client Lead": f"""Hi {sender_name},

Thank you for your interest!

Please share some details about your requirement (scope, timeline, and budget) so we can suggest the best next steps.

Best regards,
[Your Name]
""",
    }

    return templates.get(category, f"""Hi {sender_name},

Thank you for your email. We have received your message and will look into it shortly.

Best regards,
[Your Name]
""")


# ----------------- UTILS -----------------

def decode_str(s):
    if not s:
        return ""
    parts = decode_header(s)
    out = []
    for text, enc in parts:
        if isinstance(text, bytes):
            try:
                out.append(text.decode(enc or "utf-8", errors="ignore"))
            except Exception:
                out.append(text.decode("utf-8", errors="ignore"))
        else:
            out.append(text)
    return " ".join(out)


def extract_body(msg: email.message.Message):
    if msg.is_multipart():
        for part in msg.walk():
            ctype = part.get_content_type()
            disp = (part.get("Content-Disposition") or "").lower()
            if ctype == "text/plain" and "attachment" not in disp:
                try:
                    return part.get_payload(decode=True).decode(errors="ignore")
                except Exception:
                    continue
    else:
        if msg.get_content_type() == "text/plain":
            try:
                return msg.get_payload(decode=True).decode(errors="ignore")
            except Exception:
                pass
    return ""


def save_attachments(msg: email.message.Message, uid: str):
    saved = []
    folder = os.path.join(ATTACH_DIR, f"email_{uid}")
    os.makedirs(folder, exist_ok=True)

    if msg.is_multipart():
        for part in msg.walk():
            disp = (part.get("Content-Disposition") or "").lower()
            if "attachment" in disp:
                filename = part.get_filename()
                filename = decode_str(filename) if filename else f"file_{int(time.time())}.bin"
                path = os.path.join(folder, filename)
                try:
                    with open(path, "wb") as f:
                        f.write(part.get_payload(decode=True))
                    saved.append(path)
                except Exception:
                    continue
    return saved


def ensure_log_csv():
    if not os.path.exists(LOG_CSV_PATH):
        with open(LOG_CSV_PATH, "w", newline="", encoding="utf-8") as f:
            w = csv.writer(f)
            w.writerow([
                "timestamp",
                "from",
                "subject",
                "category",
                "urgent",
                "attachments",
                "mode",  # manual/auto
            ])


# ----------------- PDF LOG SUMMARY -----------------

def generate_pdf_log_summary(csv_path: str, pdf_path: str):
    if not os.path.exists(csv_path):
        raise FileNotFoundError("No log CSV found.")

    counts = {c: 0 for c in CATEGORIES}
    urgent_count = 0
    total = 0

    with open(csv_path, "r", encoding="utf-8") as f:
        r = csv.DictReader(f)
        for row in r:
            total += 1
            cat = row.get("category", "Other")
            if cat not in counts:
                counts[cat] = 0
            counts[cat] += 1
            if row.get("urgent", "").lower() in ("1", "true", "yes"):
                urgent_count += 1

    pdf = FPDF()
    pdf.add_page()
    pdf.set_font("Arial", "B", 14)
    pdf.cell(0, 10, "Email Log Summary", 0, 1, "C")
    pdf.ln(4)

    pdf.set_font("Arial", "", 11)
    pdf.cell(0, 8, f"Total logged replies: {total}", 0, 1)
    pdf.cell(0, 8, f"Urgent emails: {urgent_count}", 0, 1)
    pdf.ln(4)

    pdf.set_font("Arial", "B", 12)
    pdf.cell(0, 8, "Category breakdown:", 0, 1)
    pdf.set_font("Arial", "", 11)

    for cat, n in counts.items():
        pdf.cell(0, 7, f"{cat}: {n}", 0, 1)

    pdf.output(pdf_path)


# ----------------- PROGRESS POPUP -----------------

class LoadingPopup:
    def __init__(self, parent, title="Processing", status="Please wait..."):
        self.win = ctk.CTkToplevel(parent)
        self.win.title(title)
        self.win.geometry("320x140")
        self.win.resizable(False, False)
        self.win.grab_set()

        self.label = ctk.CTkLabel(self.win, text=status, font=ctk.CTkFont(size=14, weight="bold"))
        self.label.pack(pady=(15, 5))

        self.bar = ctk.CTkProgressBar(self.win, width=260)
        self.bar.pack(pady=(0, 10))
        self.bar.set(0)

        self.percent = ctk.CTkLabel(self.win, text="0%", font=ctk.CTkFont(size=12))
        self.percent.pack()

        self.win.update()

    def set(self, value: float, text: str | None = None):
        if value < 0:
            value = 0
        if value > 1:
            value = 1
        self.bar.set(value)
        self.percent.configure(text=f"{int(value * 100)}%")
        if text:
            self.label.configure(text=text)
        self.win.update_idletasks()

    def close(self):
        try:
            self.win.grab_release()
            self.win.destroy()
        except Exception:
            pass


# ----------------- MAIN APP -----------------

class EmailAssistantPro:
    def __init__(self, root):
        self.root = root
        self.root.title("Email Assistant Pro")
        self.root.geometry("1150x650")

        ctk.set_appearance_mode("System")
        ctk.set_default_color_theme("blue")

        self.imap_conn = None
        self.smtp_conn = None

        self.emails = []  # list of dicts with keys: sub, from, body, date, uid, cat, urgent, attachments, replied
        self.selected_index = None

        self.auto_check_enabled = False
        self.auto_check_interval_min = 5

        ensure_log_csv()
        self._build_ui()

        # Try autofill password when email field loses focus
        self.entry_email.bind("<FocusOut>", self.autofill_password)

    # ------------- UI -------------
    def _build_ui(self):
        main = ctk.CTkFrame(self.root, corner_radius=10)
        main.pack(fill="both", expand=True, padx=10, pady=10)

        # Top connection panel
        top = ctk.CTkFrame(main)
        top.pack(fill="x", padx=8, pady=(8, 4))

        self.entry_email = ctk.CTkEntry(top, placeholder_text="Email address", width=220)
        self.entry_email.grid(row=0, column=0, padx=5, pady=4)

        self.entry_pass = ctk.CTkEntry(top, placeholder_text="App password", show="*", width=220)
        self.entry_pass.grid(row=0, column=1, padx=5, pady=4)

        self.entry_imap = ctk.CTkEntry(top, placeholder_text="IMAP server (e.g. imap.gmail.com)", width=220)
        self.entry_imap.grid(row=0, column=2, padx=5, pady=4)

        self.entry_smtp = ctk.CTkEntry(top, placeholder_text="SMTP server (e.g. smtp.gmail.com)", width=220)
        self.entry_smtp.grid(row=0, column=3, padx=5, pady=4)

        self.entry_imap_port = ctk.CTkEntry(top, placeholder_text="993", width=80)
        self.entry_imap_port.grid(row=1, column=0, padx=5, pady=4)

        self.entry_smtp_port = ctk.CTkEntry(top, placeholder_text="465", width=80)
        self.entry_smtp_port.grid(row=1, column=1, padx=5, pady=4)

        self.var_ssl = ctk.BooleanVar(value=True)
        chk_ssl = ctk.CTkCheckBox(top, text="Use SSL", variable=self.var_ssl)
        chk_ssl.grid(row=1, column=2, padx=5, pady=4, sticky="w")

        btn_connect = ctk.CTkButton(top, text="Connect", width=140,
                                    command=lambda: self.run_async(self.connect_accounts))
        btn_connect.grid(row=1, column=3, padx=5, pady=4)

        # Middle area: left list + center details + right dashboard
        middle = ctk.CTkFrame(main)
        middle.pack(fill="both", expand=True, padx=8, pady=(4, 8))

        # Left: email list & buttons
        left = ctk.CTkFrame(middle, width=380)
        left.pack(side="left", fill="both", expand=False, padx=(0, 8), pady=4)

        lbl_list = ctk.CTkLabel(left, text="Emails", font=ctk.CTkFont(size=14, weight="bold"))
        lbl_list.pack(pady=(6, 4))

        self.text_list = ctk.CTkTextbox(left, width=360, height=360, state="disabled")
        self.text_list.pack(fill="both", expand=True, padx=6, pady=4)

        btn_row = ctk.CTkFrame(left)
        btn_row.pack(fill="x", padx=6, pady=(4, 6))

        btn_fetch = ctk.CTkButton(btn_row, text="Fetch latest",
                                  command=lambda: self.run_async(self.fetch_emails), width=120)
        btn_fetch.pack(side="left", padx=3)

        btn_classify = ctk.CTkButton(btn_row, text="Classify all",
                                     command=self.classify_all, width=100)
        btn_classify.pack(side="left", padx=3)

        btn_reply = ctk.CTkButton(btn_row, text="Auto-reply selected",
                                  command=lambda: self.run_async(self.auto_reply_selected), width=150)
        btn_reply.pack(side="left", padx=3)

        # Center: email details
        center = ctk.CTkFrame(middle)
        center.pack(side="left", fill="both", expand=True, padx=(0, 8), pady=4)

        lbl_detail = ctk.CTkLabel(center, text="Email Details", font=ctk.CTkFont(size=14, weight="bold"))
        lbl_detail.pack(pady=(6, 4))

        self.lbl_subject = ctk.CTkLabel(center, text="Subject: ", anchor="w")
        self.lbl_subject.pack(fill="x", padx=6, pady=2)

        self.lbl_from = ctk.CTkLabel(center, text="From: ", anchor="w")
        self.lbl_from.pack(fill="x", padx=6, pady=2)

        self.lbl_meta = ctk.CTkLabel(center, text="Category:  | Urgency:  | Attachments: 0", anchor="w")
        self.lbl_meta.pack(fill="x", padx=6, pady=(2, 6))

        self.text_body = ctk.CTkTextbox(center, state="disabled", height=260)
        self.text_body.pack(fill="both", expand=True, padx=6, pady=4)

        btn_log_pdf = ctk.CTkButton(center, text="Generate PDF Log Summary",
                                    command=lambda: self.run_async(self.generate_pdf_log))
        btn_log_pdf.pack(pady=(4, 6))

        # Right: dashboard & auto-check
        right = ctk.CTkFrame(middle, width=230)
        right.pack(side="left", fill="y", expand=False, padx=(0, 0), pady=4)

        lbl_dash = ctk.CTkLabel(right, text="Dashboard", font=ctk.CTkFont(size=14, weight="bold"))
        lbl_dash.pack(pady=(6, 4))

        self.lbl_total = ctk.CTkLabel(right, text="Total loaded: 0", anchor="w")
        self.lbl_total.pack(fill="x", padx=6, pady=2)

        self.lbl_replied = ctk.CTkLabel(right, text="Total replied: 0", anchor="w")
        self.lbl_replied.pack(fill="x", padx=6, pady=2)

        self.lbl_urgent = ctk.CTkLabel(right, text="Urgent emails: 0", anchor="w")
        self.lbl_urgent.pack(fill="x", padx=6, pady=2)

        self.lbl_cat_stats = ctk.CTkLabel(right, text="By category:\n", anchor="w", justify="left")
        self.lbl_cat_stats.pack(fill="x", padx=6, pady=(4, 8))

        # Auto-check controls
        sep = ctk.CTkLabel(right, text="────────────", anchor="center")
        sep.pack(pady=(4, 4))

        lbl_auto = ctk.CTkLabel(right, text="Auto-check (fetch+classify)", anchor="w")
        lbl_auto.pack(fill="x", padx=6, pady=(4, 2))

        self.var_auto_check = ctk.BooleanVar(value=False)
        chk_auto = ctk.CTkCheckBox(right, text="Enable auto-check", variable=self.var_auto_check,
                                   command=self.toggle_auto_check)
        chk_auto.pack(padx=6, pady=(2, 2), anchor="w")

        self.entry_interval = ctk.CTkEntry(right, placeholder_text="Interval (min, default 5)", width=180)
        self.entry_interval.pack(padx=6, pady=(2, 4))

        # Status bar
        self.lbl_status = ctk.CTkLabel(main, text="Ready.", anchor="w")
        self.lbl_status.pack(fill="x", padx=8, pady=(0, 4))

        # Bind clicks on list pane
        self.text_list.bind("<Button-1>", self.on_list_click)

    # ------------- GENERAL HELPERS -------------

    def set_status(self, text: str):
        self.lbl_status.configure(text=text)
        self.root.update_idletasks()

    def run_async(self, func, *args, **kwargs):
        t = threading.Thread(target=func, args=args, kwargs=kwargs, daemon=True)
        t.start()

    def autofill_password(self, event=None):
        email_addr = self.entry_email.get().strip()
        if not email_addr:
            return
        stored = keyring.get_password(APP_NAME, email_addr)
        if stored and not self.entry_pass.get().strip():
            self.entry_pass.insert(0, stored)

    # ------------- CONNECTION -------------

    def connect_accounts(self):
        popup = LoadingPopup(self.root, "Connecting", "Starting connection...")
        try:
            email_addr = self.entry_email.get().strip()
            password = self.entry_pass.get().strip()
            imap_host = self.entry_imap.get().strip()
            smtp_host = self.entry_smtp.get().strip()
            imap_port = int(self.entry_imap_port.get().strip() or "993")
            smtp_port = int(self.entry_smtp_port.get().strip() or "465")

            if not (email_addr and imap_host and smtp_host):
                popup.close()
                messagebox.showerror("Missing fields", "Email, IMAP and SMTP are required.")
                return

            if not password:
                stored = keyring.get_password(APP_NAME, email_addr)
                if stored:
                    password = stored
                    self.entry_pass.insert(0, stored)

            if not password:
                popup.close()
                messagebox.showerror("Missing password", "Enter the app password at least once.")
                return

            popup.set(0.2, "Connecting IMAP...")
            if self.var_ssl.get():
                self.imap_conn = imaplib.IMAP4_SSL(imap_host, imap_port)
            else:
                self.imap_conn = imaplib.IMAP4(imap_host, imap_port)
            self.imap_conn.login(email_addr, password)

            popup.set(0.6, "Connecting SMTP...")
            if self.var_ssl.get():
                context = ssl.create_default_context()
                self.smtp_conn = smtplib.SMTP_SSL(smtp_host, smtp_port, context=context)
            else:
                self.smtp_conn = smtplib.SMTP(smtp_host, smtp_port)
                self.smtp_conn.starttls()
            self.smtp_conn.login(email_addr, password)

            keyring.set_password(APP_NAME, email_addr, password)

            popup.set(1.0, "Connected ✓")
            time.sleep(0.4)
            popup.close()
            self.set_status("Connected.")
            messagebox.showinfo("Connected", "IMAP and SMTP connected.\nPassword saved securely.")
        except Exception as e:
            popup.close()
            traceback.print_exc()
            self.set_status("Connection failed.")
            messagebox.showerror("Connection error", str(e))

    # ------------- FETCH EMAILS -------------

    def fetch_emails(self, limit=20):
        if not self.imap_conn:
            messagebox.showerror("Not connected", "Connect before fetching emails.")
            return

        popup = LoadingPopup(self.root, "Fetching", "Reading inbox...")
        try:
            self.set_status("Fetching emails...")
            self.imap_conn.select("INBOX")
            status, data = self.imap_conn.search(None, "ALL")
            if status != "OK":
                popup.close()
                raise RuntimeError("IMAP search failed")

            ids = data[0].split()
            if not ids:
                popup.close()
                self.set_status("No emails found.")
                return

            ids_to_fetch = ids[-limit:]
            self.emails.clear()
            self.text_list.configure(state="normal")
            self.text_list.delete("0.0", "end")

            n = len(ids_to_fetch)
            for idx, mail_id in enumerate(reversed(ids_to_fetch), start=1):
                frac = idx / n
                popup.set(frac * 0.9, f"Fetching {idx}/{n}...")
                status, msg_data = self.imap_conn.fetch(mail_id, "(RFC822)")
                if status != "OK":
                    continue
                raw = msg_data[0][1]
                msg = email.message_from_bytes(raw)

                subject = decode_str(msg.get("Subject", ""))
                from_ = decode_str(msg.get("From", ""))
                date = decode_str(msg.get("Date", ""))
                body = extract_body(msg)
                uid = mail_id.decode() if isinstance(mail_id, bytes) else str(mail_id)

                attach_paths = save_attachments(msg, uid)

                self.emails.append({
                    "uid": uid,
                    "subject": subject,
                    "from": from_,
                    "date": date,
                    "body": body,
                    "category": "Unclassified",
                    "urgent": False,
                    "attachments": attach_paths,
                    "replied": False,
                })

                self.text_list.insert("end", f"[{idx}] {subject} | {from_}\n")

            self.text_list.configure(state="disabled")
            popup.set(1.0, "Done ✓")
            time.sleep(0.4)
            popup.close()
            self.set_status(f"Fetched {len(self.emails)} emails.")
            self.update_dashboard()
        except Exception as e:
            popup.close()
            traceback.print_exc()
            self.set_status("Fetch failed.")
            messagebox.showerror("Fetch error", str(e))

    # ------------- CLASSIFY -------------

    def classify_all(self):
        if not self.emails:
            messagebox.showwarning("No emails", "Fetch emails before classifying.")
            return

        self.set_status("Classifying emails...")
        try:
            for mail in self.emails:
                cat, urg = classify_email(mail["subject"], mail["body"])
                mail["category"] = cat
                mail["urgent"] = urg

            self.text_list.configure(state="normal")
            self.text_list.delete("0.0", "end")
            for idx, mail in enumerate(self.emails, start=1):
                tag = " [URGENT]" if mail["urgent"] else ""
                self.text_list.insert(
                    "end",
                    f"[{idx}] {mail['subject']} | {mail['from']} | {mail['category']}{tag}\n"
                )
            self.text_list.configure(state="disabled")

            self.set_status("Classification complete.")
            self.update_dashboard()
        except Exception as e:
            traceback.print_exc()
            self.set_status("Classification failed.")
            messagebox.showerror("Classification error", str(e))

    # ------------- LIST CLICK -------------

    def on_list_click(self, event):
        index = self.text_list.index(f"@{event.x},{event.y}")
        try:
            line = int(str(index).split(".")[0]) - 1
        except Exception:
            return
        if 0 <= line < len(self.emails):
            self.selected_index = line
            self.show_email_detail(line)

    def show_email_detail(self, idx: int):
        mail = self.emails[idx]
        self.lbl_subject.configure(text=f"Subject: {mail['subject']}")
        self.lbl_from.configure(text=f"From: {mail['from']}")
        urg = "URGENT" if mail["urgent"] else "Normal"
        self.lbl_meta.configure(
            text=f"Category: {mail['category']}   |   Urgency: {urg}   |   Attachments: {len(mail['attachments'])}"
        )

        self.text_body.configure(state="normal")
        self.text_body.delete("0.0", "end")
        self.text_body.insert("0.0", mail["body"] or "(No text content)")
        self.text_body.configure(state="disabled")

    # ------------- AUTO REPLY -------------

    def auto_reply_selected(self, mode="manual"):
        if not self.smtp_conn:
            messagebox.showerror("Not connected", "Connect before sending replies.")
            return
        if self.selected_index is None:
            messagebox.showwarning("No selection", "Click an email from the list first.")
            return

        mail = self.emails[self.selected_index]
        sender_raw = mail["from"]
        if "<" in sender_raw and ">" in sender_raw:
            addr = sender_raw.split("<")[-1].split(">")[0].strip()
            name = sender_raw.split("<")[0].strip().strip('"')
        else:
            addr = sender_raw.strip()
            name = None

        if not addr or "@" not in addr:
            messagebox.showerror("Invalid sender", f"Cannot parse email from: {sender_raw}")
            return

        if mail["category"] == "Unclassified":
            cat, urg = classify_email(mail["subject"], mail["body"])
            mail["category"] = cat
            mail["urgent"] = urg

        category = mail["category"]
        reply_body = build_reply(category, name)
        from_addr = self.entry_email.get().strip()
        if not from_addr:
            messagebox.showerror("Missing from address", "Your email address is missing.")
            return

        popup = LoadingPopup(self.root, "Sending reply", "Preparing message...")
        try:
            msg = MIMEText(reply_body)
            msg["Subject"] = f"Re: {mail['subject'] or ''}"
            msg["From"] = from_addr
            msg["To"] = addr

            popup.set(0.5, "Sending...")
            self.smtp_conn.send_message(msg)

            mail["replied"] = True
            self.log_reply(mail, mode=mode)
            self.update_dashboard()

            popup.set(1.0, "Sent ✓")
            time.sleep(0.4)
            popup.close()
            self.set_status(f"Reply sent to {addr}")
            messagebox.showinfo("Sent", f"Auto reply sent to {addr}\nCategory: {category}")
        except Exception as e:
            popup.close()
            traceback.print_exc()
            self.set_status("Reply failed.")
            messagebox.showerror("Reply error", str(e))

    # ------------- LOGGING -------------

    def log_reply(self, mail: dict, mode: str = "manual"):
        ensure_log_csv()
        with open(LOG_CSV_PATH, "a", newline="", encoding="utf-8") as f:
            w = csv.writer(f)
            w.writerow([
                datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                mail.get("from", ""),
                mail.get("subject", ""),
                mail.get("category", ""),
                "1" if mail.get("urgent") else "0",
                len(mail.get("attachments", [])),
                mode,
            ])

    def generate_pdf_log(self):
        try:
            pdf_path = os.path.join(LOG_DIR, "email_log_summary.pdf")
            generate_pdf_log_summary(LOG_CSV_PATH, pdf_path)
            self.set_status(f"PDF log created: {pdf_path}")
            messagebox.showinfo("PDF created", f"Summary saved to:\n{pdf_path}")
        except Exception as e:
            traceback.print_exc()
            messagebox.showerror("PDF error", str(e))

    # ------------- DASHBOARD -------------

    def update_dashboard(self):
        total = len(self.emails)
        replied = sum(1 for m in self.emails if m.get("replied"))
        urgent = sum(1 for m in self.emails if m.get("urgent"))

        cats = {c: 0 for c in CATEGORIES}
        for m in self.emails:
            c = m.get("category", "Other")
            if c not in cats:
                cats[c] = 0
            cats[c] += 1

        self.lbl_total.configure(text=f"Total loaded: {total}")
        self.lbl_replied.configure(text=f"Total replied: {replied}")
        self.lbl_urgent.configure(text=f"Urgent emails: {urgent}")

        lines = ["By category:"]
        for c in CATEGORIES:
            lines.append(f"- {c}: {cats.get(c, 0)}")
        self.lbl_cat_stats.configure(text="\n".join(lines))

    # ------------- AUTO-CHECK -------------

    def toggle_auto_check(self):
        self.auto_check_enabled = self.var_auto_check.get()
        if self.auto_check_enabled:
            try:
                interval = int(self.entry_interval.get().strip() or "5")
                if interval < 1:
                    interval = 5
            except ValueError:
                interval = 5
            self.auto_check_interval_min = interval
            self.set_status(f"Auto-check enabled ({interval} min).")
            self.schedule_auto_check()
        else:
            self.set_status("Auto-check disabled.")

    def schedule_auto_check(self):
        if not self.auto_check_enabled:
            return
        # schedule next run
        self.root.after(self.auto_check_interval_min * 60 * 1000,
                        lambda: self.run_async(self.auto_check_cycle))

    def auto_check_cycle(self):
        if not self.auto_check_enabled:
            return
        try:
            # Silent fetch + classify (no popups, just status)
            self.set_status("Auto-check: fetching + classifying...")
            self.fetch_emails(limit=10)
            self.classify_all()
            self.set_status("Auto-check done.")
        except Exception:
            traceback.print_exc()
        finally:
            # schedule next
            self.schedule_auto_check()


# ----------------- RUN -----------------

def main():
    root = ctk.CTk()
    app = EmailAssistantPro(root)
    root.mainloop()


if __name__ == "__main__":
    main()
