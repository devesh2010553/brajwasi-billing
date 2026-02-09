import io
import os
import math
from datetime import datetime, timedelta, time
import json

from flask import Flask, render_template, request, redirect, session
from openpyxl import load_workbook
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload, MediaIoBaseUpload
from google.oauth2.service_account import Credentials

# ----------------------------
# GOOGLE SERVICE ACCOUNT SETUP
# ----------------------------
SERVICE_ACCOUNT_INFO = os.environ.get("GOOGLE_SERVICE_ACCOUNT_JSON")
if not SERVICE_ACCOUNT_INFO:
    raise Exception("Service account JSON not set in environment variable")

SERVICE_ACCOUNT_DICT = json.loads(SERVICE_ACCOUNT_INFO)
SCOPES = ["https://www.googleapis.com/auth/drive"]
creds = Credentials.from_service_account_info(SERVICE_ACCOUNT_DICT, scopes=SCOPES)

# Google Drive service
service = build('drive', 'v3', credentials=creds)

# ----------------------------
# FLASK APP SETUP
# ----------------------------
app = Flask(__name__)
app.secret_key = os.environ.get("FLASK_SECRET_KEY", "supersecretkey")

# ----------------------------
# DRIVER DATA
# ----------------------------
DRIVER_DATA = {
    "UP80ET4509": {"file_id": "1qiwgJAwv6CRdPlrU8eHL5AEyO8keSzZD", "sheet": "UP80ET4509", "code": "1234"},
    "UP80JT7912": {"file_id": "1qiwgJAwv6CRdPlrU8eHL5AEyO8keSzZD", "sheet": "UP80JT7912", "code": "5678"},
    "UP79AT9051": {"file_id": "1qiwgJAwv6CRdPlrU8eHL5AEyO8keSzZD", "sheet": "UP79AT9051", "code": "9012"},
    "UP80JT5884": {"file_id": "1Fu08Gou2DB4YmZi1Z4pXKON8GCVUlsdG", "sheet": "UP80JT5884", "code": "3456"},
    "UP80GT6593": {"file_id": "1Fu08Gou2DB4YmZi1Z4pXKON8GCVUlsdG", "sheet": "UP80GT6593", "code": "7890"}
}

DRIVE_FOLDER_ID = os.environ.get("DRIVE_FOLDER_ID")  # Optional backup folder

# ----------------------------
# HELPER FUNCTIONS
# ----------------------------
def today_date():
    return datetime.now().date()

def parse_time(t):
    return datetime.strptime(t, "%H:%M").time()

def hours_between(start, end):
    dt1 = datetime.combine(today_date(), start)
    dt2 = datetime.combine(today_date(), end)
    if dt2 < dt1:
        dt2 += timedelta(days=1)
    return (dt2 - dt1).total_seconds() / 3600

def calculate_ot(start, end):
    hrs = hours_between(start, end)
    extra = hrs - 12
    if extra <= 0:
        return 0
    elif extra > 0.5:
        return math.ceil(extra)
    return 0

def is_night(start, end):
    return start < time(5, 0) or end >= time(22, 0)

def get_remarks(start, end, entry_date):
    night = is_night(start, end)
    sunday = entry_date.weekday() == 6
    if night and sunday:
        return "Night/Sunday"
    if night:
        return "Night"
    if sunday:
        return "Sunday"
    return ""

def find_row(ws, target):
    for r in range(9, ws.max_row + 1):
        cell = ws.cell(row=r, column=2).value
        if isinstance(cell, datetime) and cell.date() == target:
            return r
        if isinstance(cell, str):
            try:
                if datetime.strptime(cell.strip(), "%d-%b-%y").date() == target:
                    return r
            except Exception:
                pass
    return None

# ----------------------------
# GOOGLE DRIVE FUNCTIONS
# ----------------------------
def download_excel(file_id):
    request = service.files().get_media(fileId=file_id)
    fh = io.BytesIO()
    downloader = MediaIoBaseDownload(fh, request)
    done = False
    while not done:
        status, done = downloader.next_chunk()
    fh.seek(0)
    return fh

def upload_excel(file_obj, name, folder_id=None):
    file_metadata = {"name": name}
    if folder_id:
        file_metadata["parents"] = [folder_id]
    media = MediaIoBaseUpload(file_obj, mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    service.files().create(body=file_metadata, media_body=media, fields="id").execute()

# ----------------------------
# FLASK ROUTES
# ----------------------------
@app.route("/", methods=["GET", "POST"])
def login():
    msg = ""
    if request.method == "POST":
        code = request.form.get("code")
        for car, info in DRIVER_DATA.items():
            if info.get("code") == code:
                session["car"] = car
                return redirect("/entry")
        msg = "Invalid code"
    return render_template("login.html", msg=msg)

@app.route("/entry", methods=["GET", "POST"])
def entry():
    if "car" not in session:
        return redirect("/")
    car = session["car"]
    info = DRIVER_DATA[car]
    msg = ""
    cls = "success"
    if request.method == "POST":
        try:
            opening = int(request.form["opening"])
            closing = int(request.form["closing"])
            start = parse_time(request.form["start"])
            end = parse_time(request.form["end"])

            excel_stream = download_excel(info["file_id"])
            wb = load_workbook(filename=excel_stream)
            ws = wb[info["sheet"]]

            row = find_row(ws, today_date())
            if not row:
                msg = "Date row not found"
                cls = "error"
            else:
                if ws.cell(row=row, column=3).value:
                    msg = "Entry already exists"
                    cls = "error"
                else:
                    ws.cell(row=row, column=3).value = opening
                    ws.cell(row=row, column=4).value = closing
                    ws.cell(row=row, column=5).value = closing - opening
                    ws.cell(row=row, column=6).value = start.strftime("%I:%M %p")
                    ws.cell(row=row, column=7).value = end.strftime("%I:%M %p")
                    ws.cell(row=row, column=8).value = calculate_ot(start, end)
                    ws.cell(row=row, column=9).value = get_remarks(start, end, today_date())

                    out = io.BytesIO()
                    wb.save(out)
                    out.seek(0)
                    upload_excel(out, os.path.basename(info["sheet"] + ".xlsx"), folder_id=DRIVE_FOLDER_ID)
                    msg = "Saved & backed up to Drive"
                    cls = "success"
        except Exception as e:
            msg = f"Error: {e}"
            cls = "error"
    return render_template("entry.html", car=car, msg=msg, cls=cls)

# ----------------------------
# RUN
# ----------------------------
if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)), debug=True)
