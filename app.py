import os
import io
import math
import json
from datetime import datetime, timedelta, time
from flask import Flask, render_template, request, redirect, session
from openpyxl import load_workbook
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload, MediaIoBaseUpload
from google.oauth2.service_account import Credentials

# -------------------- CONFIG --------------------
APP_SECRET = "your_secret_key_here"
DRIVER_JSON = "driver.json"   # Your driver.json file
SCOPES = ["https://www.googleapis.com/auth/drive"]
SERVICE_ACCOUNT_FILE = "driver.json"  # Must be your Google service account JSON
DRIVE_FOLDER_ID = "1PZXxUVvB7IOIEG3uVsjUOyEo1A1zPZoP"  # Folder with Excel files

app = Flask(__name__)
app.secret_key = APP_SECRET

# Load driver data
with open(DRIVER_JSON) as f:
    DRIVER_DATA = json.load(f)

# Initialize Google Drive API
creds = Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=SCOPES)
drive_service = build("drive", "v3", credentials=creds)

# -------------------- HELPER FUNCTIONS --------------------
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
    for r in range(8, ws.max_row + 1):  # Start from row 8 (first data row)
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

def download_excel(file_id):
    request_drive = drive_service.files().get_media(fileId=file_id)
    fh = io.BytesIO()
    downloader = MediaIoBaseDownload(fh, request_drive)
    done = False
    while not done:
        status, done = downloader.next_chunk()
    fh.seek(0)
    return fh

def upload_excel(stream, name, folder_id):
    file_metadata = {
        "name": name,
        "parents": [folder_id]
    }
    media = MediaIoBaseUpload(stream, mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", resumable=True)
    drive_service.files().create(body=file_metadata, media_body=media).execute()

# -------------------- ROUTES --------------------
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

            # Download workbook from Drive
            excel_stream = download_excel(info["file_id"])
            wb = load_workbook(filename=excel_stream)
            ws = wb[info["sheet"]]

            # Find today's row
            row = find_row(ws, today_date())
            if not row:
                msg = "Date row not found in sheet"
                cls = "error"
            else:
                if ws.cell(row=row, column=3).value:
                    msg = "Entry already exists"
                    cls = "error"
                else:
                    # Fill data
                    ws.cell(row=row, column=3).value = opening
                    ws.cell(row=row, column=4).value = closing
                    ws.cell(row=row, column=5).value = closing - opening
                    ws.cell(row=row, column=6).value = start.strftime("%I:%M %p")
                    ws.cell(row=row, column=7).value = end.strftime("%I:%M %p")
                    ws.cell(row=row, column=8).value = calculate_ot(start, end)
                    ws.cell(row=row, column=9).value = get_remarks(start, end, today_date())

                    # Upload back to Drive
                    out = io.BytesIO()
                    wb.save(out)
                    out.seek(0)
                    upload_excel(out, f"{info['sheet']}.xlsx", folder_id=DRIVE_FOLDER_ID)

                    msg = "Saved successfully & backed up to Drive"
                    cls = "success"

        except Exception as e:
            msg = f"Error: {e}"
            cls = "error"

    return render_template("entry.html", car=car, msg=msg, cls=cls)

@app.route("/logout")
def logout():
    session.pop("car", None)
    return redirect("/")

# -------------------- RUN --------------------
if __name__ == "__main__":
    app.run(debug=True)
