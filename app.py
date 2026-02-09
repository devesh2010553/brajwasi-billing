from flask import Flask, request, render_template, session, redirect, send_file
import io, os, json, math
from datetime import datetime, timedelta, time
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload, MediaFileUpload
import pandas as pd
from openpyxl import load_workbook

app = Flask(__name__)
# Use environment variable for secret key, fallback to a default for dev (change in production!)
app.secret_key = os.environ.get("FLASK_SECRET_KEY", "supersecretkey")

SCOPES = ['https://www.googleapis.com/auth/drive']
DRIVE_FOLDER_ID = os.environ.get("DRIVE_FOLDER_ID", "1PZXxUVvB7IOIEG3uVsjUOyEo1A1zPZoP")

# Load driver mapping JSON file once on startup
with open("driver.json") as f:
    DRIVER_DATA = json.load(f)

def get_drive_service():
    creds_json = os.environ.get("GOOGLE_APPLICATION_CREDENTIALS_JSON")
    if not creds_json:
        raise Exception("Google service account JSON not found in environment variable GOOGLE_APPLICATION_CREDENTIALS_JSON")
    creds = service_account.Credentials.from_service_account_info(
        json.loads(creds_json), scopes=SCOPES
    )
    return build('drive', 'v3', credentials=creds)

def download_excel(file_id):
    service = get_drive_service()
    request = service.files().export_media(
        fileId=file_id,
        mimeType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    fh = io.BytesIO()
    downloader = MediaIoBaseDownload(fh, request)
    done = False
    while not done:
        _, done = downloader.next_chunk()
    fh.seek(0)
    return fh

def upload_excel(file_bytes, file_name, folder_id=None):
    service = get_drive_service()
    temp_path = f"/tmp/{file_name}"
    with open(temp_path, "wb") as f:
        f.write(file_bytes.getbuffer())

    file_metadata = {'name': file_name}
    if folder_id:
        file_metadata['parents'] = [folder_id]

    media = MediaFileUpload(
        temp_path,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    uploaded = service.files().create(body=file_metadata, media_body=media, fields="id").execute()
    return uploaded.get("id")

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

# ROUTES
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
                    upload_excel(out, os.path.basename(info["file"] + ".xlsx"), folder_id=DRIVE_FOLDER_ID)
                    msg = "Saved & backed up to Drive"
                    cls = "success"
        except Exception as e:
            msg = f"Error: {e}"
            cls = "error"
    return render_template("entry.html", car=car, msg=msg, cls=cls)

@app.route("/admin", methods=["GET"])
def admin_panel():
    return render_template("admin.html")

@app.route("/download/<file_key>", methods=["GET"])
def download_file(file_key):
    if file_key not in ["file1", "file2"]:
        return "Invalid file", 404
    file_info = {
        "file1": "S&T BT February bill 2026.xlsx",
        "file2": "BT bill February 2026(common).xlsx"
    }
    file_name = file_info[file_key]
    file_id = None
    for drv in DRIVER_DATA.values():
        if drv["file"].endswith(file_name):
            file_id = drv.get("file_id")
            break
    if not file_id:
        return "File not found", 404
    stream = download_excel(file_id)
    return send_file(stream, download_name=file_name, as_attachment=True)

@app.route("/logout")
def logout():
    session.pop("car", None)
    return redirect("/")

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)))
