from flask import Flask, request, render_template, session, redirect, send_file
from openpyxl import load_workbook
from datetime import datetime, timedelta, time
import math, json, os

# --------- Google Drive Setup ---------
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload

app = Flask(__name__)
app.secret_key = "supersecretkey"

SCOPES = ['https://www.googleapis.com/auth/drive.file']

def authenticate_drive():
    # Read service account JSON from environment variable
    creds_json = os.environ.get("GOOGLE_APPLICATION_CREDENTIALS_JSON")
    if not creds_json:
        raise Exception("Service account JSON not found in environment variable")
    creds_dict = json.loads(creds_json)
    creds = service_account.Credentials.from_service_account_info(creds_dict, scopes=SCOPES)
    service = build('drive', 'v3', credentials=creds)
    return service

def upload_file_to_drive(file_path, folder_id=None):
    service = authenticate_drive()
    file_name = os.path.basename(file_path)
    file_metadata = {'name': file_name}
    if folder_id:
        file_metadata['parents'] = [folder_id]
    media = MediaFileUpload(file_path, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    file = service.files().create(body=file_metadata, media_body=media, fields='id').execute()
    return file.get('id')


# --------- Load driver data ---------
with open("driver.json") as f:
    DRIVER_DATA = json.load(f)


# --------- Helper functions ---------
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
    if extra <= 0: return 0
    elif extra > 0.5: return math.ceil(extra)
    return 0

def is_night(start, end):
    return start < time(5,0) or end >= time(22,0)

def get_remarks(start, end, entry_date):
    night = is_night(start, end)
    sunday = entry_date.weekday() == 6
    if night and sunday: return "Night/Sunday"
    elif night: return "Night"
    elif sunday: return "Sunday"
    return ""

def find_row_by_date(ws, target_date):
    for r in range(9, 40):
        cell = ws.cell(row=r, column=2).value
        if isinstance(cell, datetime) and cell.date() == target_date:
            return r
        if isinstance(cell, str):
            try:
                if datetime.strptime(cell.strip(), "%d-%b-%y").date() == target_date:
                    return r
            except: pass
    return None

def is_row_locked(ws, row):
    return ws.cell(row=row, column=3).value not in (None, "")


# --------- Routes ---------
@app.route("/", methods=["GET","POST"])
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

@app.route("/entry", methods=["GET","POST"])
def entry():
    if "car" not in session: return redirect("/")
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

            wb = load_workbook(info["file"])
            ws = wb[info["sheet"]]

            row = find_row_by_date(ws, today_date())
            if not row:
                msg = "Date row not found"
                cls = "error"
            elif is_row_locked(ws, row):
                msg = "Entry already saved 🔒"
                cls = "error"
            else:
                ws.cell(row=row, column=3).value = opening
                ws.cell(row=row, column=4).value = closing
                ws.cell(row=row, column=5).value = closing - opening
                ws.cell(row=row, column=6).value = start.strftime("%I:%M %p")
                ws.cell(row=row, column=7).value = end.strftime("%I:%M %p")
                ws.cell(row=row, column=8).value = calculate_ot(start,end)
                ws.cell(row=row, column=9).value = get_remarks(start,end,today_date())

                # Save locally
                wb.save(info["file"])

                # Upload to Google Drive
                upload_file_to_drive(info["file"])

                msg = "Saved & backed up to Drive ✅"
        except Exception as e:
            msg = f"Error: {e}"
            cls = "error"
    return render_template("entry.html", car=car, msg=msg, cls=cls)

@app.route("/admin")
def admin():
    return render_template("admin.html")

@app.route("/download/<file_id>")
def download(file_id):
    file_path = ""
    if file_id == "file1":
        file_path = "excel/S&T BT February bill 2026.xlsx"
    elif file_id == "file2":
        file_path = "excel/BT bill February 2026(common).xlsx"
    return send_file(file_path, as_attachment=True)

@app.route("/logout")
def logout():
    session.pop("car", None)
    return redirect("/")


# --------- Run server ---------
if __name__ == "__main__":
    port = int(os.environ.get("PORT",5000))
    app.run(host="0.0.0.0", port=port, debug=True)
