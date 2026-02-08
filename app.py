from flask import Flask, request, render_template, session, redirect, send_file
from openpyxl import load_workbook
from datetime import datetime, timedelta, time
import math, json, os

app = Flask(__name__)
app.secret_key = "supersecretkey"

# ------------------- Load Driver JSON -------------------
with open("driver.json") as f:
    DRIVER_DATA = json.load(f)

# ------------------- Helper Functions -------------------
def today_date():
    return datetime.now().date()

def parse_time(t):
    return datetime.strptime(t, "%H:%M").time()

def hours_between(start, end):
    d = today_date()
    dt1 = datetime.combine(d, start)
    dt2 = datetime.combine(d, end)
    if dt2 < dt1:
        dt2 += timedelta(days=1)
    return (dt2 - dt1).total_seconds() / 3600  # float hours

def calculate_ot(start, end):
    hrs = hours_between(start, end)
    extra = hrs - 12
    if extra <= 0:
        return 0
    elif extra > 0.5:
        return math.ceil(extra)
    else:
        return 0

def is_night(start, end):
    return start < time(5,0) or end >= time(22,0)

def get_remarks(start, end, entry_date):
    night = is_night(start, end)
    sunday = entry_date.weekday() == 6
    if night and sunday:
        return "Night/Sunday"
    elif night:
        return "Night"
    elif sunday:
        return "Sunday"
    else:
        return ""

def find_row_by_date(ws, target_date):
    """Find the row for today's date, starting from row 8"""
    for r in range(8, 100):
        cell = ws.cell(row=r, column=2).value
        if isinstance(cell, datetime) and cell.date() == target_date:
            return r
        if isinstance(cell, str):
            try:
                if datetime.strptime(cell.strip(), "%d-%b-%y").date() == target_date:
                    return r
            except:
                pass
    return None

def is_row_locked(ws, row):
    return ws.cell(row=row, column=3).value not in (None, "")

# ------------------- Routes -------------------
@app.route("/", methods=["GET", "POST"])
def login():
    msg = ""
    cls = "error"
    if request.method == "POST":
        code = request.form.get("code")
        for car, info in DRIVER_DATA.items():
            if info.get("code") == code:
                session["car"] = car
                return redirect("/entry")
        msg = "Invalid code"
    return render_template("login.html", msg=msg, cls=cls)

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

            wb = load_workbook(info["file"])
            ws = wb[info["sheet"]]

            row = find_row_by_date(ws, today_date())
            if not row:
                msg = "Today's date row not found"
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
                ws.cell(row=row, column=8).value = calculate_ot(start, end)
                ws.cell(row=row, column=9).value = get_remarks(start, end, today_date())
                wb.save(info["file"])
                msg = "Saved successfully ✅"
        except Exception as e:
            msg = f"Error: {e}"
            cls = "error"

    return render_template("entry.html", car=car, msg=msg, cls=cls)

@app.route("/admin", methods=["GET"])
def admin():
    """Simple admin page to download both files"""
    # You can add authentication here if needed
    return render_template("admin.html")

@app.route("/download/<file_id>")
def download(file_id):
    """Download the full Excel workbooks"""
    if file_id == "file1":
        path = "excel/S&T BT February bill 2026.xlsx"
    elif file_id == "file2":
        path = "excel/BT bill February 2026(common).xlsx"
    else:
        return "File not found", 404

    return send_file(path, as_attachment=True)

@app.route("/logout")
def logout():
    session.pop("car", None)
    return redirect("/")

# ------------------- Run -------------------
if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=True)
