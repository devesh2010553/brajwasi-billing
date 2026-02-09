from flask import Flask, render_template, request, redirect, session
from datetime import datetime, timedelta, time
import math, io, os
import openpyxl
from openpyxl import load_workbook
import json

app = Flask(__name__)
app.secret_key = "supersecretkey"

# Load driver data
with open("driver.json") as f:
    DRIVER_DATA = json.load(f)

DRIVE_FOLDER_ID = "1PZXxUVvB7IOIEG3uVsjUOyEo1A1zPZoP"  # Google Drive folder if using API

# Helper functions
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
    """Find the row number for today's date. If column 2 is empty, just return the next empty row from 8."""
    for r in range(8, ws.max_row + 1):
        cell = ws.cell(row=r, column=2).value
        if isinstance(cell, datetime) and cell.date() == target:
            return r
        if isinstance(cell, str):
            try:
                if datetime.strptime(cell.strip(), "%d-%b-%y").date() == target:
                    return r
            except Exception:
                pass
    # If no date in column 2, just return the first empty row starting from 8
    for r in range(8, ws.max_row + 1):
        if not ws.cell(row=r, column=3).value:
            return r
    return ws.max_row + 1  # append at the end if all filled

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

    # Load local Excel file (replace with Google Drive API if needed)
    file_path = f"{info['sheet']}.xlsx"
    if not os.path.exists(file_path):
        msg = "Excel file not found locally."
        cls = "error"
        return render_template("entry.html", car=car, msg=msg, cls=cls)

    wb = load_workbook(file_path)
    ws = wb[info["sheet"]]

    if request.method == "POST":
        try:
            opening = int(request.form["opening"])
            closing = int(request.form["closing"])
            start = parse_time(request.form["start"])
            end = parse_time(request.form["end"])

            row = find_row(ws, today_date())
            if ws.cell(row=row, column=3).value:
                msg = "Entry already exists for today."
                cls = "error"
            else:
                ws.cell(row=row, column=3).value = opening
                ws.cell(row=row, column=4).value = closing
                ws.cell(row=row, column=5).value = closing - opening
                ws.cell(row=row, column=6).value = start.strftime("%I:%M %p")
                ws.cell(row=row, column=7).value = end.strftime("%I:%M %p")
                ws.cell(row=row, column=8).value = calculate_ot(start, end)
                ws.cell(row=row, column=9).value = get_remarks(start, end, today_date())

                wb.save(file_path)
                msg = "Saved successfully."
                cls = "success"

        except Exception as e:
            msg = f"Error: {e}"
            cls = "error"

    return render_template("entry.html", car=car, msg=msg, cls=cls)

@app.route("/logout")
def logout():
    session.clear()
    return redirect("/")

if __name__ == "__main__":
    app.run(debug=True)
