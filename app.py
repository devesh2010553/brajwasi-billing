from flask import Flask, render_template, request, redirect, session
import json, os, math
from datetime import datetime, timedelta, time
from google.oauth2 import service_account
from googleapiclient.discovery import build

app = Flask(__name__)
app.secret_key = "supersecretkey"

# ---------- Load drivers ----------
with open("driver.json", "r") as f:
    DRIVERS = json.load(f)

# ---------- Google Auth ----------
sa_json = os.getenv("GOOGLE_SERVICE_ACCOUNT_JSON")
if not sa_json:
    raise RuntimeError("GOOGLE_SERVICE_ACCOUNT_JSON not set")

creds = service_account.Credentials.from_service_account_info(
    json.loads(sa_json),
    scopes=[
        "https://www.googleapis.com/auth/spreadsheets"
    ]
)

sheets = build("sheets", "v4", credentials=creds)

# ---------- Helpers ----------
def today_date():
    return datetime.now().date()

def parse_time(t):
    return datetime.strptime(t, "%H:%M").time()

def hours_between(start, end):
    d1 = datetime.combine(today_date(), start)
    d2 = datetime.combine(today_date(), end)
    if d2 < d1:
        d2 += timedelta(days=1)
    return (d2 - d1).total_seconds() / 3600

def calculate_ot(start, end):
    hrs = hours_between(start, end)
    extra = hrs - 12
    if extra <= 0:
        return 0
    if extra > 0.5:
        return math.ceil(extra)
    return 0

def is_night(start, end):
    return start < time(5, 0) or end >= time(22, 0)

def get_remarks(start, end, date):
    night = is_night(start, end)
    sunday = date.weekday() == 6
    if night and sunday:
        return "Night/Sunday"
    if night:
        return "Night"
    if sunday:
        return "Sunday"
    return ""

# ---------- Routes ----------
@app.route("/", methods=["GET", "POST"])
def login():
    msg = ""
    if request.method == "POST":
        code = request.form["code"]
        for car, info in DRIVERS.items():
            if info["code"] == code:
                session["car"] = car
                return redirect("/entry")
        msg = "Invalid code"
    return render_template("login.html", msg=msg)

@app.route("/entry", methods=["GET", "POST"])
def entry():
    if "car" not in session:
        return redirect("/")

    car = session["car"]
    info = DRIVERS[car]
    msg = ""
    cls = "success"

    if request.method == "POST":
        try:
            opening = int(request.form["opening"])
            closing = int(request.form["closing"])
            start = parse_time(request.form["start"])
            end = parse_time(request.form["end"])

            today = today_date()
            remarks = get_remarks(start, end, today)
            ot = calculate_ot(start, end)

            # Example row (adjust if needed)
            row = today.day + 8
            rng = f"{info['sheet']}!C{row}:I{row}"

            values = [[
                opening,
                closing,
                closing - opening,
                start.strftime("%I:%M %p"),
                end.strftime("%I:%M %p"),
                ot,
                remarks
            ]]

            sheets.spreadsheets().values().update(
                spreadsheetId=info["file_id"],
                range=rng,
                valueInputOption="USER_ENTERED",
                body={"values": values}
            ).execute()

            msg = "Saved successfully"

        except Exception as e:
            msg = str(e)
            cls = "error"

    return render_template("entry.html", car=car, msg=msg, cls=cls)

@app.route("/logout")
def logout():
    session.clear()
    return redirect("/")

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000)
