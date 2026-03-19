from flask import Flask, render_template, request, redirect, session, send_from_directory
import json, os, math
from datetime import datetime, timedelta, time
from google.oauth2 import service_account
from googleapiclient.discovery import build

app = Flask(__name__)
app.secret_key = "supersecretkey"

# ---------- SESSION LIFETIME (10 YEARS) ----------
app.permanent_session_lifetime = timedelta(days=3650)

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
        return 1
    return 0

def get_remarks(start, end, date):
    night_start = start < time(5, 0)       # started before 5 AM
    night_end = end >= time(22, 0)          # closed at or after 10 PM
    sunday = date.weekday() == 6

    parts = []

    if night_start and night_end:
        parts.append("Night Night")
    elif night_start or night_end:
        parts.append("Night")

    if sunday:
        parts.append("Sunday")

    return "/".join(parts)

# ---------- PWA Routes ----------
@app.route('/manifest.json')
def manifest():
    return send_from_directory(os.getcwd(), 'manifest.json')

@app.route('/service-worker.js')
def sw():
    return send_from_directory(os.getcwd(), 'service-worker.js')

# ---------- Routes ----------
@app.route("/", methods=["GET", "POST"])
def login():

    # Skip login if already logged in
    if "car" in session:
        return redirect("/entry")

    msg = ""

    if request.method == "POST":
        code = request.form["code"]

        for car, info in DRIVERS.items():
            if info["code"] == code:

                session.permanent = True   # keeps login for 10 years
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

            entry_date_str = request.form.get("entry_date", "")
            entry_date = datetime.strptime(entry_date_str, "%Y-%m-%d").date() if entry_date_str else today_date()

            remarks = get_remarks(start, end, entry_date)
            ot = calculate_ot(start, end)

            row = entry_date.day + 7
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

    return render_template("entry.html", car=car, msg=msg, cls=cls, today=today_date().isoformat())

@app.route("/check-entry", methods=["POST"])
def check_entry():
    if "car" not in session:
        return {"filled": False}

    car = session["car"]
    info = DRIVERS[car]

    try:
        entry_date_str = request.json.get("entry_date", "")
        entry_date = datetime.strptime(entry_date_str, "%Y-%m-%d").date()
        row = entry_date.day + 7
        rng = f"{info['sheet']}!C{row}"

        result = sheets.spreadsheets().values().get(
            spreadsheetId=info["file_id"],
            range=rng
        ).execute()

        values = result.get("values", [])
        filled = bool(values and values[0] and str(values[0][0]).strip() != "")
        return {"filled": filled}
    except Exception as e:
        return {"filled": False, "error": str(e)}

@app.route("/logout")
def logout():
    session.clear()
    return redirect("/")

@app.route("/ping")
def ping():
    return "ok"
if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000)