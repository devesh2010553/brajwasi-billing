from flask import Flask, render_template, request, redirect, session, send_from_directory, jsonify
import json, os, math, calendar, base64
from datetime import datetime, timedelta, time
from google.oauth2 import service_account
from googleapiclient.discovery import build
from pywebpush import webpush, WebPushException

app = Flask(__name__)
app.secret_key = "supersecretkey"
ADMIN_CODE = os.getenv("ADMIN_CODE", "admin1234")

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
    scopes=["https://www.googleapis.com/auth/spreadsheets"]
)
sheets = build("sheets", "v4", credentials=creds)

# ---------- VAPID Keys ----------
VAPID_PUBLIC_KEY  = os.getenv("VAPID_PUBLIC_KEY",  "BN2Fj08Qoy8TsV0glz8yxg7x6vSkqrJTEWR7lumMnHmXdb0DY1zKOzDY2rCqw-T58LcGU2xNgVnAmXPh6sjO_ok")
VAPID_PRIVATE_KEY = os.getenv("VAPID_PRIVATE_KEY", "-SQGm7yHMjF27t8E7rnDLVUEI1yk-DfDqyGpiE5g6bw")
VAPID_EMAIL       = os.getenv("VAPID_EMAIL", "mailto:admin@brajwasitravels.com")

# ---------- Push subscription store (file-based) ----------
SUBS_FILE = "subscriptions.json"

def load_subs():
    if os.path.exists(SUBS_FILE):
        with open(SUBS_FILE) as f:
            return json.load(f)
    return {}

def save_subs(subs):
    with open(SUBS_FILE, "w") as f:
        json.dump(subs, f, indent=2)

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
    full_hours = int(extra)
    fraction   = extra - full_hours
    if fraction > 0.5:
        return full_hours + 1
    elif full_hours == 0:
        return 0
    else:
        return full_hours

def get_remarks(start, end, date):
    night_start = start < time(5, 0)
    night_end   = end >= time(22, 0)
    sunday      = date.weekday() == 6
    parts = []
    if night_start and night_end:
        parts.append("Night/Night")
    elif night_start or night_end:
        parts.append("Night")
    if sunday:
        parts.append("Sunday")
    return "/".join(parts)

def ordinal_suffix(n):
    if 11 <= n <= 13:
        return "th"
    return {1: "st", 2: "nd", 3: "rd"}.get(n % 10, "th")

# ---------- PWA Routes ----------
@app.route('/manifest.json')
def manifest():
    return send_from_directory(os.getcwd(), 'manifest.json')

@app.route('/service-worker.js')
def sw():
    return send_from_directory(os.getcwd(), 'service-worker.js')

# ---------- Login ----------
@app.route("/", methods=["GET", "POST"])
def login():
    if "car" in session:
        return redirect("/entry")
    msg = ""
    if request.method == "POST":
        code = request.form["code"]
        for car, info in DRIVERS.items():
            if info["code"] == code:
                session.permanent = True
                session["car"] = car
                return redirect("/entry")
        msg = "Invalid code"
    return render_template("login.html", msg=msg)

# ---------- Entry ----------
@app.route("/entry", methods=["GET", "POST"])
def entry():
    if "car" not in session:
        return redirect("/")
    car  = session["car"]
    info = DRIVERS[car]
    msg  = ""
    cls  = "success"

    if request.method == "POST":
        try:
            opening = int(request.form["opening"])
            closing = int(request.form["closing"])
            start   = parse_time(request.form["start"])
            end     = parse_time(request.form["end"])

            entry_date_str = request.form.get("entry_date", "")
            entry_date = datetime.strptime(entry_date_str, "%Y-%m-%d").date() if entry_date_str else today_date()

            remarks  = get_remarks(start, end, entry_date)
            ot       = calculate_ot(start, end)
            total_km = closing - opening

            row = entry_date.day + 7
            rng = f"{info['sheet']}!C{row}:I{row}"

            values = [[
                opening, closing, total_km,
                start.strftime("%I:%M %p"),
                end.strftime("%I:%M %p"),
                ot, remarks
            ]]

            sheets.spreadsheets().values().update(
                spreadsheetId=info["file_id"],
                range=rng,
                valueInputOption="USER_ENTERED",
                body={"values": values}
            ).execute()

            msg = f"Saved successfully ✅ | Total KMs: {total_km} km"

        except Exception as e:
            msg = str(e)
            cls = "error"

    return render_template("entry.html", car=car, msg=msg, cls=cls,
                           today=today_date().isoformat(),
                           vapid_public_key=VAPID_PUBLIC_KEY)

# ---------- Check entry ----------
@app.route("/check-entry", methods=["POST"])
def check_entry():
    if "car" not in session:
        return jsonify({"filled": False})
    car  = session["car"]
    info = DRIVERS[car]
    try:
        entry_date_str = request.json.get("entry_date", "")
        entry_date = datetime.strptime(entry_date_str, "%Y-%m-%d").date()
        row = entry_date.day + 7
        rng = f"{info['sheet']}!C{row}"
        result = sheets.spreadsheets().values().get(
            spreadsheetId=info["file_id"], range=rng).execute()
        values = result.get("values", [])
        filled = bool(values and values[0] and str(values[0][0]).strip() != "")
        return jsonify({"filled": filled})
    except Exception as e:
        return jsonify({"filled": False, "error": str(e)})

# ---------- Get last closing KM ----------
@app.route("/get-last-closing", methods=["POST"])
def get_last_closing():
    if "car" not in session:
        return jsonify({"closing": None})
    car  = session["car"]
    info = DRIVERS[car]
    try:
        entry_date_str = request.json.get("entry_date", "")
        entry_date = datetime.strptime(entry_date_str, "%Y-%m-%d").date()
        for i in range(1, 8):
            prev_date = entry_date - timedelta(days=i)
            prev_row  = prev_date.day + 7
            rng       = f"{info['sheet']}!D{prev_row}"
            result    = sheets.spreadsheets().values().get(
                spreadsheetId=info["file_id"], range=rng).execute()
            values = result.get("values", [])
            if values and values[0] and str(values[0][0]).strip() != "":
                return jsonify({"closing": values[0][0]})
        return jsonify({"closing": None})
    except Exception as e:
        return jsonify({"closing": None, "error": str(e)})

# ---------- Groq voice transcription ----------
@app.route("/transcribe", methods=["POST"])
def transcribe():
    if "car" not in session:
        return jsonify({"error": "Not logged in"}), 401
    try:
        import requests as req_lib
        audio_file = request.files.get("audio")
        if not audio_file:
            return jsonify({"error": "No audio"}), 400

        groq_key = os.getenv("GROQ_API_KEY")
        if not groq_key:
            return jsonify({"error": "GROQ_API_KEY not set"}), 500

        # Send to Groq Whisper
        resp = req_lib.post(
            "https://api.groq.com/openai/v1/audio/transcriptions",
            headers={"Authorization": f"Bearer {groq_key}"},
            files={"file": (audio_file.filename or "audio.webm", audio_file.read(), audio_file.content_type or "audio/webm")},
            data={"model": "whisper-large-v3", "language": "hi", "response_format": "text"}
        )
        if resp.status_code != 200:
            return jsonify({"error": resp.text}), 500

        raw_text = resp.text.strip()

        # Parse the transcription into a number/time using Groq LLM
        parse_resp = req_lib.post(
            "https://api.groq.com/openai/v1/chat/completions",
            headers={"Authorization": f"Bearer {groq_key}", "Content-Type": "application/json"},
            json={
                "model": "llama-3.3-70b-versatile",
                "temperature": 0,
                "messages": [
                    {"role": "system", "content": """You are a number/time parser for a vehicle log entry app.
The user speaks in Hindi, English, or mixed (Hinglish). Your job is to extract ONLY the number or time they said.

RULES:
- For KM numbers: return ONLY the integer. Examples:
  "char hajar nau sao nabbe" → 4990
  "teen hajar paanch sau" → 3500
  "4990" → 4990
  "बारह हज़ार" → 12000
  "do hajar tin sau sattavan" → 2357

- For times: return ONLY in HH:MM (24hr) format. Examples:
  "paanch bajke pandrah minute" → 05:15
  "shaam ke saat baje" → 19:00
  "raat ke das baje" → 22:00
  "subah chhe bajkar bis minute" → 06:20
  "3 bajke 45 minute" → 03:45
  "dopahar ke 2 baje" → 14:00
  "raat ke 11 baje" → 23:00
  "5:15 PM" → 17:15
  "6 AM" → 06:00

- If the input is irrelevant noise, unclear, or not a number/time, return: INVALID

Return ONLY the number or HH:MM or INVALID. No explanation, no extra words."""},
                    {"role": "user", "content": raw_text}
                ]
            }
        )
        if parse_resp.status_code != 200:
            return jsonify({"raw": raw_text, "parsed": None})

        parsed = parse_resp.json()["choices"][0]["message"]["content"].strip()
        if parsed == "INVALID":
            parsed = None

        return jsonify({"raw": raw_text, "parsed": parsed})

    except Exception as e:
        return jsonify({"error": str(e)}), 500

# ---------- Push notification subscription ----------
@app.route("/subscribe-push", methods=["POST"])
def subscribe_push():
    if "car" not in session:
        return jsonify({"error": "Not logged in"}), 401
    car  = session["car"]
    sub  = request.json
    subs = load_subs()
    subs[car] = sub
    save_subs(subs)
    return jsonify({"ok": True})

# ---------- Admin ----------
@app.route("/admin", methods=["GET", "POST"])
def admin():
    msg = ""
    cls = "error"

    if request.method == "POST":
        action = request.form.get("action", "reset")
        code   = request.form.get("code", "")

        if code != ADMIN_CODE:
            msg = "Invalid admin code"
        elif action == "notify":
            try:
                target  = request.form.get("target", "all")  # "all" or car key
                message = request.form.get("message", "").strip()
                if not message:
                    raise ValueError("Message cannot be empty")

                subs = load_subs()
                sent = 0
                failed = 0

                targets = subs.keys() if target == "all" else ([target] if target in subs else [])

                for car_key in targets:
                    sub = subs[car_key]
                    try:
                        webpush(
                            subscription_info=sub,
                            data=json.dumps({"title": "Brajwasi Travels 🚗", "body": message}),
                            vapid_private_key=VAPID_PRIVATE_KEY,
                            vapid_claims={"sub": VAPID_EMAIL}
                        )
                        sent += 1
                    except WebPushException as e:
                        failed += 1
                        # Remove dead subscriptions
                        if "410" in str(e) or "404" in str(e):
                            subs.pop(car_key, None)
                            save_subs(subs)

                msg = f"✅ Notification sent to {sent} driver(s). {f'({failed} failed)' if failed else ''}"
                cls = "success"
            except Exception as e:
                msg = f"Error: {e}"

        elif action == "reset":
            try:
                month         = int(request.form["month"])
                year          = int(request.form["year"])
                days_in_month = calendar.monthrange(year, month)[1]
                month_name    = datetime(year, month, 1).strftime("%B")
                first_ord     = f"1{ordinal_suffix(1)}"
                last_ord      = f"{days_in_month}{ordinal_suffix(days_in_month)}"
                title_text    = (f" Vehicle Bill for the period from "
                                 f"{first_ord} {month_name} {year} "
                                 f"to {last_ord} {month_name} {year}")

                with open("driver.json") as f:
                    drivers = json.load(f)

                for car, info in drivers.items():
                    sheet   = info["sheet"]
                    file_id = info["file_id"]

                    sheets.spreadsheets().values().update(
                        spreadsheetId=file_id,
                        range=f"{sheet}!A3",
                        valueInputOption="USER_ENTERED",
                        body={"values": [[title_text]]}
                    ).execute()

                    date_values = [
                        [day, datetime(year, month, day).strftime("%d-%b-%y")]
                        for day in range(1, days_in_month + 1)
                    ]
                    sheets.spreadsheets().values().update(
                        spreadsheetId=file_id,
                        range=f"{sheet}!A8:B{7 + days_in_month}",
                        valueInputOption="USER_ENTERED",
                        body={"values": date_values}
                    ).execute()

                    if days_in_month < 31:
                        sheets.spreadsheets().values().clear(
                            spreadsheetId=file_id,
                            range=f"{sheet}!A{8 + days_in_month}:I{7 + 31}"
                        ).execute()

                    sheets.spreadsheets().values().clear(
                        spreadsheetId=file_id,
                        range=f"{sheet}!C8:I{7 + days_in_month}"
                    ).execute()

                msg = (f"✅ All sheets updated for {month_name} {year} "
                       f"({days_in_month} days). Ready for new entries.")
                cls = "success"

            except Exception as e:
                msg = f"Error: {e}"

    now      = datetime.now()
    subs     = load_subs()
    drivers  = list(DRIVERS.keys())

    return render_template("admin.html", msg=msg, cls=cls,
                           cur_month=now.month, cur_year=now.year,
                           drivers=drivers, subs=subs)

# ---------- Logout ----------
@app.route("/logout")
def logout():
    session.clear()
    return redirect("/")

@app.route("/ping")
def ping():
    return "ok"

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000)