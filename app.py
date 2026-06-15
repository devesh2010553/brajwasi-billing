from flask import Flask, render_template, request, redirect, session, send_from_directory, jsonify
import json, os, math, calendar
from datetime import datetime, timedelta, time
from google.oauth2 import service_account
from googleapiclient.discovery import build
from pywebpush import webpush, WebPushException
from pymongo import MongoClient

app = Flask(__name__)
app.secret_key = "supersecretkey"
ADMIN_CODE = os.getenv("ADMIN_CODE", "admin1234")

app.permanent_session_lifetime = timedelta(days=3650)

with open("driver.json", "r") as f:
    DRIVERS = json.load(f)

sa_json = os.getenv("GOOGLE_SERVICE_ACCOUNT_JSON")
if not sa_json:
    raise RuntimeError("GOOGLE_SERVICE_ACCOUNT_JSON not set")

creds = service_account.Credentials.from_service_account_info(
    json.loads(sa_json),
    scopes=["https://www.googleapis.com/auth/spreadsheets"]
)
sheets = build("sheets", "v4", credentials=creds)

# ---------- VAPID ----------
VAPID_PUBLIC_KEY  = os.getenv("VAPID_PUBLIC_KEY",  "BI495kOQBEdy0aSHpJfT1bpzOmlryP__uLVreDYYrA2zsqEirtPSDrX7CGTvol6oygSPvTcBQ_RLwXDMPqcFDQo")
VAPID_PRIVATE_KEY = os.getenv("VAPID_PRIVATE_KEY", "Hg1eASB4wKsblPHdoZyvEVihWjYQxbNNaBPUkS7Rxzs")
VAPID_EMAIL       = os.getenv("VAPID_EMAIL", "mailto:admin@brajwasitravels.com")

# ---------- MongoDB ----------
MONGO_URI  = os.getenv("MONGO_URI", "")
_mongo_client = None
_mongo_col    = None

def get_col():
    """Return MongoDB collection, creating connection once. Returns None if unavailable."""
    global _mongo_client, _mongo_col
    if _mongo_col is not None:
        return _mongo_col
    if not MONGO_URI:
        print("⚠️  MONGO_URI not set — using file fallback")
        return None
    try:
        _mongo_client = MongoClient(MONGO_URI, serverSelectionTimeoutMS=5000)
        _mongo_client.admin.command("ping")   # verify connection
        _mongo_col = _mongo_client["brajwasi"]["subscriptions"]
        print("✅ MongoDB connected")
        return _mongo_col
    except Exception as e:
        print(f"❌ MongoDB failed: {e}")
        return None

# ---------- Subscription CRUD ----------
SUBS_FILE = "subscriptions.json"

def _file_load():
    if os.path.exists(SUBS_FILE):
        with open(SUBS_FILE) as f:
            return json.load(f)
    return {}

def _file_save(subs):
    with open(SUBS_FILE, "w") as f:
        json.dump(subs, f, indent=2)

def load_subs():
    col = get_col()
    if col is not None:
        try:
            return {d["_id"]: d["sub"] for d in col.find()}
        except Exception as e:
            print(f"❌ load_subs MongoDB error: {e}")
    return _file_load()

def save_sub(car_key, sub_info):
    col = get_col()
    if col is not None:
        try:
            col.update_one(
                {"_id": car_key},
                {"$set": {"sub": sub_info, "updated_at": datetime.utcnow()}},
                upsert=True
            )
            print(f"✅ MongoDB: saved sub for {car_key}")
            return
        except Exception as e:
            print(f"❌ save_sub MongoDB error: {e}")
    # file fallback
    subs = _file_load()
    subs[car_key] = sub_info
    _file_save(subs)
    print(f"✅ File: saved sub for {car_key}")

def delete_sub(car_key):
    col = get_col()
    if col is not None:
        try:
            col.delete_one({"_id": car_key})
            print(f"🗑️  MongoDB: deleted sub for {car_key}")
            return
        except Exception as e:
            print(f"❌ delete_sub MongoDB error: {e}")
    subs = _file_load()
    subs.pop(car_key, None)
    _file_save(subs)

# ---------- Push ----------
def send_push(sub_info, message, title="Brajwasi Travels 🚗"):
    """Send a web push notification. sub_info is the full subscription dict."""
    webpush(
        subscription_info=sub_info,
        data=json.dumps({"title": title, "body": message}),
        vapid_private_key=VAPID_PRIVATE_KEY,   # pass raw base64 string directly
        vapid_claims={"sub": VAPID_EMAIL}
    )

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

# ---------- PWA ----------
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

# ---------- Last closing KM ----------
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

        resp = req_lib.post(
            "https://api.groq.com/openai/v1/audio/transcriptions",
            headers={"Authorization": f"Bearer {groq_key}"},
            files={"file": (audio_file.filename or "audio.webm",
                            audio_file.read(),
                            audio_file.content_type or "audio/webm")},
            data={"model": "whisper-large-v3-turbo",
                  "language": "hi",
                  "response_format": "text"}
        )
        if resp.status_code != 200:
            return jsonify({"error": resp.text}), 500

        raw_text = resp.text.strip()

        parse_resp = req_lib.post(
            "https://api.groq.com/openai/v1/chat/completions",
            headers={"Authorization": f"Bearer {groq_key}",
                     "Content-Type": "application/json"},
            json={
                "model": "llama-3.3-70b-versatile",
                "temperature": 0,
                "messages": [
                    {"role": "system", "content": """You are a number and time parser for a vehicle daily log app used by Indian drivers.
Input is spoken Hindi, English, Hinglish, or mixed language. Extract ONLY the number or time being said.

═══ HINDI NUMBER UNITS — memorize all ═══
ek=1, do=2, teen=3, char=4, paanch=5, chhe/chheh/chhah=6, saat=7, aath=8, nau=9, das=10
gyarah=11, barah=12, terah=13, chaudah=14, pandrah=15, solah=16, satrah=17, atharah=18, unnis=19
bees=20, ikkees=21, baees=22, teis=23, chaubees=24, pachchees=25, chhabbees=26, sattaees=27, athaees=28, untees=29
tees=30, ikattees=31, battees=32, taintees=33, chautees=34, paintees=35, chhattees=36, saintees=37, artees=38, untaalees=39
chaalees=40, iktaalees=41, bayaalees=42, taintaalees=43, chavaalees=44, paintaalees=45, chhiyaalees=46, saintaalees=47, artaalees=48, unchaas=49
pachaas=50, ikyavan=51, baavan=52, tirpan=53, chauvan=54, pachpan=55, chhappan=56, sattavan=57, atthavan=58, unsath=59
saath/saathi=60, iksath=61, basath=62, tirsath=63, chausath=64, painsath=65, chhiyasath=66, sarsath=67, arsath=68, unhattar=69
sattar=70, ikhattar=71, bahattar=72, tihattar=73, chauhattar=74, pachhattar=75, chhihattar=76, sathattar=77, athahattar=78, unnaasee=79
assi=80, ikyaasee=81, bayaasee=82, tiraasee=83, chauraasee=84, pachaasee=85, chhiyaasee=86, sataasee=87, athaasee=88, navaasee=89
nabbe=90, ikyaanave=91, baanave=92, tiraanave=93, chauraanave=94, panchaanave=95, chhiyaanave=96, sataanave=97, athaanave=98, ninaanave=99
sau=100, hazaar/hazar=1000, lakh/laakh=100000, crore=10000000

CRITICAL — lakh and laakh are IDENTICAL, both = 100,000:
  "nau lakh" = "nau laakh" = 9,00,000 = 900000
  "ek lakh" = "ek laakh" = 1,00,000 = 100000
  "paanch lakh" = "paanch laakh" = 5,00,000 = 500000
  "saat laakh" = "saat lakh" = 7,00,000 = 700000

Number examples:
  "char hajar nau sau nabbe" → 4990
  "char hajar nau soo nabbe" → 4990
  "teen hazaar paanch sau" → 3500
  "barah hazaar teen sau pachaas" → 12350
  "nau lakh" → 900000
  "nau laakh" → 900000
  "paanch lakh bees hazaar" → 520000
  "do laakh pachaas hazaar" → 250000
  "ek lakh pachaas hazaar" → 150000
  "saat lakh" → 700000
  "bees hazaar" → 20000
  "teen sau" → 300
  "pachpan" → 55
  "baees hazaar char sau sattavan" → 22457
  "1 lakh" → 100000
  "9 lakh" → 900000
  "9 laakh" → 900000
  "12345" → 12345
  "ek crore" → 10000000

═══ TIME PARSING RULES ═══
Return 24-hour HH:MM format.
Time words: subah/sawere=morning AM, dopahar=afternoon 12-4PM, shaam=evening 4-8PM, raat=night PM
Special forms:
  saade X = X hours 30 min (saade aath = 8:30, saade nau = 9:30)
  paune X = X hours minus 15 min (paune nau = 8:45, paune das = 9:45)
  sawa X = X hours 15 min (sawa chhe = 6:15, sawa saat = 7:15)
  dhaai = 2:30

Time examples:
  "paanch bajke pandrah minute" → 05:15
  "shaam ke saat baje" → 19:00
  "raat ke das baje" → 22:00
  "subah chhe bajkar bis minute" → 06:20
  "dopahar ke do baje" → 14:00
  "raat ke gyarah baje" → 23:00
  "teen bajke chaalees minute" → 03:40
  "saat bajke paanch minute shaam" → 19:05
  "saade aath subah" → 08:30
  "paune nau raat" → 20:45
  "sawa chhe subah" → 06:15
  "10 baje raat" → 22:00
  "5:15 PM" → 17:15
  "6 AM" → 06:00
  "7 baje" → 07:00

═══ OUTPUT RULES ═══
- Return ONLY: an integer number OR HH:MM time OR the word INVALID
- No explanation, no units (km, घंटे), no extra words whatsoever
- Ignore filler words: um, uh, matlab, yaani, woh, toh
- If input is noise, irrelevant, or unparseable: INVALID"""},
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

# ---------- Push subscription ----------
@app.route("/subscribe-push", methods=["POST"])
def subscribe_push():
    if "car" not in session:
        return jsonify({"error": "Not logged in"}), 401
    car = session["car"]
    sub = request.json
    save_sub(car, sub)
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
                target  = request.form.get("target", "all")
                message = request.form.get("message", "").strip()
                if not message:
                    raise ValueError("Message cannot be empty")

                subs   = load_subs()
                sent   = 0
                failed = 0
                dead   = []

                targets = list(subs.keys()) if target == "all" else ([target] if target in subs else [])

                for car_key in targets:
                    try:
                        send_push(subs[car_key], message)
                        sent += 1
                        print(f"✅ Push sent to {car_key}")
                    except WebPushException as e:
                        failed += 1
                        err_str = str(e)
                        print(f"❌ WebPushException for {car_key}: {err_str}")
                        if "410" in err_str or "404" in err_str:
                            dead.append(car_key)
                    except Exception as e:
                        failed += 1
                        print(f"❌ Push error for {car_key}: {e}")

                for k in dead:
                    delete_sub(k)

                if sent == 0 and failed == 0:
                    msg = "⚠️ No subscribed drivers found. Drivers must allow notifications first."
                    cls = "error"
                else:
                    msg = f"✅ Sent to {sent} driver(s)." + (f" {failed} failed — check Render logs." if failed else "")
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
                        spreadsheetId=file_id, range=f"{sheet}!A3",
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
                       f"({days_in_month} days).")
                cls = "success"
            except Exception as e:
                msg = f"Error: {e}"

    now     = datetime.now()
    subs    = load_subs()
    drivers = list(DRIVERS.keys())

    return render_template("admin.html", msg=msg, cls=cls,
                           cur_month=now.month, cur_year=now.year,
                           drivers=drivers, subs=subs)

@app.route("/mongo-test")
def mongo_test():
    col = get_db()
    if col is None:
        return jsonify({"status": "❌ MongoDB not connected", "uri_set": bool(MONGO_URI)})
    try:
        count = col.count_documents({})
        docs  = [{"car": d["_id"], "updated_at": str(d.get("updated_at", ""))} for d in col.find()]
        return jsonify({"status": "✅ MongoDB connected", "subscriptions": count, "drivers": docs})
    except Exception as e:
        return jsonify({"status": f"❌ Error: {e}"})

@app.route("/logout")
def logout():
    session.clear()
    return redirect("/")

@app.route("/ping")
def ping():
    return "ok"

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000)