from flask import Flask, request, send_file, redirect, session
import pandas as pd
import re
import os
from datetime import datetime
from werkzeug.security import generate_password_hash, check_password_hash

app = Flask(__name__)
app.secret_key = "CHANGE_THIS_SECRET"

# ✅ USERS
USERS = {
    "joanine": generate_password_hash("Duvesco123"),
    "leani": generate_password_hash("Password456")
}

ADMIN_USER = "joanine"

# ✅ LOGIN
@app.route("/login", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        u = request.form["username"]
        p = request.form["password"]

        if u in USERS and check_password_hash(USERS[u], p):
            session["user"] = u
            return redirect("/")
        else:
            return "<h3 style='text-align:center;color:red;'>Invalid login</h3>"

    return """
    <html>
    <body style="font-family:Arial;background:#f4f6f9;">
    <div style="width:320px;margin:120px auto;background:white;padding:30px;border-radius:10px;text-align:center;">
        <h2>Login</h2>
        <form method="post">
            <input name="username" placeholder="Username" required><br><br>
            <input type="password" name="password" placeholder="Password" required><br><br>
            <button style="background:#0078D4;color:white;padding:10px 20px;border:none;border-radius:6px;">Login</button>
        </form>
    </div>
    </body>
    </html>
    """

# ✅ LOGOUT
@app.route("/logout")
def logout():
    session.clear()
    return redirect("/login")

# ✅ MAIN
@app.route("/", methods=["GET", "POST"])
def index():

    # ✅ FORCE LOGIN
    if "user" not in session:
        return redirect("/login")

    summary_html = ""
    chart_html = ""

    if request.method == "POST":

        files = request.files.getlist("files")

        pattern = re.compile(r'([A-Z]{2,3})\s*(\d{9,10})')

        valid, suspense = [], []
        total_records = 0
        total_statement = 0
        daily_data = {}

        for file in files:
            filename = file.filename.lower()
            df = pd.read_excel(file)

            # ✅ DATE FIX (ALL BANKS)
            for col in df.columns:
                if "date" in col.lower():
                    df[col] = pd.to_datetime(df[col], format="%d/%m/%Y", errors="coerce")
                    mask = df[col].isna()
                    df.loc[mask, col] = pd.to_datetime(df.loc[mask, col], errors="coerce")
                    df[col] = df[col].dt.strftime("%Y/%m/%d")

            total_statement += df.get("Amount", pd.Series()).sum()
            total_records += len(df)

            # ✅ BANK DETECTION
            bank = (
                "Std B 4174" if "4174" in filename else
                "Std B 6831" if "6831" in filename else
                "Std B 1039" if "1039" in filename else
                "Capitec 8868" if "8868" in filename else
                "Unknown Bank"
            )

            for _, row in df.iterrows():
                desc = str(row.get("Description", "")).upper()
                matches = pattern.findall(desc)

                date = row.get("Date")
                amount = row.get("Amount", 0)

                if date:
                    daily_data[date] = daily_data.get(date, 0) + amount

                if matches:
                    prefix, digits = matches[0]

                    if len(prefix) == 2 and digits.startswith("0"):
                        digits = "O" + digits[1:]

                    row["Matter Number"] = prefix + digits
                    row["Description 1"] = bank

                    valid.append(row.drop(labels=["Description"]))
                else:
                    row["Reason"] = "No valid Matter Number found"
                    suspense.append(row)

        valid_df = pd.DataFrame(valid)
        suspense_df = pd.DataFrame(suspense)

        valid_amt = valid_df.get("Amount", pd.Series()).sum()
        suspense_amt = suspense_df.get("Amount", pd.Series()).sum()
        processed_total = valid_amt + suspense_amt
        diff = total_statement - processed_total

        # ✅ SAVE OUTPUT FILE
        output_file = f"processed_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        with pd.ExcelWriter(output_file) as writer:
            valid_df.to_excel(writer, "Valid", index=False)
            suspense_df.to_excel(writer, "Suspense", index=False)

        session["file"] = output_file

        # ✅ AUDIT LOG
        log_file = "audit_log.xlsx"
        log_data = {
            "Timestamp": datetime.now(),
            "User": session["user"],
            "Records": total_records,
            "Valid": len(valid_df),
            "Suspense": len(suspense_df),
            "Difference": diff
        }

        log_df = pd.DataFrame([log_data])

        if os.path.exists(log_file):
            old = pd.read_excel(log_file)
            log_df = pd.concat([old, log_df], ignore_index=True)

        log_df.to_excel(log_file, index=False)

        # ✅ SUMMARY (CENTERED)
        summary_html = f"""
        <div class="card">
            <h2>Processing Summary</h2>

            <p><b>Total Records:</b> {total_records}</p>
            <p class="green"><b>Valid:</b> {len(valid_df)}</p>
            <p class="red"><b>Suspense:</b> {len(suspense_df)}</p>

            <hr>

            <p>Total Amount: R {processed_total:,.2f}</p>
            <p class="green">Valid: R {valid_amt:,.2f}</p>
            <p class="red">Suspense: R {suspense_amt:,.2f}</p>

            <hr>

            <h3>Reconciliation</h3>
            <p>Statement: R {total_statement:,.2f}</p>
            <p>Processed: R {processed_total:,.2f}</p>
            <p class="{ 'green' if diff==0 else 'red' }"><b>Difference: R {diff:,.2f}</b></p>
        </div>
        """

        # ✅ CHART
        labels = list(daily_data.keys())
        values = list(daily_data.values())

        chart_html = f"""
        <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
        <canvas id="chart" style="max-width:600px;margin:20px auto;"></canvas>
        <script>
        new Chart(document.getElementById('chart'), {{
            type:'line',
            data:{{
                labels:{labels},
                datasets:[{{label:'Daily Totals',data:{values},borderColor:'blue'}}]
            }}
        }});
        </script>
        """

    # ✅ FINAL PAGE
    return f"""
    <html>
    <head>
    <style>

    body {{
        margin:0;
        font-family:Arial;
        background:#f4f6f9;
    }}

    .sidebar {{
        position:fixed;
        width:220px;
        height:100%;
        background:#0a2540;
        color:white;
        padding:20px;
    }}

    .logo img {{
        max-width:140px;
        display:block;
        margin:auto;
        margin-bottom:20px;
    }}

    .sidebar a {{
        display:block;
        color:white;
        text-decoration:none;
        padding:10px;
    }}

    .sidebar a:hover {{
        background:#144066;
    }}

    .main {{
        margin-left:240px;
        padding:40px;
    }}

    .content {{
        max-width:600px;
        margin:auto;
        text-align:center;
    }}

    .card {{
        background:white;
        padding:25px;
        border-radius:10px;
        margin-top:20px;
        box-shadow:0 2px 6px rgba(0,0,0,0.1);
    }}

    .green {{ color:green; }}
    .red {{ color:red; }}

    </style>
    </head>

    <body>

    <div class="sidebar">
        <div class="logo">
            <img src="/static/logo.png">
        </div>

        <a href="/">Dashboard</a>
        <a href="/download">Download</a>
        <a href="/logout">Logout</a>
    </div>

    <div class="main">

        <div class="content">

            <h1>Bank Allocation System</h1>

            <div class="card">

                <form method="post" enctype="multipart/form-data">
                    <input type="file" name="files" multiple required><br><br>
                    <button style="padding:10px 20px;background:#0078D4;color:white;border:none;border-radius:6px;">Process</button>
                </form>

            </div>

            {summary_html}

            {chart_html}

        </div>

    </div>

    </body>
    </html>
    """

# ✅ DOWNLOAD
@app.route("/download")
def download():
    return send_file(session.get("file"), as_attachment=True)

# ✅ RUN
if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)))