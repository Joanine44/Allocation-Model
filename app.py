from flask import Flask, request, send_file, redirect, session, render_template_string
import pandas as pd
import re
import os
from datetime import datetime
from werkzeug.security import generate_password_hash, check_password_hash

app = Flask(__name__)
app.secret_key = "CHANGE_THIS_SECRET_KEY"

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
        username = request.form.get("username")
        password = request.form.get("password")

        if username in USERS and check_password_hash(USERS[username], password):
            session["user"] = username
            return redirect("/")
        else:
            return "<h3 style='text-align:center;color:red;'>Invalid login</h3>"

    return """
    <html>
    <body style="font-family:Arial;background:#f4f6f9;">
        <div style="width:350px;margin:120px auto;background:white;padding:30px;border-radius:10px;text-align:center;">
            <h2>Login</h2>
            <form method="post">
                <input name="username" placeholder="Username" required><br><br>
                <input type="password" name="password" placeholder="Password" required><br><br>
                <button style="padding:10px 20px;background:#0078D4;color:white;border:none;">Login</button>
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

# ✅ MAIN SYSTEM
@app.route("/", methods=["GET", "POST"])
def index():

    # 🔐 FORCE LOGIN
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

            # ✅ DATE FIX
            for col in df.columns:
                if "date" in col.lower():
                    df[col] = pd.to_datetime(df[col], format="%d/%m/%Y", errors="coerce")
                    mask = df[col].isna()
                    df.loc[mask, col] = pd.to_datetime(df.loc[mask, col], errors="coerce")
                    df[col] = df[col].dt.strftime("%Y/%m/%d")

            total_statement += df.get("Amount", pd.Series()).sum()
            total_records += len(df)

            # ✅ Bank detection
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
                    p, d = matches[0]

                    if len(p) == 2 and d.startswith("0"):
                        d = "O" + d[1:]

                    row["Matter Number"] = p + d
                    row["Description 1"] = bank
                    valid.append(row.drop(labels=["Description"]))
                else:
                    row["Reason"] = "No valid Matter Number"
                    suspense.append(row)

        valid_df = pd.DataFrame(valid)
        suspense_df = pd.DataFrame(suspense)

        valid_amt = valid_df.get("Amount", pd.Series()).sum()
        suspense_amt = suspense_df.get("Amount", pd.Series()).sum()
        processed_total = valid_amt + suspense_amt
        diff = total_statement - processed_total

        # ✅ SAVE FILE
        output_file = f"processed_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        with pd.ExcelWriter(output_file) as writer:
            valid_df.to_excel(writer, "Valid", index=False)
            suspense_df.to_excel(writer, "Suspense", index=False)

        session["file"] = output_file

        # ✅ SUMMARY UI
        summary_html = f"""
        <div class="card">
            <h2>Processing Summary</h2>

            <p><b>Total Records:</b> {total_records}</p>
            <p style="color:green;"><b>Valid:</b> {len(valid_df)}</p>
            <p style="color:red;"><b>Suspense:</b> {len(suspense_df)}</p>

            <hr>

            <p>Total Amount: R {processed_total:,.2f}</p>
            <p style="color:green;">Valid: R {valid_amt:,.2f}</p>
            <p style="color:red;">Suspense: R {suspense_amt:,.2f}</p>

            <hr>

            <h3>Reconciliation</h3>
            <p>Statement: R {total_statement:,.2f}</p>
            <p>Processed: R {processed_total:,.2f}</p>
            <p style="color:{'green' if diff == 0 else 'red'};">
                <b>Difference: R {diff:,.2f}</b>
            </p>
        </div>
        """

        # ✅ GRAPH
        labels = list(daily_data.keys())
        values = list(daily_data.values())

        chart_html = f"""
        <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
        <canvas id="chart"></canvas>
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

    return f"""
    <html>
    <head>
    <style>
    body {{ margin:0;font-family:Arial;background:#f4f6f9; }}

    .sidebar {{
        width:220px;height:100%;position:fixed;
        background:#0a2540;color:white;padding:20px;
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
        margin:10px 0;
        padding:10px;
        border-radius:5px;
    }}

    .sidebar a:hover {{
        background:#144066;
    }}

    .main {{ margin-left:240px;padding:20px; }}

    .card {{
        background:white;
        padding:20px;
        border-radius:10px;
        margin-top:20px;
    }}
    </style>
    </head>

    <body>

    <div class="sidebar">

        <!-- ✅ FIXED LOGO -->
        <div class="logo">
            <img src="/static/logo.png">
        </div>

        <!-- ✅ FIXED LINKS -->
        <a href="/">Dashboard</a>
        <a href="/download">Download</a>
        <a href="/logout">Logout</a>

    </div>

    <div class="main">

        <h1>Bank Allocation System</h1>

        <form method="post" enctype="multipart/form-data">
            <input type="file" name="files" multiple required>
            <button>Process</button>
        </form>

        {summary_html}

        {chart_html}

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