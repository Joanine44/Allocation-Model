from flask import Flask, request, send_file, redirect, session, render_template_string
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
        return "<h3>Login failed</h3>"

    return """
    <form method="post" style="text-align:center;margin-top:100px;">
        <h2>Login</h2>
        <input name="username"><br><br>
        <input type="password" name="password"><br><br>
        <button>Login</button>
    </form>
    """

@app.route("/logout")
def logout():
    session.clear()
    return redirect("/login")

# ✅ MAIN SYSTEM
@app.route("/", methods=["GET","POST"])
def index():

    if "user" not in session:
        return redirect("/login")

    chart_script = ""
    report_link = ""

    if request.method == "POST":

        files = request.files.getlist("files")

        pattern = re.compile(r'([A-Z]{2,3})\s*(\d{9,10})')

        valid, suspense = [], []
        total_records = 0
        total_statement = 0
        daily_data = {}

        for file in files:

            df = pd.read_excel(file)

            for col in df.columns:
                if "date" in col.lower():
                    df[col] = pd.to_datetime(df[col], format="%d/%m/%Y", errors="coerce")
                    mask = df[col].isna()
                    df.loc[mask, col] = pd.to_datetime(df.loc[mask,col], errors="coerce")
                    df[col] = df[col].dt.strftime("%Y/%m/%d")

            total_statement += df["Amount"].sum()
            total_records += len(df)

            for _,row in df.iterrows():

                desc = str(row["Description"]).upper()
                matches = pattern.findall(desc)

                date = row["Date"]
                amount = row["Amount"]

                if date not in daily_data:
                    daily_data[date] = 0
                daily_data[date] += amount

                if matches:
                    prefix, digits = matches[0]
                    row["Matter Number"] = prefix + digits
                    valid.append(row.drop(labels=["Description"]))
                else:
                    suspense.append(row)

        valid_df = pd.DataFrame(valid)
        suspense_df = pd.DataFrame(suspense)

        valid_amt = valid_df.get("Amount", pd.Series()).sum()
        suspense_amt = suspense_df.get("Amount", pd.Series()).sum()
        processed_total = valid_amt + suspense_amt
        diff = total_statement - processed_total

        name = f"report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"

        with pd.ExcelWriter(name) as w:
            valid_df.to_excel(w, "Valid", index=False)
            suspense_df.to_excel(w, "Suspense", index=False)

        session["file"] = name
        report_link = "/download"

        # ✅ GRAPH DATA
        labels = list(daily_data.keys())
        values = list(daily_data.values())

        chart_script = f"""
        <canvas id='chart'></canvas>
        <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
        <script>
        new Chart(document.getElementById('chart'), {{
            type: 'line',
            data: {{
                labels: {labels},
                datasets: [{{
                    label: 'Daily Total',
                    data: {values},
                    borderColor: 'blue'
                }}]
            }}
        }});
        </script>
        """

        summary_html = f"""
        <div class="card">
        <h2>Summary</h2>

        <div class="row">
        <div>Total Records<br><b>{total_records}</b></div>
        <div style="color:green;">Valid<br><b>{len(valid_df)}</b></div>
        <div style="color:red;">Suspense<br><b>{len(suspense_df)}</b></div>
        </div>

        <hr>

        <div class="row">
        <div>Total R {processed_total:,.2f}</div>
        <div style="color:green;">Valid R {valid_amt:,.2f}</div>
        <div style="color:red;">Suspense R {suspense_amt:,.2f}</div>
        </div>

        <hr>

        <h3>Reconciliation</h3>
        <p>Statement: R {total_statement:,.2f}</p>
        <p>Processed: R {processed_total:,.2f}</p>
        <p style="color:{'green' if diff==0 else 'red'};"><b>Diff: R {diff:,.2f}</b></p>
        </div>
        """
    else:
        summary_html = ""

    return f"""
    <html>
    <head>
    <style>
    body {{ font-family: Arial; margin:0; }}

    .sidebar {{
        position:fixed;
        width:200px;
        height:100%;
        background:#0a2540;
        color:white;
        padding:20px;
    }}

    .sidebar a {{
        display:block;
        color:white;
        margin:10px 0;
        text-decoration:none;
    }}

    .main {{
        margin-left:220px;
        padding:20px;
    }}

    .header {{
        font-size:20px;
        font-weight:bold;
        margin-bottom:20px;
    }}

    .card {{
        background:#f4f6f9;
        padding:20px;
        margin-bottom:20px;
        border-radius:8px;
    }}

    .row {{
        display:flex;
        justify-content:space-around;
    }}

    button {{
        padding:10px 20px;
    }}
    </style>
    </head>

    <body>

    <div class="sidebar">
        <img src="/static/logo.png" width="150"><br><br>

        <a href="/">Dashboard</a>
        <a href="/download">Download</a>
        <a href="/logout">Logout</a>
    </div>

    <div class="main">

    <div class="header">Bank Allocation System</div>

    <form method="post" enctype="multipart/form-data">
        <input type="file" name="files" multiple required>
        <button>Process</button>
    </form>

    {summary_html}

    {chart_script}

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