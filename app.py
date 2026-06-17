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

# ✅ LOGIN
@app.route("/login", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        u = request.form["username"]
        p = request.form["password"]

        if u in USERS and check_password_hash(USERS[u], p):
            session["user"] = u
            return redirect("/")
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

@app.route("/logout")
def logout():
    session.clear()
    return redirect("/login")


# ✅ MAIN SYSTEM
@app.route("/", methods=["GET", "POST"])
def index():

    if "user" not in session:
        return redirect("/login")

    summary_html = ""
    chart_html = ""

    if request.method == "POST":

        try:
            files = request.files.getlist("files")

            pattern = re.compile(r'([A-Z]{2,3})\s*(\d{9,10})')

            valid, suspense = [], []
            total_records = 0
            total_statement = 0
            daily_data = {}

            for file in files:
                filename = file.filename.lower()
                df = pd.read_excel(file)

                # ✅ CLEAN COLUMN NAMES
                df.columns = df.columns.str.strip()

                # ✅ FIND COLUMNS DYNAMICALLY
                date_col = next((c for c in df.columns if "date" in c.lower()), None)
                desc_col = next((c for c in df.columns if "desc" in c.lower()), None)
                amt_col = next((c for c in df.columns if "amount" in c.lower() or "value" in c.lower()), None)

                if not date_col or not desc_col or not amt_col:
                    return "<h3>ERROR: File must contain Date, Description and Amount columns</h3>"

                # ✅ DATE FIX
                df[date_col] = pd.to_datetime(df[date_col], errors="coerce")
                df[date_col] = df[date_col].dt.strftime("%Y/%m/%d")

                total_statement += df[amt_col].sum()
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

                    desc = str(row[desc_col]).upper()
                    amount = row[amt_col]
                    date = row[date_col]

                    if date:
                        daily_data[date] = daily_data.get(date, 0) + amount

                    matches = pattern.findall(desc)

                    if matches:
                        prefix, digits = matches[0]

                        if len(prefix) == 2 and digits.startswith("0"):
                            digits = "O" + digits[1:]

                        # ✅ CREATE CLEAN ROW
                        clean_row = row.copy()
                        clean_row = clean_row[[date_col, amt_col]]
                        clean_row["Matter Number"] = prefix + digits
                        clean_row["Description 1"] = bank
                        valid.append(clean_row)
                    else:
                        clean_row = row.copy()
                        clean_row = clean_row[[date_col, amt_col]]
                        clean_row["Matter Number"] = ""
                        clean_row["Description 1"] = bank
                        clean_row["Reason"] = "No valid Matter Number"
                        suspense.append(clean_row)
                # continue processing after rows collected
                valid_df = pd.DataFrame(valid)
                suspense_df = pd.DataFrame(suspense)

                # ✅ CLEAN STANDARD COLUMN NAMES
                if not valid_df.empty:
                    valid_df.columns = ["Date", "Amount", "Matter Number", "Description 1"]

                if not suspense_df.empty:
                    suspense_df.columns = ["Date", "Amount", "Matter Number", "Description 1", "Reason"]

                valid_amt = valid_df.get(amt_col, pd.Series(dtype=float)).sum()
                suspense_amt = suspense_df.get(amt_col, pd.Series(dtype=float)).sum()

                processed_total = valid_amt + suspense_amt
                diff = total_statement - processed_total

                # ✅ AUDIT LOG
            log_file = "audit_log.xlsx"
            log_df = pd.DataFrame([{
                "Timestamp": datetime.now(),
                "User": session["user"],
                "Records": total_records,
                "Valid": len(valid_df),
                "Suspense": len(suspense_df),
                "Difference": diff
            }])

            if os.path.exists(log_file):
                old = pd.read_excel(log_file)
                log_df = pd.concat([old, log_df], ignore_index=True)

            log_df.to_excel(log_file, index=False)

            valid_preview = ""
            suspense_preview = ""

            if not valid_df.empty:
                valid_preview = valid_df.head(10).to_html(index=False)
            else:
                valid_preview = "<p>No valid records found</p>"

            if not suspense_df.empty:
                suspense_preview = suspense_df.head(10).to_html(index=False)
            else:
                suspense_preview = "<p>No suspense records found</p>"

            # ✅ SUMMARY
            summary_html = f"""
            <div class="card">
                <h2>Processing Summary</h2>

                <p>Total Records: <b>{total_records}</b></p>
                <p class="green">Valid: <b>{len(valid_df)}</b></p>
                <p class="red">Suspense: <b>{len(suspense_df)}</b></p>

                <hr>

                <p>Total Amount: R {processed_total:,.2f}</p>
                <p class="green">Valid: R {valid_amt:,.2f}</p>
                <p class="red">Suspense: R {suspense_amt:,.2f}</p>

                <hr>

                <h3>Reconciliation</h3>
                <p>Statement: R {total_statement:,.2f}</p>
                <p>Processed: R {processed_total:,.2f}</p>
                <p class="{ 'green' if diff==0 else 'red' }">
                    Difference: R {diff:,.2f}
                </p>
            </div>
            """

            # ✅ GRAPH
            labels = list(daily_data.keys())
            values = list(daily_data.values())

            chart_html = f"""
            https://cdn.jsdelivr.net/npm/chart.js
            <canvas id="chart" style="max-width:600px;margin:20px auto;"></canvas>
            <script>
            new Chart(document.getElementById('chart'), {{
                type: 'line',
                data: {{
                    labels: {labels},
                    datasets: [{{
                        label: 'Daily Totals',
                        data: {values},
                        borderColor: 'blue'
                    }}]
                }}
            }});
            </script>
            """

        except Exception as e:
            summary_html = f"<p style='color:red;'>ERROR: {str(e)}</p>"

    return f"""
   <html>
<head>
<style>

body {
    margin:0;
    font-family:Arial;
    background:#f4f6f9;
}

.sidebar {
    position:fixed;
    width:220px;
    height:100%;
    background:#0a2540;
    padding:20px;
}

.logo img {
    max-width:140px;
    display:block;
    margin:auto;
    margin-bottom:20px;
}

.sidebar a {
    display:block;
    color:white;
    text-decoration:none;
    padding:10px;
}

.sidebar a:hover {
    background:#144066;
}

.main {
    margin-left:240px;
    padding:40px;
}

.content {
    max-width:600px;
    margin:auto;
    text-align:center;
}

.card {
    background:white;
    padding:25px;
    border-radius:10px;
    margin-top:20px;
    box-shadow:0 2px 6px rgba(0,0,0,0.1);
}

.green { color:green; }
.red { color:red; }

/* ✅ SPINNER ANIMATION (FIXED LOCATION) */
@keyframes spin {{
    0% {{ transform: rotate(0deg); }}
    100% {{ transform: rotate(360deg); }}
}}

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

            <!-- ✅ FORM -->
            <form method="post" enctype="multipart/form-data" onsubmit="showLoader()">
                <input type="file" name="files" multiple required><br><br>
                <button id="processBtn">Process</button>
            </form>

            <br>

            <!-- ✅ LOADING SPINNER -->
            <div id="loader" style="display:none; text-align:center; margin-top:20px;">
                
                <div style="
                    border:6px solid #f3f3f3;
                    border-top:6px solid #0078D4;
                    border-radius:50%;
                    width:40px;
                    height:40px;
                    animation: spin 1s linear infinite;
                    margin:auto;
                "></div>

                <p>Processing... Please wait</p>

            </div>

        </div>

    </div>
</div>

<!-- ✅ SCRIPT (MUST BE AT BOTTOM) -->
<script>
function showLoader() {
    document.getElementById("loader").style.display = "block";
    document.getElementById("processBtn").disabled = true;
}
</script>

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