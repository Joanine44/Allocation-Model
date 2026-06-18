from flask import Flask, request, send_file, redirect, session
import pandas as pd
import re
import os
from datetime import datetime
from werkzeug.security import generate_password_hash, check_password_hash

app = Flask(__name__)
app.debug = True
app.secret_key = "CHANGE_THIS_SECRET"

# ✅ USERS
from werkzeug.security import generate_password_hash, check_password_hash

USERS = {
    "joanine": generate_password_hash("Duvesco123"),
    "leani": generate_password_hash("Password456")
}


# ✅ LOGIN
from flask import request, redirect, session

@app.route("/login", methods=["GET", "POST"])
def login():
   
 # ✅ If already logged in → skip login
    if "user" in session:
        return redirect("/")

    error = ""

    if request.method == "POST":
        u = request.form["username"]
        p = request.form["password"]

        if u in USERS and check_password_hash(USERS[u], p):
            session["user"] = u
            return redirect("/")
        else:
            error = "Invalid username or password"

    return f"""
    <html>
    <body style="font-family:Arial;background:#f4f6f9;">
    
    <div style="
        width:320px;
        margin:120px auto;
        background:white;
        padding:30px;
        border-radius:10px;
        text-align:center;
    ">
    
        <h2>Login</h2>

        <p style="color:red;">{error}</p>

        <form method="post">
            <input name="username" placeholder="Username" required><br><br>
            <input type="password" name="password" placeholder="Password" required><br><br>

            <button style="
                background:#0078D4;
                color:white;
                padding:10px 20px;
                border:none;
                border-radius:6px;
                cursor:pointer;
            ">
                Login
            </button>
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
    page_html = ""

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

            # ✅ CLEAN COLUMN NAMES
            df.columns = df.columns.str.strip()

            # ✅ FIND COLUMNS DYNAMICALLY
            date_col = next((c for c in df.columns if "date" in c.lower()), None)
            desc_col = next((c for c in df.columns if "desc" in c.lower()), None)
            amt_col = next((c for c in df.columns if "amount" in c.lower() or "value" in c.lower()), None)

            if not date_col or not desc_col or not amt_col:
                return "<h3>ERROR: File must contain Date, Description and Amount columns</h3>"

            # ✅ DATE FIX (SA FORMAT)
            df[date_col] = pd.to_datetime(df[date_col], dayfirst=True, errors="coerce")
            df[date_col] = df[date_col].dt.strftime("%Y/%m/%d")
            df[date_col] = df[date_col].astype(str)

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
                    clean_row["Matter Number"] = row[desc_col]
                    clean_row["Description 1"] = bank
                    clean_row["Reason"] = "No valid Matter Number"
                    suspense.append(clean_row)
        
        valid_df = pd.DataFrame(valid)
        suspense_df = pd.DataFrame(suspense)

        # ✅ CREATE TIMESTAMP
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")

        # ✅ CREATE FULL FILE PATH (THIS IS WHERE YOUR LINE GOES)
        output_file = os.path.join(os.getcwd(), f"processed_{timestamp}.xlsx")

        # ✅ WRITE EXCEL FILE
        with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
            valid_df.to_excel(writer, sheet_name="Valid", index=False)
            suspense_df.to_excel(writer, sheet_name="Suspense", index=False)

        # ✅ STORE FILE IN SESSION
        session["file"] = output_file

        # ✅ CLEAN STANDARD COLUMN NAMES
        if not valid_df.empty:
            valid_df.columns = ["Date", "Amount", "Matter Number", "Description 1"]

        if not suspense_df.empty:
            suspense_df.columns = ["Date", "Amount", "Matter Number", "Description 1", "Reason"]

        valid_amt = valid_df["Amount"].sum() if "Amount" in valid_df.columns else 0
        suspense_amt = suspense_df["Amount"].sum() if "Amount" in suspense_df.columns else 0

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

        labels = list(daily_data.keys())
        values = list(daily_data.values())

        chart_html = """
            <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
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

    page_html = """
<html>
<head>
<style>

body {
    margin:0;
    font-family:Arial;
}

.sidebar {
    width:220px;
}

@keyframes spin {
    0% { transform: rotate(0deg); }
    100% { transform: rotate(360deg); }
}

.card {
    margin-top:20px;
    padding:20px;
    border:1px solid #ddd;
    border-radius:10px;
    width:90%;
    max-width:600px;
    text-align:left;
}

.green { color:green; }
.red { color:red; }

</style>
</head>

<body>

<div style="
    display:flex;
    justify-content:center;
    align-items:center;
    min-height:80vh;
    flex-direction:column;
    text-align:center;
">

<h2>Bank Allocation System</h2>

<form method="post" enctype="multipart/form-data" onsubmit="showLoader()">
    <input type="file" name="files" multiple required><br><br>
    <button id="processBtn" style="
        padding:10px 20px;
        border:none;
        background:#0078D4;
        color:white;
        border-radius:6px;
        cursor:pointer;
    ">
        Process
    </button>
</form>
"""

    page_html += summary_html
    page_html += chart_html

    if session.get("file"):
        page_html += """
    <br><br>
    <a href="/download">
        <button style="
            background:#28a745;
            color:white;
            padding:10px 20px;
            border:none;
            border-radius:6px;
            cursor:pointer;
        ">
            Download File
        </button>
    </a>
"""

    page_html += """
<br>
<div id="loader" style="display:none;text-align:center;">
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

<script>
function showLoader() {
    document.getElementById("loader").style.display = "block";
    document.getElementById("processBtn").disabled = true;
}
</script>
</body>
</html>
"""

    return page_html

@app.route("/download")
def download():
    file_path = session.get("file")

    if not file_path or not os.path.exists(file_path):
        return "<h3>File not found</h3>"

    return send_file(
        file_path,
        as_attachment=True,
        download_name=os.path.basename(file_path),
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

# ✅ RUN
if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)))