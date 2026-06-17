from flask import Flask, request, send_file, render_template_string, redirect, session
import pandas as pd
import re
import os
from datetime import datetime
from werkzeug.security import generate_password_hash, check_password_hash

app = Flask(__name__)
app.secret_key = "VERY_STRONG_SECRET_KEY_CHANGE_THIS"

# ✅ USERS (SECURE)
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

        stored_password = USERS.get(username)

        if stored_password and check_password_hash(stored_password, password):
            session["user"] = username
            return redirect("/")
        else:
            return "<h3>Invalid login</h3>/loginTry again</a>"

    return """
    <html><body style="font-family: Arial; text-align:center; margin-top:100px;">
        <h2>Login</h2>
        <form method="post">
            <input type="text" name="username" placeholder="Username" required><br><br>
            <input type="password" name="password" placeholder="Password" required><br><br>
            <input type="submit" value="Login">
        </form>
    </body></html>
    """

# ✅ LOGOUT
@app.route("/logout")
def logout():
    session.clear()
    return redirect("/login")

# ✅ MAIN SYSTEM
@app.route("/", methods=["GET", "POST"])
def upload_file():

    if "user" not in session:
        return redirect("/login")

    if request.method == "POST":

        files = request.files.getlist("files")

        all_valid = []
        all_suspense = []
        total_records = 0

        pattern = re.compile(r'([A-Z]{2,3})\s*(\d{9,10})([A-Z]?)')

        for file in files:
            filename = file.filename.lower()
            df = pd.read_excel(file, engine="openpyxl")

            # ✅ CORRECT DATE FORMAT (SA FIX)
            for col in df.columns:
                if "date" in col.lower():
                    df[col] = pd.to_datetime(
                        df[col].astype(str),
                        format="%d/%m/%Y",
                        errors='coerce'
                    ).dt.strftime("%Y/%m/%d")

            total_records += len(df)

            # ✅ BANK DETECTION
            if "4174" in filename:
                bank_description = "Std B 4174"
            elif "6831" in filename:
                bank_description = "Std B 6831"
            elif "1039" in filename:
                bank_description = "Std B 1039"
            elif "8868" in filename:
                bank_description = "Capitec 8868"
            else:
                bank_description = "Unknown Bank"

            for _, row in df.iterrows():
                description = str(row["Description"]).upper()
                match = pattern.search(description)

                if match:
                    prefix = match.group(1)
                    digits = match.group(2)

                    if len(prefix) == 2 and digits.startswith("0"):
                        digits = "O" + digits[1:]

                    row["Matter Number"] = prefix + digits
                    row["Description 1"] = bank_description
                    all_valid.append(row.drop(labels=["Description"]))
                else:
                    row["Reason"] = "No valid Matter Number found"
                    all_suspense.append(row)

        valid_df = pd.DataFrame(all_valid)
        suspense_df = pd.DataFrame(all_suspense)

        # ✅ UNIQUE FILE
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_file = f"processed_{timestamp}.xlsx"

        with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
            valid_df.to_excel(writer, sheet_name="Valid", index=False)
            suspense_df.to_excel(writer, sheet_name="Suspense", index=False)

        # ✅ AUDIT LOG
        log_file = "audit_log.xlsx"

        log_data = {
            "Timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "User": session["user"],
            "Files Uploaded": len(files),
            "Total Records": total_records,
            "Valid Records": len(valid_df),
            "Suspense Records": len(suspense_df)
        }

        log_df = pd.DataFrame([log_data])

        if os.path.exists(log_file):
            existing_log = pd.read_excel(log_file)
            log_df = pd.concat([existing_log, log_df], ignore_index=True)

        log_df.to_excel(log_file, index=False)

        session["last_file"] = output_file

        # ✅ DASHBOARD DATA
        valid_count = len(valid_df)
        suspense_count = len(suspense_df)

        total_amount = valid_df.get("Amount", pd.Series()).sum() + suspense_df.get("Amount", pd.Series()).sum()
        valid_amount = valid_df.get("Amount", pd.Series()).sum()
        suspense_amount = suspense_df.get("Amount", pd.Series()).sum()

        valid_preview = valid_df.head(10).to_html(index=False) if not valid_df.empty else "<p>No valid records</p>"
        suspense_preview = suspense_df.head(10).to_html(index=False) if not suspense_df.empty else "<p>No suspense records</p>"

        return f"""
        <html>
        <head>
        <style>
        body {{ font-family: Arial; background:#f4f6f9; text-align:center; }}
        .card {{
            background:white;
            padding:30px;
            margin:auto;
            width:900px;
            border-radius:10px;
        }}
        .valid {{ color:green; }}
        .suspense {{ color:red; }}
        table {{ border-collapse:collapse; width:100%; }}
        th, td {{ border:1px solid #ccc; padding:6px; }}
        th {{ background:#0078D4; color:white; }}
        </style>
        </head>

        <body>
        <h1>Processing Summary</h1>

        <div class="card">
        <p>User: <b>{session['user']}</b></p>
        <p>Total Records: <b>{total_records}</b></p>
        <p class="valid">Valid: <b>{valid_count}</b></p>
        <p class="suspense">Suspense: <b>{suspense_count}</b></p>

        <hr>

        <p>Total Amount: <b>R {total_amount:,.2f}</b></p>
        <p class="valid">Valid Amount: <b>R {valid_amount:,.2f}</b></p>
        <p class="suspense">Suspense Amount: <b>R {suspense_amount:,.2f}</b></p>

        <hr>

        <h3>Valid Preview</h3>
        {valid_preview}

        <h3>Suspense Preview</h3>
        {suspense_preview}

        <br>
        /downloadDownload File</a><br><br>
        /adminAdmin Dashboard</a> |
        /logoutLogout</a>

        </div>
        </body>
        </html>
        """

    return f"""
    <html><body style="text-align:center;">
    <h1>Bank Allocation Tool</h1>
    <p>Logged in as: {session['user']}</p>

    <form method="post" enctype="multipart/form-data">
        <input type="file" name="files" multiple required><br><br>
        <input type="submit" value="Process Files">
    </form>

    <br>
    /adminAdmin Dashboard</a> |
    /logoutLogout</a>
    </body></html>
    """

# ✅ DOWNLOAD
@app.route("/download")
def download():
    if "last_file" in session:
        return send_file(session["last_file"], as_attachment=True)
    return "No file available"

# ✅ ADMIN
@app.route("/admin")
def admin():

    if "user" not in session or session["user"] != ADMIN_USER:
        return "Access denied"

    log_file = "audit_log.xlsx"

    if not os.path.exists(log_file):
        return "No audit data yet"

    df = pd.read_excel(log_file)

    return f"""
    <html><body style="font-family: Arial;">
    <h2>Audit Log</h2>
    {df.to_html(index=False)}
    /Back</a>
    </body></html>
    """

# ✅ RUN
if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)