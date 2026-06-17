from flask import Flask, request, send_file, render_template_string, redirect, session
import pandas as pd
import re
import os
from datetime import datetime

app = Flask(__name__)
app.secret_key = "CHANGE_THIS_TO_STRONG_SECRET"

# ✅ MULTI-USER LOGIN
USERS = {
    "joanine": "Duvesco123",
    "leani": "Password456"
}

# ✅ LOGIN PAGE
@app.route("/login", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        username = request.form.get("username")
        password = request.form.get("password")

        if USERS.get(username) == password:
            session["user"] = username
            return redirect("/")
        else:
            return "<h3>Invalid login</h3><a href='/login'>Try again</a>"

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

# ✅ MAIN APP
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

            # ✅ CORRECT DATE FORMAT
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

                    # ✅ Fix O vs 0 for 2-letter prefix
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

        # ✅ UNIQUE FILE NAME (no overwrite)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_file = f"processed_{timestamp}.xlsx"

        with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
            valid_df.to_excel(writer, sheet_name="Valid", index=False)
            suspense_df.to_excel(writer, sheet_name="Suspense", index=False)

        # ✅ ✅ AUDIT LOG
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

        # ✅ STORE FILE FOR DOWNLOAD
        session["last_file"] = output_file

        return f"""
        <html>
        <body style="font-family: Arial; text-align:center;">
            <h2>Processing Complete</h2>

            <p>User: {session['user']}</p>
            <p>Total Records: {total_records}</p>
            <p>Valid: {len(valid_df)}</p>
            <p>Suspense: {len(suspense_df)}</p>

            <br>
            <a href="/download">Download File</a><br><br>
            <a href="/logout">Logout</a>
        </body>
        </html>
        """

    return f"""
    <html>
    <body style="font-family: Arial; text-align:center; margin-top:40px;">
        <h1>Bank Allocation Tool</h1>
        <p>Logged in as: {session['user']}</p>

        <form method="post" enctype="multipart/form-data">
            <input type="file" name="files" multiple required><br><br>
            <input type="submit" value="Process Files">
        </form>

        <br><br>
        <a href="/logout">Logout</a>
    </body>
    </html>
    """

# ✅ DOWNLOAD LAST GENERATED FILE
@app.route("/download")
def download():
    if "last_file" in session:
        return send_file(session["last_file"], as_attachment=True)
    return "No file available"

# ✅ RENDER START
if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)