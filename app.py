from flask import Flask, request, send_file, render_template_string, redirect, url_for, session
import pandas as pd
import re
import os

app = Flask(__name__)
app.secret_key = "supersecretkey"  # 🔴 Change this in production

# ✅ SIMPLE LOGIN PASSWORD
PASSWORD = "Duvesco123"   # 🔴 Change this

# ✅ LOGIN ROUTE
@app.route("/login", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        password = request.form.get("password")

        if password == PASSWORD:
            session["logged_in"] = True
            return redirect("/")
        else:
            return render_template_string("""
            <h3>Invalid password</h3>
            <a href="/login">Try again</a>
            """)

    return """
    <html>
    <body style="font-family: Arial; text-align:center; margin-top:100px;">
        <h2>Login Required</h2>
        <form method="post">
            <input type="password" name="password" placeholder="Enter password" required><br><br>
            <input type="submit" value="Login">
        </form>
    </body>
    </html>
    """

# ✅ LOGOUT
@app.route("/logout")
def logout():
    session.clear()
    return redirect("/login")

# ✅ MAIN ROUTE (PROTECTED)
@app.route("/", methods=["GET", "POST"])
def upload_file():

    # ✅ CHECK LOGIN
    if not session.get("logged_in"):
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

            # ✅ FIX DATE FORMAT
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

        # ✅ SAVE FILE
        output_file = "processed_transactions.xlsx"
        with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
            valid_df.to_excel(writer, sheet_name="Valid", index=False)
            suspense_df.to_excel(writer, sheet_name="Suspense", index=False)

        valid_preview = valid_df.head(10).to_html(index=False)
        suspense_preview = suspense_df.head(10).to_html(index=False)

        return f"""
        <html>
        <body style="font-family: Arial; text-align:center;">
        <h1>Processing Complete</h1>

        <div>Total Records: {total_records}</div>
        <div>Valid: {len(valid_df)}</div>
        <div>Suspense: {len(suspense_df)}</div>

        <h3>Valid Preview</h3>
        {valid_preview}

        <h3>Suspense Preview</h3>
        {suspense_preview}

        <br>
        <a href="/download">Download File</a><br><br>
        <a href="/logout">Logout</a>
        </body>
        </html>
        """

    return """
    <html>
    <body style="font-family: Arial; text-align:center; margin-top:40px;">
        <h1>Bank Allocation Tool</h1>

        <form method="post" enctype="multipart/form-data">
            <input type="file" name="files" multiple required><br><br>
            <input type="submit" value="Process Files">
        </form>

        <br><br>
        <a href="/logout">Logout</a>
    </body>
    </html>
    """

@app.route("/download")
def download():
    return send_file("processed_transactions.xlsx", as_attachment=True)

# ✅ RENDER CONFIG
if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)