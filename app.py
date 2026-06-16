from flask import Flask, request, send_file, render_template_string
import pandas as pd
import re
import os

app = Flask(__name__)

@app.route("/", methods=["GET", "POST"])
def upload_file():
    if request.method == "POST":

        files = request.files.getlist("files")

        all_valid = []
        all_suspense = []
        total_records = 0

        # ✅ Pattern for extracting matter number
        pattern = re.compile(r'([A-Z]{2,3})\s*(\d{9,10})([A-Z]?)')

        for file in files:
            filename = file.filename.lower()

            df = pd.read_excel(file, engine="openpyxl")

            # ✅ ✅ FINAL DATE FIX (FORCE DD/MM/YYYY)
            for col in df.columns:
                if "date" in col.lower():
                    df[col] = pd.to_datetime(
                        df[col].astype(str),
                        format="%d/%m/%Y",
                        errors='coerce'
                    ).dt.strftime("%Y/%m/%d")

            total_records += len(df)

            # ✅ Bank detection
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

            # ✅ Process rows
            for _, row in df.iterrows():
                description = str(row["Description"]).upper()

                match = pattern.search(description)

                if match:
                    prefix = match.group(1)
                    digits = match.group(2)

                    # ✅ Fix O vs 0 ONLY for 2-letter prefix
                    if len(prefix) == 2 and digits.startswith("0"):
                        digits = "O" + digits[1:]

                    matter_number = prefix + digits

                    row["Matter Number"] = matter_number
                    row["Description 1"] = bank_description

                    row_valid = row.drop(labels=["Description"])
                    all_valid.append(row_valid)

                else:
                    row["Reason"] = "No valid Matter Number found"
                    all_suspense.append(row)

        valid_df = pd.DataFrame(all_valid)
        suspense_df = pd.DataFrame(all_suspense)

        # ✅ ✅ APPLY DATE FIX AGAIN (OUTPUT SAFETY)
        for col in valid_df.columns:
            if "date" in col.lower():
                valid_df[col] = pd.to_datetime(
                    valid_df[col].astype(str),
                    format="%Y/%m/%d",
                    errors='coerce'
                ).dt.strftime("%Y/%m/%d")

        for col in suspense_df.columns:
            if "date" in col.lower():
                suspense_df[col] = pd.to_datetime(
                    suspense_df[col].astype(str),
                    format="%Y/%m/%d",
                    errors='coerce'
                ).dt.strftime("%Y/%m/%d")

        # ✅ Save file
        output_file = "processed_transactions.xlsx"

        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            valid_df.to_excel(writer, sheet_name='Valid', index=False)
            suspense_df.to_excel(writer, sheet_name='Suspense', index=False)

        # ✅ Summary
        valid_count = len(valid_df)
        suspense_count = len(suspense_df)

        total_amount = valid_df.get("Amount", pd.Series()).sum() + suspense_df.get("Amount", pd.Series()).sum()
        valid_amount = valid_df.get("Amount", pd.Series()).sum()
        suspense_amount = suspense_df.get("Amount", pd.Series()).sum()

        valid_preview = valid_df.head(10).to_html(index=False) if not valid_df.empty else "<p>No valid records</p>"
        suspense_preview = suspense_df.head(10).to_html(index=False) if not suspense_df.empty else "<p>No suspense records</p>"

        return render_template_string(f"""
<html>
<head>
<title>Processing Summary</title>
<style>
body {{ font-family: Arial; background:#f4f6f9; text-align:center; }}
.card {{
    background:white;
    padding:30px;
    margin:auto;
    width:900px;
    border-radius:10px;
}}
.stat {{ margin:6px; font-size:16px; }}
.valid {{ color:green; }}
.suspense {{ color:red; }}
.button {{
    background:#0078D4;
    color:white;
    padding:10px 20px;
    text-decoration:none;
    border-radius:5px;
    display:inline-block;
    margin-top:15px;
}}
table {{
    border-collapse:collapse;
    width:100%;
    margin-top:10px;
}}
th, td {{
    border:1px solid #ccc;
    padding:6px;
    font-size:12px;
}}
th {{
    background:#0078D4;
    color:white;
}}
</style>
</head>

<body>
<h1>Processing Summary</h1>

<div class="card">

<div class="stat">Total Records: <b>{total_records}</b></div>
<div class="stat valid">Valid: <b>{valid_count}</b></div>
<div class="stat suspense">Suspense: <b>{suspense_count}</b></div>

<hr>

<div class="stat">Total Amount: <b>R {total_amount:,.2f}</b></div>
<div class="stat valid">Valid Amount: <b>R {valid_amount:,.2f}</b></div>
<div class="stat suspense">Suspense Amount: <b>R {suspense_amount:,.2f}</b></div>

<hr>

<h3>Valid Preview</h3>
{valid_preview}

<h3>Suspense Preview</h3>
{suspense_preview}

<br>

<a href="/download" class="button">Download Combined File</a>

</div>
</body>
</html>
""")

    return """
<html>
<head><title>Bank Allocation Tool</title></head>
<body style="font-family: Arial; text-align:center; background:#f4f6f9;">
<h1>Multi-Bank Allocation Tool</h1>

<form method="post" enctype="multipart/form-data" style="margin-top:40px;">
<input type="file" name="files" multiple required><br><br>
<input type="submit" value="Process Files">
</form>
</body>
</html>
"""

@app.route("/download")
def download():
    return send_file("processed_transactions.xlsx", as_attachment=True)

# ✅ Render-ready config
if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)