from flask import Flask, render_template, request, send_file
import pandas as pd
import io
import re

app = Flask(__name__)

# --- Helpers ---
def clean_code(s):
    s = str(s).replace('\u200b', '').strip()
    s = re.sub(r'[^A-Za-z0-9]', '', s)
    return s.upper()

def clean_mobile(s):
    s = str(s).replace('\u200b', '').strip()
    s = re.sub(r'\D', '', s)
    if len(s) > 10:
        s = s[-10:]
    return s

# --- Routes ---
@app.route('/')
def index():
    return render_template('index.html')

@app.route('/compare', methods=['POST'])
def compare():
    file = request.files['file']
    pasted_data = request.form['pasted_data']

    # ✅ 1️⃣ Read all sheets and build ClientCodes + Mobile numbers set
    xl = pd.ExcelFile(file)
    reference_codes = set()
    reference_mobiles = set()

    for sheet_name in xl.sheet_names:
        df = xl.parse(sheet_name, header=None, usecols=[0, 1])
        df[0] = df[0].apply(clean_code)
        df[1] = df[1].apply(clean_mobile)

        reference_codes.update(df[0].dropna().unique())
        reference_mobiles.update(df[1].dropna().unique())

    # ✅ 2️⃣ Prepare pasted data
    rows = pasted_data.strip().split('\n')
    master = []
    for row in rows:
        parts = row.strip().split()
        if len(parts) >= 2:
            code = clean_code(parts[0])
            mobile = clean_mobile(parts[1])
            master.append({'ClientCode': code, 'Mobile': mobile})

    # ✅ 3️⃣ Compare
    matched = []
    unmatched = []

    for entry in master:
        if entry['ClientCode'] in reference_codes or entry['Mobile'] in reference_mobiles:
            matched.append(entry)
        else:
            unmatched.append(entry)

    return render_template(
        'index.html',
        matched=matched,
        unmatched=unmatched,
        matched_count=len(matched),
        unmatched_count=len(unmatched),
        pasted_data=pasted_data
    )

@app.route('/download', methods=['POST'])
def download():
    matched = request.form['matched'].split('||')
    unmatched = request.form['unmatched'].split('||')
    file_name = request.form['file_name']

    matched_rows = [x.split('|') for x in matched if x]
    unmatched_rows = [x.split('|') for x in unmatched if x]

    matched_df = pd.DataFrame(matched_rows, columns=['ClientCode', 'Mobile'])
    unmatched_df = pd.DataFrame(unmatched_rows, columns=['ClientCode', 'Mobile'])

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        matched_df.to_excel(writer, index=False, sheet_name='Matched')
        unmatched_df.to_excel(writer, index=False, sheet_name='NotMatched')
    output.seek(0)

    return send_file(
        output,
        as_attachment=True,
        download_name=f"{file_name}.xlsx",
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )

if __name__ == '__main__':
    app.run(debug=True)
