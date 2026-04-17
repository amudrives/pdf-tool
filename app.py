from flask import Flask, render_template, request, send_file
import pdfplumber
import re
import pandas as pd

app = Flask(__name__)

# -------- FINAL LIST --------
def extract_final(pdf):
    data = []
    with pdfplumber.open(pdf) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            if text:
                matches = re.findall(r'(\d{7})\s+([A-Z]+)', text)
                data.extend(matches)
    return data

# -------- SELECTED --------
def extract_selected(pdf):
    data = {}
    with pdfplumber.open(pdf) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            if text:
                matches = re.findall(r'(\d{7})\s+(S\d+)', text)
                for roll, rank in matches:
                    data[roll] = rank
    return data

# -------- WAITING --------
def extract_waiting(pdf):
    data = {}
    with pdfplumber.open(pdf) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            if text:
                matches = re.findall(r'(\d{7})\s+(C\d+)', text)
                for roll, rank in matches:
                    data[roll] = rank
    return data

# -------- SPECIAL --------
def extract_special(pdf):
    data = {}
    with pdfplumber.open(pdf) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            if text:
                matches = re.findall(r'(\d{7})\s+([A-Z]+)(\d+)([#\$])', text)
                for roll, cat, num, sign in matches:
                    data[roll] = cat + num
    return data

# -------- HOME --------
@app.route('/')
def home():
    return render_template('index.html')

# -------- GENERATE --------
@app.route('/generate', methods=['POST'])
def generate():

    final_pdf = request.files['final']
    selected_pdf = request.files['selected']
    waiting_pdf = request.files['waiting']
    special_pdf = request.files['special']

    branch_filter = request.form.get("branch")
    status_filter = request.form.get("status")

    final_data = extract_final(final_pdf)
    selected_data = extract_selected(selected_pdf)
    waiting_data = extract_waiting(waiting_pdf)
    special_data = extract_special(special_pdf)

    result = []

    for roll, branch in final_data:

        sel = selected_data.get(roll, "-")
        wait = waiting_data.get(roll, "-")
        cat = special_data.get(roll, "-")

        status_list = []

        if sel != "-":
            status_list.append("Selected")
        if wait != "-":
            status_list.append("Waiting")
        if cat != "-":
            status_list.append("Special")

        status = ", ".join(status_list) if status_list else "Not Found"

        # FILTERS
        if branch_filter and branch_filter != "ALL":
            if branch != branch_filter:
                continue

        if status_filter and status_filter != "ALL":
            if status != status_filter:
                continue

        result.append([roll, branch, sel, wait, cat, status])

    # -------- EXCEL OUTPUT --------
    df = pd.DataFrame(result, columns=[
        "Roll No", "Branch", "Selected Rank", "Waiting Rank", "Category", "Status"
    ])

    file = "output.xlsx"
    df.to_excel(file, index=False)

    return send_file(file, as_attachment=True)


if __name__ == '__main__':
    app.run(debug=True)