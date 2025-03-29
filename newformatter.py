import os
import tempfile
from flask import Flask, request, render_template_string, send_file, redirect, url_for, flash
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

app = Flask(__name__)
app.secret_key = "replace_with_a_secure_key"  # Needed for flashing messages

# Temporary folder to store uploaded and processed files
UPLOAD_FOLDER = tempfile.gettempdir()
app.config["UPLOAD_FOLDER"] = UPLOAD_FOLDER

# HTML template for file upload (with Bootstrap styling)
upload_template = """
<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <title>Excel Formatter - Upload</title>
  <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
</head>
<body class="bg-light">
  <div class="container mt-5">
    <h2 class="mb-4">Upload Excel File</h2>
    <form method="post" enctype="multipart/form-data" action="{{ url_for('upload') }}">
      <div class="mb-3">
        <input class="form-control" type="file" name="file" required>
      </div>
      <button type="submit" class="btn btn-primary">Upload</button>
    </form>
    {% with messages = get_flashed_messages() %}
      {% if messages %}
        <div class="mt-3">
          {% for message in messages %}
            <div class="alert alert-danger">{{ message }}</div>
          {% endfor %}
        </div>
      {% endif %}
    {% endwith %}
  </div>
</body>
</html>
"""

# HTML template for header selection (with Bootstrap styling)
select_template = """
<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <title>Select Headers</title>
  <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
  <style>
    .header-checkbox { margin-bottom: 10px; }
  </style>
</head>
<body class="bg-light">
  <div class="container mt-5">
    <h2 class="mb-4">Select the Headers to Keep</h2>
    <form method="post" action="{{ url_for('process') }}">
      <div class="mb-3">
        {% for header in headers %}
          <div class="form-check header-checkbox">
            <input class="form-check-input" type="checkbox" name="selected" value="{{ header }}">
            <label class="form-check-label">{{ header }}</label>
          </div>
        {% endfor %}
      </div>
      <!-- Pass the uploaded filename as a hidden field -->
      <input type="hidden" name="filename" value="{{ filename }}">
      <button type="submit" class="btn btn-success">Process File</button>
    </form>
  </div>
</body>
</html>
"""

def process_excel(in_filepath, selected_headers):
    """
    Process the Excel file:
      1. Keep only the selected headers (in their original order).
      2. Always add a first column "S. No." and fill it with serial numbers.
      3. Convert all text to proper case.
      4. Replace:
           "Bachelors Of Commerce - Commerce" -> "B.Com. (Hons.)"
           "Bachelors Of Arts - Humanities" -> "B.A. Hons. Economics"
           "Cgpa" -> "CGPA"
      5. If a header "Name" exists, sort the data rows alphabetically (case‑insensitively) by that column.
      6. Apply formatting (headers, data, margins, title row with merged cells).
      7. Delete extra rows/columns and hide gridlines.
      8. Auto-adjust column widths (for data columns) so that text does not wrap.
    Returns the path to the processed file.
    """
    wb = openpyxl.load_workbook(in_filepath)
    sheet = wb.active

    # Create new workbook for output
    new_wb = openpyxl.Workbook()
    new_sheet = new_wb.active
    new_sheet.title = "Formatted Data"

    # Build a mapping from header to original column index and list headers in order from the original sheet
    original_headers = {}
    orig_header_list = []
    for col_index, cell in enumerate(sheet[1], start=1):
        if cell.value is not None:
            header_val = cell.value
            original_headers[header_val] = col_index
            orig_header_list.append(header_val)

    # Build new header order:
    # Always add "S. No." as first column,
    # then include headers (in original order) that are selected.
    new_order = {}
    new_order["S. No."] = 1
    col_num = 2
    for header in orig_header_list:
        if header in selected_headers:
            new_order[header] = col_num
            col_num += 1

    # Write header row to new_sheet (row 1)
    for header, new_col in new_order.items():
        new_sheet.cell(row=1, column=new_col, value=header)

    # Copy data for each selected header (skip "S. No.")
    max_row_original = sheet.max_row
    for header, new_col in new_order.items():
        if header == "S. No.":
            continue
        orig_col_index = original_headers[header]
        for row in range(2, max_row_original + 1):
            value = sheet.cell(row=row, column=orig_col_index).value
            new_sheet.cell(row=row, column=new_col, value=value)

    # Determine used area (data columns)
    data_max_col = max(new_order.values())
    data_max_row = new_sheet.max_row

    # --- Text Transformation ---
    def transform_text(text):
        """Convert text to proper case and perform specified replacements."""
        new_text = text.title()  # Proper case conversion
        if new_text == "Bachelors Of Commerce - Commerce":
            return "B.Com. (Hons.)"
        elif new_text == "Bachelors Of Arts - Humanities":
            return "B.A. Hons. Economics"
        elif new_text == "Cgpa":
            return "CGPA"
        elif new_text == "Na":
            return "NA"
        elif new_text == "Roll No":
            return "Roll. No."
        else:
            return new_text

    # Process header row (row 1)
    for col in range(1, data_max_col + 1):
        cell = new_sheet.cell(row=1, column=col)
        if isinstance(cell.value, str):
            cell.value = transform_text(cell.value)

    # Process data rows and build list for sorting.
    data_rows = []
    for row in new_sheet.iter_rows(min_row=2, max_row=data_max_row, max_col=data_max_col, values_only=True):
        row_vals = []
        # For "S. No." column, leave as None (to be filled later)
        row_vals.append(None)
        for val in row[1:]:
            if isinstance(val, str):
                row_vals.append(transform_text(val))
            else:
                row_vals.append(val)
        # If this row has a value in the "Roll No" column, convert it to uppercase.
        if "Roll No" in new_order:
            roll_no_index = new_order["Roll No"] - 1  # zero-indexed position
            if roll_no_index < len(row_vals) and isinstance(row_vals[roll_no_index], str):
                row_vals[roll_no_index] = row_vals[roll_no_index].upper()
        data_rows.append(row_vals)

    # If "Name" is among the selected headers, sort data rows alphabetically by that column.
    if "Name" in new_order:
        name_index = new_order["Name"] - 1
        data_rows.sort(key=lambda r: (r[name_index] or "").lower())

    # Write sorted data rows back to new_sheet (starting at row 2)
    for i, row_data in enumerate(data_rows, start=2):
        for j, value in enumerate(row_data, start=1):
            new_sheet.cell(row=i, column=j, value=value)
    data_max_row = len(data_rows) + 1

    # Fill in Serial Numbers in column 1 (S. No.)
    for i in range(2, data_max_row + 1):
        new_sheet.cell(row=i, column=1, value=i - 1)

    # --- Formatting ---
    header_light_font = Font(name="Trebuchet MS", size=12, bold=True, color="FFFFFF")
    header_font = Font(name="Trebuchet MS", size=11, bold=True, color="000000")
    header_light_fill = PatternFill(start_color="1b3055", end_color="1b3055", fill_type="solid")
    header_fill = PatternFill(start_color="c9daf8", end_color="c9daf8", fill_type="solid")
    header_alignment = Alignment(horizontal="center", vertical="center")
    thin_side = Side(style="thin", color="000000")
    header_border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)

    data_font = Font(name="Trebuchet MS", size=10)
    data_alignment = Alignment(horizontal="center", vertical="center")
    data_border = header_border

    # Format header row (row 1)
    for header, new_col in new_order.items():
        cell = new_sheet.cell(row=1, column=new_col)
        cell.font = header_light_font
        cell.fill = header_light_fill
        cell.alignment = header_alignment
        cell.border = header_border
        new_sheet.column_dimensions[get_column_letter(new_col)].width = 20
    new_sheet.row_dimensions[1].height = 20

    # Format data cells (rows 2 to data_max_row)
    for row in range(2, data_max_row + 1):
        for col in range(1, data_max_col + 1):
            cell = new_sheet.cell(row=row, column=col)
            cell.font = data_font
            cell.alignment = data_alignment
            cell.border = data_border

    # --- Insert Top Margins and Title Row ---
    new_sheet.insert_rows(1, amount=2)
    new_sheet.row_dimensions[1].height = 10
    new_sheet.row_dimensions[3].height = 20
    for col in range(1, data_max_col + 1):
        cell = new_sheet.cell(row=2, column=col)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = header_alignment
        cell.border = header_border

    # --- Insert Left, Right, and Bottom Margins ---
    new_sheet.insert_cols(1)
    new_sheet.column_dimensions["A"].width = 2
    right_margin_index = data_max_col + 2
    right_margin_index1 = data_max_col + 1
    # new_sheet.column_dimensions[get_column_letter(right_margin_index1)].width = 50
    # right_margin_index1.width = 100
    new_sheet.insert_cols(right_margin_index)
    new_sheet.column_dimensions[get_column_letter(right_margin_index)].width = 2
    blank_row = [None] * new_sheet.max_column
    new_sheet.append(blank_row)
    new_sheet.row_dimensions[new_sheet.max_row].height = 10

    used_last_row = new_sheet.max_row
    used_last_col = new_sheet.max_column

    # --- Merge Title Row (Row 2) ---
    new_sheet.merge_cells(start_row=2, start_column=2, end_row=2, end_column=used_last_col)

    # --- Delete Excess Rows and Columns ---
    # total_rows = new_sheet.max_row
    # if total_rows > used_last_row:
    #     new_sheet.delete_rows(used_last_row + 1, total_rows - used_last_row)
    # total_cols = new_sheet.max_column
    # if total_cols > used_last_col:
    #     new_sheet.delete_cols(used_last_col + 1, total_cols - used_last_col)

    # --- Auto-fit Column Widths for Data Columns (skip margin columns) ---
    # Data columns are from col 2 to used_last_col - 1 (right margin)
    for col in range(2, used_last_col+1):
        max_length = 0
        col_letter = get_column_letter(col)
        for row in new_sheet.iter_rows(min_row=1, max_row=used_last_row, min_col=col, max_col=col):
            for cell in row:
                if cell.value:
                    cell_length = len(str(cell.value))
                    if cell_length > max_length:
                        max_length = cell_length
        # If this is the last data column (i.e. col == used_last_col - 1), add extra padding.
        if col == used_last_col - 1:
            new_sheet.column_dimensions[col_letter].width = max_length + 4
        else:
            new_sheet.column_dimensions[col_letter].width = max_length + 2

    # --- Hide Gridlines ---
    new_sheet.sheet_view.showGridLines = False
    new_sheet.page_setup.printArea = f"A1:{get_column_letter(used_last_col)}{used_last_row}"

    # Save processed file to temporary folder.
    out_filepath = os.path.join(app.config["UPLOAD_FOLDER"], "processed.xlsx")
    new_wb.save(out_filepath)
    wb.close()
    new_wb.close()
    return out_filepath

@app.route("/", methods=["GET"])
def index():
    return render_template_string(upload_template)

@app.route("/upload", methods=["POST"])
def upload():
    if "file" not in request.files:
        flash("No file part")
        return redirect(url_for("index"))
    file = request.files["file"]
    if file.filename == "":
        flash("No selected file")
        return redirect(url_for("index"))
    filepath = os.path.join(app.config["UPLOAD_FOLDER"], file.filename)
    file.save(filepath)
    try:
        wb = openpyxl.load_workbook(filepath)
        sheet = wb.active
        headers = [cell.value for cell in sheet[1] if cell.value is not None]
        wb.close()
    except Exception as e:
        flash(f"Error reading Excel file: {e}")
        return redirect(url_for("index"))
    return render_template_string(select_template, headers=headers, filename=file.filename)

@app.route("/process", methods=["POST"])
def process():
    selected_headers = request.form.getlist("selected")
    if not selected_headers:
        flash("Please select at least one header.")
        return redirect(url_for("index"))
    filename = request.form.get("filename")
    filepath = os.path.join(app.config["UPLOAD_FOLDER"], filename)
    try:
        out_filepath = process_excel(filepath, selected_headers)
    except Exception as e:
        flash(f"Error processing file: {e}")
        return redirect(url_for("index"))
    return send_file(out_filepath, as_attachment=True, download_name="Formatted.xlsx")

if __name__ == "__main__":
    app.run(debug=True)
