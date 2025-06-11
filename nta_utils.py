import os
import json
import csv
import openpyxl
from openpyxl.styles import Alignment
from openpyxl.utils import get_column_letter
from openpyxl import load_workbook

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# Load config on module load
CONFIG_PATH = os.path.join(BASE_DIR, "config.json")
SETTINGS_PATH = os.path.join(BASE_DIR, "settings.json")

def load_config():
    if os.path.exists(CONFIG_PATH):
        with open(CONFIG_PATH, "r") as f:
            return json.load(f)
    return {}

config = load_config()

TEMPLATE_PATH = os.path.join(BASE_DIR, config.get("template_path", ""))

def load_csv_blocks(csv_path):
    blocks = []
    current_block = []
    with open(csv_path, newline='') as f:
        reader = csv.reader(f)
        for row in reader:
            # Detect blank line = new block
            if not any(cell.strip() for cell in row):
                if current_block:
                    blocks.append(current_block)
                    current_block = []
            else:
                # Limit columns to first 12
                current_block.append(row[:12])
        if current_block:
            blocks.append(current_block)
    return blocks

def process_csv_to_template(
    csv_path,
    template_path,
    output_path,
    num_pseudotypes,
    pseudotype_texts,
    assay_title_text,
    sample_id_text
):
    blocks = load_csv_blocks(csv_path)
    if not os.path.exists(template_path):
        raise FileNotFoundError(f"Template not found at {template_path}")

    wb = openpyxl.load_workbook(template_path)
    template_sheet = wb.active

    # Split pseudotypes and sample_ids properly by line or comma, strip spaces
    pseudotype_list = [pt.strip() for line in pseudotype_texts.splitlines() for pt in line.split(",") if pt.strip()]
    sample_id_list = [sid.strip() for line in sample_id_text.splitlines() for sid in line.split(",") if sid.strip()]


    sample_index = 0

    for i, block in enumerate(blocks):
        sheet_title = f"Plate{i+1}"
        new_sheet = wb.copy_worksheet(template_sheet)
        new_sheet.title = sheet_title

        # Paste block data into rows 5–12, columns 2–13 (B5:M12)
        for r in range(8):
            for c in range(12):
                cell = new_sheet.cell(row=5 + r, column=2 + c)
                try:
                    val = block[r][c]
                    cell.value = float(val) if val.replace('.', '', 1).isdigit() else val
                except IndexError:
                    cell.value = ""

        ws = new_sheet

        # Set assay title in B2
        ws['B2'] = assay_title_text

        # Apply pseudotype and sample_id placement logic
        if num_pseudotypes == 1:
            val = pseudotype_list[0] if len(pseudotype_list) > 0 else ''
            for cell in ['B3', 'E3', 'H3', 'K3']:
                ws[cell] = val
            sample_cells = ['B4', 'E4', 'H4', 'K4']
            for cell in sample_cells:
                if sample_index < len(sample_id_list):
                    ws[cell] = sample_id_list[sample_index]
                    sample_index += 1
                else:
                    ws[cell] = ''
        elif num_pseudotypes == 2:
            ws['B3'] = pseudotype_list[0] if len(pseudotype_list) > 0 else ''
            ws['E3'] = pseudotype_list[0] if len(pseudotype_list) > 0 else ''
            ws['H3'] = pseudotype_list[1] if len(pseudotype_list) > 1 else ''
            ws['K3'] = pseudotype_list[1] if len(pseudotype_list) > 1 else ''
            val1 = sample_id_list[sample_index] if sample_index < len(sample_id_list) else ''
            if sample_index < len(sample_id_list): sample_index += 1
            val2 = sample_id_list[sample_index] if sample_index < len(sample_id_list) else ''
            if sample_index < len(sample_id_list): sample_index += 1

            ws['B4'] = val1
            ws['H4'] = val1
            ws['E4'] = val2
            ws['K4'] = val2

        elif num_pseudotypes == 3:
            ws['B3'] = pseudotype_list[0] if len(pseudotype_list) > 0 else ''
            ws['E3'] = pseudotype_list[1] if len(pseudotype_list) > 1 else ''
            ws['H3'] = pseudotype_list[2] if len(pseudotype_list) > 2 else ''
            ws['K3'] = ''
            val = sample_id_list[sample_index] if sample_index < len(sample_id_list) else ''
            if sample_index < len(sample_id_list): sample_index += 1
            ws['B4'] = val
            ws['E4'] = val
            ws['H4'] = val
        elif num_pseudotypes == 4:
            for idx, cell in enumerate(['B3', 'E3', 'H3', 'K3']):
                ws[cell] = pseudotype_list[idx] if idx < len(pseudotype_list) else ''
            val = sample_id_list[sample_index] if sample_index < len(sample_id_list) else ''
            if sample_index < len(sample_id_list): sample_index += 1
            for cell in ['B4', 'E4', 'H4', 'K4']:
                ws[cell] = val
        else:
            for cell in ['B3', 'E3', 'H3', 'K3', 'B4', 'E4', 'H4', 'K4']:
                ws[cell] = ''


    wb.remove(template_sheet)  # Remove original template sheet
    wb.save(output_path)


def extract_final_titres_openpyxl(output_path):
    wb = load_workbook(output_path)

    if "Summary" in wb.sheetnames:
        wb.remove(wb["Summary"])

    summary_ws = wb.create_sheet("Summary", 0)

    summary_ws.append([
        "Plate", "Pseudotype", "Sample ID", 
        "NT 90% Replicate 1", "NT 90% Replicate 2", "NT 90% Replicate 3", "NT 90%",
        "NT 50% Replicate 1", "NT 50% Replicate 2", "NT 50% Replicate 3", "NT 50%"
    ])

    # Define cell addresses on the plate sheets
    nt90_cells = [
        ["B14", "C14", "D14"],
        ["E14", "F14", "G14"],
        ["H14", "I14", "J14"],
        ["K14", "L14", "M14"],
    ]
    nt50_cells = [
        ["B16", "C16", "D16"],
        ["E16", "F16", "G16"],
        ["H16", "I16", "J16"],
        ["K16", "L16", "M16"],
    ]
    pseudotype_cells = ["B3", "E3", "H3", "K3"]
    sample_id_cells = ["B4", "E4", "H4", "K4"]
    nt90_avg_cells = ["D14", "G14", "J14", "M14"]
    nt50_avg_cells = ["D16", "G16", "J16", "M16"]

    for sheet_name in wb.sheetnames:
        if not sheet_name.startswith("Plate"):
            continue

        for i in range(4):  # For up to 4 pseudotypes
            pt_formula = f"={sheet_name}!{pseudotype_cells[i]}"
            sid_formula = f"={sheet_name}!{sample_id_cells[i]}"
            nt90_formulas = [f"={sheet_name}!{cell}" for cell in nt90_cells[i]]
            nt90_avg = f"={sheet_name}!{nt90_avg_cells[i]}"
            nt50_formulas = [f"={sheet_name}!{cell}" for cell in nt50_cells[i]]
            nt50_avg = f"={sheet_name}!{nt50_avg_cells[i]}"


            summary_ws.append([
                sheet_name,
                pt_formula,
                sid_formula,
                *nt90_formulas,
                nt90_avg,
                *nt50_formulas,
                nt50_avg
            ])

            last_row = summary_ws.max_row
            for col in range(4, 12):  # Columns D to K
                cell = summary_ws.cell(row=last_row, column=col)
                cell.number_format = '0'

    wb.save(output_path)


    add_default_to_final_titres(output_path)


def add_default_to_final_titres(output_path):
    wb = openpyxl.load_workbook(output_path)
    summary = wb["Summary"]
    plate1 = wb["Plate1"]
    a5_val = plate1["A5"].value

    try:
        a5_val = round(float(a5_val))
    except:
        a5_val = ""

    for row in summary.iter_rows(min_row=2, max_row=summary.max_row, min_col=4, max_col=10):
        for cell in row:
            if cell.value in (None, "") and a5_val:
                cell.value = f"<{a5_val}"

    # Format numeric cells as integers and center-align all summary cells
    for row in summary.iter_rows(min_row=2, max_row=summary.max_row, min_col=1, max_col=11):
        for cell in row:
            try:
                if isinstance(cell.value, (float, int)):
                    cell.value = int(round(cell.value))

            except:
                pass
            cell.alignment = Alignment(horizontal="center", vertical="center")

    # Set column widths
    for col in range(1, 12):
        summary.column_dimensions[get_column_letter(col)].width = 15

    wb.save(output_path)

def save_template_path(path, config_file=CONFIG_PATH):
    config = load_config()
    config["template_path"] = path
    with open(config_file, "w") as f:
        json.dump(config, f, indent=4)

def load_template_path(config_file=CONFIG_PATH):
    config = load_config()
    template_path = config.get("template_path")
    if not template_path or not os.path.exists(template_path):
        raise FileNotFoundError("Saved template path not found or does not exist.")
    return template_path

def save_settings(settings, settings_file=SETTINGS_PATH):
    with open(settings_file, "w") as f:
        json.dump(settings, f, indent=4)

def load_settings(settings_file=SETTINGS_PATH):
    if os.path.exists(settings_file):
        with open(settings_file, "r") as f:
            return json.load(f)
    return {}
