import os
import json
import pandas as pd
from openpyxl import load_workbook
import xlwings as xw
from uuid import uuid4

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
CONFIG_FILE = os.path.join(BASE_DIR, "config.json")
SETTINGS_FILE = os.path.join(BASE_DIR, "settings.json")

# ------------------- TEMPLATE PATH -------------------
def save_template_path(path):
    rel_path = os.path.relpath(path, BASE_DIR)
    config = {"template_path": rel_path}
    with open(CONFIG_FILE, "w") as f:
        json.dump(config, f)
        
def load_template_path():
    if os.path.exists(CONFIG_FILE):
        with open(CONFIG_FILE, "r") as f:
            config = json.load(f)
            rel_path = config.get("template_path")
            if rel_path:
                abs_path = os.path.join(BASE_DIR, rel_path)
                if os.path.exists(abs_path):
                    return abs_path
    return ""


# ------------------- SETTINGS -------------------
def load_settings():
    default_settings = {
        "timestamp_in_title": True,
        "Q1_color": "#ff0000",
        "Q2_color": "#0000ff",
        "Q3_color": "#00ff00",
        "Q4_color": "#800080",
        "presets": {}
    }
    if os.path.exists(SETTINGS_FILE):
        with open(SETTINGS_FILE, "r") as f:
            try:
                settings = json.load(f)
                return {**default_settings, **settings}
            except json.JSONDecodeError:
                return default_settings
    return default_settings

def save_settings(data):
    with open(SETTINGS_FILE, "w") as f:
        json.dump(data, f, indent=4)

# ------------------- DATA PROCESSING -------------------
def process_8x12_blocks_with_template(csv_path, template_path, output_path, num_pseudotypes, pseudotype_texts, assay_title_text, sample_id_text):
    from math import isnan
    try:
        data = pd.read_csv(csv_path, header=None, dtype=str).dropna(how="all").reset_index(drop=True)
        blocks = [data.iloc[i:i+8, :12] for i in range(0, len(data), 8) if not data.iloc[i:i+8, :12].isnull().all().all()]

        if len(blocks) == 0:
            raise ValueError("No valid 8x12 blocks found.")

        workbook = load_workbook(template_path)
        template_sheet = workbook.active

        sample_ids = [s.strip() for s in sample_id_text.split(",") if s.strip()]
        sample_chunks = [sample_ids[i:i+4] for i in range(0, len(sample_ids), 4)]

        col_positions = {
            1: [2, 5, 8, 11],
            2: [2, 5, 8, 11],
            3: [2, 5, 8, None],
            4: [2, 5, 8, 11],
        }

        for i, block in enumerate(blocks):
            new_sheet = workbook.copy_worksheet(template_sheet)
            new_sheet.title = f"Plate {i + 1}"

            for r_idx, row in enumerate(block.values):
                for c_idx, val in enumerate(row):
                    try:
                        new_sheet.cell(row=5 + r_idx, column=2 + c_idx, value=float(val))
                    except:
                        new_sheet.cell(row=5 + r_idx, column=2 + c_idx, value=val)

            samples = []
            if num_pseudotypes == 2:
                if sample_chunks and sample_chunks[0]:
                    samples.append(sample_chunks[0].pop(0))
                if sample_chunks and sample_chunks[0]:
                    samples.append(sample_chunks[0].pop(0))
            elif num_pseudotypes == 1:
                for _ in range(4):
                    if sample_chunks and sample_chunks[0]:
                        samples.append(sample_chunks[0].pop(0))
            elif num_pseudotypes in (3, 4):
                if sample_chunks and sample_chunks[0]:
                    samples.append(sample_chunks[0].pop(0))

            if sample_chunks and not sample_chunks[0]:
                sample_chunks.pop(0)

            for col in col_positions[num_pseudotypes]:
                if col is None:
                    continue
                new_sheet.merge_cells(start_row=4, start_column=col, end_row=4, end_column=col + 2)
                idx = col_positions[num_pseudotypes].index(col)
                if idx < len(samples):
                    new_sheet.cell(row=4, column=col, value=samples[idx])

            pseudotype_positions = {
                1: ["B3", "E3", "H3", "K3"],
                2: ["B3", "E3", "H3", "K3"],
                3: ["B3", "E3", "H3"],
                4: ["B3", "E3", "H3", "K3"],
            }
            pseudotypes = [p.strip() for p in pseudotype_texts.split(",") if p.strip()] + ["Unlabelled"] * 4

            for idx, cell in enumerate(pseudotype_positions[num_pseudotypes]):
                if num_pseudotypes == 1:
                    new_sheet[cell].value = pseudotypes[0]
                elif num_pseudotypes == 2:
                    new_sheet[cell].value = pseudotypes[idx // 2]
                else:
                    new_sheet[cell].value = pseudotypes[idx]

            if assay_title_text:
                new_sheet["B2"].value = assay_title_text

        del workbook[template_sheet.title]
        workbook.save(output_path)

    except Exception as e:
        raise Exception(f"Data processing error: {e}")

# ------------------- FINAL TITRES -------------------
def extract_final_titres_xlwings(output_path):
    app = None
    try:
        app = xw.App(visible=False)
        wb = app.books.open(output_path)
        if "Summary" not in [s.name for s in wb.sheets]:
            summary = wb.sheets.add("Summary")
        else:
            summary = wb.sheets["Summary"]
            summary.clear()

        summary.range("A1").value = ["Plate", "Pseudotype", "Sample ID", "NT 90% Rep 1", "NT 90% Rep 2", "NT 90% Rep 3", "NT 90%", "NT 50% Rep 1", "NT 50% Rep 2", "NT 50% Rep 3", "NT 50%"]
        row_idx = 2

        for sheet in wb.sheets:
            if sheet.name.startswith("Plate"):
                ranges = {
                    "nt90": ["B14:D14", "E14:G14", "H14:J14", "K14:M14"],
                    "nt50": ["B16:D16", "E16:G16", "H16:J16", "K16:M16"]
                }
                info = sheet.range("F26:I29").value or []

                if not isinstance(info, list):
                    continue

                for i in range(min(4, len(info))):
                    row = info[i] if isinstance(info[i], list) else []
                    while len(row) < 4:
                        row.append("")
                    nt90 = sheet.range(ranges["nt90"][i]).value or ["", "", ""]
                    nt50 = sheet.range(ranges["nt50"][i]).value or ["", "", ""]

                    nt90 = [v if isinstance(v, (int, float)) else "" for v in nt90]
                    nt50 = [v if isinstance(v, (int, float)) else "" for v in nt50]

                    summary.range(f"A{row_idx}").value = [sheet.name, row[0], row[1], *nt90[:3], row[2], *nt50[:3], row[3]]
                    row_idx += 1

        wb.save()
        wb.close()
        app.quit()

        add_default_to_final_titres(output_path)

    except Exception as e:
        app.quit()
        raise Exception(f"Titre extraction error: {e}")

def add_default_to_final_titres(output_path):
    app = xw.App(visible=False)
    try:
        wb = app.books.open(output_path)
        summary = wb.sheets["Summary"]
        a5_val = round(wb.sheets["Plate 1"].range("A5").value)
        last_row = summary.range("A" + str(summary.cells.last_cell.row)).end("up").row

        for col in range(4, 11):
            for row in range(2, last_row + 1):
                cell = summary.range((row, col))
                if not cell.value:
                    cell.value = f"<{a5_val}"

        summary.range(f"A2:K{last_row}").number_format = "0"
        summary.range(f"A1:K{last_row}").api.HorizontalAlignment = -4108
        summary.range(f"A1:K{last_row}").api.VerticalAlignment = -4108

        wb.save()
        wb.close()
        app.quit()

    except Exception as e:
        app.quit()
        raise Exception(f"Default fill error: {e}")