from flask import Flask, render_template, request, redirect, url_for, send_file, flash, jsonify, send_from_directory, session
import os
import uuid
import json
import subprocess
from datetime import datetime
from io import BytesIO
import tempfile
from openpyxl import load_workbook
from openpyxl.drawing.image import Image as XLImage 
from PIL import Image as PILImage  


from nta_utils import (
    process_csv_to_template,
    extract_final_titres_openpyxl as extract_final_titres_xlwings,
    save_template_path,
    load_template_path,
    load_settings,
    save_settings,
    generate_sigmoid_csv,
    flag_triplicate_errors,
    count_errors_from_workbook,
    validate_csv_mode,
)

in_memory_files = {}  # Key: UUID, Value: BytesIO

app = Flask(__name__)
app.secret_key = "your-secret-key"



@app.route("/")
def index():
    settings = load_settings()
    settings["template_path"] = load_template_path()
    return render_template("index.html", settings=settings)


@app.route("/help")
def help_page():
    return "", 204

@app.route("/save_timestamp_setting", methods=["POST"])
def save_timestamp_setting():
    data = request.get_json()
    settings = load_settings()
    settings["timestamp_in_filename"] = bool(data.get("enabled", False))
    save_settings(settings)
    return jsonify({"status": "ok"})

@app.route("/settings", methods=["GET", "POST"])
def settings():
    current_settings = load_settings()

    default_templates = {
        "NTA Template (dil 50-36450)": "excel_templates/NTA_Template.xlsx",
        "Measles Template (dil 32-65536)": "excel_templates/Measles_NTA_Template.xlsx",
        "Backup NTA Template": "excel_templates/Backup_NTA_Template.xlsx",
    }

    try:
        current_template_path = load_template_path()
    except FileNotFoundError:
        current_template_path = None

    if request.method == "POST":
        selected_template_key = request.form.get("default_template_select")
        if selected_template_key in default_templates:
            selected_template_path = default_templates[selected_template_key]
            save_template_path(selected_template_path)
            flash(f"Template set to {selected_template_key}.", "success")
        else:
            file = request.files.get("template_file")
            if file and file.filename.endswith(".xlsx"):
                filename = f"template_{uuid.uuid4().hex}.xlsx"
                filepath = os.path.join("excel_templates", filename)
                os.makedirs("excel_templates", exist_ok=True)
                file.save(filepath)
                save_template_path(filepath)
                flash("New template uploaded and path updated.", "success")

        timestamp_flag = request.form.get("timestamp_in_filename") == "on"
        error_flagging_flag = request.form.get("error_flagging") == "on"
        new_settings = current_settings.copy()
        new_settings["timestamp_in_filename"] = timestamp_flag
        new_settings["error_flagging"] = error_flagging_flag
        save_settings(new_settings)
        flash("Settings saved.", "success")
        return redirect(url_for("settings"))

    current_settings["template_path"] = current_template_path

    return render_template("settings.html", settings=current_settings, default_templates=default_templates)


@app.route("/process", methods=["POST"])
def process():
    file = request.files["csv_file"]
    if not file:
        flash("No CSV file uploaded.", "danger")
        return redirect(url_for("index"))

    assay_title = request.form.get("assay_title", "")
    pseudotypes = request.form.get("pseudotype_text", "").strip()
    sample_ids = request.form.get("sample_id_text", "")
    data_mode = request.form.get("data_mode", "standard")

    if data_mode not in ("data_only", "standard"):
        data_mode = "standard"

    if not pseudotypes:
        flash("Please enter at least one pseudotype name.", "danger")
        return redirect(url_for("index"))

    try:
        num_pseudotypes = int(request.form.get("num_pseudotypes", "1"))
        if num_pseudotypes not in [1, 2, 3, 4]:
            raise ValueError()
    except ValueError:
        flash("Invalid pseudotype count. Must be 1–4.", "danger")
        return redirect(url_for("index"))

    settings = load_settings()
    safe_title = assay_title.strip().replace(" ", "_")
    timestamp = datetime.now().strftime("%Y-%m-%d")
    filename = f"{safe_title}_{timestamp}.xlsx" if settings.get("timestamp_in_filename", True) else f"{safe_title}.xlsx"

    csv_bytes = BytesIO(file.read())
    csv_bytes.seek(0)

    try:
        template_path = load_template_path()
    except Exception as e:
        flash(str(e), "danger")
        return redirect(url_for("index"))

    output_bytes = BytesIO()
    process_csv_to_template(
        csv_path=csv_bytes,
        template_path=template_path,
        output_path=output_bytes,
        num_pseudotypes=num_pseudotypes,
        pseudotype_texts=pseudotypes,
        assay_title_text=assay_title,
        sample_id_text=sample_ids,
        data_mode=data_mode,
    )

    extract_final_titres_xlwings(output_bytes)

    from nta_utils import add_default_to_final_titres
    add_default_to_final_titres(output_bytes)

    # ── Error flagging (if enabled in settings) ──
    if settings.get("error_flagging", False):
        flag_triplicate_errors(output_bytes)

    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp_excel:
        tmp_excel.write(output_bytes.getvalue())
        excel_path = tmp_excel.name

    r_script = os.path.join(os.getcwd(), "process_data.R")
    presets = settings.get("presets", {})
    active_preset_name = settings.get("selected_preset", None)
    default_colours = {"Q1": "#ff7e79", "Q2": "#ffd479", "Q3": "#009193", "Q4": "#d783ff"}
    colours = presets.get(active_preset_name, default_colours)

    q1_colour = colours.get("Q1", default_colours["Q1"])
    q2_colour = colours.get("Q2", default_colours["Q2"])
    q3_colour = colours.get("Q3", default_colours["Q3"])
    q4_colour = colours.get("Q4", default_colours["Q4"])

    quadrants = settings.get("quadrants", {"Q1": True, "Q2": True, "Q3": True, "Q4": True})
    q1_flag = str(quadrants.get("Q1", True)).lower()
    q2_flag = str(quadrants.get("Q2", True)).lower()
    q3_flag = str(quadrants.get("Q3", True)).lower()
    q4_flag = str(quadrants.get("Q4", True)).lower()

    plot_title = os.path.splitext(filename)[0]
    with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as tmp_png:
        output_plot_path = tmp_png.name

    subprocess.run(
        [
            "Rscript", r_script,
            excel_path, output_plot_path,
            str(settings.get("timestamp_in_filename", True)).lower(),
            q1_colour, q2_colour, q3_colour, q4_colour,
            plot_title,
            q1_flag, q2_flag, q3_flag, q4_flag
        ],
        check=True
    )

    from openpyxl import load_workbook
    from openpyxl.drawing.image import Image as XLImage
    from PIL import Image as PILImage

    wb_temp = load_workbook(excel_path)
    wb_temp.save(excel_path)

    wb = load_workbook(excel_path)

    ws = wb.create_sheet("Summary Plots")
    img = XLImage(output_plot_path)
    ws.add_image(img, "A1")

    plate_dir = os.path.dirname(output_plot_path)
    for sheet_name in wb.sheetnames:
        if not sheet_name.startswith("Plate"):
            continue
        plate_png = os.path.join(plate_dir, f"{sheet_name}.png")
        if os.path.exists(plate_png):
            ws_plate = wb[sheet_name]
            pil_img = PILImage.open(plate_png)
            temp_png = os.path.join(plate_dir, f"temp_{sheet_name}.png")
            pil_img.save(temp_png, "PNG")
            img_plate = XLImage(temp_png)
            img_plate.anchor = "B33"
            ws_plate.add_image(img_plate)

    wb.save(excel_path)

    with open(excel_path, "rb") as f:
        final_bytes = BytesIO(f.read())

    os.remove(excel_path)
    os.remove(output_plot_path)

    file_id = uuid.uuid4().hex
    in_memory_files[file_id] = {"data": final_bytes.getvalue(), "name": filename}
    session["file_id"] = file_id

    # ── NEW: Redirect to Data Analysis instead of old results page ──
    return render_template(
        "analysis_hub.html",
        excel_file_id=file_id,
        filename=filename,
    )


# ════════════════════════════════════════════════════════════════
# NEW ROUTES: Data Analysis and dedicated analysis pages
# ════════════════════════════════════════════════════════════════

@app.route("/hub/<file_id>")
def analysis_hub(file_id):
    """Data Analysis — the central page with three analysis options."""
    if file_id not in in_memory_files:
        flash("Results not found. Please process your data again.", "danger")
        return redirect(url_for("index"))

    file_info = in_memory_files[file_id]
    return render_template(
        "analysis_hub.html",
        excel_file_id=file_id,
        filename=file_info.get("name", "results.xlsx"),
    )


@app.route("/linear/<file_id>")
def linear_results(file_id):
    """Dedicated Linear Interpolation results page."""
    if file_id not in in_memory_files:
        flash("Results not found. Please process your data again.", "danger")
        return redirect(url_for("index"))

    file_info = in_memory_files[file_id]
    return render_template(
        "linear_results.html",
        excel_file_id=file_id,
        filename=file_info.get("name", "results.xlsx"),
    )


@app.route("/linear_summary/<file_id>")
def linear_summary_data(file_id):
    """JSON API returning plate/pseudotype/sample/titre counts for Data Summary card."""
    if file_id not in in_memory_files:
        return jsonify({"status": "error", "message": "File not found"})

    try:
        file_info = in_memory_files[file_id]
        file_bytes = file_info["data"]
        wb = load_workbook(BytesIO(file_bytes), data_only=True)

        plate_sheets = [s for s in wb.sheetnames if s.startswith("Plate")]

        # Only count plates that contain actual numeric well data (B5:M12).
        # This excludes any extra/blank plates the plate reader appended.
        def _plate_has_data(ws):
            for row in range(5, 13):
                for col in ['B','C','D','E','F','G','H','I','J','K','L','M']:
                    try:
                        float(ws[f'{col}{row}'].value)
                        return True
                    except (ValueError, TypeError):
                        pass
            return False

        num_plates = sum(1 for s in plate_sheets if _plate_has_data(wb[s]))

        pseudotypes = set()
        num_quadrants = 0
        num_labelled = 0
        has_any_label = False
        all_labelled = True

        for sheet_name in plate_sheets:
            ws = wb[sheet_name]
            for pt_cell, sid_cell in [('B3','B4'), ('E3','E4'), ('H3','H4'), ('K3','K4')]:
                pt_val = ws[pt_cell].value
                sid_val = ws[sid_cell].value
                if pt_val and str(pt_val).strip():
                    pseudotypes.add(str(pt_val).strip())
                    num_quadrants += 1
                    if sid_val and str(sid_val).strip():
                        has_any_label = True
                        num_labelled += 1
                    else:
                        all_labelled = False

        # Determine labelling status
        if num_quadrants == 0:
            label_status = "none"
        elif all_labelled:
            label_status = "labelled"
        elif has_any_label:
            label_status = "partial"
        else:
            label_status = "unlabelled"

        # Count flagged errors (if Errors sheet exists)
        error_count, has_errors_sheet = count_errors_from_workbook(file_bytes)

        return jsonify({
            "status": "success",
            "num_plates": num_plates,
            "num_pseudotypes": len(pseudotypes),
            "num_samples": num_quadrants,
            "num_labelled": num_labelled,
            "label_status": label_status,
            "error_count": error_count,
            "error_flagging_enabled": has_errors_sheet,
        })
    except Exception as e:
        return jsonify({"status": "error", "message": str(e)})


@app.route("/boxplot_data/<file_id>")
def boxplot_data(file_id):
    """JSON API: delegates NT linear interpolation to boxplot_nt50.R,
    which averages 3 replicates per quadrant and returns 1 value per
    sample×pseudotype, grouped by pseudotype for the boxplot.

    Query params:
        threshold: 50 (default) or 90
    """
    if file_id not in in_memory_files:
        return jsonify({"status": "error", "message": "File not found"})

    threshold = request.args.get("threshold", "50")
    if threshold not in ("50", "90"):
        threshold = "50"

    q_active = {
        'Q1': request.args.get('q1', 'true').lower() != 'false',
        'Q2': request.args.get('q2', 'true').lower() != 'false',
        'Q3': request.args.get('q3', 'true').lower() != 'false',
        'Q4': request.args.get('q4', 'true').lower() != 'false',
    }

    try:
        file_info = in_memory_files[file_id]
        file_bytes = file_info["data"]

        # Write Excel to a temp file so R can read it
        with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp_xl:
            tmp_xl.write(file_bytes)
            excel_path = tmp_xl.name

        # Temp file for the JSON output from R
        with tempfile.NamedTemporaryFile(delete=False, suffix=".json") as tmp_json:
            json_path = tmp_json.name

        r_script = os.path.join(os.getcwd(), "boxplot_nt50.R")

        result = subprocess.run(
            ["Rscript", r_script, excel_path, json_path, threshold],
            check=True,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
        )

        import json
        with open(json_path, "r") as f:
            r_result = json.load(f)

        # Clean up temp files
        os.remove(excel_path)
        os.remove(json_path)

        # Filter pseudotypes to only those belonging to active quadrants
        if r_result.get('status') == 'success' and 'data' in r_result:
            q_pt_cells = {'Q1': 'B3', 'Q2': 'E3', 'Q3': 'H3', 'Q4': 'K3'}
            wb_f = load_workbook(BytesIO(file_bytes), data_only=True)
            allowed = set()
            for sheet in [s for s in wb_f.sheetnames if s.startswith("Plate")]:
                ws = wb_f[sheet]
                for q, cell in q_pt_cells.items():
                    if q_active[q]:
                        val = ws[cell].value
                        if val and str(val).strip():
                            allowed.add(str(val).strip())
            if allowed:
                r_result['data'] = {k: v for k, v in r_result['data'].items() if k in allowed}

        return jsonify(r_result)

    except subprocess.CalledProcessError as e:
        # Clean up on error
        for p in [excel_path, json_path]:
            if os.path.exists(p):
                os.remove(p)
        return jsonify({
            "status": "error",
            "message": f"R script failed: {e.stderr or e.stdout or str(e)}"
        })
    except Exception as e:
        return jsonify({"status": "error", "message": str(e)})


@app.route("/compare_titres_page/<file_id>")
def compare_titres_page(file_id):
    """
    Titre comparison — triggered from Data Analysis.
    Expects ?fitting_id=<id> in query string.
    Immediately runs the comparison and shows results.
    """
    fitting_id = request.args.get("fitting_id")
    if not file_id or file_id not in in_memory_files:
        flash("Excel results not found.", "danger")
        return redirect(url_for("index"))
    if not fitting_id or fitting_id not in in_memory_files:
        flash("Curve fitting results not found. Please perform curve fitting first.", "danger")
        return redirect(url_for("analysis_hub", file_id=file_id))

    # Delegate to the existing compare logic
    return _run_comparison(file_id, fitting_id)


# ════════════════════════════════════════════════════════════════
# KEEP: Existing download / utility routes (unchanged)
# ════════════════════════════════════════════════════════════════

@app.route("/download_memory/<file_id>")
def download_memory(file_id):
    file_info = in_memory_files.get(file_id)
    if not file_info:
        flash("File not found in memory.", "danger")
        return redirect(url_for("index"))

    file_stream = BytesIO(file_info["data"])
    file_stream.seek(0)

    return send_file(
        file_stream,
        as_attachment=True,
        download_name=file_info["name"],
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )


@app.route("/generate_graphs", methods=["POST"])
def generate_graphs():
    file_id = request.form.get("file_id")
    if not file_id or file_id not in in_memory_files:
        flash("No Excel file found for graph generation.", "danger")
        return redirect(url_for("index"))

    file_info = in_memory_files[file_id]
    file_bytes = file_info["data"]
    filename = file_info["name"]

    r_script = os.path.join(os.getcwd(), "process_data.R")

    settings = load_settings()
    include_timestamp = settings.get("timestamp_in_filename", True)
    default_colours = {"Q1": "#ff7e79", "Q2": "#ffd479", "Q3": "#009193", "Q4": "#d783ff"}

    graph_preset    = request.form.get("graph_preset", "").strip()
    graph_quadrants = request.form.get("graph_quadrants", "").strip()

    if graph_preset and graph_preset in settings.get("presets", {}):
        colours = settings["presets"][graph_preset]
    else:
        active_preset_name = settings.get("selected_preset", None)
        colours = settings.get("presets", {}).get(active_preset_name, default_colours)

    q1_colour = colours.get("Q1", default_colours["Q1"])
    q2_colour = colours.get("Q2", default_colours["Q2"])
    q3_colour = colours.get("Q3", default_colours["Q3"])
    q4_colour = colours.get("Q4", default_colours["Q4"])

    if graph_quadrants:
        try:
            qdata  = json.loads(graph_quadrants)
            q1_flag = str(qdata.get("Q1", True)).lower()
            q2_flag = str(qdata.get("Q2", True)).lower()
            q3_flag = str(qdata.get("Q3", True)).lower()
            q4_flag = str(qdata.get("Q4", True)).lower()
        except (json.JSONDecodeError, ValueError):
            quadrants = settings.get("quadrants", {"Q1": True, "Q2": True, "Q3": True, "Q4": True})
            q1_flag = str(quadrants.get("Q1", True)).lower()
            q2_flag = str(quadrants.get("Q2", True)).lower()
            q3_flag = str(quadrants.get("Q3", True)).lower()
            q4_flag = str(quadrants.get("Q4", True)).lower()
    else:
        quadrants = settings.get("quadrants", {"Q1": True, "Q2": True, "Q3": True, "Q4": True})
        q1_flag = str(quadrants.get("Q1", True)).lower()
        q2_flag = str(quadrants.get("Q2", True)).lower()
        q3_flag = str(quadrants.get("Q3", True)).lower()
        q4_flag = str(quadrants.get("Q4", True)).lower()

    try:
        file_stream = BytesIO(file_bytes)
        file_stream.seek(0)

        with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp_input:
            tmp_input.write(file_stream.read())
            input_path = tmp_input.name

        with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as tmp_output:
            output_plot_path = tmp_output.name

        plot_title = os.path.splitext(filename)[0]
        subprocess.run(
            [
                "Rscript", r_script,
                input_path, output_plot_path,
                str(include_timestamp).lower(),
                q1_colour, q2_colour, q3_colour, q4_colour,
                plot_title,
                q1_flag, q2_flag, q3_flag, q4_flag
            ],
            check=True,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
        )

        with open(output_plot_path, "rb") as f:
            image_bytes = BytesIO(f.read())

        image_bytes.seek(0)
        os.remove(input_path)
        os.remove(output_plot_path)

        png_filename = os.path.splitext(filename)[0] + ".png"
        return send_file(
            image_bytes,
            mimetype="image/png",
            as_attachment=True,
            download_name=png_filename
        )

    except subprocess.CalledProcessError as e:
        flash(f"R script failed: {e.stderr}", "danger")
        return redirect(url_for("index"))

    except Exception as e:
        flash(f"Unexpected error: {str(e)}", "danger")
        return redirect(url_for("index"))


@app.route("/save_quadrants", methods=["POST"])
def save_quadrants():
    quadrants = request.get_json()
    settings = load_settings()
    settings["quadrants"] = quadrants
    save_settings(settings)
    return "Quadrant settings saved", 200


@app.route("/get_settings")
def get_settings():
    settings = load_settings()
    return jsonify(settings)

@app.route("/get_template_dilutions")
def get_template_dilutions():
    try:
        template_path = load_template_path()
        wb = load_workbook(template_path, data_only=True)
        ws = wb.active
        
        dilutions = []
        for row in range(5, 13):
            cell_value = ws[f'A{row}'].value
            try:
                num_val = float(cell_value)
                if num_val == 0:
                    dilutions.append("0")
                elif num_val >= 1000:
                    dilutions.append(f"{int(num_val):,}")
                elif num_val == int(num_val):
                    dilutions.append(str(int(num_val)))
                else:
                    dilutions.append(str(num_val))
            except (ValueError, TypeError):
                dilutions.append(str(cell_value) if cell_value else "—")
        
        return jsonify({"dilutions": dilutions, "status": "success"})
    
    except FileNotFoundError:
        return jsonify({"dilutions": [], "status": "error", "message": "No template selected"})
    except Exception as e:
        return jsonify({"dilutions": [], "status": "error", "message": str(e)})


@app.route("/validate_csv_mode", methods=["POST"])
def validate_csv_mode_route():
    """JSON API: check whether the uploaded CSV matches the selected data mode."""
    file = request.files.get("csv_file")
    selected_mode = request.form.get("data_mode", "standard")

    if not file:
        return jsonify({"ok": True})

    csv_bytes = BytesIO(file.read())
    csv_bytes.seek(0)

    ok, detected, message = validate_csv_mode(csv_bytes, selected_mode)
    return jsonify({
        "ok": ok,
        "detected_mode": detected,
        "message": message,
    })


@app.route("/save_preset", methods=["POST"])
def save_preset():
    data = request.get_json()
    settings = load_settings()
    settings["presets"][data["name"]] = data["colours"]
    save_settings(settings)
    return jsonify({"status": "success", "message": "Preset saved."}), 200


@app.route("/delete_preset", methods=["POST"])
def delete_preset():
    data = request.get_json()
    settings = load_settings()
    settings["presets"].pop(data["name"], None)
    save_settings(settings)
    return jsonify({"status": "success", "message": "Preset deleted."}), 200


@app.route("/set_active_preset", methods=["POST"])
def set_active_preset():
    data = request.get_json()
    name = data.get("name")
    if not name:
        return "No preset name provided", 400

    settings = load_settings()
    if "presets" not in settings or name not in settings["presets"]:
        return "Preset not found", 404

    settings["selected_preset"] = name
    save_settings(settings)

    return "Preset updated", 200


@app.route("/perform_curve_fitting", methods=["POST"])
def perform_curve_fitting():
    file_id = request.form.get("file_id")
    if not file_id or file_id not in in_memory_files:
        flash("No Excel file found for curve fitting.", "danger")
        return redirect(url_for("index"))

    file_info = in_memory_files[file_id]
    file_bytes = file_info["data"]
    filename = file_info["name"]

    try:
        file_stream = BytesIO(file_bytes)
        file_stream.seek(0)

        with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp_excel:
            tmp_excel.write(file_stream.read())
            excel_path = tmp_excel.name

        output_dir = tempfile.mkdtemp(prefix="sigmoid_")
        
        sigmoid_csv_path = os.path.join(output_dir, "sigmoidData.csv")
        generate_sigmoid_csv(excel_path, sigmoid_csv_path)

        settings = load_settings()
        include_timestamp = settings.get("timestamp_in_filename", True)
        
        assay_title = os.path.splitext(filename)[0]
        import re
        assay_title = re.sub(r'_\d{4}-\d{2}-\d{2}$', '', assay_title)
        
        if include_timestamp:
            timestamp = datetime.now().strftime("%Y-%m-%d")
        else:
            timestamp = ""

        r_script = os.path.join(os.getcwd(), "fit_sigmoids.R")
        
        result = subprocess.run(
            ["Rscript", r_script, sigmoid_csv_path, output_dir, assay_title, timestamp],
            check=True,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
        )

        if assay_title and timestamp:
            ic50_filename = f"IC50s_{assay_title}_{timestamp}.csv"
        elif assay_title:
            ic50_filename = f"IC50s_{assay_title}.csv"
        elif timestamp:
            ic50_filename = f"IC50s_{timestamp}.csv"
        else:
            ic50_filename = "IC50s.csv"
        
        ic50_path = os.path.join(output_dir, ic50_filename)
        
        plot_files = [f for f in os.listdir(output_dir) if f.endswith('.png')]
        
        output_files = {}
        
        if os.path.exists(ic50_path):
            with open(ic50_path, 'rb') as f:
                output_files[ic50_filename] = f.read()
        
        if os.path.exists(sigmoid_csv_path):
            with open(sigmoid_csv_path, 'rb') as f:
                output_files['sigmoidData.csv'] = f.read()
        
        for plot_file in plot_files:
            plot_path = os.path.join(output_dir, plot_file)
            with open(plot_path, 'rb') as f:
                output_files[plot_file] = f.read()
        
        os.remove(excel_path)
        for f in os.listdir(output_dir):
            os.remove(os.path.join(output_dir, f))
        os.rmdir(output_dir)
        
        fitting_id = uuid.uuid4().hex
        in_memory_files[fitting_id] = {
            "data": output_files,
            "name": "sigmoid_fitting_results",
            "type": "sigmoid_results",
            "ic50_filename": ic50_filename
        }
        
        return render_template(
            "curve_fitting_results.html",
            fitting_id=fitting_id,
            excel_file_id=file_id,
            output_files=list(output_files.keys()),
            ic50_filename=ic50_filename,
            settings=load_settings()
        )

    except subprocess.CalledProcessError as e:
        flash(f"R script failed: {e.stderr}", "danger")
        return redirect(url_for("analysis_hub", file_id=file_id))
    except Exception as e:
        flash(f"Curve fitting error: {str(e)}", "danger")
        return redirect(url_for("analysis_hub", file_id=file_id))

@app.route("/download_sigmoid/<fitting_id>/<filename>")
def download_sigmoid(fitting_id, filename):
    if fitting_id not in in_memory_files:
        flash("Fitting results not found.", "danger")
        return redirect(url_for("index"))

    file_info = in_memory_files[fitting_id]
    if file_info.get("type") != "sigmoid_results":
        flash("Invalid file type.", "danger")
        return redirect(url_for("index"))

    if filename not in file_info["data"]:
        flash("File not found.", "danger")
        return redirect(url_for("index"))

    file_bytes = BytesIO(file_info["data"][filename])
    file_bytes.seek(0)

    if filename.endswith('.csv'):
        mimetype = 'text/csv'
    elif filename.endswith('.png'):
        mimetype = 'image/png'
    else:
        mimetype = 'application/octet-stream'

    download_name = filename
    settings = load_settings()
    if settings.get("timestamp_in_filename", False):
        import re as _re
        ts = datetime.now().strftime("%Y-%m-%d")
        base, ext = os.path.splitext(filename)
        # Skip if filename already contains a date stamp (e.g. IC50 files)
        if not _re.search(r'\d{4}-\d{2}-\d{2}', base):
            download_name = f"{base}_{ts}{ext}"

    return send_file(
        file_bytes,
        as_attachment=True,
        download_name=download_name,
        mimetype=mimetype
    )


@app.route("/generate_sigmoid_graph/<fitting_id>")
def generate_sigmoid_graph(fitting_id):
    if fitting_id not in in_memory_files:
        return jsonify({"error": "Fitting results not found"}), 404

    file_info = in_memory_files[fitting_id]
    if file_info.get("type") != "sigmoid_results":
        return jsonify({"error": "Invalid type"}), 400

    show_good     = request.args.get("good",     "true").lower() == "true"
    show_unstable = request.args.get("unstable", "true").lower() == "true"

    tmp_dir = None
    try:
        tmp_dir = tempfile.mkdtemp(prefix="sigplot_")

        raw_csv = os.path.join(tmp_dir, "sigmoidData.csv")
        with open(raw_csv, "wb") as f:
            f.write(file_info["data"]["sigmoidData.csv"])

        ic50_filename = file_info["ic50_filename"]
        ic50_csv = os.path.join(tmp_dir, ic50_filename)
        with open(ic50_csv, "wb") as f:
            f.write(file_info["data"][ic50_filename])

        output_png = os.path.join(tmp_dir, "sigmoid_combined.png")

        r_script = os.path.join(os.getcwd(), "plot_sigmoids.R")
        subprocess.run(
            ["Rscript", r_script, raw_csv, ic50_csv, output_png,
             str(show_good).lower(), str(show_unstable).lower()],
            check=True,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
        )

        with open(output_png, "rb") as f:
            png_bytes = f.read()

        # Store latest render for download
        in_memory_files[fitting_id]["data"]["sigmoid_combined.png"] = png_bytes

        return send_file(BytesIO(png_bytes), mimetype="image/png")

    except subprocess.CalledProcessError as e:
        return jsonify({"error": f"R script failed: {e.stderr}"}), 500
    except Exception as e:
        return jsonify({"error": str(e)}), 500
    finally:
        if tmp_dir and os.path.exists(tmp_dir):
            for fn in os.listdir(tmp_dir):
                try:
                    os.remove(os.path.join(tmp_dir, fn))
                except Exception:
                    pass
            try:
                os.rmdir(tmp_dir)
            except Exception:
                pass


# ════════════════════════════════════════════════════════════════
# Titre Comparison (refactored into reusable helper)
# ════════════════════════════════════════════════════════════════

@app.route("/compare_titres", methods=["POST"])
def compare_titres():
    """Legacy POST route — still works for form submissions."""
    excel_file_id = request.form.get("excel_file_id")
    fitting_id = request.form.get("fitting_id")
    
    if not excel_file_id or excel_file_id not in in_memory_files:
        flash("Excel results not found.", "danger")
        return redirect(url_for("index"))
    
    if not fitting_id or fitting_id not in in_memory_files:
        flash("Curve fitting results not found. Please perform curve fitting first.", "danger")
        return redirect(url_for("index"))

    return _run_comparison(excel_file_id, fitting_id)


def _run_comparison(excel_file_id, fitting_id):
    """
    Shared comparison logic used by both POST and GET routes.
    
    R now handles both the NT50 linear interpolation (reading Plate sheets
    directly from the Excel file) and the IC50 comparison, so Python just
    orchestrates temp files and reads back the CSV results.
    """
    try:
        excel_info = in_memory_files[excel_file_id]
        excel_bytes = BytesIO(excel_info["data"])

        fitting_info = in_memory_files[fitting_id]
        if fitting_info.get("type") != "sigmoid_results":
            flash("Invalid fitting results.", "danger")
            return redirect(url_for("index"))

        ic50_filename = fitting_info.get("ic50_filename", "IC50s.csv")
        if ic50_filename not in fitting_info["data"]:
            flash("IC50 file not found in fitting results.", "danger")
            return redirect(url_for("index"))

        ic50_bytes = fitting_info["data"][ic50_filename]

        # Write Excel and IC50 CSV to temp files for R
        with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp_excel:
            tmp_excel.write(excel_bytes.getvalue())
            excel_path = tmp_excel.name

        with tempfile.NamedTemporaryFile(delete=False, suffix=".csv") as tmp_ic50:
            tmp_ic50.write(ic50_bytes)
            ic50_path = tmp_ic50.name

        output_dir = tempfile.mkdtemp(prefix="comparison_")

        # R now receives the Excel file directly (not a pre-computed NT50 CSV)
        r_script = os.path.join(os.getcwd(), "compare_titres.R")

        result = subprocess.run(
            ["Rscript", r_script, excel_path, ic50_path, output_dir],
            check=True,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
        )

        # Collect output files
        output_files = {}
        for fname in ['comparison_stats.csv', 'merged_titres.csv',
                       'top_mismatches.csv', 'titre_comparison.png',
                       'titre_comparison_interactive.html']:
            filepath = os.path.join(output_dir, fname)
            if os.path.exists(filepath):
                with open(filepath, 'rb') as f:
                    output_files[fname] = f.read()

        # Clean up temp files
        os.remove(excel_path)
        os.remove(ic50_path)
        for f in os.listdir(output_dir):
            os.remove(os.path.join(output_dir, f))
        os.rmdir(output_dir)

        # Store results in memory
        comparison_id = uuid.uuid4().hex
        in_memory_files[comparison_id] = {
            "data": output_files,
            "name": "titre_comparison",
            "type": "comparison_results"
        }

        # Parse stats and mismatches for template rendering
        import csv as py_csv

        stats = {}
        if 'comparison_stats.csv' in output_files:
            csv_data = output_files['comparison_stats.csv'].decode('utf-8')
            reader = py_csv.DictReader(csv_data.splitlines())
            stats = next(reader)

        mismatches = []
        if 'top_mismatches.csv' in output_files:
            csv_data = output_files['top_mismatches.csv'].decode('utf-8')
            reader = py_csv.DictReader(csv_data.splitlines())
            mismatches = list(reader)

        return render_template(
            "titre_comparison_results.html",
            comparison_id=comparison_id,
            excel_file_id=excel_file_id,
            stats=stats,
            mismatches=mismatches,
            has_plot='titre_comparison.png' in output_files,
            settings=load_settings()
        )

    except subprocess.CalledProcessError as e:
        flash(f"R script failed: {e.stderr}", "danger")
        return redirect(url_for("analysis_hub", file_id=excel_file_id))
    except Exception as e:
        flash(f"Comparison error: {str(e)}", "danger")
        return redirect(url_for("analysis_hub", file_id=excel_file_id))


@app.route("/download_comparison/<comparison_id>/<filename>")
def download_comparison(comparison_id, filename):
    if comparison_id not in in_memory_files:
        flash("Comparison results not found.", "danger")
        return redirect(url_for("index"))
    
    file_info = in_memory_files[comparison_id]
    if file_info.get("type") != "comparison_results":
        flash("Invalid file type.", "danger")
        return redirect(url_for("index"))
    
    if filename not in file_info["data"]:
        flash("File not found.", "danger")
        return redirect(url_for("index"))
    
    file_bytes = BytesIO(file_info["data"][filename])
    file_bytes.seek(0)
    
    if filename.endswith('.csv'):
        mimetype = 'text/csv'
        as_attachment = True
    elif filename.endswith('.png'):
        mimetype = 'image/png'
        as_attachment = True
    elif filename.endswith('.html'):
        mimetype = 'text/html'
        as_attachment = request.args.get('download') == '1'
    else:
        mimetype = 'application/octet-stream'
        as_attachment = True
    
    download_name = filename
    if as_attachment:
        settings = load_settings()
        if settings.get("timestamp_in_filename", False):
            ts = datetime.now().strftime("%Y-%m-%d")
            base, ext = os.path.splitext(filename)
            download_name = f"{base}_{ts}{ext}"

    return send_file(
        file_bytes,
        as_attachment=as_attachment,
        download_name=download_name if as_attachment else None,
        mimetype=mimetype
    )


# ════════════════════════════════════════════════════════════════
# LEGACY: Keep /results/<file_id> route as redirect to hub
# ════════════════════════════════════════════════════════════════

@app.route("/results/<file_id>")
def view_results(file_id):
    """Legacy route — redirects to Data Analysis."""
    return redirect(url_for("analysis_hub", file_id=file_id))


if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port)