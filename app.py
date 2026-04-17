from flask import Flask, render_template, request, redirect, url_for, send_file, flash, jsonify, send_from_directory, session
import os
import uuid
import json
import subprocess
import time
from datetime import datetime
from io import BytesIO
import re
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
    DEFAULT_SETTINGS,
)

in_memory_files = {}  # Key: UUID, Value: BytesIO

app = Flask(__name__)
app.secret_key = "your-secret-key"
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024  # 50MB upload limit



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
    # Append user-created templates (only if the file still exists)
    for name, path in current_settings.get("custom_templates", {}).items():
        if os.path.exists(path):
            default_templates[name] = path

    try:
        current_template_path = load_template_path()
    except FileNotFoundError:
        current_template_path = None

    if request.method == "POST":
        selected_template_key = request.form.get("default_template_select")
        if selected_template_key in default_templates:
            selected_template_path = default_templates[selected_template_key]
            save_template_path(selected_template_path)
        else:
            file = request.files.get("template_file")
            if file and file.filename.endswith(".xlsx"):
                filename = f"template_{uuid.uuid4().hex}.xlsx"
                filepath = os.path.join("excel_templates", filename)
                os.makedirs("excel_templates", exist_ok=True)
                file.save(filepath)
                save_template_path(filepath)

        timestamp_flag = request.form.get("timestamp_in_filename") == "on"
        error_flagging_flag = request.form.get("error_flagging") == "on"
        new_settings = current_settings.copy()
        new_settings["timestamp_in_filename"] = timestamp_flag
        new_settings["error_flagging"] = error_flagging_flag
        new_settings["default_data_mode"] = request.form.get("default_data_mode", "standard")
        try:
            new_settings["default_num_pseudotypes"] = int(request.form.get("default_num_pseudotypes", 1))
        except ValueError:
            new_settings["default_num_pseudotypes"] = 1
        try:
            new_settings["outlier_threshold_log2"] = float(request.form.get("outlier_threshold_log2", 1.0))
        except ValueError:
            new_settings["outlier_threshold_log2"] = 1.0
        try:
            new_settings["sigmoid_r2_threshold"] = float(request.form.get("sigmoid_r2_threshold", 0.5))
        except ValueError:
            new_settings["sigmoid_r2_threshold"] = 0.5
        new_settings["lod_censor_include"] = request.form.get("lod_censor_include") == "on"
        try:
            new_settings["comparison_disagreement_threshold"] = float(request.form.get("comparison_disagreement_threshold", 1.0))
        except ValueError:
            new_settings["comparison_disagreement_threshold"] = 1.0
        save_settings(new_settings)
        flash("Settings saved.", "success")
        return redirect(url_for("settings"))

    current_settings["template_path"] = current_template_path

    return render_template("settings.html", settings=current_settings, default_templates=default_templates)


@app.route("/reset_settings", methods=["POST"])
def reset_settings():
    """Reset threshold/toggle settings to defaults, preserving presets and custom templates."""
    current = load_settings()
    reset_keys = [
        "timestamp_in_filename", "error_flagging", "default_data_mode",
        "default_num_pseudotypes", "outlier_threshold_log2",
        "sigmoid_r2_threshold", "lod_censor_include", "comparison_disagreement_threshold",
    ]
    for key in reset_keys:
        current[key] = DEFAULT_SETTINGS[key]
    save_settings(current)
    flash("Settings reset to defaults.", "success")
    return redirect(url_for("settings"))


@app.route("/save_template_selection", methods=["POST"])
def save_template_selection():
    """Lightweight endpoint for the template selector + upload only."""
    current_settings = load_settings()
    default_templates = {
        "NTA Template (dil 50-36450)": "excel_templates/NTA_Template.xlsx",
        "Measles Template (dil 32-65536)": "excel_templates/Measles_NTA_Template.xlsx",
        "Backup NTA Template": "excel_templates/Backup_NTA_Template.xlsx",
    }
    for name, path in current_settings.get("custom_templates", {}).items():
        if os.path.exists(path):
            default_templates[name] = path

    selected_key = request.form.get("default_template_select")
    if selected_key in default_templates:
        save_template_path(default_templates[selected_key])
    else:
        file = request.files.get("template_file")
        if file and file.filename.endswith(".xlsx"):
            filename = f"template_{uuid.uuid4().hex}.xlsx"
            filepath = os.path.join("excel_templates", filename)
            os.makedirs("excel_templates", exist_ok=True)
            file.save(filepath)
            save_template_path(filepath)

    flash("Template saved.", "success")
    return redirect(url_for("settings"))


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

    raw_np = request.form.get("num_pseudotypes", "1")
    if raw_np == "2alt":
        num_pseudotypes = "2alt"
    else:
        try:
            num_pseudotypes = int(raw_np)
            if num_pseudotypes not in [1, 2, 3, 4]:
                raise ValueError()
        except ValueError:
            flash("Invalid pseudotype count. Must be 1–4.", "danger")
            return redirect(url_for("index"))

    # Per-plate config (optional — sent as JSON when custom per-plate mode is active)
    plate_configs = None
    raw_pc = request.form.get("plate_configs", "").strip()
    if raw_pc:
        try:
            plate_configs = json.loads(raw_pc)
            if not isinstance(plate_configs, list):
                plate_configs = None
        except (ValueError, TypeError):
            plate_configs = None

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

    _proc_start = time.time()
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
        plate_configs=plate_configs,
    )

    extract_final_titres_xlwings(output_bytes)

    from nta_utils import add_default_to_final_titres
    add_default_to_final_titres(output_bytes)

    # ── Error flagging (if enabled in settings) ──
    if settings.get("error_flagging", False):
        flag_triplicate_errors(output_bytes, threshold_log2=settings.get("outlier_threshold_log2", 1.0))

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

    with open(output_plot_path, "rb") as f:
        summary_plot_bytes = f.read()

    os.remove(excel_path)
    os.remove(output_plot_path)

    file_id = uuid.uuid4().hex
    in_memory_files[file_id] = {"data": final_bytes.getvalue(), "name": filename, "summary_plot": summary_plot_bytes}
    session["file_id"] = file_id

    # ── NEW: Redirect to Data Analysis instead of old results page ──
    _proc_elapsed = round(time.time() - _proc_start, 1)
    return render_template(
        "analysis_hub.html",
        excel_file_id=file_id,
        filename=filename,
        processing_time=_proc_elapsed,
        fitting_id=None,
        settings=load_settings(),
    )


# ════════════════════════════════════════════════════════════════
# Data Analysis
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
        fitting_id=file_info.get("fitting_id"),
        comparison_id=file_info.get("comparison_id"),
        settings=load_settings(),
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
        settings=load_settings(),
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


@app.route("/curve_fitting_results/<fitting_id>")
def curve_fitting_results(fitting_id):
    """Serve cached sigmoid curve fitting results without reprocessing."""
    if fitting_id not in in_memory_files:
        flash("Curve fitting results not found. Please run curve fitting again.", "danger")
        return redirect(url_for("index"))
    info = in_memory_files[fitting_id]
    excel_file_id = info.get("excel_file_id")
    if not excel_file_id or excel_file_id not in in_memory_files:
        flash("Original Excel results not found.", "danger")
        return redirect(url_for("index"))
    return render_template(
        "curve_fitting_results.html",
        fitting_id=fitting_id,
        excel_file_id=excel_file_id,
        output_files=list(info["data"].keys()),
        ic50_filename=info["ic50_filename"],
        settings=load_settings(),
        processing_time=None,
    )


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

    # Return cached comparison results if already computed for this file
    existing_cmp_id = in_memory_files[file_id].get("comparison_id")
    if existing_cmp_id and existing_cmp_id in in_memory_files:
        cached = in_memory_files[existing_cmp_id]
        return render_template(
            "titre_comparison_results.html",
            comparison_id=existing_cmp_id,
            excel_file_id=file_id,
            stats=cached["stats"],
            mismatches=cached["mismatches"],
            has_plot=cached["has_plot"],
            settings=load_settings(),
            processing_time=None,
        )

    # Delegate to the existing compare logic
    return _run_comparison(file_id, fitting_id)


# ════════════════════════════════════════════════════════════════
# download / utility routes
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


@app.route("/summary_plot/<file_id>")
def summary_plot(file_id):
    if file_id not in in_memory_files:
        return "", 404
    plot_bytes = in_memory_files[file_id].get("summary_plot")
    if not plot_bytes:
        return "", 404
    return send_file(BytesIO(plot_bytes), mimetype="image/png")


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
                    dilutions.append(str(int(num_val)))
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


@app.route("/create_template_variant", methods=["POST"])
def create_template_variant():
    """
    Duplicate the current template, replace A5:A12 with a new dilution series,
    and save the result as a new named template. The source is never modified.
    """
    data = request.get_json()
    template_name = (data.get("template_name") or "").strip()
    dilutions_raw = data.get("dilutions", [])

    if not template_name:
        return jsonify({"status": "error", "message": "Template name is required."}), 400
    if len(dilutions_raw) != 8:
        return jsonify({"status": "error", "message": "Exactly 8 dilution values are required (one per row A5:A12)."}), 400

    try:
        dilutions = [float(str(d).replace(",", "")) for d in dilutions_raw]
    except (ValueError, TypeError):
        return jsonify({"status": "error", "message": "All dilution values must be numbers."}), 400

    try:
        source_path = load_template_path()
    except FileNotFoundError:
        return jsonify({"status": "error", "message": "No source template is currently selected."}), 400

    safe_name = re.sub(r"[^\w\s-]", "", template_name).strip().replace(" ", "_")
    if not safe_name:
        return jsonify({"status": "error", "message": "Template name contains no valid characters."}), 400

    os.makedirs("excel_templates", exist_ok=True)
    output_path = os.path.join("excel_templates", f"{safe_name}.xlsx")

    if os.path.exists(output_path):
        return jsonify({"status": "error", "message": f"A file named '{safe_name}.xlsx' already exists. Choose a different name."}), 400

    try:
        import openpyxl as _xl
        wb = _xl.load_workbook(source_path)
        ws = wb.active
        for i, val in enumerate(dilutions):
            ws[f"A{5 + i}"] = val
        wb.save(output_path)
    except Exception as e:
        return jsonify({"status": "error", "message": f"Failed to write template: {e}"}), 500

    # Register in settings so it appears in the dropdown
    settings = load_settings()
    if "custom_templates" not in settings:
        settings["custom_templates"] = {}
    settings["custom_templates"][template_name] = output_path
    save_settings(settings)

    return jsonify({"status": "success", "name": template_name, "path": output_path})


@app.route("/delete_custom_template", methods=["POST"])
def delete_custom_template():
    """Remove a user-created template from settings (optionally deletes the file too)."""
    data = request.get_json()
    name = (data.get("name") or "").strip()
    if not name:
        return jsonify({"status": "error", "message": "No name provided."}), 400

    settings = load_settings()
    custom = settings.get("custom_templates", {})
    if name not in custom:
        return jsonify({"status": "error", "message": "Template not found."}), 404

    path = custom.pop(name)
    save_settings(settings)

    # Delete the file if it's in excel_templates/ and still exists
    if path.startswith("excel_templates/") and os.path.exists(path):
        try:
            os.remove(path)
        except OSError:
            pass

    return jsonify({"status": "success"})


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

    # Return cached results immediately if already computed for this file
    existing_fitting_id = file_info.get("fitting_id")
    if existing_fitting_id and existing_fitting_id in in_memory_files:
        return redirect(url_for("curve_fitting_results", fitting_id=existing_fitting_id))

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
        assay_title = re.sub(r'_\d{4}-\d{2}-\d{2}$', '', assay_title)
        
        if include_timestamp:
            timestamp = datetime.now().strftime("%Y-%m-%d")
        else:
            timestamp = ""

        r_script = os.path.join(os.getcwd(), "fit_sigmoids.R")
        r2_threshold = str(settings.get("sigmoid_r2_threshold", 0.5))
        include_lod = "TRUE" if settings.get("lod_censor_include", False) else "FALSE"

        _proc_start = time.time()
        result = subprocess.run(
            ["Rscript", r_script, sigmoid_csv_path, output_dir, assay_title, timestamp, r2_threshold, include_lod],
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
            "ic50_filename": ic50_filename,
            "excel_file_id": file_id,
        }
        # Store fitting_id server-side so analysis hub can unlock comparison card
        in_memory_files[file_id]["fitting_id"] = fitting_id
        
        _proc_elapsed = round(time.time() - _proc_start, 1)
        return render_template(
            "curve_fitting_results.html",
            fitting_id=fitting_id,
            excel_file_id=file_id,
            output_files=list(output_files.keys()),
            ic50_filename=ic50_filename,
            settings=load_settings(),
            processing_time=_proc_elapsed,
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
        ts = datetime.now().strftime("%Y-%m-%d")
        base, ext = os.path.splitext(filename)
        # Skip if filename already contains a date stamp (e.g. IC50 files)
        if not re.search(r'\d{4}-\d{2}-\d{2}', base):
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
    show_poor_fit = request.args.get("poor_fit", "true").lower() == "true"

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
        _plot_settings = load_settings()
        show_lod = "true" if _plot_settings.get("lod_censor_include", False) else "false"
        subprocess.run(
            ["Rscript", r_script, raw_csv, ic50_csv, output_png,
             str(show_good).lower(), str(show_unstable).lower(), show_lod,
             str(show_poor_fit).lower()],
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
        cmp_settings = load_settings()
        disagreement_threshold = str(cmp_settings.get("comparison_disagreement_threshold", 1.0))

        _proc_start = time.time()
        result = subprocess.run(
            ["Rscript", r_script, excel_path, ic50_path, output_dir, disagreement_threshold],
            check=True,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
            timeout=180,
        )

        if result.stderr:
            print(f"[compare_titres.R stderr]:\n{result.stderr}", flush=True)

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
        if 'merged_titres.csv' in output_files:
            csv_data = output_files['merged_titres.csv'].decode('utf-8')
            reader = py_csv.DictReader(csv_data.splitlines())
            all_rows = list(reader)
            for row in all_rows:
                try:
                    lfd = float(row.get('log2_fold_difference', 0) or 0)
                except ValueError:
                    lfd = 0
                if abs(lfd) >= 2:
                    mismatches.append({
                        'Sample':               row.get('Sample_ID', ''),
                        'Virus':                row.get('Pseudotype', ''),
                        'NT50':                 row.get('NT50 (Linear Interpolation)', ''),
                        'IC50_Titre':           row.get('NT50 / IC50 (Curve Fitting)', ''),
                        'Log2_Fold_Difference': lfd,
                        'Quality':              row.get('Sigmoid Quality', ''),
                    })
            mismatches.sort(key=lambda r: abs(r['Log2_Fold_Difference']), reverse=True)

        # Cache parsed render data so future visits skip reprocessing
        has_plot = 'titre_comparison.png' in output_files
        in_memory_files[comparison_id]["stats"]      = stats
        in_memory_files[comparison_id]["mismatches"] = mismatches
        in_memory_files[comparison_id]["has_plot"]   = has_plot
        # Store back-reference so the hub and compare_titres_page can find the cache
        in_memory_files[excel_file_id]["comparison_id"] = comparison_id

        _proc_elapsed = round(time.time() - _proc_start, 1)
        return render_template(
            "titre_comparison_results.html",
            comparison_id=comparison_id,
            excel_file_id=excel_file_id,
            stats=stats,
            mismatches=mismatches,
            has_plot=has_plot,
            settings=load_settings(),
            processing_time=_proc_elapsed,
        )

    except subprocess.TimeoutExpired:
        flash("Comparison timed out (>3 min). Check your terminal for R output.", "danger")
        return redirect(url_for("analysis_hub", file_id=excel_file_id))
    except subprocess.CalledProcessError as e:
        print(f"[compare_titres.R FAILED]\nSTDOUT: {e.stdout}\nSTDERR: {e.stderr}", flush=True)
        flash(f"R script failed — see terminal for details. Error: {(e.stderr or e.stdout or '').strip()[:300]}", "danger")
        return redirect(url_for("analysis_hub", file_id=excel_file_id))
    except Exception as e:
        print(f"[_run_comparison Exception]: {e}", flush=True)
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


# ════════════════════════════════════════════════════════════════
# PLATE MAPPER
# ════════════════════════════════════════════════════════════════

PM_PRESETS_FILE = os.path.join(os.path.dirname(__file__), 'plate_mapper_presets.json')


def pm_load_presets():
    if os.path.exists(PM_PRESETS_FILE):
        with open(PM_PRESETS_FILE, 'r') as f:
            return json.load(f)
    return {}


def _pm_write_presets(presets):
    with open(PM_PRESETS_FILE, 'w') as f:
        json.dump(presets, f, indent=4)


def pm_save_preset(name, settings):
    presets = pm_load_presets()
    presets[name] = settings
    _pm_write_presets(presets)
    return True


def pm_delete_preset(name):
    presets = pm_load_presets()
    if name in presets:
        del presets[name]
        _pm_write_presets(presets)
        return True
    return False


def pm_format_id(entry, prefix, suffix='', id_length=0, pad_char='0'):
    if not entry or entry.strip() == "":
        return None
    if entry.startswith('!'):
        return entry[1:]
    if id_length > 0:
        entry = entry.zfill(id_length) if pad_char == '0' else entry.rjust(id_length, pad_char)
    return f"{prefix}{entry}{suffix}"


def pm_build_grid(sample_ids, replicate_count, layout_mode,
                  pos_label='PosCtrl', neg_label='NegCtrl', ctrl3_label='Ctrl3',
                  include_pos=True, include_neg=True, include_ctrl3=True):
    grid = [["" for _ in range(12)] for _ in range(8)]
    if layout_mode == 'vertical':
        fill_positions = []
        control_labels = []
        if include_pos:
            control_labels += [pos_label] * replicate_count
        if include_neg:
            control_labels += [neg_label] * replicate_count
        if include_ctrl3:
            control_labels += [ctrl3_label] * replicate_count
        control_col = 0
        control_row = 0
        for label in control_labels:
            if control_row >= 8:
                control_row = 0
                control_col += 1
                if control_col >= 12:
                    break
            grid[control_row][control_col] = label
            control_row += 1
        for col in range(12):
            for row in range(8):
                if grid[row][col] == "":
                    fill_positions.append((row, col))
        idx = 0
        for sid in sample_ids:
            for _ in range(replicate_count):
                if idx >= len(fill_positions):
                    break
                r, c = fill_positions[idx]
                grid[r][c] = sid
                idx += 1
    elif layout_mode == 'horizontal':
        current_row = 0
        col = 0
        if include_pos:
            for rep in range(replicate_count):
                grid[current_row][col + rep] = pos_label
            current_row += 1
        if include_neg:
            for rep in range(replicate_count):
                grid[current_row][col + rep] = neg_label
            current_row += 1
        if include_ctrl3:
            for rep in range(replicate_count):
                grid[current_row][col + rep] = ctrl3_label
            current_row += 1
        sample_idx = 0
        while sample_idx < len(sample_ids):
            if current_row >= 8:
                current_row = 0
                col += replicate_count
                if col + replicate_count > 12:
                    break
            for rep in range(replicate_count):
                if col + rep < 12:
                    grid[current_row][col + rep] = sample_ids[sample_idx]
            current_row += 1
            sample_idx += 1
    return grid


@app.route('/plate_mapper', methods=['GET', 'POST'])
def plate_mapper():
    if request.method == 'POST':
        prefix = request.form.get('prefix', '')
        suffix = request.form.get('suffix', '')
        pos_label = request.form.get('pos_label', '').strip() or "Control 1"
        neg_label = request.form.get('neg_label', '').strip() or "Control 2"
        ctrl3_label = request.form.get('ctrl3_label', '').strip() or "Control 3"
        include_pos = 'include_pos' in request.form
        include_neg = 'include_neg' in request.form
        include_ctrl3 = 'include_ctrl3' in request.form
        replicate_count = int(request.form.get('replicate_count', 2) or 2)
        id_length = int(request.form.get('id_length', 0) or 0)
        pad_char = (request.form.get('pad_char', '0') or '0')[:1] or '0'
        layout_mode = request.form.get('layout_mode', 'vertical')
        plate_count = int(request.form.get('plate_count', 1) or 1)

        plate_labels = []
        for i in range(1, plate_count + 1):
            label = request.form.get(f'plate_label_{i}', '').strip()
            plate_labels.append(label or f"Plate {i}")

        plates = []
        for i in range(1, plate_count + 1):
            raw = request.form.get(f'plate_{i}', '')
            nums = [s.strip() for s in raw.replace(';', ',').split(',') if s.strip()]
            formatted = [
                pm_format_id(n, prefix, suffix, id_length, pad_char)
                for n in nums
                if pm_format_id(n, prefix, suffix, id_length, pad_char)
            ]
            plates.append(pm_build_grid(
                formatted, replicate_count, layout_mode,
                pos_label, neg_label, ctrl3_label,
                include_pos, include_neg, include_ctrl3
            ))

        control_labels = set()
        if include_pos:
            control_labels.add(pos_label)
        if include_neg:
            control_labels.add(neg_label)
        if include_ctrl3:
            control_labels.add(ctrl3_label)

        return render_template('plate_mapper_results.html', plates=plates, plate_labels=plate_labels,
                               control_labels=control_labels)
    return render_template('plate_mapper.html')


@app.route('/plate_mapper/settings')
def plate_mapper_settings():
    return render_template('plate_mapper_settings.html')


@app.route('/plate_mapper/api/presets', methods=['GET'])
def pm_api_get_presets():
    return jsonify(pm_load_presets())


@app.route('/plate_mapper/api/presets', methods=['POST'])
def pm_api_save_preset():
    data = request.json
    if not data or "name" not in data or "settings" not in data:
        return jsonify({"error": "Invalid data"}), 400
    if pm_save_preset(data["name"], data["settings"]):
        return jsonify({"message": f'Preset "{data["name"]}" saved.'})
    return jsonify({"error": "Failed to save preset"}), 500


@app.route('/plate_mapper/api/presets/<name>', methods=['DELETE'])
def pm_api_delete_preset(name):
    if pm_delete_preset(name):
        return jsonify({"message": f'Preset "{name}" deleted.'})
    return jsonify({"error": "Preset not found"}), 404


if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port)