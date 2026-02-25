from flask import Flask, render_template, request, redirect, url_for, send_file, flash, jsonify, send_from_directory, session
import os
import uuid
import json
import subprocess
from datetime import datetime
from io import BytesIO
import tempfile


from nta_utils import (
    process_csv_to_template,
    extract_final_titres_openpyxl as extract_final_titres_xlwings,
    save_template_path,
    load_template_path,
    load_settings,
    save_settings,
    generate_sigmoid_csv,
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
    return render_template("help.html")

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
        new_settings = current_settings.copy()
        new_settings["timestamp_in_filename"] = timestamp_flag
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
    )

    extract_final_titres_xlwings(output_bytes)

    from nta_utils import add_default_to_final_titres
    add_default_to_final_titres(output_bytes)

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

    return render_template(
        "results.html",
        excel_file_id=file_id,
        filename=filename,
        settings=settings,
        plot_file=None
    )


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

    # ── Read session-only overrides from the results page ──
    # Field names must exactly match the hidden input name="" attributes in results.html
    graph_preset    = request.form.get("graph_preset", "").strip()
    graph_quadrants = request.form.get("graph_quadrants", "").strip()

    # Colours: use the override preset if it exists, otherwise fall back to saved
    if graph_preset and graph_preset in settings.get("presets", {}):
        colours = settings["presets"][graph_preset]
    else:
        active_preset_name = settings.get("selected_preset", None)
        colours = settings.get("presets", {}).get(active_preset_name, default_colours)

    q1_colour = colours.get("Q1", default_colours["Q1"])
    q2_colour = colours.get("Q2", default_colours["Q2"])
    q3_colour = colours.get("Q3", default_colours["Q3"])
    q4_colour = colours.get("Q4", default_colours["Q4"])

    # Quadrants: use override JSON if provided, otherwise fall back to saved
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
    """
    Perform sigmoid curve fitting on the processed Excel data.
    
    This generates sigmoidData.csv, runs fit_sigmoids.R, and returns
    downloadable outputs (IC50s.csv and plot images).
    """
    file_id = request.form.get("file_id")
    if not file_id or file_id not in in_memory_files:
        flash("No Excel file found for curve fitting.", "danger")
        return redirect(url_for("index"))

    file_info = in_memory_files[file_id]
    file_bytes = file_info["data"]
    filename = file_info["name"]

    try:
        # Create temporary files
        file_stream = BytesIO(file_bytes)
        file_stream.seek(0)

        with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp_excel:
            tmp_excel.write(file_stream.read())
            excel_path = tmp_excel.name

        # Create temporary directory for outputs
        output_dir = tempfile.mkdtemp(prefix="sigmoid_")
        
        # Generate sigmoidData.csv
        sigmoid_csv_path = os.path.join(output_dir, "sigmoidData.csv")
        generate_sigmoid_csv(excel_path, sigmoid_csv_path)

        # Run R script for sigmoid fitting
        r_script = os.path.join(os.getcwd(), "fit_sigmoids.R")
        
        result = subprocess.run(
            ["Rscript", r_script, sigmoid_csv_path, output_dir],
            check=True,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
        )

        # Collect output files
        ic50_path = os.path.join(output_dir, "IC50s.csv")
        
        # Find all generated PNG files
        plot_files = [f for f in os.listdir(output_dir) if f.endswith('.png')]
        
        # Store files in memory for download
        output_files = {}
        
        # Store IC50s.csv
        if os.path.exists(ic50_path):
            with open(ic50_path, 'rb') as f:
                output_files['IC50s.csv'] = f.read()
        
        # Store all plot files
        for plot_file in plot_files:
            plot_path = os.path.join(output_dir, plot_file)
            with open(plot_path, 'rb') as f:
                output_files[plot_file] = f.read()
        
        # Clean up temporary files
        os.remove(excel_path)
        for f in os.listdir(output_dir):
            os.remove(os.path.join(output_dir, f))
        os.rmdir(output_dir)
        
        # Store output files in memory with new UUID
        fitting_id = uuid.uuid4().hex
        in_memory_files[fitting_id] = {
            "data": output_files,
            "name": "sigmoid_fitting_results",
            "type": "sigmoid_results"
        }
        
        return render_template(
            "curve_fitting_results.html",
            fitting_id=fitting_id,
            output_files=list(output_files.keys()),
            settings=load_settings()
        )

    except subprocess.CalledProcessError as e:
        flash(f"R script failed: {e.stderr}", "danger")
        return redirect(url_for("index"))
    except Exception as e:
        flash(f"Curve fitting error: {str(e)}", "danger")
        return redirect(url_for("index"))


@app.route("/download_sigmoid/<fitting_id>/<filename>")
def download_sigmoid(fitting_id, filename):
    """Download individual sigmoid fitting output files"""
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
    
    # Determine mimetype
    if filename.endswith('.csv'):
        mimetype = 'text/csv'
    elif filename.endswith('.png'):
        mimetype = 'image/png'
    else:
        mimetype = 'application/octet-stream'
    
    return send_file(
        file_bytes,
        as_attachment=True,
        download_name=filename,
        mimetype=mimetype
    )


if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port)