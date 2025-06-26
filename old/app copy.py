from flask import Flask, render_template, request, redirect, url_for, send_from_directory, flash, jsonify
import os
import uuid
import subprocess
from datetime import datetime


from nta_utils import (
    process_csv_to_template,
    extract_final_titres_openpyxl as extract_final_titres_xlwings,
    save_template_path,
    load_template_path,
    load_settings,
    save_settings,
)

UPLOAD_FOLDER = os.path.join("static", "uploads")
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

app = Flask(__name__)
app.secret_key = "your-secret-key"
app.config["UPLOAD_FOLDER"] = UPLOAD_FOLDER

@app.route("/")
def index():
    return render_template("index.html")

@app.route("/help")
def help_page():
    return render_template("help.html")

@app.route("/settings", methods=["GET", "POST"])
def settings():
    current_settings = load_settings()

    if request.method == "POST":
        file = request.files.get("template_file")
        if file and file.filename.endswith(".xlsx"):
            filename = f"template_{uuid.uuid4().hex}.xlsx"
            filepath = os.path.join("excel_templates", filename)
            os.makedirs("excel_templates", exist_ok=True)
            file.save(filepath)
            save_template_path(filepath)
            flash("New template saved and path updated.", "success")

        timestamp_flag = request.form.get("timestamp_in_filename") == "on"
        new_settings = current_settings.copy()
        new_settings["timestamp_in_filename"] = timestamp_flag

        save_settings(new_settings)
        flash("Settings saved.", "success")
        return redirect(url_for("settings"))

    return render_template("settings.html", settings=current_settings)


@app.route("/process", methods=["POST"])
def process():
    file = request.files["csv_file"]
    if not file:
        flash("No CSV file uploaded.", "danger")
        return redirect(url_for("index"))
    
    # Retrieve assay inputs
    assay_title = request.form.get("assay_title", "")
    pseudotypes = request.form.get("pseudotype_text", "")
    sample_ids = request.form.get("sample_id_text", "")
    try:
        num_pseudotypes = int(request.form.get("num_pseudotypes", "1"))
        if num_pseudotypes not in [1, 2, 3, 4]:
            raise ValueError()
    except ValueError:
        flash("Invalid pseudotype count. Must be 1–4.", "danger")
        return redirect(url_for("index"))
    
    # Load user settings to check if timestamp should be included in filename
    settings = load_settings()
    timestamp_in_filename = settings.get("timestamp_in_filename", True)

    safe_title = assay_title.strip().replace(" ", "_")

    if timestamp_in_filename:
        timestamp = datetime.now().strftime("%Y-%m-%d")
        filename = f"{safe_title}_{timestamp}.xlsx"
    else:
        filename = f"{safe_title}.xlsx"
    
    # Save the uploaded CSV file
    csv_filename = filename.replace(".xlsx", ".csv")
    csv_path = os.path.join(app.config["UPLOAD_FOLDER"], csv_filename)
    file.save(csv_path)

    # Set output Excel path
    output_path = os.path.join(app.config["UPLOAD_FOLDER"], filename)

    # Load the template path
    try:
        template_path = load_template_path()
    except Exception as e:
        flash(str(e), "danger")
        return redirect(url_for("index"))
    
    # Process data
    try:
        process_csv_to_template(
            csv_path=csv_path,
            template_path=template_path,
            output_path=output_path,
            num_pseudotypes=num_pseudotypes,
            pseudotype_texts=pseudotypes,
            assay_title_text=assay_title,
            sample_id_text=sample_ids,
        )
        extract_final_titres_xlwings(output_path)

        return render_template(
            "results.html",
            excel_file=filename,
            plot_file=None,
            settings=settings
        )
    
    except Exception as e:
        flash(f"Data processing error: {e}", "danger")
        return redirect(url_for("index"))

@app.route("/generate_graphs", methods=["POST"])
def generate_graphs():
    excel_file = request.form.get("excel_file")
    if not excel_file:
        flash("No Excel file found for graph generation.", "danger")
        return redirect(url_for("index"))

    input_path = os.path.join(app.config["UPLOAD_FOLDER"], excel_file)
    base_name = os.path.splitext(excel_file)[0]
    output_plot = f"{base_name}.png"
    output_path = os.path.join(app.config["UPLOAD_FOLDER"], output_plot)

    r_script = os.path.join(os.getcwd(), "process_data.R")

    # Load settings to get timestamp preference and presets
    settings = load_settings()
    include_timestamp = settings.get("timestamp_in_filename", True)

    # Get the active preset name
    active_preset_name = settings.get("selected_preset", None)
    presets = settings.get("presets", {})

    # Default colours if no presets or selected preset
    default_colours = {
        "Q1": "#ff7e79",
        "Q2": "#ffd479",
        "Q3": "#009193",
        "Q4": "#d783ff"
    }

    # Get colours from active preset or fallback to defaults
    if active_preset_name and active_preset_name in presets:
        colours = presets[active_preset_name]
    else:
        colours = default_colours

    q1_colour = colours.get("Q1", default_colours["Q1"])
    q2_colour = colours.get("Q2", default_colours["Q2"])
    q3_colour = colours.get("Q3", default_colours["Q3"])
    q4_colour = colours.get("Q4", default_colours["Q4"])

    try:
        subprocess.run(
            [
                "Rscript",
                r_script,
                input_path,
                output_path,
                str(include_timestamp).lower(),
                q1_colour,
                q2_colour,
                q3_colour,
                q4_colour
            ],
            check=True,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
        )
        flash("Graph generated successfully.", "success")

        return render_template(
            "results.html",
            excel_file=excel_file,
            plot_file=output_plot,
            settings=settings
        )

    except subprocess.CalledProcessError as e:
        flash(f"R script failed: {e.stderr}", "danger")
        return redirect(url_for("index"))

@app.route("/download/<filename>")
def download(filename):
    return send_from_directory(app.config["UPLOAD_FOLDER"], filename, as_attachment=True)

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
    return jsonify({"status": "success", "message": "Preset saved."}), 200


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



if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port)
