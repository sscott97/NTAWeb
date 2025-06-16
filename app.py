from flask import Flask, render_template, request, redirect, url_for, send_from_directory, flash
import os
import uuid
import subprocess
from datetime import datetime
import re
import json



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

PRESETS_FILE = "colour_presets.json"

def load_colour_presets():
    if os.path.exists(PRESETS_FILE):
        with open(PRESETS_FILE, "r") as f:
            return json.load(f)
    else:
        # Return a default preset if file not found
        return {
            "Default": {
                "Q1_colour": "#ff0000",
                "Q2_colour": "#0000ff",
                "Q3_colour": "#00ff00",
                "Q4_colour": "#800080",
            }
        }

def save_colour_presets(presets):
    with open(PRESETS_FILE, "w") as f:
        json.dump(presets, f, indent=2)

@app.route("/")
def index():
    return render_template("index.html")


@app.route("/help")
def help_page():
    return render_template("help.html")


@app.route("/settings", methods=["GET", "POST"])
def settings():
    presets = load_colour_presets()

    if request.method == "POST":
        # --- Handle deletion ---
        delete_preset = request.form.get("delete_preset")
        if delete_preset:
            if delete_preset in presets:
                del presets[delete_preset]
                save_colour_presets(presets)
                flash(f"Preset '{delete_preset}' deleted.", "success")
            else:
                flash(f"Preset '{delete_preset}' not found.", "danger")
            return redirect(url_for("settings"))

        # --- Handle saving ---
        timestamp_in_filename = request.form.get("timestamp_in_filename") == "on"

        new_settings = {
            "timestamp_in_filename": timestamp_in_filename,
            "Q1_colour": request.form.get("Q1_colour", "#ff0000"),
            "Q2_colour": request.form.get("Q2_colour", "#0000ff"),
            "Q3_colour": request.form.get("Q3_colour", "#00ff00"),
            "Q4_colour": request.form.get("Q4_colour", "#800080"),
        }

        file = request.files.get("template_file")
        if file and file.filename.endswith(".xlsx"):
            filename = f"template_{uuid.uuid4().hex}.xlsx"
            filepath = os.path.join("Templates", filename)
            os.makedirs("Templates", exist_ok=True)
            file.save(filepath)
            save_template_path(filepath)
            flash("New template saved and path updated.", "success")

        new_preset_name = request.form.get("new_preset_name", "").strip()
        if new_preset_name:
            presets[new_preset_name] = {
                "Q1_colour": new_settings["Q1_colour"],
                "Q2_colour": new_settings["Q2_colour"],
                "Q3_colour": new_settings["Q3_colour"],
                "Q4_colour": new_settings["Q4_colour"],
            }
            save_colour_presets(presets)
            flash(f"Preset '{new_preset_name}' saved.", "success")

        save_settings(new_settings)
        flash("Settings saved.", "success")
        return redirect(url_for("settings"))

    current_settings = load_settings()
    return render_template("settings.html", settings=current_settings, presets=presets)


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
    
    # Save the uploaded CSV file with filename (use .csv extension for CSV)
    csv_filename = filename.replace(".xlsx", ".csv")
    csv_path = os.path.join(app.config["UPLOAD_FOLDER"], csv_filename)
    file.save(csv_path)

    # Set output Excel path and name (xlsx)
    output_filename = filename
    output_path = os.path.join(app.config["UPLOAD_FOLDER"], output_filename)

    # Load the template path
    try:
        template_path = load_template_path()
    except Exception as e:
        flash(str(e), "danger")
        return redirect(url_for("index"))
    
    # Process data and render results
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
            excel_file=output_filename,
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

    # Load settings and get timestamp inclusion flag
    settings = load_settings()

    base_name = os.path.splitext(excel_file)[0]
    output_plot = f"{base_name}.png"
    output_path = os.path.join(app.config["UPLOAD_FOLDER"], output_plot)

    q1 = settings.get("Q1_colour", "#ff0000")
    q2 = settings.get("Q2_colour", "#0000ff")
    q3 = settings.get("Q3_colour", "#00ff00")
    q4 = settings.get("Q4_colour", "#800080")

    r_script = os.path.join(os.getcwd(), "process_data.R")

    try:
        subprocess.run(
            [
                "Rscript",
                r_script,
                input_path,
                output_path,
                q1,
                q2,
                q3,
                q4,
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


if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port)