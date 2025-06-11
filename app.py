from flask import Flask, render_template, request, redirect, url_for, send_from_directory, flash
import os
import uuid
import subprocess


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
    if request.method == "POST":
        timestamp = request.form.get("timestamp_in_title") == "on"
        new_settings = {
            "timestamp_in_title": timestamp,
            "Q1_color": request.form.get("Q1_color", "#ff0000"),
            "Q2_color": request.form.get("Q2_color", "#0000ff"),
            "Q3_color": request.form.get("Q3_color", "#00ff00"),
            "Q4_color": request.form.get("Q4_color", "#800080"),
        }

        file = request.files.get("template_file")
        if file and file.filename.endswith(".xlsx"):
            filename = f"template_{uuid.uuid4().hex}.xlsx"
            filepath = os.path.join("Templates", filename)
            os.makedirs("Templates", exist_ok=True)
            file.save(filepath)
            save_template_path(filepath)
            flash("New template saved and path updated.", "success")

        save_settings(new_settings)
        flash("Settings saved.", "success")
        return redirect(url_for("settings"))

    return render_template("settings.html", settings=load_settings())


@app.route("/process", methods=["POST"])
def process():
    file = request.files["csv_file"]
    if not file:
        flash("No CSV file uploaded.", "danger")
        return redirect(url_for("index"))

    filename = f"input_{uuid.uuid4().hex}.csv"
    csv_path = os.path.join(app.config["UPLOAD_FOLDER"], filename)
    file.save(csv_path)

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

    # Get template path
    try:
        template_path = load_template_path()
    except Exception as e:
        flash(str(e), "danger")
        return redirect(url_for("index"))

    # Set output Excel path
    output_filename = f"processed_{uuid.uuid4().hex}.xlsx"
    output_path = os.path.join(app.config["UPLOAD_FOLDER"], output_filename)

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
            settings=load_settings()
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
    output_plot = f"graph_{uuid.uuid4().hex}.png"
    output_path = os.path.join(app.config["UPLOAD_FOLDER"], output_plot)

    include_timestamp = "include_timestamp" in request.form

    q1 = request.form.get("Q1_color", "#ff0000")
    q2 = request.form.get("Q2_color", "#0000ff")
    q3 = request.form.get("Q3_color", "#00ff00")
    q4 = request.form.get("Q4_color", "#800080")

    r_script = os.path.join(os.getcwd(), "process_data.R")

    try:
        subprocess.run(
            [
                "Rscript",
                r_script,
                input_path,
                output_path,
                str(include_timestamp).upper(),
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
            settings=load_settings()
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