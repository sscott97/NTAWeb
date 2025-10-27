# NTAWeb
Neutralising Titre Automator (NTA) Web Application

Automate the processing of raw neutralisation assay data into Excel files and generate customizable graphs.

## Overview

The Neutralising Titre Automator (NTA) is a Flask-based web application designed to streamline the analysis of raw CSV data exported from the Kaleido software on the Perkin-Elmer Luminometer. It processes this data using customizable Excel templates to automatically calculate neutralising antibody titres, then generates detailed graphs based on the processed results.

## Features

- Upload raw CSV files (8x12 data blocks, Data Only mode).
- Select or upload Excel template files to customize output.
- Input assay title, pseudotype IDs, and sample IDs for labeling.
- Automatically generates Excel reports with organized plate layouts.
- Supports 1–4 pseudotypes per plate with dynamic plate labeling logic.
- Create graphs with configurable colour presets.
- Download processed Excel files and generated graphs directly.
- Save and manage color presets and app settings.

## Requirements

### Python 3.8+

### R with the following packages installed:
- readxl
- ggplot2
- dplyr
- tidyr

Install required R packages in your R environment:

install.packages(c("readxl", "ggplot2", "dplyr", "tidyr"))

### Flask and Python dependencies

Install via pip:

pip install flask openpyxl


## Setup

Clone this repository:

git clone https://github.com/sscott97/NTAWeb

cd ntaweb

Place your default Excel templates in the excel_templates/ directory. You can add multiple templates and select the active one in the Settings page.

Ensure the process_data.R script is present in the project root.

Run the Flask app:
python app.py

Open your browser and navigate to http://localhost:5000

## Usage

Home Page: Upload your raw CSV file and enter assay information.

Settings: Select an Excel template or upload a new one, toggle filename timestamp option, and manage graph color presets.

Process: After uploading, the app processes the CSV, generates an Excel report, and provides a download link.

Graphs: Generate graphs based on processed data with the selected color presets. Graphs download as PNG images.


## How It Works

The app reads CSV data, splits it into 8x12 blocks.

Each block is pasted into a new Excel sheet copied from the selected template.

Plate layouts update pseudotype labels and sample IDs according to the number of pseudotypes chosen.

The processed workbook includes a summary sheet and neutralising titres.

An R script generates graphs from the processed Excel file, respecting color presets and including dilution info from the template.


## Customization

Excel Templates: You can create and upload your own templates matching the required layout.

Color Presets: Save and apply custom graph color presets in settings.

Filename Settings: Toggle timestamp inclusion in output filenames.

Pseudotype Counts: Choose between 1 to 4 pseudotypes affecting layout.

## File Structure

<pre>
/app.py                # Main Flask app
/nta_utils.py          # Utility functions for processing Excel and CSV
/process_data.R        # R script for graph generation
/excel_templates/      # Folder for Excel template files
/templates/            # HTML templates (index.html, settings.html, help.html, results.html)
/static/               # Static assets (CSS, JS, images)
/settings.json         # JSON file to store user settings and presets
/config.json           # JSON file to store template location
</pre>

## Troubleshooting

Ensure R and required packages are installed and accessible in your system PATH.

Excel templates must match expected formatting and layout (8x12 data starting at B5).

CSV files must be saved in "Data Only" mode from Kaleido software.

Temporary files are cleaned up automatically; if you encounter issues, check file permissions.


## Contact

For questions or issues, contact via email: Sam.Scott.2@glasgow.ac.uk

