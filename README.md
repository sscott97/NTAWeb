# NTAWeb — Neutralising Titre Automator

A Flask web app that processes raw neutralisation assay CSV data into Excel reports and generates publication-quality graphs, including linear interpolation titres, sigmoid curve fitting, and titre comparison.

---

## Prerequisites

| Dependency | Minimum version | Notes |
|---|---|---|
| Python | 3.8+ | |
| R | 4.0+ | Must be on your system `PATH` |
| pip packages | — | see below |
| R packages | — | see below |

**Python packages**

```bash
pip install flask openpyxl Pillow
```

**R packages** — run once inside an R session:

```r
install.packages(c("ggplot2", "dplyr", "tidyr", "readxl", "cowplot", "grid"))
```

---

## Setup

```bash
git clone https://github.com/sscott97/NTAWeb
cd NTAWeb
pip install flask openpyxl Pillow
```

No further configuration is needed. `settings.json` is created automatically on first run.

---

## Running the app

```bash
python app.py
```

Then open **http://localhost:5000** in your browser.

---

## Workflow

### 1. Home page — upload and label your data

- **Input mode**
  - *Standard* — CSV exported from Kaleido in standard format (includes a header block before the 8×12 data)
  - *Data Only* — CSV exported with "Data Only" option (raw 8×12 numbers, no header)
- **Pseudotype count** — 1–4 pseudotypes per plate. Controls how plates are labelled and grouped.
- **Assay title, Pseudotype IDs, Sample IDs** — used to label the output Excel and graphs.
- Upload the CSV and click **Process**.

### 2. Data Analysis

After processing you land on the Data Analysis Hub. Three options:

| Option | What it does |
|---|---|
| **Linear Interpolation** | Calculates NT50 / NT90 by interpolating between measured dilution points. Default method. |
| **Sigmoid Curve Fitting** | Fits a four-parameter logistic curve to each sample. Calculates IC50. Requires R. |
| **Titre Comparison** | Compares NT50 (linear) vs IC50 (sigmoid). Requires curve fitting to be run first. |

### 3. Results pages

Each analysis page lets you:
- View per-plate graphs inline (filterable by pseudotype/sample/quality)
- Download an Excel file of results
- Download individual or all graphs as PNGs

---

## Excel templates

Templates live in `excel_templates/`. The active template is selected in **Settings**.

- **Data range**: raw plate data is written into cells **B5:M12** of each sheet.
- **Dilution series**: stored in **A5:A12** — the app reads these to label graphs and calculate titres.
- You can upload your own `.xlsx` template through the Settings page.
- **Create Dilution Variant**: duplicate the active template with a different dilution series (A5:A12) without modifying the original.

---

## Settings

Navigate to **Settings** (top-right) to configure:

| Setting | Description |
|---|---|
| Active template | Which Excel template is used for processing |
| Include timestamp in filename | Appends `_YYYYMMDD_HHMMSS` to output filenames |
| Flag triplicate errors | Highlights wells where one replicate deviates > threshold (log₂) |
| Outlier threshold | Log₂ fold-change cutoff for error flagging |
| Default CSV mode | Standard or Data Only — pre-selects on the home page |
| Default pseudotype count | Pre-selects the pseudotype pill on the home page |
| Sigmoid R² threshold | Fits below this R² are marked "Unstable" |
| Comparison disagreement threshold | Log₂ fold-change above which NT50 vs IC50 are flagged as disagreeing |
| Graph colour presets | Save/delete named colour schemes for graphs |
| Theme | Dark (default), Dark Paper, Light, or Forest |

---

## File structure

```
app.py                        # Flask routes and main logic
nta_utils.py                  # Data processing utilities and settings helpers
process_data.R                # Graph generation (per-plate NT50 plots + summary)
fit_sigmoids.R                # Four-parameter logistic curve fitting
plot_sigmoids.R               # Sigmoid curve graph generation
compare_titres.R              # NT50 vs IC50 titre comparison plots
boxplot_NT50.R                # Boxplot generation for linear results

excel_templates/              # Built-in and user-uploaded Excel templates
templates/                    # Jinja2 HTML templates
static/                       # CSS themes, favicon

settings.json                 # Auto-created; stores all user settings and presets
config.json                   # Stores active template path
```

---

## Troubleshooting

**R not found**
Ensure `Rscript` is on your `PATH`. Test with `Rscript --version` in a terminal.

**Missing R packages**
Run `install.packages(c("ggplot2", "dplyr", "tidyr", "readxl", "cowplot", "grid"))` in R.

**CSV not processing**
- Check you selected the correct input mode (Standard vs Data Only).
- The data block must be exactly 8 rows × 12 columns of numeric values.

**Template errors**
- Templates must be `.xlsx` files.
- Raw data is written to B5:M12; dilution series must be in A5:A12.

**Sigmoid curve fitting fails**
- Requires at least 4 dilution points with a measurable signal drop.
- Samples that fail to converge are marked "Unstable" and excluded from plots by default.

---

## Contact

Sam Scott — Sam.Scott.2@glasgow.ac.uk
