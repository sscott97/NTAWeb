# ==============================================================================
# ===== Compare NT50 and IC50 Titres ===========================================
# ==============================================================================

suppressPackageStartupMessages({
  library(tidyverse)
  library(scales)
  library(plotly)
  library(readxl)
})

# ----- Command Line Arguments -------------------------------------------------

args <- commandArgs(trailingOnly = TRUE)

if (length(args) < 3) {
  stop("Usage: Rscript compare_titres.R <excel_file> <ic50_csv> <output_dir>")
}

excel_file <- args[1]
ic50_csv   <- args[2]
output_dir <- args[3]

# Create output directory if needed
if (!dir.exists(output_dir)) {
  dir.create(output_dir, recursive = TRUE)
}

# ==============================================================================
# ===== 1. Extract NT50 by Linear Interpolation from Plate Sheets =============
# ==============================================================================

cat("── Extracting NT50 by linear interpolation from Excel ──\n")

sheets <- excel_sheets(excel_file)
plate_sheets <- sheets[grepl("^Plate\\d+$", sheets)]
plate_numbers <- as.numeric(gsub("^Plate", "", plate_sheets))
plate_sheets <- plate_sheets[order(plate_numbers)]

if (length(plate_sheets) == 0) {
  stop("No Plate sheets found in Excel file.")
}

# Quadrant definitions: pseudotype cell, sample cell, data columns (1-indexed from A=1)
quad_defs <- list(
  list(pt_cell = "B3", sid_cell = "B4", data_cols = 2:4),   # B, C, D
  list(pt_cell = "E3", sid_cell = "E4", data_cols = 5:7),   # E, F, G
  list(pt_cell = "H3", sid_cell = "H4", data_cols = 8:10),  # H, I, J
  list(pt_cell = "K3", sid_cell = "K4", data_cols = 11:13)   # K, L, M
)

# Helper: read a single cell value from a sheet
read_cell <- function(sheet_name, cell_ref) {
  val <- tryCatch(
    suppressMessages(read_excel(excel_file, sheet = sheet_name, range = cell_ref, col_names = FALSE)[[1]]),
    error = function(e) NA
  )
  if (length(val) == 0 || is.na(val) || trimws(as.character(val)) == "") return(NA_character_)
  return(trimws(as.character(val)))
}

unlabelled_counter <- 1
nt50_rows <- list()

for (sheet_name in plate_sheets) {

  # Read dilution series from A5:A12 (8 rows)
  dilutions <- tryCatch(
    suppressMessages(read_excel(excel_file, sheet = sheet_name, range = "A5:A12", col_names = FALSE)[[1]]),
    error = function(e) rep(NA_real_, 8)
  )
  dilutions <- as.numeric(dilutions)

  # Read full data block B5:M12 (8 rows x 12 cols)
  raw_block <- tryCatch(
    suppressMessages(read_excel(excel_file, sheet = sheet_name, range = "B5:M12", col_names = FALSE)),
    error = function(e) NULL
  )
  if (is.null(raw_block)) next

  # Coerce to numeric matrix
  data_matrix <- as.data.frame(lapply(raw_block, as.numeric))

  for (quad in quad_defs) {

    pt_val  <- read_cell(sheet_name, quad$pt_cell)
    sid_val <- read_cell(sheet_name, quad$sid_cell)

    # Skip unused quadrants (no pseudotype name)
    if (is.na(pt_val)) next

    pseudotype_name <- pt_val

    if (is.na(sid_val)) {
      sample_name <- paste0("Unlabelled", unlabelled_counter)
      unlabelled_counter <- unlabelled_counter + 1
    } else {
      sample_name <- sid_val
    }

    # Data columns are offset by 1 because raw_block starts at col B (col index 1 in data_matrix)
    col_indices <- quad$data_cols - 1  # convert from spreadsheet col to data_matrix col

    nt50_replicates <- c()

    for (ci in col_indices) {
      lum <- data_matrix[[ci]]  # 8 values (rows 5-12)

      if (length(lum) < 8 || is.na(lum[8])) next

      nsc_value <- lum[8]
      if (is.na(nsc_value) || nsc_value <= 0) next

      target <- nsc_value * 0.5
      nt50_val <- NA_real_

      # Linear interpolation: scan adjacent pairs in rows 1-7 (dilution rows 5-11)
      for (i in 1:7) {
        if (i >= 8) break
        y1 <- lum[i]
        y2 <- lum[i + 1]
        x1 <- dilutions[i]
        x2 <- dilutions[i + 1]

        if (is.na(y1) || is.na(y2) || is.na(x1) || is.na(x2)) next

        # Check if the target falls between these two points
        crosses <- (y1 >= target & y2 <= target) | (y1 <= target & y2 >= target)
        if (crosses && y2 != y1) {
          nt50_val <- (target - y1) / (y2 - y1) * (x2 - x1) + x1
          break
        }
      }

      # Boundary cases (if no crossing found)
      if (is.na(nt50_val)) {
        valid_lum <- lum[1:7]
        valid_lum <- valid_lum[!is.na(valid_lum)]
        if (length(valid_lum) > 0 && all(valid_lum < target)) {
          nt50_val <- dilutions[1]
        } else if (length(valid_lum) > 0 && all(valid_lum > target)) {
          nt50_val <- dilutions[7]
        }
      }

      if (!is.na(nt50_val)) {
        nt50_replicates <- c(nt50_replicates, nt50_val)
      }
    }

    if (length(nt50_replicates) > 0) {
      nt50_avg <- mean(nt50_replicates)
      nt50_rows[[length(nt50_rows) + 1]] <- data.frame(
        Pseudotype = pseudotype_name,
        Sample_ID  = sample_name,
        NT50       = round(nt50_avg),
        stringsAsFactors = FALSE
      )
    }
  }
}

if (length(nt50_rows) == 0) {
  stop("No NT50 data could be calculated from Plate sheets.")
}

nt50_data <- bind_rows(nt50_rows) %>%
  filter(!is.na(NT50)) %>%
  mutate(NT50_numeric = as.numeric(NT50))

cat(sprintf("  Extracted %d NT50 values from %d plate(s)\n",
            nrow(nt50_data), length(plate_sheets)))


# ==============================================================================
# ===== 2. Load IC50 Data =====================================================
# ==============================================================================

ic50_data <- read_csv(ic50_csv, show_col_types = FALSE) %>%
  select(Sample, Virus, IC50_Titre = Titre, Quality,
         Lower, Upper, Slope, IC50, LOD_Flag) %>%
  filter(!is.na(IC50_Titre))


# ==============================================================================
# ===== 3. Merge Datasets =====================================================
# ==============================================================================

nt50_data <- nt50_data %>%
  select(Pseudotype, Sample_ID, NT50_numeric)

merged_data <- nt50_data %>%
  inner_join(
    ic50_data,
    by = c("Pseudotype" = "Virus", "Sample_ID" = "Sample"),
    relationship = "many-to-many"
  )

# If no matches found, try case-insensitive matching
if (nrow(merged_data) == 0) {
  cat("Warning: No exact matches found. Trying case-insensitive matching...\n")

  nt50_data <- nt50_data %>%
    mutate(
      Pseudotype_lower = tolower(trimws(Pseudotype)),
      Sample_ID_lower  = tolower(trimws(Sample_ID))
    )

  ic50_data <- ic50_data %>%
    mutate(
      Virus_lower  = tolower(trimws(Virus)),
      Sample_lower = tolower(trimws(Sample))
    )

  merged_data <- nt50_data %>%
    inner_join(
      ic50_data,
      by = c("Pseudotype_lower" = "Virus_lower", "Sample_ID_lower" = "Sample_lower"),
      relationship = "many-to-many"
    ) %>%
    select(Pseudotype, Sample_ID, NT50_numeric, Virus, Sample,
           IC50_Titre, Quality, Lower, Upper, Slope, IC50, LOD_Flag)
}

# Remove rows that cannot be analysed
merged_data <- merged_data %>%
  filter(!is.na(NT50_numeric), !is.na(IC50_Titre),
         NT50_numeric > 0, IC50_Titre > 0) %>%
  mutate(
    log10_NT50 = log10(NT50_numeric),
    log10_IC50 = log10(IC50_Titre),
    log2_fold_difference     = log2(NT50_numeric / IC50_Titre),
    abs_log2_fold_difference = abs(log2_fold_difference),
    disagreement = abs_log2_fold_difference > 1
  )


# ==============================================================================
# ===== 4. Curve Fit Hybrid NT50 ==============================================
# ==============================================================================

merged_data <- merged_data %>%
  mutate(
    hybrid_log2 = case_when(
      is.na(Lower) | is.na(Upper) | is.na(Slope) | is.na(IC50) ~ NA_real_,
      Lower >= 50  ~ NA_real_,
      Upper <= 50  ~ NA_real_,
      Slope == 0   ~ NA_real_,
      TRUE ~ IC50 - (1 / Slope) * log((Upper - Lower) / (50 - Lower) - 1)
    ),
    `Curve Fit Hybrid NT50` = ifelse(is.na(hybrid_log2), NA_real_, round(2^(-hybrid_log2), 2))
  ) %>%
  select(-hybrid_log2)

cat(sprintf("Successfully merged %d samples across %d viruses\n",
            nrow(merged_data),
            length(unique(merged_data$Pseudotype))))

hybrid_valid <- sum(!is.na(merged_data$`Curve Fit Hybrid NT50`))
cat(sprintf("Curve Fit Hybrid NT50 calculated for %d / %d samples\n",
            hybrid_valid, nrow(merged_data)))


# ==============================================================================
# ===== 5. Statistics ==========================================================
# ==============================================================================

if (nrow(merged_data) > 0) {

  correlation <- cor(merged_data$log10_NT50, merged_data$log10_IC50,
                     use = "complete.obs")

  lm_fit    <- lm(log10_IC50 ~ log10_NT50, data = merged_data)
  r_squared <- summary(lm_fit)$r.squared

  top_mismatches <- merged_data %>%
    arrange(desc(abs_log2_fold_difference)) %>%
    head(5) %>%
    select(Sample = Sample_ID, Virus = Pseudotype, NT50 = NT50_numeric,
           IC50_Titre, Log2_Fold_Difference = log2_fold_difference, Quality)

  stats <- data.frame(
    n_samples            = nrow(merged_data),
    correlation          = correlation,
    r_squared            = r_squared,
    n_disagreements      = sum(merged_data$disagreement),
    percent_disagreement = 100 * mean(merged_data$disagreement),
    median_abs_log2_fold = median(merged_data$abs_log2_fold_difference)
  )

  # ── Static Plot (PNG) ──────────────────────────────────────────────────────

  plot_file <- file.path(output_dir, "titre_comparison.png")

  p <- ggplot(merged_data, aes(x = NT50_numeric, y = IC50_Titre)) +
    geom_point(aes(color = disagreement), size = 3, alpha = 0.7) +
    geom_abline(slope = 1, intercept = 0, linetype = "dashed",
                color = "gray40", linewidth = 0.8) +
    geom_smooth(method = "lm", se = TRUE, color = "#28a745",
                fill = "#28a74533", linewidth = 1) +
    scale_x_log10(
      breaks = scales::trans_breaks("log10", function(x) 10^x),
      labels = scales::trans_format("log10", scales::math_format(10^.x))
    ) +
    scale_y_log10(
      breaks = scales::trans_breaks("log10", function(x) 10^x),
      labels = scales::trans_format("log10", scales::math_format(10^.x))
    ) +
    scale_color_manual(
      values = c("FALSE" = "#3498db", "TRUE" = "#e74c3c"),
      labels = c("FALSE" = "Agreement (\u22642-fold)", "TRUE" = "Disagreement (>2-fold)"),
      name = ""
    ) +
    annotation_logticks() +
    labs(
      title    = "NT50 vs IC50 Titre Comparison",
      subtitle = sprintf("n = %d | r = %.3f | R\u00b2 = %.3f",
                         nrow(merged_data), correlation, r_squared),
      x = "NT50 (Linear Interpolation)",
      y = "NT50 / IC50 (Curve Fitting)"
    ) +
    theme_minimal(base_size = 12) +
    theme(
      panel.grid.minor = element_blank(),
      panel.border     = element_rect(color = "black", fill = NA),
      plot.title       = element_text(hjust = 0.5, face = "bold", size = 14),
      plot.subtitle    = element_text(hjust = 0.5, color = "gray40", size = 11),
      legend.position  = "bottom"
    )

  ggsave(plot_file, p, width = 8, height = 7, dpi = 300, units = "in")

  # ── Interactive Plot (HTML) ────────────────────────────────────────────────

  interactive_plot_file <- file.path(output_dir, "titre_comparison_interactive.html")

  hover_text <- paste0(
    "<b>Sample:</b> ", merged_data$Sample_ID, "<br>",
    "<b>Pseudotype:</b> ", merged_data$Pseudotype, "<br>",
    "<b>NT50:</b> ", round(merged_data$NT50_numeric, 1), "<br>",
    "<b>IC50:</b> ", round(merged_data$IC50_Titre, 1), "<br>",
    "<b>Log\u2082 Fold Diff:</b> ", round(merged_data$log2_fold_difference, 2), "<br>",
    "<b>Quality:</b> ", merged_data$Quality
  )

  disagreement_label <- ifelse(merged_data$disagreement,
                               "Disagreement (>2-fold)",
                               "Agreement (\u22642-fold)")
  legend_label <- paste0(merged_data$Pseudotype, " - ", disagreement_label)

  x_range <- seq(log10(min(merged_data$NT50_numeric)),
                 log10(max(merged_data$NT50_numeric)),
                 length.out = 100)
  y_fitted <- predict(lm_fit, newdata = data.frame(log10_NT50 = x_range))
  regression_data <- data.frame(x = 10^x_range, y = 10^y_fitted)

  p_interactive <- plot_ly() %>%
    add_trace(
      x = merged_data$NT50_numeric,
      y = merged_data$IC50_Titre,
      color = legend_label,
      text = hover_text, hoverinfo = "text",
      type = "scatter", mode = "markers",
      marker = list(size = 8, opacity = 0.7)
    ) %>%
    add_segments(
      x = min(merged_data$NT50_numeric),
      y = min(merged_data$NT50_numeric),
      xend = max(merged_data$NT50_numeric),
      yend = max(merged_data$NT50_numeric),
      line = list(color = "gray", dash = "dash", width = 2),
      showlegend = FALSE, hoverinfo = "skip",
      name = "Perfect Agreement", inherit = FALSE
    ) %>%
    add_lines(
      data = regression_data, x = ~x, y = ~y,
      line = list(color = "#28a745", width = 2),
      showlegend = FALSE, hoverinfo = "skip",
      name = "Linear Regression", inherit = FALSE
    ) %>%
    layout(
      title = list(
        text = sprintf(
          "NT50 vs IC50 Titre Comparison<br><sub>n = %d | r = %.3f | R\u00b2 = %.3f</sub>",
          nrow(merged_data), correlation, r_squared
        ),
        x = 0.5, xanchor = "center"
      ),
      xaxis = list(title = "NT50 (Linear Interpolation)", type = "log",
                   showgrid = FALSE, showline = TRUE, linecolor = "black",
                   linewidth = 1, ticks = "outside", exponentformat = "power"),
      yaxis = list(title = "NT50 / IC50 (Curve Fitting)", type = "log",
                   showgrid = FALSE, showline = TRUE, linecolor = "black",
                   linewidth = 1, ticks = "outside", exponentformat = "power"),
      legend = list(title = list(text = ""), orientation = "h",
                    y = -0.15, x = 0.5, xanchor = "center"),
      hovermode = "closest",
      plot_bgcolor = "white", paper_bgcolor = "white"
    ) %>%
    config(displayModeBar = TRUE, displaylogo = FALSE)

  plotly_json <- plotly::plotly_json(p_interactive, jsonedit = FALSE)

  html_template <- sprintf('
<!DOCTYPE html>
<html>
<head>
    <meta charset="utf-8" />
    <script src="https://cdn.plot.ly/plotly-2.27.0.min.js" charset="utf-8"></script>
</head>
<body>
    <div id="plot" style="width:100%%%%;height:600px;"></div>
    <script>
        var plotData = %s;
        Plotly.newPlot("plot", plotData.data, plotData.layout, plotData.config);
    </script>
</body>
</html>
', plotly_json)

  writeLines(html_template, interactive_plot_file)

  # ── Write Outputs ──────────────────────────────────────────────────────────

  merged_export <- merged_data %>%
    select(
      Pseudotype,
      Sample_ID,
      `NT50 (Linear Interpolation)` = NT50_numeric,
      `NT50 / IC50 (Curve Fitting)` = IC50_Titre,
      `Curve Fit Hybrid NT50`,
      `Sigmoid Quality` = Quality,
      log10_NT50,
      log10_IC50,
      log2_fold_difference,
      disagreement
    )

  write_csv(stats,          file.path(output_dir, "comparison_stats.csv"))
  write_csv(merged_export,  file.path(output_dir, "merged_titres.csv"))
  write_csv(top_mismatches, file.path(output_dir, "top_mismatches.csv"))

  cat("Comparison complete!\n")
  cat("Correlation:", correlation, "\n")
  cat("R-squared:", r_squared, "\n")
  cat("Static plot saved to:", plot_file, "\n")
  cat("Interactive plot saved to:", interactive_plot_file, "\n")

} else {
  stop("No matching data found between NT50 and IC50 datasets")
}