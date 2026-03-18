# ==============================================================================
# ===== Boxplot NT — Linear Interpolation from Plate Sheets ===================
# ==============================================================================
#
# Usage: Rscript boxplot_nt50.R <excel_file> <output_json> [threshold_pct]
#
# threshold_pct: 50 (default) for NT50, or 90 for NT90.
#
# Reads every Plate sheet, computes NT by linear interpolation for each
# replicate column, averages the 3 replicates per quadrant, and writes a
# JSON file of { pseudotype: [averaged_nt, ...], ... } for the boxplot.
# ==============================================================================

suppressPackageStartupMessages({
  library(readxl)
  library(jsonlite)
})

# ----- Command Line Arguments -------------------------------------------------

args <- commandArgs(trailingOnly = TRUE)

if (length(args) < 2) {
  stop("Usage: Rscript boxplot_nt50.R <excel_file> <output_json> [threshold_pct]")
}

excel_file  <- args[1]
output_json <- args[2]
threshold_pct <- if (length(args) >= 3) as.numeric(args[3]) else 50

# Validate threshold
if (is.na(threshold_pct) || !(threshold_pct %in% c(50, 90))) {
  threshold_pct <- 50
}

# The fraction of NSC that defines the target luminescence
# NT50 → neutralises 50% → luminescence drops to 50% of NSC → target = 0.5
# NT90 → neutralises 90% → luminescence drops to 10% of NSC → target = 0.1
target_fraction <- (100 - threshold_pct) / 100

titre_label <- paste0("NT", threshold_pct)

# ==============================================================================
# ===== Extract NT by Linear Interpolation =====================================
# ==============================================================================

sheets <- excel_sheets(excel_file)
plate_sheets <- sheets[grepl("^Plate\\d+$", sheets)]
plate_numbers <- as.numeric(gsub("^Plate", "", plate_sheets))
plate_sheets <- plate_sheets[order(plate_numbers)]

if (length(plate_sheets) == 0) {
  writeLines(toJSON(list(status = "success", data = list(), titre_label = titre_label), auto_unbox = TRUE), output_json)
  quit(save = "no", status = 0)
}

# Quadrant definitions
quad_defs <- list(
  list(pt_cell = "B3", sid_cell = "B4", data_cols = 2:4),
  list(pt_cell = "E3", sid_cell = "E4", data_cols = 5:7),
  list(pt_cell = "H3", sid_cell = "H4", data_cols = 8:10),
  list(pt_cell = "K3", sid_cell = "K4", data_cols = 11:13)
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

# Collect averaged NT values grouped by pseudotype
grouped <- list()

for (sheet_name in plate_sheets) {

  dilutions <- tryCatch(
    suppressMessages(read_excel(excel_file, sheet = sheet_name, range = "A5:A12", col_names = FALSE)[[1]]),
    error = function(e) rep(NA_real_, 8)
  )
  dilutions <- as.numeric(dilutions)

  raw_block <- tryCatch(
    suppressMessages(read_excel(excel_file, sheet = sheet_name, range = "B5:M12", col_names = FALSE)),
    error = function(e) NULL
  )
  if (is.null(raw_block)) next

  data_matrix <- as.data.frame(lapply(raw_block, as.numeric))

  for (quad in quad_defs) {

    pt_val <- read_cell(sheet_name, quad$pt_cell)
    if (is.na(pt_val)) next

    pseudotype_name <- pt_val
    col_indices <- quad$data_cols - 1

    nt_replicates <- c()

    for (ci in col_indices) {
      lum <- data_matrix[[ci]]

      if (length(lum) < 8 || is.na(lum[8])) next

      nsc_value <- lum[8]
      if (is.na(nsc_value) || nsc_value <= 0) next

      target <- nsc_value * target_fraction
      nt_val <- NA_real_

      for (i in 1:7) {
        y1 <- lum[i]
        y2 <- lum[i + 1]
        x1 <- dilutions[i]
        x2 <- dilutions[i + 1]

        if (is.na(y1) || is.na(y2) || is.na(x1) || is.na(x2)) next

        crosses <- (y1 >= target & y2 <= target) | (y1 <= target & y2 >= target)
        if (crosses && y2 != y1) {
          nt_val <- (target - y1) / (y2 - y1) * (x2 - x1) + x1
          break
        }
      }

      if (is.na(nt_val)) {
        valid_lum <- lum[1:7]
        valid_lum <- valid_lum[!is.na(valid_lum)]
        if (length(valid_lum) > 0 && all(valid_lum < target)) {
          nt_val <- dilutions[1]
        } else if (length(valid_lum) > 0 && all(valid_lum > target)) {
          nt_val <- dilutions[7]
        }
      }

      if (!is.na(nt_val)) {
        nt_replicates <- c(nt_replicates, nt_val)
      }
    }

    if (length(nt_replicates) > 0) {
      avg_nt <- round(mean(nt_replicates), 1)

      if (is.null(grouped[[pseudotype_name]])) {
        grouped[[pseudotype_name]] <- c()
      }
      grouped[[pseudotype_name]] <- c(grouped[[pseudotype_name]], avg_nt)
    }
  }
}

result <- list(status = "success", data = grouped, titre_label = titre_label)
writeLines(toJSON(result, auto_unbox = TRUE), output_json)

cat(sprintf("Boxplot data (%s): %d pseudotypes, %d total values\n",
            titre_label,
            length(grouped),
            sum(sapply(grouped, length))))