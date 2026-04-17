# ==============================================================================
# ===== Fitting Sigmoids in R ==================================================
# ==============================================================================

# ------------------------------------------------------------------------------
# ----- 0. Initialisation ------------------------------------------------------
# ------------------------------------------------------------------------------

# ----- 0.1. Description -------------------------------------------------------

# This script fits sigmoids to virus neutralisation assay data collected with a
# 2-fold dilution series.

# The process uses a non-linear least squares approach, provided by the nlsLM
# function of the minpack.lm package.

# ----- 0.2. Dependencies ------------------------------------------------------

suppressPackageStartupMessages({
  library(tidyverse)
  library(minpack.lm)
})

# ----- 0.3. Command Line Arguments --------------------------------------------

args <- commandArgs(trailingOnly = TRUE)

if (length(args) < 2) {
  stop("Usage: Rscript fit_sigmoids.R <input_csv> <output_dir> [assay_title] [timestamp]")
}

input_csv <- args[1]
output_dir <- args[2]

# Optional: assay title, timestamp, R² threshold, LOD censor mode
assay_title <- if (length(args) >= 3 && args[3] != "") args[3] else NULL
timestamp <- if (length(args) >= 4 && args[4] != "") args[4] else NULL
r2_threshold <- if (length(args) >= 5 && args[5] != "") as.numeric(args[5]) else 0.5
include_lod <- if (length(args) >= 6) tolower(args[6]) == "true" else FALSE

# Create output directory if it doesn't exist
if (!dir.exists(output_dir)) {
  dir.create(output_dir, recursive = TRUE)
}


# ----- 0.4. Load Data ---------------------------------------------------------

data <- read_csv(input_csv, show_col_types = FALSE)

# Rename columns to match expected format
data <- data %>%
  rename(
    Sample = Sample,
    Virus = Virus,
    DilutionLog2 = DilutionLog2,
    Neutralisation = Neutralisation
  )

# ------------------------------------------------------------------------------
# ----- 1. Fitting Sigmoids ----------------------------------------------------
# ------------------------------------------------------------------------------

# ----- 1.2. Define Sigmoid ----------------------------------------------------

sigmoid <- function(x, Lower, Upper, Slope, IC50) {
  Lower + (Upper - Lower) / (1 + exp(-Slope * (x - IC50)))
}

# ----- 1.3. Set starting values and bounds ------------------------------------

start_values <- list(Lower = 0,   Upper = 100, Slope = 1,   IC50 = -8)
Lower_bounds <-    c(Lower = -50, Upper = 0,   Slope = 0.001, IC50 = -20)
Upper_bounds <-    c(Lower = 80,  Upper = 100, Slope = 10,    IC50 = 0)

# Pre-calculate dilution range for LOD boundary checks
dil_min <- min(data$DilutionLog2, na.rm = TRUE)
dil_max <- max(data$DilutionLog2, na.rm = TRUE)

# ----- 1.4. Fit Sigmoids ------------------------------------------------------

# Each unique Plate+Quadrant+Sample+Virus combination is fitted independently.
# This prevents quadrants with duplicate sample names from being merged.
combos <- unique(data[, c("Plate", "Quadrant", "Sample", "Virus")])

results <- data.frame()

for (i in seq_len(nrow(combos))) {
  p <- combos$Plate[i]
  q <- combos$Quadrant[i]
  s <- combos$Sample[i]
  v <- combos$Virus[i]

  # Collect only that plate x quadrant x sample x virus data
  current_data <- filter(data, Plate == p, Quadrant == q, Sample == s, Virus == v)

  # Skip if no data for this combination
  if (nrow(current_data) == 0) {
    next
  }

  # Try to fit a sigmoid statistically
  fit <- try(nlsLM(Neutralisation ~ sigmoid(DilutionLog2, Lower, Upper, Slope, IC50),
                   data    = current_data,
                   start   = start_values,
                   lower   = Lower_bounds,
                   upper   = Upper_bounds,
                   control = nls.lm.control(maxiter = 200)),
             silent = TRUE)

  # If the process fails, add a row with NAs to results table
  if (inherits(fit, "try-error")) {

    results <- rbind(results,
      data.frame(Plate = p, Quadrant = q, Sample = s, Virus = v,
                 Lower = NA, Upper = NA, Slope = NA, IC50 = NA, R2 = NA))

  # Otherwise, add parameter estimates to results table
  } else {

    # Calculate pseudo R squared
    RSS <- sum((current_data$Neutralisation - predict(fit))^2)
    TSS <- sum((current_data$Neutralisation - mean(current_data$Neutralisation))^2)
    R2 <- pmax(0, 1 - RSS / TSS)

    # Get parameter estimates
    coefs <- coef(fit)

    # Append to results
    results <- rbind(results,
                     data.frame(Plate = p, Quadrant = q, Sample = s, Virus = v,
                                Lower = coefs["Lower"], Upper = coefs["Upper"],
                                Slope = coefs["Slope"], IC50 = coefs["IC50"], R2 = R2))
  }
}

# ----- 1.5. LOD Detection and Quality Flagging --------------------------------

# Outside LOD is determined by whether the fitted IC50 falls outside the
# measured dilution range — directly analogous to the linear interpolation
# approach where a titre outside the range is clamped to the boundary.
#
#   IC50 < dil_min → serum is too potent to bracket → >Upper LOD
#   IC50 > dil_max → serum is too weak to bracket   → <Lower LOD
#
# Samples where the fit failed entirely (no IC50) or where the IC50 is within
# range but R² is poor are classified as "Poor Fit" — we cannot determine
# whether they are outside LOD, so they are not assigned a boundary value.

results$LOD_Flag <- NA_character_

for (i in seq_len(nrow(results))) {
  ic50 <- results$IC50[i]

  if (is.na(ic50)) {
    # Fit failed entirely — cannot determine LOD
    next
  }

  if (ic50 < dil_min) {
    # IC50 is beyond the most-diluted well → above upper LOD
    results$LOD_Flag[i] <- ">Upper LOD"
    results$Lower[i]    <- NA
    results$Upper[i]    <- NA
    results$Slope[i]    <- NA
    results$IC50[i]     <- NA
    results$R2[i]       <- NA
  } else if (ic50 > dil_max) {
    # IC50 is before the least-diluted well → below lower LOD
    results$LOD_Flag[i] <- "<Lower LOD"
    results$Lower[i]    <- NA
    results$Upper[i]    <- NA
    results$Slope[i]    <- NA
    results$IC50[i]     <- NA
    results$R2[i]       <- NA
  } else if (results$R2[i] < r2_threshold) {
    # IC50 within range but poor fit quality — not LOD, just unreliable
    results$Lower[i] <- NA
    results$Upper[i] <- NA
    results$Slope[i] <- NA
    results$IC50[i]  <- NA
    results$R2[i]    <- NA
  }
  # Otherwise: IC50 within range and R² acceptable — leave params as-is
}


# ----- 1.6. LOD Censoring (optional) -----------------------------------------
# When include_lod = TRUE, samples flagged as outside LOD are assigned the
# boundary dilution value as their IC50/Titre rather than being left as NA.
# <Lower LOD → boundary = A5 dilution (max DilutionLog2, least negative)
# >Upper LOD → boundary = A11 dilution (min DilutionLog2, most negative)
# Lower/Upper/Slope remain NA to show no real sigmoid was fitted.

if (include_lod) {
  lod_lower_ic50 <- max(data$DilutionLog2, na.rm = TRUE)
  lod_upper_ic50 <- min(data$DilutionLog2, na.rm = TRUE)

  for (i in seq_len(nrow(results))) {
    if (!is.na(results$LOD_Flag[i])) {
      if (results$LOD_Flag[i] == "<Lower LOD") {
        results$IC50[i] <- round(lod_lower_ic50, 4)
      } else if (results$LOD_Flag[i] == ">Upper LOD") {
        results$IC50[i] <- round(lod_upper_ic50, 4)
      }
    }
  }
}

# ----- 1.8. Quality Call ------------------------------------------------------

# Sigmoids with a parameter estimate pressed against the lower or upper bounds
# should be treated with suspicion. These samples are often either erroneous or
# have multiple possible solutions to their sigmoid.

results$Quality <- NA

for (i in seq_len(nrow(results))) {

  # Outside LOD: IC50 was outside the dilution range
  if (!is.na(results$LOD_Flag[i])) {
    results$Quality[i] <- "Outside LOD"
    next
  }

  # Poor Fit: fit failed entirely or IC50 within range but R² too low
  # (these have NA IC50 but no LOD_Flag)
  if (is.na(results$IC50[i])) {
    results$Quality[i] <- "Poor Fit"
    next
  }

  checking <- results[i,]

  failed <- any(c(checking$Lower == Lower_bounds["Lower"],
                  checking$Lower == Upper_bounds["Lower"],
                  checking$IC50 == Lower_bounds["IC50"],
                  checking$IC50 == Upper_bounds["IC50"],
                  checking$Slope == Lower_bounds["Slope"],
                  checking$Slope == Upper_bounds["Slope"]))

  if (failed) {
    results$Quality[i] <- "Unstable"
  } else {
    results$Quality[i] <- "Good"
  }
}
    
# ----- 1.10. Write Output -----------------------------------------------------

# Round all numbers in numeric columns to 4dp (only for non-NA values)
numeric_cols <- c("Lower", "Upper", "Slope", "IC50", "R2")
for (col in numeric_cols) {
  results[[col]] <- round(results[[col]], 4)
}

# Calculate titres from IC50 values (NA for LOD and Poor Fit samples)
results$Titre <- ifelse(is.na(results$IC50), NA, round(2^(-results$IC50), 2))

# Reorder columns: Plate, Quadrant, Sample, Virus, Lower, Upper, Slope, IC50, Titre, R2, Quality, LOD_Flag
results <- results %>%
  select(Plate, Quadrant, Sample, Virus, Lower, Upper, Slope, IC50, Titre, R2, Quality, LOD_Flag)

# Build custom filename based on assay title and timestamp
if (!is.null(assay_title) && assay_title != "" && !is.null(timestamp) && timestamp != "") {
  ic50_filename <- paste0("IC50s_", assay_title, "_", timestamp, ".csv")
} else if (!is.null(assay_title) && assay_title != "") {
  ic50_filename <- paste0("IC50s_", assay_title, ".csv")
} else if (!is.null(timestamp) && timestamp != "") {
  ic50_filename <- paste0("IC50s_", timestamp, ".csv")
} else {
  ic50_filename <- "IC50s.csv"
}

ic50_output_path <- file.path(output_dir, ic50_filename)
write_csv(results, ic50_output_path)

cat("\nSigmoid fitting complete!\n")
cat("Results saved to:", ic50_output_path, "\n")
cat("Plots saved to:", output_dir, "\n")

# Summary
n_good     <- sum(results$Quality == "Good",        na.rm = TRUE)
n_unstable <- sum(results$Quality == "Unstable",    na.rm = TRUE)
n_lod      <- sum(results$Quality == "Outside LOD", na.rm = TRUE)
n_poor     <- sum(results$Quality == "Poor Fit",    na.rm = TRUE)
cat(sprintf("  Good: %d | Unstable: %d | Outside LOD: %d | Poor Fit: %d\n",
            n_good, n_unstable, n_lod, n_poor))