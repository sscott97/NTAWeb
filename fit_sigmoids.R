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

# Optional: assay title and timestamp
assay_title <- if (length(args) >= 3 && args[3] != "") args[3] else NULL
timestamp <- if (length(args) >= 4 && args[4] != "") args[4] else NULL

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

# ----- 1.4. Fit Sigmoids ------------------------------------------------------

samples <- unique(data$Sample)
viruses <- unique(data$Virus)

results <- data.frame()

# For each sample and virus
for (s in samples) {
  for (v in viruses) {
    
    # Collect only that sample x virus data
    current_data <- filter(data, Sample == s, Virus == v)
    
    # Skip if no data for this combination
    if (nrow(current_data) == 0) {
      next
    }
    
    # Get mean neutralisation (useful for failed curves)
    u <- mean(current_data$Neutralisation)
    
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
        data.frame(Sample = s, Virus = v, u_Neutralisation = u, Lower = NA,
                   Upper = NA, Slope = NA, IC50 = NA, R2 = NA))
      
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
                       data.frame(Sample = s, Virus = v, u_Neutralisation = u, Lower = coefs["Lower"],
                                  Upper = coefs["Upper"], Slope = coefs["Slope"], IC50 = coefs["IC50"],
                                  R2 = R2))
    }
  }
}

# ----- 1.5. Handle Flat Titrations (Outside Limit of Detection) ---------------

# Samples that failed to converge or yielded R² < 0.5 cannot support a
# meaningful sigmoid fit. Rather than forcing IC50 = 0 (which produces a
# misleading titre of 1), these are flagged as outside the limit of detection. 
# REMOVED BY SAM, CHANGED TO R^2<0.5 = OUTSIDE LIMIT OF DETECTION

# Determine LOD direction based on mean neutralisation:
#   - If mean neutralisation > 50% → sample is highly neutralising but the curve
#     shape couldn't be resolved → titre is above the upper LOD (> highest dilution)
#   - If mean neutralisation ≤ 50% → sample shows little neutralisation →
#     titre is below the lower LOD (< lowest dilution)

for (i in seq_len(nrow(results))) {
  
  if (is.na(results$R2[i]) || results$R2[i] < 0.5) {
    
    mean_neut <- results$u_Neutralisation[i]
    
    if (mean_neut > 50) {
      # High neutralisation but flat/failed curve → above upper LOD
      results$Lower[i]  <- NA
      results$Upper[i]  <- NA
      results$Slope[i]  <- NA
      results$IC50[i]   <- NA
      results$R2[i]     <- NA
    } else {
      # Low neutralisation → below lower LOD
      results$Lower[i]  <- NA
      results$Upper[i]  <- NA
      results$Slope[i]  <- NA
      results$IC50[i]   <- NA
      results$R2[i]     <- NA
    }
  }
}

# Add LOD flag column based on whether parameters are NA after the above step
# (distinguishing from fits that simply failed to converge, which also have NAs)
results$LOD_Flag <- NA_character_

for (i in seq_len(nrow(results))) {
  if (is.na(results$IC50[i])) {
    mean_neut <- results$u_Neutralisation[i]
    if (mean_neut > 50) {
      results$LOD_Flag[i] <- ">Upper LOD"
    } else {
      results$LOD_Flag[i] <- "<Lower LOD"
    }
  }
}


# ----- 1.8. Quality Call ------------------------------------------------------

# Sigmoids with a parameter estimate pressed against the lower or upper bounds
# should be treated with suspicion. These samples are often either erroneous or
# have multiple possible solutions to their sigmoid.

results$Quality <- NA

for (i in seq_len(nrow(results))) {
  
  # Samples outside LOD get their own quality label
  if (!is.na(results$LOD_Flag[i])) {
    results$Quality[i] <- "Outside LOD"
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
numeric_cols <- c("u_Neutralisation", "Lower", "Upper", "Slope", "IC50", "R2")
for (col in numeric_cols) {
  results[[col]] <- round(results[[col]], 4)
}

# Mean neutralisation across dilutions now irrelevant
results$u_Neutralisation <- NULL

# Calculate titres from IC50 values (NA for LOD samples)
results$Titre <- ifelse(is.na(results$IC50), NA, round(2^(-results$IC50), 2))

# Reorder columns: Sample, Virus, Lower, Upper, Slope, IC50, Titre, R2, Quality, LOD_Flag
results <- results %>%
  select(Sample, Virus, Lower, Upper, Slope, IC50, Titre, R2, Quality, LOD_Flag)

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
n_good <- sum(results$Quality == "Good", na.rm = TRUE)
n_unstable <- sum(results$Quality == "Unstable", na.rm = TRUE)
n_lod <- sum(results$Quality == "Outside LOD", na.rm = TRUE)
cat(sprintf("  Good: %d | Unstable: %d | Outside LOD: %d\n", n_good, n_unstable, n_lod))