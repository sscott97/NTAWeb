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
  stop("Usage: Rscript fit_sigmoids.R <input_csv> <output_dir>")
}

input_csv <- args[1]
output_dir <- args[2]

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
Lower_bounds <-    c(Lower = -50, Upper = 0,  Slope = 0.001, IC50 = -20)
Upper_bounds <-    c(Lower = 50,  Upper = 100, Slope  = 10, IC50 = 0)

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

# ----- 1.5. Fix Flat Titrations -----------------------------------------------

# Parameter estimates are unstable when data are flat (i.e., low Slope).

# These samples either produce a low R2 or NA estimates, indicating they are better
# described with a straight horizontal line through their mean neutralisation.

# This is applied here:

for (i in c(1:nrow(results))){
  
  if (is.na(results$R2[i]) || results$R2[i] < 0.5){
    results$Upper[i] <- results$u_Neutralisation[i]
    results$Lower[i] <- results$u_Neutralisation[i]
    results$IC50[i]  <- 0
    results$Slope[i] <- 1e-8
  }
  
}

# ----- 1.6. Use Parameter Estimates to Generate Predicted Curves --------------

x_values <- seq(min(data$DilutionLog2) - 1, max(data$DilutionLog2) + 1, length.out = 100)

sigmoids <- data.frame()

for (i in c(1:nrow(results))) {
  
  y_values <- sigmoid(x_values, Lower = results$Lower[i], Upper = results$Upper[i],
                      Slope = results$Slope[i], IC50  = results$IC50[i])
  
  sigmoids <- rbind(sigmoids,
                    data.frame(Sample = results$Sample[i],
                               Virus  = results$Virus[i],
                               DilutionLog2 = x_values,
                               Neutralisation = y_values))
}

# ----- 1.7. Plot Fitted Sigmoid Curves ----------------------------------------

# Create combined label with both Virus and Sample for faceting
data$Facet_Label <- paste0("Virus: ", data$Virus, "\nSample: ", data$Sample)
sigmoids$Facet_Label <- paste0("Virus: ", sigmoids$Virus, "\nSample: ", sigmoids$Sample)

# Keep original Sample column for joining with results
results$Sample <- results$Sample

# ----- 1.8. Quality Call ------------------------------------------------------

# Sigmoids with a parameter estimate pressed against the lower or upper bounds
# should be treated with suspicion. These samples are often either erroneous or
# have multiple possible solutions to their sigmoid.

results$Quality <- NA

for (i in c(1:nrow(results))){
  
  checking <- results[i,]
  
  failed <- any(c(checking$Lower == Lower_bounds["Lower"],
                  checking$Lower == Upper_bounds["Lower"],
                  checking$IC50 == Lower_bounds["IC50"],
                  checking$IC50 == Upper_bounds["IC50"],
                  checking$Slope == Lower_bounds["Slope"],
                  checking$Slope == Upper_bounds["Slope"]))
  
  if (failed){
    results$Quality[i] <- "Unstable"
  } else {
    results$Quality[i] <- "Good"
  }

}
    
# ----- 1.9. Quality Plot ------------------------------------------------------

# Join quality data using original Sample column (without prefix)
sigmoids <- left_join(sigmoids, results, by = c("Sample", "Virus"))
data <- left_join(data, results, by = c("Sample", "Virus"))

# Create plot for each virus
unique_viruses <- unique(data$Virus)

for (virus_name in unique_viruses) {
  
  virus_data <- data %>% filter(Virus == virus_name)
  virus_sigmoids <- sigmoids %>% filter(Virus == virus_name)
  
  # Count unique samples for this virus to calculate optimal dimensions
  n_samples <- length(unique(virus_data$Sample))
  
  # Calculate optimal plot dimensions
  # Use 4 columns, calculate rows needed
  ncol_facets <- 4
  nrow_facets <- ceiling(n_samples / ncol_facets)
  
  # Base dimensions per facet
  facet_width <- 3.5   # inches per column
  facet_height <- 2.8  # inches per row (slightly taller for 2-line labels)
  
  # Calculate total dimensions with some padding
  plot_width <- ncol_facets * facet_width + 1.5  # extra for margins/legend
  plot_height <- nrow_facets * facet_height + 0.5  # reduced top padding (no title)
  
  # Set minimum dimensions
  plot_width <- max(plot_width, 12)
  plot_height <- max(plot_height, 8)
  
  # Create safe filename
  safe_virus_name <- gsub("[^A-Za-z0-9_-]", "_", virus_name)
  
  p <- ggplot() +
    geom_line(data = virus_sigmoids,
              aes(x = DilutionLog2,
                  y = Neutralisation,
                  color = Quality,
                  group = Facet_Label)) +
    geom_point(data = virus_data,
               aes(x = DilutionLog2,
                   y = Neutralisation,
                   color = Quality),
               size = 2) +  # Slightly larger points for visibility
    facet_wrap(~Facet_Label, ncol = ncol_facets) +
    scale_x_continuous(name = "Dilution (log2)") +
    scale_y_continuous(name = "Neutralisation (%)") +
    scale_color_manual(values = c("Good" = "#3498db", "Unstable" = "#e74c3c")) +
    theme_bw() +
    theme(
      panel.grid = element_blank(),
      strip.text = element_text(size = 8.5, lineheight = 0.9),  # Facet labels (2 lines)
      axis.text = element_text(size = 8),   # Axis tick labels
      axis.title = element_text(size = 10), # Axis titles
      legend.position = "bottom",
      plot.margin = margin(0.2, 0.5, 0.2, 0.2, "cm")  # Reduced top margin
    )
  
  # Save plot with dynamic dimensions
  ggsave(
    filename = file.path(output_dir, paste0("sigmoid_fit_", safe_virus_name, ".png")),
    plot = p,
    width = plot_width,
    height = plot_height,
    dpi = 300
  )
  
  cat(paste0("  Saved plot for ", virus_name, " (", n_samples, " samples, ", 
             plot_width, "x", plot_height, " inches)\n"))
}

# ----- 1.10. Write Output -----------------------------------------------------

# Round all numbers in columns 3-8 to 4dp
results[,c(3:8)] <- round(results[,c(3:8)], 4)

# Mean neutralisation across dilutions now irrelevant
results$u_Neutralisation <- NULL

write_csv(results, file.path(output_dir, "IC50s.csv"))

cat("Sigmoid fitting complete!\n")
cat("Results saved to:", file.path(output_dir, "IC50s.csv"), "\n")
cat("Plots saved to:", output_dir, "\n")