# plot_sigmoids.R
# On-demand combined sigmoid plot from stored CSVs.
# Called by Flask /generate_sigmoid_graph/<fitting_id>
#
# Args:
#   1: path to sigmoidData.csv   (raw data points)
#   2: path to IC50s*.csv        (fitted parameters + quality)
#   3: output PNG path
#   4: show_good    ("true"/"false")
#   5: show_unstable ("true"/"false")

suppressPackageStartupMessages({
  library(ggplot2)
  library(dplyr)
  library(readr)
})

args <- commandArgs(trailingOnly = TRUE)
raw_csv      <- args[1]
ic50_csv     <- args[2]
output_png   <- args[3]
show_good    <- tolower(args[4]) == "true"
show_unstable <- tolower(args[5]) == "true"

sigmoid <- function(x, Lower, Upper, Slope, IC50) {
  Lower + (Upper - Lower) / (1 + exp(-Slope * (x - IC50)))
}

raw_data  <- read_csv(raw_csv,  show_col_types = FALSE)
ic50_data <- read_csv(ic50_csv, show_col_types = FALSE)

# Build quality filter
allowed_quality <- c()
if (show_good)     allowed_quality <- c(allowed_quality, "Good")
if (show_unstable) allowed_quality <- c(allowed_quality, "Unstable")
if (length(allowed_quality) == 0) stop("No quality selected")

ic50_filtered <- ic50_data %>% filter(Quality %in% allowed_quality)
if (nrow(ic50_filtered) == 0) stop("No samples match the selected quality filters")

# Data points: join raw data with quality info
plot_points <- raw_data %>%
  inner_join(ic50_filtered %>% select(Sample, Virus, Quality), by = c("Sample", "Virus")) %>%
  mutate(Facet_Label = paste0(Virus, "\n", Sample))

# Fitted curves: reconstruct from parameters for samples with valid IC50
x_range <- seq(min(raw_data$DilutionLog2) - 0.5, max(raw_data$DilutionLog2) + 0.5, length.out = 120)

curves <- bind_rows(lapply(seq_len(nrow(ic50_filtered)), function(i) {
  row <- ic50_filtered[i, ]
  if (is.na(row$IC50)) return(NULL)
  data.frame(
    Sample        = row$Sample,
    Virus         = row$Virus,
    Quality       = row$Quality,
    DilutionLog2  = x_range,
    Neutralisation = sigmoid(x_range, row$Lower, row$Upper, row$Slope, row$IC50),
    Facet_Label   = paste0(row$Virus, "\n", row$Sample)
  )
}))

if (!is.null(curves) && nrow(curves) > 0) {
  plot_points$Facet_Label <- factor(plot_points$Facet_Label, levels = unique(plot_points$Facet_Label))
  curves$Facet_Label      <- factor(curves$Facet_Label,      levels = levels(plot_points$Facet_Label))
}

# Facet dimensions — 4 columns, dynamic height
n_facets   <- length(unique(plot_points$Facet_Label))
ncol_facets <- min(4, n_facets)
nrow_facets <- ceiling(n_facets / ncol_facets)
plot_w <- max(min(ncol_facets * 3.5 + 1.5, 48), 10)
plot_h <- max(min(nrow_facets * 2.8 + 1.2, 48), 6)

p <- ggplot() +
  geom_point(data  = plot_points,
             aes(x = DilutionLog2, y = Neutralisation, color = Quality),
             size = 1.8, alpha = 0.85)

if (!is.null(curves) && nrow(curves) > 0) {
  p <- p + geom_line(data  = curves,
                     aes(x = DilutionLog2, y = Neutralisation,
                         color = Quality, group = Facet_Label),
                     linewidth = 0.8)
}

p <- p +
  facet_wrap(~Facet_Label, ncol = ncol_facets) +
  scale_x_continuous(name = "Dilution (log\u2082)") +
  scale_y_continuous(name = "Neutralisation (%)", limits = c(NA, 105)) +
  scale_color_manual(name   = "Fit quality",
                     values = c("Good" = "#3498db", "Unstable" = "#e74c3c")) +
  theme_bw(base_size = 10) +
  theme(
    panel.grid        = element_blank(),
    strip.background  = element_rect(fill = "#f5f5f5", color = "#cccccc"),
    strip.text        = element_text(size = 8, lineheight = 0.95, face = "bold",
                                     family = "sans"),
    axis.text         = element_text(size = 8,  family = "sans", color = "black"),
    axis.title        = element_text(size = 10, family = "sans", color = "black"),
    legend.position   = "bottom",
    legend.text       = element_text(family = "sans"),
    legend.title      = element_text(family = "sans"),
    plot.background   = element_rect(fill = "white", color = NA),
    panel.background  = element_rect(fill = "white"),
    text              = element_text(family = "sans"),
    plot.margin       = margin(0.3, 0.5, 0.3, 0.3, "cm")
  )

ggsave(output_png, p, width = plot_w, height = plot_h,
       dpi = 150, limitsize = FALSE, bg = "white")
