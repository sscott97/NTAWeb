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
show_good     <- tolower(args[4]) == "true"
show_unstable <- tolower(args[5]) == "true"
show_lod      <- if (length(args) >= 6) tolower(args[6]) == "true" else FALSE
show_poor_fit <- if (length(args) >= 7) tolower(args[7]) == "true" else FALSE

sigmoid <- function(x, Lower, Upper, Slope, IC50) {
  Lower + (Upper - Lower) / (1 + exp(-Slope * (x - IC50)))
}

raw_data  <- read_csv(raw_csv,  show_col_types = FALSE)
ic50_data <- read_csv(ic50_csv, show_col_types = FALSE)

# Build quality filter
allowed_quality <- c()
if (show_good)     allowed_quality <- c(allowed_quality, "Good")
if (show_unstable) allowed_quality <- c(allowed_quality, "Unstable")
if (show_lod)      allowed_quality <- c(allowed_quality, "Outside LOD")
if (show_poor_fit) allowed_quality <- c(allowed_quality, "Poor Fit")
if (length(allowed_quality) == 0) stop("No quality selected")

ic50_filtered <- ic50_data %>% filter(Quality %in% allowed_quality)
if (nrow(ic50_filtered) == 0) stop("No samples match the selected quality filters")

# Build a Color_Group column: LOD samples split by direction, others use Quality
# LOD_Flag values: "<Lower LOD" (green) or ">Upper LOD" (orange)
ic50_filtered <- ic50_filtered %>%
  mutate(Color_Group = case_when(
    Quality == "Outside LOD" & !is.na(LOD_Flag) ~ LOD_Flag,
    TRUE ~ Quality
  ))

# Data points: join raw data with quality + color info
# Join on Plate+Quadrant+Sample+Virus so duplicate sample names are never merged
plot_points <- raw_data %>%
  inner_join(ic50_filtered %>% select(Plate, Quadrant, Sample, Virus, Color_Group),
             by = c("Plate", "Quadrant", "Sample", "Virus")) %>%
  mutate(Facet_Label = paste0(Virus, "\n", Sample, "\n(", Plate, " ", Quadrant, ")"))

# Fitted curves: reconstruct from parameters for samples with valid sigmoid params
x_range <- seq(min(raw_data$DilutionLog2) - 0.5, max(raw_data$DilutionLog2) + 0.5, length.out = 120)

curves <- bind_rows(lapply(seq_len(nrow(ic50_filtered)), function(i) {
  row <- ic50_filtered[i, ]
  # Skip LOD censored samples — no real sigmoid, shown as vline instead
  if (is.na(row$IC50) || is.na(row$Lower) || is.na(row$Upper) || is.na(row$Slope)) return(NULL)
  data.frame(
    Plate          = row$Plate,
    Quadrant       = row$Quadrant,
    Sample         = row$Sample,
    Virus          = row$Virus,
    Color_Group    = row$Color_Group,
    DilutionLog2   = x_range,
    Neutralisation = sigmoid(x_range, row$Lower, row$Upper, row$Slope, row$IC50),
    Facet_Label    = paste0(row$Virus, "\n", row$Sample, "\n(", row$Plate, " ", row$Quadrant, ")")
  )
}))

if (!is.null(curves) && nrow(curves) > 0) {
  plot_points$Facet_Label <- factor(plot_points$Facet_Label, levels = unique(plot_points$Facet_Label))
  curves$Facet_Label      <- factor(curves$Facet_Label,      levels = levels(plot_points$Facet_Label))
}

# Facet dimensions — 4 columns, dynamic height
n_facets    <- length(unique(plot_points$Facet_Label))
ncol_facets <- min(4, n_facets)
nrow_facets <- ceiling(n_facets / ncol_facets)
plot_w <- max(ncol_facets * 4.0 + 1.5, 10)
plot_h <- max(nrow_facets * 3.5 + 1.2, 6)

# LOD boundary vlines — one dashed line per censored sample, coloured by direction
lod_vlines <- ic50_filtered %>%
  filter(Quality == "Outside LOD", !is.na(IC50)) %>%
  mutate(Facet_Label = paste0(Virus, "\n", Sample, "\n(", Plate, " ", Quadrant, ")"))

p <- ggplot() +
  geom_point(data  = plot_points,
             aes(x = DilutionLog2, y = Neutralisation, color = Color_Group),
             size = 1.8, alpha = 0.85)

if (!is.null(curves) && nrow(curves) > 0) {
  p <- p + geom_line(data  = curves,
                     aes(x = DilutionLog2, y = Neutralisation,
                         color = Color_Group, group = Facet_Label),
                     linewidth = 0.8)
}

if (nrow(lod_vlines) > 0) {
  p <- p + geom_vline(data = lod_vlines,
                      aes(xintercept = IC50, color = Color_Group),
                      linetype = "dashed", linewidth = 0.7, alpha = 0.85)
}

lod_color_values <- c(
  "Good"         = "#648fff",
  "Unstable"     = "#ffb000",
  "<Lower LOD"   = "#785ef0",
  ">Upper LOD"   = "#fe6100",
  "Poor Fit"     = "#dc267f"
)
lod_color_labels <- c(
  "Good"         = "Good",
  "Unstable"     = "Unstable",
  "<Lower LOD"   = "<Lower LOD (boundary)",
  ">Upper LOD"   = ">Upper LOD (boundary)",
  "Poor Fit"     = "Poor Fit"
)

p <- p +
  facet_wrap(~Facet_Label, ncol = ncol_facets) +
  scale_x_reverse(name = "Dilution (log\u2082)") +
  scale_y_reverse(name = "Neutralisation (%)", limits = c(105, NA)) +
  scale_color_manual(name   = "Fit quality",
                     values = lod_color_values,
                     labels = lod_color_labels) +
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
