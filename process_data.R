script_start <- Sys.time()
cat("⏱ Script started at: ", script_start, "\n")

library(readxl)
library(ggplot2)
library(dplyr)
library(tidyr)
library(cowplot)
library(grid)

####### AUTOMATION #######

args <- commandArgs(trailingOnly = TRUE)
excel_file <- args[1]
output_plot <- args[2]
include_timestamp <- tolower(args[3]) == "true"
q1_colour <- args[4]
q2_colour <- args[5]
q3_colour <- args[6]
q4_colour <- args[7]
plot_title <- args[8]
q1_flag <- tolower(args[9]) == "true"
q2_flag <- tolower(args[10]) == "true"
q3_flag <- tolower(args[11]) == "true"
q4_flag <- tolower(args[12]) == "true"

if (is.na(plot_title) || plot_title == "") {
  plot_title <- tools::file_path_sans_ext(basename(output_plot))
}

if (is.na(include_timestamp)) stop("Error: include_timestamp argument missing or invalid.")
if (!file.exists(excel_file)) stop("Error: Excel file not found.")

sheets <- excel_sheets(excel_file)
plate_sheets <- sheets[grepl("^Plate\\d+$", sheets)]
plate_numbers <- as.numeric(gsub("^Plate\\s*", "", plate_sheets))
if (any(is.na(plate_numbers))) stop("Error: Failed to extract plate numbers correctly.")
plate_sheets <- plate_sheets[order(plate_numbers)]
if (length(plate_sheets) == 0) stop("Error: No 'Plate' sheets found.")

####### FUNCTION TO PROCESS A SINGLE PLATE #######
process_plate <- function(sheet_name) {
  # Read the entire sheet once
  sheet_data <- suppressMessages(
    read_excel(excel_file, sheet = sheet_name, col_names = FALSE)
  )

  # Extract dilutions and main data
  dilutions <- as.numeric(sheet_data[5:12, 1][[1]])
  data <- sheet_data[5:12, 2:13] %>% lapply(unlist) %>% as.data.frame()
  data[] <- lapply(data, as.numeric)


  # Extract sample IDs and pseudotypes
  sample_ids <- c(
    Q1 = sheet_data[3, 2],
    Q2 = sheet_data[3, 5],
    Q3 = sheet_data[3, 8],
    Q4 = sheet_data[3, 11]
  )

  pseudotypes <- c(
    Q1 = sheet_data[4, 2],
    Q2 = sheet_data[4, 5],
    Q3 = sheet_data[4, 8],
    Q4 = sheet_data[4, 11]
  )

  # Build full labels
  full_labels <- sapply(names(sample_ids), function(q) {
    sid <- trimws(as.character(sample_ids[q]))
    psd <- trimws(as.character(pseudotypes[q]))
    if (is.na(sid) || sid == "" || is.na(psd) || psd == "") {
      return(q)
    } else {
      return(paste(sid, psd, sep = " - "))
    }
  })

  # Split data by quadrant
  Titration <- list(
    Q1 = data[, 1:3],
    Q2 = data[, 4:6],
    Q3 = data[, 7:9],
    Q4 = data[, 10:12]
  )

  # Build tidy data frame for plotting
  plot_data <- bind_rows(lapply(names(Titration), function(q) {
    data.frame(
      Dilution = dilutions,
      Mean = rowMeans(Titration[[q]], na.rm = TRUE),
      Min = apply(Titration[[q]], 1, min, na.rm = TRUE),
      Max = apply(Titration[[q]], 1, max, na.rm = TRUE),
      Titration = q,
      FullLabel = full_labels[q],
      Plate = sheet_name
    )
  }))

  return(plot_data)
}


####### PROCESS ALL PLATES #######
all_data <- bind_rows(lapply(plate_sheets, process_plate))
all_data$Plate <- factor(all_data$Plate, levels = plate_sheets, ordered = TRUE)
all_data$Dilution <- factor(all_data$Dilution, levels = unique(all_data$Dilution))

quadrant_flags <- c(Q1 = q1_flag, Q2 = q2_flag, Q3 = q3_flag, Q4 = q4_flag)
active_quadrants <- names(which(quadrant_flags))
if (length(active_quadrants) == 0) stop("Error: No quadrants selected for plotting.")

all_data <- all_data[all_data$Titration %in% active_quadrants, ]
if (nrow(all_data) == 0) stop("Error: No data left after applying quadrant filters.")

color_map <- c(Q1 = q1_colour, Q2 = q2_colour, Q3 = q3_colour, Q4 = q4_colour)
color_map <- color_map[active_quadrants]

####### UTILITY TO BUILD PLOT + LEGEND SIDE BY SIDE #######
make_fixed_plot <- function(base_plot, legend_width = 0.35, total_width = 8, height = 4) {
  legend <- cowplot::get_legend(base_plot)
  # Create white background containers
  legend_bg <- ggdraw() + theme(plot.background = element_rect(fill = "white", color = NA)) +
    draw_plot(legend, 0, 0, 1, 1)
  plot_bg <- ggdraw() + theme(plot.background = element_rect(fill = "white", color = NA)) +
    draw_plot(base_plot + theme(legend.position = "none"), 0, 0, 1, 1)
  cowplot::plot_grid(plot_bg, legend_bg, ncol = 2, rel_widths = c(1, legend_width))
}

####### SUMMARY PLOT (Q1–Q4) #######
summary_base <- ggplot(all_data, aes(x = Dilution, y = Mean, color = Titration, group = Titration)) +
  geom_line() +
  geom_point() +
  geom_errorbar(aes(ymin = Min, ymax = Max), width = 0.2) +
  scale_x_discrete(labels = function(x) ifelse(x == "0", "NSC", x)) +
  scale_y_log10(breaks = scales::log_breaks(n = 6)) +
  scale_color_manual(values = color_map) +
  labs(title = plot_title, x = "Sample Dilution", y = "Mean luminescence (cps)") +
  facet_wrap(~ Plate) +
  theme_minimal(base_size = 12) +
  theme(
    panel.grid.major = element_blank(),
    panel.grid.minor = element_blank(),
    panel.border = element_rect(color = "black", fill = NA),
    plot.background = element_rect(fill = "white"),
    panel.background = element_rect(fill = "white"),
    text = element_text(color = "black"),
    axis.text.y = element_text(color = "black"),
    axis.text.x = element_text(angle = 45, vjust = 1, hjust = 1),
    axis.ticks = element_line(color = "black"),
    strip.text = element_text(color = "black"),
    plot.title = element_text(hjust = 0.5),
    legend.position = "right",
    legend.justification = "top"
  )

summary_combined <- make_fixed_plot(summary_base, legend_width = 0.35, total_width = 12, height = 9)
ggsave(output_plot, summary_combined, width = 12, height = 9, dpi = 96, limitsize = FALSE, bg = "white")

####### PER-PLATE PLOTS #######
for (plate in plate_sheets) {
  plate_data <- all_data[all_data$Plate == plate, ]
  if (nrow(plate_data) == 0) next

  plate_data$FullLabel <- sapply(plate_data$Titration, function(q) {
    val <- unique(plate_data$FullLabel[plate_data$Titration == q])
    val <- val[!is.na(val) & val != ""]
    if (length(val) == 0) return(q)
    return(val[1])
  })

  legend_labels <- setNames(unique(plate_data$FullLabel), unique(plate_data$Titration))

  plate_base <- ggplot(plate_data, aes(x = Dilution, y = Mean, color = Titration, group = Titration)) +
    geom_line() +
    geom_point() +
    geom_errorbar(aes(ymin = Min, ymax = Max), width = 0.2) +
    scale_x_discrete(labels = function(x) ifelse(x == "0", "NSC", x)) +
    scale_y_log10(breaks = scales::log_breaks(n = 6)) +
    scale_color_manual(values = color_map, labels = legend_labels, name = "Titration") +
    labs(title = plate, x = "Sample Dilution", y = "Mean luminescence (cps)") +
    theme_minimal(base_size = 12) +
    theme(
      panel.grid.major = element_blank(),
      panel.grid.minor = element_blank(),
      panel.border = element_rect(color = "black", fill = NA),
      plot.background = element_rect(fill = "white"),
      panel.background = element_rect(fill = "white"),
      text = element_text(color = "black"),
      axis.text.y = element_text(color = "black"),
      axis.text.x = element_text(angle = 45, vjust = 1, hjust = 1),
      axis.ticks = element_line(color = "black"),
      strip.text = element_text(color = "black"),
      plot.title = element_text(hjust = 0.5),
      legend.position = "right",
      legend.justification = "top"
    )

  # Combine with fixed panel and separate white legend area
  plate_combined <- make_fixed_plot(plate_base, legend_width = 0.4, total_width = 8, height = 4)

  plate_filename <- file.path(dirname(output_plot), paste0(plate, ".png"))
  ggsave(plate_filename, plate_combined, width = 8.5, height = 4, dpi = 150, limitsize = FALSE, bg = "white")
}

cat("🔎 Checked file path in R:", excel_file, "\n")
cat("🔎 file.exists(excel_file):", file.exists(excel_file), "\n")


script_end <- Sys.time()
elapsed <- difftime(script_end, script_start, units = "secs")
cat("⏱ Script finished at: ", script_end, "\n")
cat("⏱ Total runtime: ", round(elapsed, 2), " seconds\n")
