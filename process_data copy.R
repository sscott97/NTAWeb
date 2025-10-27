library(readxl)
library(ggplot2)
library(dplyr)
library(tidyr)

####### AUTOMATION #######

# arguments from python script
args <- commandArgs(trailingOnly = TRUE)
excel_file <- args[1]
output_plot <- args[2]
include_timestamp <- tolower(args[3]) == "true"  # Convert properly
q1_colour <- args[4]
q2_colour <- args[5]
q3_colour <- args[6]
q4_colour <- args[7]
plot_title <- args[8]
q1_flag <- tolower(args[9]) == "true"
q2_flag <- tolower(args[10]) == "true"
q3_flag <- tolower(args[11]) == "true"
q4_flag <- tolower(args[12]) == "true"

# uses file name as title if none provided
if (is.na(plot_title) || plot_title == "") {
  plot_title <- tools::file_path_sans_ext(basename(output_plot))
}

print(paste("Include timestamp:", include_timestamp))

# inputs debug
if (is.na(include_timestamp)) stop("Error: include_timestamp argument missing or invalid.")
if (!file.exists(excel_file)) stop("Error: Excel file not found.")

# get plate sheet names and order them numerically
sheets <- excel_sheets(excel_file)
plate_sheets <- sheets[grepl("^Plate\\d+$", sheets)]
plate_numbers <- as.numeric(gsub("^Plate\\s*", "", plate_sheets))
if (any(is.na(plate_numbers))) stop("Error: Failed to extract plate numbers correctly.")
plate_sheets <- plate_sheets[order(plate_numbers)]
if (length(plate_sheets) == 0) stop("Error: No 'Plate' sheets found.")

# function to process each plate
process_plate <- function(sheet_name) {
  dilutions <- read_excel(excel_file, sheet = sheet_name, range = "A5:A12", col_names = FALSE)[[1]]
  data <- read_excel(excel_file, sheet = sheet_name, range = "B5:M12", col_names = FALSE)
  data[] <- lapply(data, function(x) as.numeric(x))
  
  Titration <- list(
    Q1 = data[, 1:3],
    Q2 = data[, 4:6],
    Q3 = data[, 7:9],
    Q4 = data[, 10:12]
  )
  
  plot_data <- bind_rows(lapply(names(Titration), function(q) {
    df <- data.frame(
      Dilution = dilutions,
      Mean = rowMeans(Titration[[q]], na.rm = TRUE),
      SD = apply(Titration[[q]], 1, sd, na.rm = TRUE),
      Titration = q,
      Plate = sheet_name
    )
    return(df)
  }))
  
  return(plot_data)
}

# combine all plates
all_data <- bind_rows(lapply(plate_sheets, process_plate))
all_data$Plate <- factor(all_data$Plate, levels = plate_sheets, ordered = TRUE)
all_data$Dilution <- factor(all_data$Dilution, levels = unique(all_data$Dilution))

# apply quadrants filter (ie only include selected quadrants)
quadrant_flags <- c(Q1 = q1_flag, Q2 = q2_flag, Q3 = q3_flag, Q4 = q4_flag)
active_quadrants <- names(which(quadrant_flags))
if (length(active_quadrants) == 0) stop("Error: No quadrants selected for plotting.")

all_data <- all_data[all_data$Titration %in% active_quadrants, ]
if (nrow(all_data) == 0) stop("Error: No data left after applying quadrant filters.")

# get colours for each quadrant
color_map <- c(Q1 = q1_colour, Q2 = q2_colour, Q3 = q3_colour, Q4 = q4_colour)
color_map <- color_map[active_quadrants]

##### PLOTS ##########

png(filename = output_plot, width = 12 * 96, height = 9 * 96, res = 96)  # 12x9 inches at 96 dpi

print(
  ggplot(all_data, aes(x = Dilution, y = Mean, color = Titration, group = Titration)) +
    geom_line() +
    geom_point() +
    geom_errorbar(aes(ymin = Mean - SD, ymax = Mean + SD), width = 0.2) +
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
      axis.text.y = element_text(angle = 0, hjust = 1, color = "black"),
      axis.text.x = element_text(angle = 45, vjust = 1, hjust = 1),
      axis.ticks.length = unit(0.1, "cm"),
      axis.ticks = element_line(color = "black"),
      axis.ticks.y = element_line(linewidth = 0.2),
      axis.ticks.x = element_line(linewidth = 0.2),
      strip.text = element_text(color = "black"),
      plot.title = element_text(hjust = 0.5)
    )
)

dev.off()

# --- Per-plate graphs ---
for (plate in plate_sheets) {
  plate_data <- all_data[all_data$Plate == plate, ]
  
  if (nrow(plate_data) == 0) next
  
  plate_plot <- ggplot(plate_data, aes(x = Dilution, y = Mean, color = Titration, group = Titration)) +
    geom_line() +
    geom_point() +
    geom_errorbar(aes(ymin = Mean - SD, ymax = Mean + SD), width = 0.2) +
    scale_x_discrete(labels = function(x) ifelse(x == "0", "NSC", x)) +
    scale_y_log10(breaks = scales::log_breaks(n = 6)) +
    scale_color_manual(values = color_map) +
    labs(title = paste(plate), x = "Sample Dilution", y = "Mean luminescence (cps)") +
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
      plot.title = element_text(hjust = 0.5)
    )
  
  plate_filename <- file.path(dirname(output_plot), paste0(plate, ".png"))
  ggsave(plate_filename, plate_plot, width = 6, height = 4, dpi = 150)
}



cat("🔎 Checking file path in R:", excel_file, "\n")
cat("🔎 file.exists(excel_file) returns:", file.exists(excel_file), "\n")
