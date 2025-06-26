library(readxl)
library(ggplot2)
library(dplyr)
library(tidyr)

####### AUTOMATION #######

# Read input arguments from Python
args <- commandArgs(trailingOnly = TRUE)
excel_file <- args[1]
output_plot <- args[2]
include_timestamp <- tolower(args[3]) == "true"  # Convert properly
q1_colour <- args[4]
q2_colour <- args[5]
q3_colour <- args[6]
q4_colour <- args[7]
plot_title <- args[8]

if (is.na(plot_title) || plot_title == "") {
    plot_title <- tools::file_path_sans_ext(basename(output_plot))
}


print(paste("Include timestamp:", include_timestamp))

# Ensure it’s either TRUE or FALSE, not NA
if (is.na(include_timestamp)) {
    stop("Error: include_timestamp argument is missing or invalid.")
}



if (!file.exists(excel_file)) {
    stop("Error: Excel file not found. Check the file path.")
}

# Get all sheet names
sheets <- excel_sheets(excel_file)
plate_sheets <- sheets[grepl("^Plate\\d+$", sheets)]

# Debugging print
print(paste("Detected Plate Sheets:", paste(plate_sheets, collapse=", ")))

# Extract numeric parts
plate_numbers <- as.numeric(gsub("^Plate\\s*", "", plate_sheets))

# Debugging print
print(paste("Extracted Plate Numbers:", paste(plate_numbers, collapse=", ")))

# Check if extraction worked
if (any(is.na(plate_numbers))) {
    stop("Error: Failed to extract plate numbers correctly.")
}

# Reorder based on numeric values
plate_sheets <- plate_sheets[order(plate_numbers)]

# Debugging print
print(paste("Sorted Plate Sheets:", paste(plate_sheets, collapse=", ")))

if (length(plate_sheets) == 0) {
    stop("Error: No 'Plate' sheets found in the Excel file.")
}

# Function to process each plate
process_plate <- function(sheet_name) {
    # Read A5:A12 for dilutions (column A)
    dilutions <- read_excel(excel_file, sheet = sheet_name, range = "A5:A12", col_names = FALSE)[[1]]

    # Read B5:M12 for assay data
    data <- read_excel(excel_file, sheet = sheet_name, range = "B5:M12", col_names = FALSE)

    
    # Convert all data cells to numeric, coercing any text to NA
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


# Process all plates
all_data <- bind_rows(lapply(plate_sheets, process_plate))
all_data$Plate <- factor(all_data$Plate, levels = plate_sheets, ordered = TRUE)
all_data$Dilution <- factor(all_data$Dilution, levels = unique(all_data$Dilution))





##### PLOTS ##########

# Define custom colors for Q1, Q2, Q3, and Q4
color_map <- c(Q1 = q1_colour, Q2 = q2_colour, Q3 = q3_colour, Q4 = q4_colour)

# Open PNG device, write plot, then close device (no extra files)
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


cat("🔎 Checking file path in R:", excel_file, "\n")
cat("🔎 file.exists(excel_file) returns:", file.exists(excel_file), "\n")
