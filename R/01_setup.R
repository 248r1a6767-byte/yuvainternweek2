# ==============================================================================
# Script: 01_setup.R
# Purpose: Initialize project environment, manage packages, set directories,
#          and define global visualization and formatting parameters.
# Author: Data Analyst & Statistical Specialist
# Date: 2026-09-28
# ==============================================================================

# Ensure reproducibility and consistent numerical formatting
options(scipen = 999)      # Disable scientific notation for clean financial output
options(digits = 4)        # Standardize decimal display
set.seed(42)               # Seed for reproducible jittering if used

# 1. Package Management --------------------------------------------------------
required_packages <- c(
  "readr",       # Fast, tidy tabular data import
  "dplyr",       # Data manipulation and aggregation
  "tidyr",       # Tidy data reshaping
  "lubridate",   # Robust date parsing and arithmetic
  "ggplot2",     # Declarative data visualization
  "scales",      # Financial, percentage, and date axis formatting
  "forcats",     # Factor reordering for visual clarity
  "patchwork"    # Multi-panel chart compositions
)

# Function to check and load packages
load_project_packages <- function(pkgs) {
  missing_pkgs <- pkgs[!(pkgs %in% installed.packages()[, "Package"])]
  if (length(missing_pkgs) > 0) {
    message("Installing missing packages: ", paste(missing_pkgs, collapse = ", "))
    install.packages(missing_pkgs, repos = "https://cloud.r-project.org", quiet = TRUE)
  }
  for (pkg in pkgs) {
    suppressPackageStartupMessages(library(pkg, character.only = TRUE))
  }
  message("All required packages successfully loaded.")
}

load_project_packages(required_packages)

# 2. Directory Structure Verification ------------------------------------------
project_dirs <- c(
  file.path("data", "raw"),
  file.path("data", "processed"),
  "visualizations",
  file.path("outputs", "tables"),
  file.path("outputs", "summaries"),
  "report"
)

for (d in project_dirs) {
  if (!dir.exists(d)) {
    dir.create(d, recursive = TRUE, showWarnings = FALSE)
    message("Created directory: ", d)
  }
}

# 3. Global Visualization Theme ------------------------------------------------
theme_superstore <- function(base_size = 11, base_family = "sans") {
  theme_minimal(base_size = base_size, base_family = base_family) +
    theme(
      plot.title = element_text(face = "bold", size = rel(1.2), color = "#1B263B", margin = margin(b = 6)),
      plot.subtitle = element_text(size = rel(0.95), color = "#415A77", margin = margin(b = 10)),
      plot.caption = element_text(size = rel(0.8), color = "#778DA9", hjust = 0, margin = margin(t = 10)),
      axis.title = element_text(face = "bold", size = rel(0.9), color = "#1B263B"),
      axis.text = element_text(size = rel(0.85), color = "#2B2D42"),
      panel.grid.minor = element_blank(),
      panel.grid.major.x = element_line(color = "#E0E1DD", linewidth = 0.4),
      panel.grid.major.y = element_line(color = "#E0E1DD", linewidth = 0.4),
      legend.position = "bottom",
      legend.title = element_text(face = "bold", size = rel(0.85)),
      legend.text = element_text(size = rel(0.8)),
      plot.margin = margin(15, 15, 15, 15)
    )
}

# Define standard palette
superstore_palette <- c(
  "Furniture"       = "#E07A5F", # Warm terracotta
  "Office Supplies" = "#3D5A80", # Deep corporate blue
  "Technology"      = "#2B7A78"  # Slate teal
)

message("Setup completed successfully. Environment is ready for analysis.")
