# ==============================================================================
# Script: 02_import_data.R
# Purpose: Import raw Superstore dataset, verify dimensions, columns,
#          and initial integrity without modifying raw source files.
# Author: Data Analyst & Statistical Specialist
# Date: 2026-09-28
# ==============================================================================

# Ensure setup has been executed
if (!exists("theme_superstore")) {
  source("R/01_setup.R")
}

# 1. File Path Resolution ------------------------------------------------------
raw_data_path <- file.path("data", "raw", "Superstore.csv")

if (!file.exists(raw_data_path)) {
  stop("CRITICAL ERROR: Raw data file not found at: ", raw_data_path)
}

# 2. Data Ingestion ------------------------------------------------------------
message("Ingesting raw Superstore dataset from: ", raw_data_path)

# Use readr::read_csv with explicit latin1 encoding to handle special characters cleanly
superstore_raw <- read_csv(
  file = raw_data_path,
  locale = locale(encoding = "latin1"),
  col_types = cols(
    `Row ID` = col_integer(),
    `Order ID` = col_character(),
    `Order Date` = col_character(),
    `Ship Date` = col_character(),
    `Ship Mode` = col_character(),
    `Customer ID` = col_character(),
    `Customer Name` = col_character(),
    `Segment` = col_character(),
    `Country` = col_character(),
    `City` = col_character(),
    `State` = col_character(),
    `Postal Code` = col_character(),  # Kept as character to preserve leading zeros
    `Region` = col_character(),
    `Product ID` = col_character(),
    `Category` = col_character(),
    `Sub-Category` = col_character(),
    `Product Name` = col_character(),
    `Sales` = col_double(),
    `Quantity` = col_integer(),
    `Discount` = col_double(),
    `Profit` = col_double()
  ),
  show_col_types = FALSE
)

# 3. Structural Verification ---------------------------------------------------
n_rows <- nrow(superstore_raw)
n_cols <- ncol(superstore_raw)
col_names <- colnames(superstore_raw)

cat("\n=======================================================\n")
cat("          RAW DATASET VERIFICATION SUMMARY             \n")
cat("=======================================================\n")
cat(sprintf("Observed Total Rows   : %s (Expected: 9,994)\n", format(n_rows, big.mark = ",")))
cat(sprintf("Observed Total Columns: %d (Expected: 21)\n", n_cols))
cat("-------------------------------------------------------\n")
cat("Column Names Verified:\n")
for (i in seq_along(col_names)) {
  cat(sprintf("  [%02d] %-16s | Type: %s\n", i, col_names[i], class(superstore_raw[[i]])[1]))
}
cat("=======================================================\n\n")

if (n_rows != 9994 || n_cols != 21) {
  warning("Dataset dimensions differ from expected 9,994 rows and 21 columns.")
} else {
  message("SUCCESS: Raw dataset dimensions and structure match expected specifications.")
}
