# ==============================================================================
# Script: 04_data_cleaning.R
# Purpose: Transform raw data into an analytical dataset. Engineer date features,
#          calculate derived metrics, standardize variable names, and export
#          processed clean datasets without altering raw source data.
# Author: Data Analyst & Statistical Specialist
# Date: 2026-09-28
# ==============================================================================

if (!exists("superstore_raw")) {
  source("R/02_import_data.R")
}

message("Beginning data transformation and feature engineering...")

# 1. Pipeline Transformation ---------------------------------------------------
superstore_clean <- superstore_raw %>%
  # Standardize column naming conventions
  rename(
    Row_ID        = `Row ID`,
    Order_ID      = `Order ID`,
    Order_Date    = `Order Date`,
    Ship_Date     = `Ship Date`,
    Ship_Mode     = `Ship Mode`,
    Customer_ID   = `Customer ID`,
    Customer_Name = `Customer Name`,
    Postal_Code   = `Postal Code`,
    Product_ID    = `Product ID`,
    Sub_Category  = `Sub-Category`,
    Product_Name  = `Product Name`
  ) %>%
  # Parse Dates explicitly using DD-MM-YYYY format
  mutate(
    Order_Date_Clean = lubridate::dmy(Order_Date),
    Ship_Date_Clean  = lubridate::dmy(Ship_Date)
  ) %>%
  # Feature Engineering
  mutate(
    # Temporal attributes
    Order_Year       = lubridate::year(Order_Date_Clean),
    Order_Month      = lubridate::month(Order_Date_Clean),
    Order_Month_Name = factor(
      month.abb[Order_Month],
      levels = month.abb,
      ordered = TRUE
    ),
    Order_Year_Month = format(Order_Date_Clean, "%Y-%m"),
    Order_YM_Date    = as.Date(paste0(Order_Year_Month, "-01")),
    
    # Operational attributes
    Days_to_Ship     = as.numeric(difftime(Ship_Date_Clean, Order_Date_Clean, units = "days")),
    
    # Financial attributes
    # Transaction-level profit margin (%)
    Profit_Margin    = ifelse(Sales > 0, (Profit / Sales) * 100, NA_real_),
    
    # Standardize factors
    Category         = factor(Category),
    Sub_Category     = factor(Sub_Category),
    Region           = factor(Region),
    Segment          = factor(Segment),
    Ship_Mode        = factor(Ship_Mode, levels = c("Same Day", "First Class", "Second Class", "Standard Class"))
  ) %>%
  # Reorder columns logically
  select(
    Row_ID, Order_ID, Order_Date_Clean, Ship_Date_Clean,
    Order_Year, Order_Month, Order_Month_Name, Order_Year_Month, Order_YM_Date,
    Days_to_Ship, Ship_Mode,
    Customer_ID, Customer_Name, Segment,
    Country, City, State, Postal_Code, Region,
    Product_ID, Category, Sub_Category, Product_Name,
    Sales, Quantity, Discount, Profit, Profit_Margin
  )

# 2. Validation of Cleaned Dataset --------------------------------------------
cat("\n=======================================================\n")
cat("          CLEANED DATASET VALIDATION SUMMARY           \n")
cat("=======================================================\n")
cat(sprintf("Cleaned Dataset Rows    : %s\n", format(nrow(superstore_clean), big.mark = ",")))
cat(sprintf("Cleaned Dataset Columns : %d\n", ncol(superstore_clean)))
cat(sprintf("Order Date Range        : %s to %s\n",
            min(superstore_clean$Order_Date_Clean),
            max(superstore_clean$Order_Date_Clean)))
cat(sprintf("Days to Ship Min/Max    : %.0f days to %.0f days (Mean: %.2f days)\n",
            min(superstore_clean$Days_to_Ship),
            max(superstore_clean$Days_to_Ship),
            mean(superstore_clean$Days_to_Ship)))
cat(sprintf("Profit Margin Range     : %.1f%% to %.1f%% (Median: %.2f%%)\n",
            min(superstore_clean$Profit_Margin, na.rm = TRUE),
            max(superstore_clean$Profit_Margin, na.rm = TRUE),
            median(superstore_clean$Profit_Margin, na.rm = TRUE)))
cat("=======================================================\n\n")

# 3. Export Processed Data -----------------------------------------------------
clean_csv_path <- file.path("data", "processed", "superstore_clean.csv")
clean_rds_path <- file.path("data", "processed", "superstore_clean.rds")

write_csv(superstore_clean, clean_csv_path)
saveRDS(superstore_clean, clean_rds_path)

message("Cleaned dataset successfully saved to: ", clean_csv_path)
message("Cleaned R binary dataset saved to: ", clean_rds_path)
