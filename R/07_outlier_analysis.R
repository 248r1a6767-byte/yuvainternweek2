# ==============================================================================
# Script: 07_outlier_analysis.R
# Purpose: Method 6 Outlier & Anomaly Analysis. Detect statistical anomalies
#          using IQR criteria and extreme percentiles for Sales, Profit, and
#          Discount; inspect raw business transactions; evaluate validity.
# Author: Data Analyst & Statistical Specialist
# Date: 2026-09-28
# ==============================================================================

if (!exists("superstore_clean")) {
  if (file.exists("data/processed/superstore_clean.rds")) {
    superstore_clean <- readRDS("data/processed/superstore_clean.rds")
  } else {
    source("R/04_data_cleaning.R")
  }
}

message("Conducting systematic Outlier and Anomaly Analysis...")

# 1. IQR Outlier Boundary Calculation Function ---------------------------------
detect_iqr_outliers <- function(data, variable_name) {
  x <- data[[variable_name]]
  q25 <- quantile(x, 0.25, na.rm = TRUE)
  q75 <- quantile(x, 0.75, na.rm = TRUE)
  iqr_val <- q75 - q25
  lower_bound <- q25 - 1.5 * iqr_val
  upper_bound <- q75 + 1.5 * iqr_val
  
  is_outlier <- x < lower_bound | x > upper_bound
  outlier_count <- sum(is_outlier, na.rm = TRUE)
  outlier_pct <- (outlier_count / length(x)) * 100
  
  data.frame(
    Variable       = variable_name,
    Q1_25          = round(q25, 2),
    Median         = round(median(x, na.rm = TRUE), 2),
    Q3_75          = round(q75, 2),
    IQR            = round(iqr_val, 2),
    Lower_Limit    = round(lower_bound, 2),
    Upper_Limit    = round(upper_bound, 2),
    Outlier_Count  = outlier_count,
    Outlier_Pct    = round(outlier_pct, 2),
    Min_Observed   = round(min(x, na.rm = TRUE), 2),
    Max_Observed   = round(max(x, na.rm = TRUE), 2),
    stringsAsFactors = FALSE
  )
}

# 2. Compute Outlier Thresholds ------------------------------------------------
table7_outliers_summary <- bind_rows(
  detect_iqr_outliers(superstore_clean, "Sales"),
  detect_iqr_outliers(superstore_clean, "Profit"),
  detect_iqr_outliers(superstore_clean, "Discount"),
  detect_iqr_outliers(superstore_clean, "Quantity")
)

write_csv(table7_outliers_summary, file.path("outputs", "tables", "table7_outliers_summary.csv"))

# 3. Qualitative Inspection of Extreme Commercial Transactions -----------------
# 3A. Top 5 Highest Sales Transactions
top_sales <- superstore_clean %>%
  arrange(desc(Sales)) %>%
  slice(1:5) %>%
  select(Row_ID, Order_ID, Order_Date_Clean, Customer_Name, Category, Sub_Category,
         Product_Name, Sales, Quantity, Discount, Profit, Profit_Margin) %>%
  mutate(Transaction_Type = "Top Positive Sales")

# 3B. Top 5 Deepest Loss Transactions
top_losses <- superstore_clean %>%
  arrange(Profit) %>%
  slice(1:5) %>%
  select(Row_ID, Order_ID, Order_Date_Clean, Customer_Name, Category, Sub_Category,
         Product_Name, Sales, Quantity, Discount, Profit, Profit_Margin) %>%
  mutate(Transaction_Type = "Deepest Commercial Loss")

# 3C. Top 5 Highest Profit Transactions
top_profits <- superstore_clean %>%
  arrange(desc(Profit)) %>%
  slice(1:5) %>%
  select(Row_ID, Order_ID, Order_Date_Clean, Customer_Name, Category, Sub_Category,
         Product_Name, Sales, Quantity, Discount, Profit, Profit_Margin) %>%
  mutate(Transaction_Type = "Top Net Profit")

# Combine for forensic auditing table
top_extreme_records <- bind_rows(top_sales, top_losses, top_profits)
write_csv(top_extreme_records, file.path("outputs", "tables", "top_extreme_transactions.csv"))

# 4. Outlier Interpretation Summary to Console ---------------------------------
cat("\n=======================================================\n")
cat("          OUTLIER FORENSIC AUDIT SUMMARY               \n")
cat("=======================================================\n")
print(table7_outliers_summary[, c("Variable", "Lower_Limit", "Upper_Limit", "Outlier_Count", "Outlier_Pct")])
cat("-------------------------------------------------------\n")
cat("MAXIMUM SALES RECORD:\n")
cat(sprintf("  Row ID: %d | Product: %s\n  Sales: $%s | Profit: $%s | Discount: %.0f%%\n",
            top_sales$Row_ID[1], top_sales$Product_Name[1],
            format(top_sales$Sales[1], big.mark = ","),
            format(top_sales$Profit[1], big.mark = ","),
            top_sales$Discount[1] * 100))
cat("\nDEEPEST COMMERCIAL LOSS RECORD:\n")
cat(sprintf("  Row ID: %d | Product: %s\n  Sales: $%s | Loss: -$%s | Discount: %.0f%%\n",
            top_losses$Row_ID[1], top_losses$Product_Name[1],
            format(top_losses$Sales[1], big.mark = ","),
            format(abs(top_losses$Profit[1]), big.mark = ","),
            top_losses$Discount[1] * 100))
cat("=======================================================\n\n")

message("Outlier assessment completed. Summary saved to outputs/tables/table7_outliers_summary.csv")
