# ==============================================================================
# Script: 05_descriptive_analysis.R
# Purpose: Execute core analytical methods: descriptive statistics, frequency
#          analysis, grouped multi-dimensional aggregations, chronological
#          time-series summaries, and correlation matrices.
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

message("Executing descriptive and exploratory statistical analyses...")

# ------------------------------------------------------------------------------
# METHOD 1: Descriptive Statistics for Numerical Variables
# ------------------------------------------------------------------------------
calc_num_stats <- function(x, name) {
  q <- quantile(x, probs = c(0.25, 0.50, 0.75), na.rm = TRUE)
  data.frame(
    Variable = name,
    Count    = sum(!is.na(x)),
    Mean     = round(mean(x, na.rm = TRUE), 2),
    Median   = round(q[2], 2),
    Std_Dev  = round(sd(x, na.rm = TRUE), 2),
    Min      = round(min(x, na.rm = TRUE), 2),
    Q1_25    = round(q[1], 2),
    Q3_75    = round(q[3], 2),
    Max      = round(max(x, na.rm = TRUE), 2),
    IQR      = round(q[3] - q[1], 2),
    stringsAsFactors = FALSE
  )
}

table2_numerical_summary <- bind_rows(
  calc_num_stats(superstore_clean$Sales, "Sales ($)"),
  calc_num_stats(superstore_clean$Quantity, "Quantity (Units)"),
  calc_num_stats(superstore_clean$Discount, "Discount (Rate)"),
  calc_num_stats(superstore_clean$Profit, "Profit ($)"),
  calc_num_stats(superstore_clean$Profit_Margin, "Profit Margin (%)")
)

write_csv(table2_numerical_summary, file.path("outputs", "tables", "table2_numerical_summary.csv"))

# ------------------------------------------------------------------------------
# METHOD 2: Frequency Analysis (Dataset Composition)
# ------------------------------------------------------------------------------
freq_category <- superstore_clean %>%
  count(Category, name = "Line_Items") %>%
  mutate(Share_Pct = round((Line_Items / sum(Line_Items)) * 100, 2))

freq_region <- superstore_clean %>%
  count(Region, name = "Line_Items") %>%
  mutate(Share_Pct = round((Line_Items / sum(Line_Items)) * 100, 2))

freq_segment <- superstore_clean %>%
  count(Segment, name = "Line_Items") %>%
  mutate(Share_Pct = round((Line_Items / sum(Line_Items)) * 100, 2))

freq_shipmode <- superstore_clean %>%
  count(Ship_Mode, name = "Line_Items") %>%
  mutate(Share_Pct = round((Line_Items / sum(Line_Items)) * 100, 2))

# ------------------------------------------------------------------------------
# METHOD 3: Grouped Aggregations
# ------------------------------------------------------------------------------
# 3A. Category Aggregation
table3_category <- superstore_clean %>%
  group_by(Category) %>%
  summarise(
    Total_Sales    = round(sum(Sales), 2),
    Sales_Share    = round((sum(Sales) / sum(superstore_clean$Sales)) * 100, 2),
    Total_Profit   = round(sum(Profit), 2),
    Profit_Share   = round((sum(Profit) / sum(superstore_clean$Profit)) * 100, 2),
    Avg_Sales      = round(mean(Sales), 2),
    Avg_Profit     = round(mean(Profit), 2),
    Avg_Discount   = round(mean(Discount) * 100, 2),
    Overall_Margin = round((sum(Profit) / sum(Sales)) * 100, 2),
    Total_Quantity = sum(Quantity),
    Total_Orders   = n_distinct(Order_ID),
    .groups = "drop"
  ) %>%
  arrange(desc(Total_Sales))

write_csv(table3_category, file.path("outputs", "tables", "table3_sales_profit_by_category.csv"))

# 3B. Regional Aggregation
table4_region <- superstore_clean %>%
  group_by(Region) %>%
  summarise(
    Total_Sales    = round(sum(Sales), 2),
    Sales_Share    = round((sum(Sales) / sum(superstore_clean$Sales)) * 100, 2),
    Total_Profit   = round(sum(Profit), 2),
    Profit_Share   = round((sum(Profit) / sum(superstore_clean$Profit)) * 100, 2),
    Avg_Sales      = round(mean(Sales), 2),
    Avg_Profit     = round(mean(Profit), 2),
    Avg_Discount   = round(mean(Discount) * 100, 2),
    Overall_Margin = round((sum(Profit) / sum(Sales)) * 100, 2),
    Total_Orders   = n_distinct(Order_ID),
    .groups = "drop"
  ) %>%
  arrange(desc(Total_Sales))

write_csv(table4_region, file.path("outputs", "tables", "table4_regional_performance.csv"))

# 3C. Customer Segment Aggregation
table5_segment <- superstore_clean %>%
  group_by(Segment) %>%
  summarise(
    Total_Sales    = round(sum(Sales), 2),
    Sales_Share    = round((sum(Sales) / sum(superstore_clean$Sales)) * 100, 2),
    Total_Profit   = round(sum(Profit), 2),
    Profit_Share   = round((sum(Profit) / sum(superstore_clean$Profit)) * 100, 2),
    Avg_Sales      = round(mean(Sales), 2),
    Avg_Profit     = round(mean(Profit), 2),
    Avg_Discount   = round(mean(Discount) * 100, 2),
    Overall_Margin = round((sum(Profit) / sum(Sales)) * 100, 2),
    Total_Orders   = n_distinct(Order_ID),
    .groups = "drop"
  ) %>%
  arrange(desc(Total_Sales))

write_csv(table5_segment, file.path("outputs", "tables", "table5_segment_performance.csv"))

# 3D. Sub-Category Aggregation
table6_subcategory <- superstore_clean %>%
  group_by(Category, Sub_Category) %>%
  summarise(
    Total_Sales    = round(sum(Sales), 2),
    Total_Profit   = round(sum(Profit), 2),
    Avg_Sales      = round(mean(Sales), 2),
    Avg_Profit     = round(mean(Profit), 2),
    Avg_Discount   = round(mean(Discount) * 100, 2),
    Overall_Margin = round((sum(Profit) / sum(Sales)) * 100, 2),
    Total_Quantity = sum(Quantity),
    Total_Orders   = n_distinct(Order_ID),
    .groups = "drop"
  ) %>%
  arrange(desc(Total_Profit))

write_csv(table6_subcategory, file.path("outputs", "tables", "table6_subcategory_profitability.csv"))

# ------------------------------------------------------------------------------
# METHOD 4: Time-Series Aggregations
# ------------------------------------------------------------------------------
monthly_trend <- superstore_clean %>%
  group_by(Order_Year_Month, Order_YM_Date) %>%
  summarise(
    Total_Sales  = round(sum(Sales), 2),
    Total_Profit = round(sum(Profit), 2),
    Total_Orders = n_distinct(Order_ID),
    Total_Units  = sum(Quantity),
    Avg_Discount = round(mean(Discount) * 100, 2),
    .groups = "drop"
  ) %>%
  arrange(Order_YM_Date)

write_csv(monthly_trend, file.path("outputs", "tables", "monthly_trend_summary.csv"))

yearly_trend <- superstore_clean %>%
  group_by(Order_Year) %>%
  summarise(
    Total_Sales  = round(sum(Sales), 2),
    Total_Profit = round(sum(Profit), 2),
    Total_Orders = n_distinct(Order_ID),
    Overall_Margin = round((sum(Profit) / sum(Sales)) * 100, 2),
    .groups = "drop"
  ) %>%
  arrange(Order_Year)

write_csv(yearly_trend, file.path("outputs", "tables", "yearly_trend_summary.csv"))

# ------------------------------------------------------------------------------
# METHOD 5: Correlation Analysis
# ------------------------------------------------------------------------------
num_vars <- superstore_clean %>% select(Sales, Quantity, Discount, Profit)

pearson_corr <- round(cor(num_vars, method = "pearson"), 4)
spearman_corr <- round(cor(num_vars, method = "spearman"), 4)

write.csv(pearson_corr, file.path("outputs", "tables", "correlation_matrix_pearson.csv"))
write.csv(spearman_corr, file.path("outputs", "tables", "correlation_matrix_spearman.csv"))

# ------------------------------------------------------------------------------
# METHOD 8: Shipping Fulfillment Analysis
# ------------------------------------------------------------------------------
shipping_summary <- superstore_clean %>%
  group_by(Ship_Mode) %>%
  summarise(
    Total_Orders = n_distinct(Order_ID),
    Total_Items  = n(),
    Mean_Days    = round(mean(Days_to_Ship), 2),
    Median_Days  = median(Days_to_Ship),
    Min_Days     = min(Days_to_Ship),
    Max_Days     = max(Days_to_Ship),
    Std_Dev_Days = round(sd(Days_to_Ship), 2),
    .groups = "drop"
  )

write_csv(shipping_summary, file.path("outputs", "tables", "shipping_mode_summary.csv"))

# ------------------------------------------------------------------------------
# METHOD 9: State Performance (Top 10 and Bottom 10 States by Profit)
# ------------------------------------------------------------------------------
state_perf <- superstore_clean %>%
  group_by(State, Region) %>%
  summarise(
    Total_Sales    = round(sum(Sales), 2),
    Total_Profit   = round(sum(Profit), 2),
    Overall_Margin = round((sum(Profit) / sum(Sales)) * 100, 2),
    Avg_Discount   = round(mean(Discount) * 100, 2),
    Total_Orders   = n_distinct(Order_ID),
    .groups = "drop"
  ) %>%
  arrange(desc(Total_Profit))

top_10_states <- head(state_perf, 10) %>% mutate(Group = "Top 10 Profitable")
bottom_10_states <- tail(state_perf, 10) %>% mutate(Group = "Bottom 10 Loss-Making")
table8_top_bottom_states <- bind_rows(top_10_states, bottom_10_states)

write_csv(table8_top_bottom_states, file.path("outputs", "tables", "table8_top_bottom_states.csv"))

# ------------------------------------------------------------------------------
# METHOD 10: Discount Band Performance Analysis
# ------------------------------------------------------------------------------
table9_discount_bands <- superstore_clean %>%
  mutate(
    Discount_Band = case_when(
      Discount == 0        ~ "0% (No Discount)",
      Discount <= 0.10     ~ "0.1% - 10%",
      Discount <= 0.20     ~ "10.1% - 20%",
      Discount <= 0.30     ~ "20.1% - 30%",
      Discount <= 0.50     ~ "30.1% - 50%",
      TRUE                 ~ "> 50%"
    ),
    Discount_Band = factor(Discount_Band, levels = c(
      "0% (No Discount)", "0.1% - 10%", "10.1% - 20%", "20.1% - 30%", "30.1% - 50%", "> 50%"
    ))
  ) %>%
  group_by(Discount_Band) %>%
  summarise(
    Transactions   = n(),
    Total_Sales    = round(sum(Sales), 2),
    Sales_Share    = round((sum(Sales) / sum(superstore_clean$Sales)) * 100, 2),
    Total_Profit   = round(sum(Profit), 2),
    Profit_Share   = round((sum(Profit) / sum(superstore_clean$Profit)) * 100, 2),
    Overall_Margin = round((sum(Profit) / sum(Sales)) * 100, 2),
    Total_Orders   = n_distinct(Order_ID),
    .groups = "drop"
  )

write_csv(table9_discount_bands, file.path("outputs", "tables", "table9_discount_bands.csv"))

# ------------------------------------------------------------------------------
# METHOD 11: Sales Tier Performance Analysis
# ------------------------------------------------------------------------------
table10_sales_tiers <- superstore_clean %>%
  mutate(
    Sales_Tier = case_when(
      Sales < 100   ~ "Tier 1: < $100",
      Sales < 500   ~ "Tier 2: $100 - $499",
      Sales < 1000  ~ "Tier 3: $500 - $999",
      Sales < 5000  ~ "Tier 4: $1,000 - $4,999",
      TRUE          ~ "Tier 5: >= $5,000"
    ),
    Sales_Tier = factor(Sales_Tier, levels = c(
      "Tier 1: < $100", "Tier 2: $100 - $499", "Tier 3: $500 - $999", "Tier 4: $1,000 - $4,999", "Tier 5: >= $5,000"
    ))
  ) %>%
  group_by(Sales_Tier) %>%
  summarise(
    Transactions   = n(),
    Total_Sales    = round(sum(Sales), 2),
    Sales_Share    = round((sum(Sales) / sum(superstore_clean$Sales)) * 100, 2),
    Total_Profit   = round(sum(Profit), 2),
    Profit_Share   = round((sum(Profit) / sum(superstore_clean$Profit)) * 100, 2),
    Overall_Margin = round((sum(Profit) / sum(Sales)) * 100, 2),
    .groups = "drop"
  )

write_csv(table10_sales_tiers, file.path("outputs", "tables", "table10_sales_profit_tiers.csv"))


# ------------------------------------------------------------------------------
# Print Summary Highlights to Console
# ------------------------------------------------------------------------------
cat("\n=======================================================\n")
cat("          DESCRIPTIVE ANALYSIS COMPLETED               \n")
cat("=======================================================\n")
cat("Top Category by Sales  : ", as.character(table3_category$Category[1]),
    sprintf("($%s | Share: %.1f%%)\n", format(table3_category$Total_Sales[1], big.mark = ","), table3_category$Sales_Share[1]))
cat("Top Category by Profit : ", as.character(table3_category$Category[which.max(table3_category$Total_Profit)]),
    sprintf("($%s | Share: %.1f%%)\n", format(max(table3_category$Total_Profit), big.mark = ","), max(table3_category$Profit_Share)))
cat("Lowest Profit Category : ", as.character(table3_category$Category[which.min(table3_category$Total_Profit)]),
    sprintf("($%s | Share: %.1f%%)\n", format(min(table3_category$Total_Profit), big.mark = ","), min(table3_category$Profit_Share)))
cat("Top Region by Sales    : ", as.character(table4_region$Region[1]),
    sprintf("($%s)\n", format(table4_region$Total_Sales[1], big.mark = ",")))
cat("Most Profitable SubCat : ", as.character(table6_subcategory$Sub_Category[1]),
    sprintf("($%s)\n", format(table6_subcategory$Total_Profit[1], big.mark = ",")))
cat("Deepest Loss SubCat    : ", as.character(tail(table6_subcategory$Sub_Category, 1)),
    sprintf("($%s)\n", format(tail(table6_subcategory$Total_Profit, 1), big.mark = ",")))
cat("Pearson Corr (Disc/Prof):", pearson_corr["Discount", "Profit"], "(Observed negative association)\n")
cat("Spearman Corr (Disc/Prof):", spearman_corr["Discount", "Profit"], "(Monotonic negative association)\n")
cat("=======================================================\n\n")

message("Descriptive tables generated and saved to outputs/tables/.")
