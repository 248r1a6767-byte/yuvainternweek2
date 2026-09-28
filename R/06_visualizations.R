# ==============================================================================
# Script: 06_visualizations.R
# Purpose: Generate high-resolution (300 DPI, 10x6 in) publication-quality
#          ggplot2 visualizations addressing all core research questions.
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
if (!exists("theme_superstore")) {
  source("R/01_setup.R")
}

message("Generating professional ggplot2 visualizations (300 DPI)...")

# Ensure output directory exists
vis_dir <- "visualizations"
if (!dir.exists(vis_dir)) dir.create(vis_dir, recursive = TRUE)

# Color constants
col_navy      <- "#1D3557"
col_teal      <- "#2A9D8F"
col_coral     <- "#E76F51"
col_sand      <- "#F4A261"
col_crimson   <- "#D90429"
col_emerald   <- "#2D6A4F"
col_slate     <- "#457B9D"

# ------------------------------------------------------------------------------
# VISUALIZATION 1: Sales by Category (Horizontal Bar Chart)
# ------------------------------------------------------------------------------
sales_cat_data <- superstore_clean %>%
  group_by(Category) %>%
  summarise(Total_Sales = sum(Sales), .groups = "drop") %>%
  mutate(
    Sales_Label = paste0("$", format(round(Total_Sales / 1000, 1), nsmall = 1), "K"),
    Share_Pct   = paste0(round((Total_Sales / sum(Total_Sales)) * 100, 1), "%")
  )

p1 <- ggplot(sales_cat_data, aes(x = reorder(Category, Total_Sales), y = Total_Sales, fill = Category)) +
  geom_col(width = 0.65, show.legend = FALSE) +
  geom_text(
    aes(label = paste0(Sales_Label, " (", Share_Pct, ")")),
    hjust = -0.15, size = 4.2, fontface = "bold", color = "#1B263B"
  ) +
  coord_flip() +
  scale_y_continuous(
    labels = dollar_format(prefix = "$", scale = 1e-3, suffix = "K"),
    expand = expansion(mult = c(0, 0.22))
  ) +
  scale_fill_manual(values = superstore_palette) +
  labs(
    title = "Visualization 1: Total Sales Revenue by Product Category",
    subtitle = "Technology generates the highest cumulative sales ($836.2K, 36.4%), closely followed by Furniture and Office Supplies",
    x = "Product Category",
    y = "Total Sales Revenue (USD)",
    caption = "Source: Superstore Dataset (9,994 transactions, 2011–2014) | Aggregation: sum(Sales)"
  ) +
  theme_superstore()

ggsave(file.path(vis_dir, "01_sales_by_category.png"), plot = p1, width = 10, height = 6, dpi = 300)
message("Saved: 01_sales_by_category.png")

# ------------------------------------------------------------------------------
# VISUALIZATION 2: Profit by Category (Comparative Bar Chart)
# ------------------------------------------------------------------------------
profit_cat_data <- superstore_clean %>%
  group_by(Category) %>%
  summarise(
    Total_Profit = sum(Profit),
    Total_Sales  = sum(Sales),
    .groups = "drop"
  ) %>%
  mutate(
    Margin_Pct   = (Total_Profit / Total_Sales) * 100,
    Profit_Label = paste0("$", format(round(Total_Profit / 1000, 1), nsmall = 1), "K"),
    Margin_Label = paste0("Margin: ", round(Margin_Pct, 1), "%")
  )

p2 <- ggplot(profit_cat_data, aes(x = reorder(Category, Total_Profit), y = Total_Profit, fill = Category)) +
  geom_col(width = 0.65, show.legend = FALSE) +
  geom_text(
    aes(label = paste0(Profit_Label, "\n(", Margin_Label, ")")),
    hjust = -0.15, size = 4.0, fontface = "bold", color = "#1B263B", lineheight = 0.9
  ) +
  coord_flip() +
  scale_y_continuous(
    labels = dollar_format(prefix = "$", scale = 1e-3, suffix = "K"),
    expand = expansion(mult = c(0, 0.25))
  ) +
  scale_fill_manual(values = superstore_palette) +
  labs(
    title = "Visualization 2: Total Net Profit and Profit Margin by Product Category",
    subtitle = "Technology yields $145.5K (17.4% margin), whereas Furniture delivers only $18.5K (2.5% margin) despite $742K in sales",
    x = "Product Category",
    y = "Total Net Profit (USD)",
    caption = "Source: Superstore Dataset (9,994 transactions, 2011–2014) | Margin calculated as sum(Profit)/sum(Sales)*100"
  ) +
  theme_superstore()

ggsave(file.path(vis_dir, "02_profit_by_category.png"), plot = p2, width = 10, height = 6, dpi = 300)
message("Saved: 02_profit_by_category.png")

# ------------------------------------------------------------------------------
# VISUALIZATION 3: Sales Distribution (Histogram with Log-Scale & Median Marker)
# ------------------------------------------------------------------------------
med_sales <- median(superstore_clean$Sales)
mean_sales <- mean(superstore_clean$Sales)

# Subsetting under $1,000 for primary view with annotation on total distribution
p3 <- ggplot(superstore_clean, aes(x = Sales)) +
  geom_histogram(
    bins = 50, fill = col_slate, color = "white", alpha = 0.9
  ) +
  geom_vline(xintercept = med_sales, color = col_crimson, linetype = "dashed", linewidth = 1) +
  geom_vline(xintercept = mean_sales, color = "#1B263B", linetype = "dotted", linewidth = 1) +
  scale_x_log10(
    labels = dollar_format(prefix = "$"),
    breaks = c(1, 5, 10, 50, 100, 500, 1000, 5000, 20000)
  ) +
  scale_y_continuous(labels = comma_format(), expand = expansion(mult = c(0, 0.1))) +
  annotate(
    "text", x = med_sales * 0.4, y = 780,
    label = paste0("Median Sales: $", round(med_sales, 2)),
    color = col_crimson, fontface = "bold", size = 3.8, hjust = 1
  ) +
  annotate(
    "text", x = mean_sales * 2.2, y = 680,
    label = paste0("Mean Sales: $", round(mean_sales, 2)),
    color = "#1B263B", fontface = "bold", size = 3.8, hjust = 0
  ) +
  labs(
    title = "Visualization 3: Distribution of Individual Transaction Sales (Log10 Scale)",
    subtitle = "Marked positive skewness: 75% of orders are under $210, pulling the mean ($229.86) far above the median ($54.49)",
    x = "Transaction Sales Revenue in USD (Logarithmic Base 10 Scale)",
    y = "Number of Line-Item Transactions (Count)",
    caption = "Source: Superstore Dataset (9,994 records) | Log10 transformation used to display extreme positive tail up to $22,638.48"
  ) +
  theme_superstore()

ggsave(file.path(vis_dir, "03_sales_distribution.png"), plot = p3, width = 10, height = 6, dpi = 300)
message("Saved: 03_sales_distribution.png")

# ------------------------------------------------------------------------------
# VISUALIZATION 4: Profit Distribution (Histogram with Loss/Profit Coloring)
# ------------------------------------------------------------------------------
# Filter visual window to -$500 to +$500 to clearly show the density mass around zero
p4 <- ggplot(superstore_clean, aes(x = Profit, fill = Profit >= 0)) +
  geom_histogram(
    binwidth = 15, boundary = 0, color = "white", alpha = 0.88, show.legend = TRUE
  ) +
  geom_vline(xintercept = 0, color = "#1B263B", linewidth = 1.1) +
  scale_x_continuous(
    labels = dollar_format(prefix = "$"),
    limits = c(-500, 500),
    breaks = seq(-500, 500, by = 100)
  ) +
  scale_y_continuous(labels = comma_format(), expand = expansion(mult = c(0, 0.08))) +
  scale_fill_manual(
    name = "Transaction Outcome",
    values = c("TRUE" = col_emerald, "FALSE" = col_crimson),
    labels = c("TRUE" = "Profitable Transaction (81.3%)", "FALSE" = "Commercial Loss (18.7%)")
  ) +
  annotate(
    "label", x = -280, y = 1400,
    label = "1,871 Loss-Making Orders\nDeep losses extend to -$6,600",
    fill = "#FDE8E8", color = col_crimson, fontface = "bold", size = 3.6
  ) +
  annotate(
    "label", x = 280, y = 1400,
    label = "8,123 Profitable Orders\nHigh gains extend to +$8,400",
    fill = "#E8F5E9", color = col_emerald, fontface = "bold", size = 3.6
  ) +
  labs(
    title = "Visualization 4: Distribution of Transaction Net Profit Around Breakeven ($0)",
    subtitle = "High concentration near zero (Median: $8.67); 18.7% of line items generate losses, creating a heavy negative tail",
    x = "Transaction Net Profit in USD (Clamped to [-$500, +$500] for visual clarity; full range: [-$6,600, +$8,400])",
    y = "Number of Line-Item Transactions (Count)",
    caption = "Source: Superstore Dataset (9,994 records) | Vertical solid black line indicates breakeven threshold ($0)"
  ) +
  theme_superstore()

ggsave(file.path(vis_dir, "04_profit_distribution.png"), plot = p4, width = 10, height = 6, dpi = 300)
message("Saved: 04_profit_distribution.png")

# ------------------------------------------------------------------------------
# VISUALIZATION 5: Sales Over Time (Monthly Chronological Trend)
# ------------------------------------------------------------------------------
monthly_sales <- superstore_clean %>%
  group_by(Order_YM_Date) %>%
  summarise(
    Total_Sales = sum(Sales),
    Total_Profit = sum(Profit),
    .groups = "drop"
  ) %>%
  arrange(Order_YM_Date)

p5 <- ggplot(monthly_sales, aes(x = Order_YM_Date, y = Total_Sales)) +
  geom_area(fill = col_slate, alpha = 0.15) +
  geom_line(color = col_navy, linewidth = 1.1) +
  geom_point(color = col_navy, size = 2.4, shape = 21, fill = "white", stroke = 1.2) +
  geom_smooth(method = "loess", color = col_coral, linetype = "dashed", se = FALSE, linewidth = 0.9) +
  scale_x_date(
    date_breaks = "6 months",
    date_labels = "%b %Y",
    expand = expansion(mult = c(0.02, 0.04))
  ) +
  scale_y_continuous(
    labels = dollar_format(prefix = "$", scale = 1e-3, suffix = "K"),
    breaks = seq(0, 120000, by = 20000),
    expand = expansion(mult = c(0, 0.1))
  ) +
  annotate(
    "text", x = as.Date("2014-11-01"), y = 118400,
    label = "Nov 2014 Peak\n$118.4K",
    fontface = "bold", size = 3.5, color = col_navy, vjust = -0.5
  ) +
  labs(
    title = "Visualization 5: Chronological Monthly Sales Trend (Jan 2011 – Dec 2014)",
    subtitle = "Consistent multi-year revenue expansion accompanied by pronounced recurring Q4 seasonal peaks (Nov/Dec)",
    x = "Order Timeline (Month & Year)",
    y = "Monthly Total Sales Revenue (USD)",
    caption = "Source: Superstore Dataset (48 monthly aggregated intervals) | Dashed red curve represents LOESS smoothed trajectory"
  ) +
  theme_superstore()

ggsave(file.path(vis_dir, "05_monthly_sales_trend.png"), plot = p5, width = 10, height = 6, dpi = 300)
message("Saved: 05_monthly_sales_trend.png")

# ------------------------------------------------------------------------------
# VISUALIZATION 6: Discount vs Profit (Scatter Plot with Association Curve)
# ------------------------------------------------------------------------------
p6 <- ggplot(superstore_clean, aes(x = Discount, y = Profit)) +
  geom_hline(yintercept = 0, color = "gray40", linetype = "dashed", linewidth = 0.8) +
  geom_point(aes(color = Profit >= 0), alpha = 0.35, size = 1.8) +
  geom_smooth(method = "loess", color = col_crimson, fill = "gray80", linewidth = 1.1) +
  scale_x_continuous(
    labels = percent_format(accuracy = 1),
    breaks = seq(0, 0.8, by = 0.1)
  ) +
  scale_y_continuous(
    labels = dollar_format(prefix = "$"),
    breaks = seq(-6000, 8000, by = 2000)
  ) +
  scale_color_manual(
    name = "Transaction Profitability",
    values = c("TRUE" = col_slate, "FALSE" = col_crimson),
    labels = c("TRUE" = "Profit >= $0", "FALSE" = "Loss < $0")
  ) +
  annotate(
    "rect", xmin = 0.25, xmax = 0.82, ymin = -6800, ymax = -50,
    alpha = 0.08, fill = col_crimson, color = col_crimson, linetype = "dotted"
  ) +
  annotate(
    "text", x = 0.55, y = -4500,
    label = "High Discount Zone (>=30%)\nAccelerated Negative Margin Concentration",
    color = col_crimson, fontface = "bold", size = 3.6
  ) +
  labs(
    title = "Visualization 6: Observed Association Between Discount Level and Transaction Profit",
    subtitle = "Statistical association reveals steep deterioration in profit when discounts exceed 20% (Correlation: r = -0.219)",
    x = "Promotional Discount Applied (Rate %)",
    y = "Net Profit in USD (Per Line-Item Transaction)",
    caption = "Source: Superstore Dataset (9,994 records) | Empirical observation shows association, not direct causal mechanism"
  ) +
  theme_superstore()

ggsave(file.path(vis_dir, "06_discount_vs_profit.png"), plot = p6, width = 10, height = 6, dpi = 300)
message("Saved: 06_discount_vs_profit.png")

# ------------------------------------------------------------------------------
# VISUALIZATION 7: Sales vs Profit (Scatter Plot with Category Coloring)
# ------------------------------------------------------------------------------
p7 <- ggplot(superstore_clean, aes(x = Sales, y = Profit, color = Category)) +
  geom_hline(yintercept = 0, color = "gray30", linetype = "dashed", linewidth = 0.8) +
  geom_point(alpha = 0.45, size = 2.0) +
  scale_x_continuous(
    labels = dollar_format(prefix = "$"),
    breaks = seq(0, 24000, by = 4000)
  ) +
  scale_y_continuous(
    labels = dollar_format(prefix = "$"),
    breaks = seq(-6000, 8000, by = 2000)
  ) +
  scale_color_manual(values = superstore_palette) +
  annotate(
    "text", x = 18000, y = 7800,
    label = "Top Profit: Technology Copiers\n(+$8,400 max transaction profit)",
    color = "#2B7A78", fontface = "bold", size = 3.4, hjust = 0.5
  ) +
  annotate(
    "text", x = 11000, y = -6200,
    label = "Deepest Loss: Technology Machines\n(-$6,600 loss at 70% discount)",
    color = col_crimson, fontface = "bold", size = 3.4, hjust = 0.5
  ) +
  labs(
    title = "Visualization 7: Bivariate Relationship Between Sales Revenue and Net Profit",
    subtitle = "The spread of profit widens dramatically as sales increase; high sales do not guarantee high profits",
    x = "Transaction Sales Revenue (USD)",
    y = "Transaction Net Profit (USD)",
    caption = "Source: Superstore Dataset (9,994 records) | Widening dispersion illustrates increased variance at higher price points"
  ) +
  theme_superstore()

ggsave(file.path(vis_dir, "07_sales_vs_profit.png"), plot = p7, width = 10, height = 6, dpi = 300)
message("Saved: 07_sales_vs_profit.png")

# ------------------------------------------------------------------------------
# VISUALIZATION 8: Profit by Sub-Category (Horizontal Diverging Bar Chart)
# ------------------------------------------------------------------------------
subcat_data <- superstore_clean %>%
  group_by(Sub_Category, Category) %>%
  summarise(Total_Profit = sum(Profit), .groups = "drop") %>%
  mutate(
    Is_Profitable = Total_Profit >= 0,
    Profit_Label  = paste0(ifelse(Total_Profit >= 0, "+$", "-$"),
                           format(abs(round(Total_Profit / 1000, 1)), nsmall = 1), "K")
  )

p8 <- ggplot(subcat_data, aes(x = reorder(Sub_Category, Total_Profit), y = Total_Profit, fill = Is_Profitable)) +
  geom_col(width = 0.7) +
  geom_hline(yintercept = 0, color = "black", linewidth = 0.9) +
  geom_text(
    aes(
      label = Profit_Label,
      hjust = ifelse(Total_Profit >= 0, -0.15, 1.15)
    ),
    size = 3.5, fontface = "bold"
  ) +
  coord_flip() +
  scale_y_continuous(
    labels = dollar_format(prefix = "$", scale = 1e-3, suffix = "K"),
    expand = expansion(mult = c(0.18, 0.22))
  ) +
  scale_fill_manual(
    name = "Performance Status",
    values = c("TRUE" = col_teal, "FALSE" = col_crimson),
    labels = c("TRUE" = "Net Profitable Sub-Category", "FALSE" = "Net Deficit / Loss Sub-Category")
  ) +
  labs(
    title = "Visualization 8: Net Profitability Across Product Sub-Categories (Diverging)",
    subtitle = "Copiers lead all sub-categories (+$55.6K), while Tables (-$17.7K), Bookcases (-$3.5K), and Supplies (-$1.2K) operate at net losses",
    x = "Product Sub-Category",
    y = "Cumulative Net Profit (USD)",
    caption = "Source: Superstore Dataset (17 sub-categories, 2011–2014) | Diverging display emphasizes structural loss centers"
  ) +
  theme_superstore()

ggsave(file.path(vis_dir, "08_profit_by_subcategory.png"), plot = p8, width = 10, height = 7, dpi = 300)
message("Saved: 08_profit_by_subcategory.png")

# ------------------------------------------------------------------------------
# VISUALIZATION 9: Profit Distribution by Category (Box Plot with Interquartiles)
# ------------------------------------------------------------------------------
p9 <- ggplot(superstore_clean, aes(x = Category, y = Profit, fill = Category)) +
  geom_hline(yintercept = 0, color = "gray40", linetype = "dashed", linewidth = 0.8) +
  geom_boxplot(
    outlier.alpha = 0.25, outlier.size = 1.4, outlier.color = "gray30",
    width = 0.5, show.legend = FALSE
  ) +
  coord_cartesian(ylim = c(-300, 300)) + # Zoom in on the central 95% of data
  scale_y_continuous(
    labels = dollar_format(prefix = "$"),
    breaks = seq(-300, 300, by = 100)
  ) +
  scale_fill_manual(values = superstore_palette) +
  annotate(
    "text", x = 1, y = -260,
    label = "Furniture: Median $7.78\nWide negative quartile spread",
    fontface = "bold", size = 3.5, color = superstore_palette["Furniture"]
  ) +
  annotate(
    "text", x = 2, y = 260,
    label = "Office Supplies: Median $6.88\nTight IQR around breakeven",
    fontface = "bold", size = 3.5, color = superstore_palette["Office Supplies"]
  ) +
  annotate(
    "text", x = 3, y = 260,
    label = "Technology: Median $25.02\nHighest median & positive skew",
    fontface = "bold", size = 3.5, color = superstore_palette["Technology"]
  ) +
  labs(
    title = "Visualization 9: Transaction Profit Distribution Across Product Categories (Box Plot)",
    subtitle = "Technology demonstrates substantially higher median profitability and upper quartile reach than Furniture",
    x = "Product Category",
    y = "Transaction Net Profit in USD (Y-axis clamped to [-$300, +$300] to reveal central box distributions)",
    caption = "Source: Superstore Dataset | Thick bar = Median, Box = Interquartile Range (Q1–Q3), Points = Statistical Outliers"
  ) +
  theme_superstore()

ggsave(file.path(vis_dir, "09_profit_boxplot_category.png"), plot = p9, width = 10, height = 6, dpi = 300)
message("Saved: 09_profit_boxplot_category.png")

# ------------------------------------------------------------------------------
# VISUALIZATION 10: Regional Performance (Dual Sales & Profit Comparison)
# ------------------------------------------------------------------------------
region_summary <- superstore_clean %>%
  group_by(Region) %>%
  summarise(
    Sales  = sum(Sales),
    Profit = sum(Profit),
    .groups = "drop"
  ) %>%
  tidyr::pivot_longer(cols = c(Sales, Profit), names_to = "Metric", values_to = "Amount")

p10 <- ggplot(region_summary, aes(x = reorder(Region, -Amount), y = Amount, fill = Metric)) +
  geom_col(position = position_dodge(width = 0.75), width = 0.65) +
  geom_text(
    aes(label = paste0("$", format(round(Amount / 1000, 1), nsmall = 1), "K")),
    position = position_dodge(width = 0.75),
    vjust = -0.4, size = 3.6, fontface = "bold"
  ) +
  scale_y_continuous(
    labels = dollar_format(prefix = "$", scale = 1e-3, suffix = "K"),
    expand = expansion(mult = c(0, 0.15))
  ) +
  scale_fill_manual(
    name = "Financial Metric",
    values = c("Sales" = col_navy, "Profit" = col_teal)
  ) +
  labs(
    title = "Visualization 10: Geographic Sales and Net Profit Performance by US Region",
    subtitle = "West leads both total revenue ($725.5K) and profit ($108.4K); Central suffers margin erosion ($39.7K profit on $501.2K sales)",
    x = "Geographic Region",
    y = "Total Financial Value (USD)",
    caption = "Source: Superstore Dataset (4 geographic regions, 2011–2014) | Central profit margin is only 7.9% vs West margin of 14.9%"
  ) +
  theme_superstore()

ggsave(file.path(vis_dir, "10_regional_performance.png"), plot = p10, width = 10, height = 6, dpi = 300)
message("Saved: 10_regional_performance.png")

# ------------------------------------------------------------------------------
# VISUALIZATION 11: Shipping Fulfillment Duration by Ship Mode (Optional Analysis)
# ------------------------------------------------------------------------------
p11 <- ggplot(superstore_clean, aes(x = Ship_Mode, y = Days_to_Ship, fill = Ship_Mode)) +
  geom_boxplot(width = 0.5, alpha = 0.85, show.legend = FALSE, outlier.color = "gray40") +
  stat_summary(fun = mean, geom = "point", shape = 18, size = 3.5, color = col_crimson) +
  scale_y_continuous(breaks = 0:7, limits = c(0, 7.5)) +
  scale_fill_brewer(palette = "Blues") +
  annotate(
    "text", x = 1:4, y = c(0.8, 3.0, 4.1, 5.9),
    label = c("Mean: 0.04 d", "Mean: 2.18 d", "Mean: 3.24 d", "Mean: 5.01 d"),
    fontface = "bold", size = 3.4, color = "#1B263B"
  ) +
  labs(
    title = "Visualization 11: Order Fulfillment Duration Across Standardized Shipping Modes",
    subtitle = "Strict adherence to delivery tiers: Same Day fulfills within 24h, whereas Standard Class averages 5.0 days (Max: 7 days)",
    x = "Logistics Shipping Mode Tier",
    y = "Days Elapsed from Order to Shipment (Days)",
    caption = "Source: Superstore Dataset (9,994 orders) | Red diamonds represent arithmetic mean duration"
  ) +
  theme_superstore()

ggsave(file.path(vis_dir, "11_shipping_time_by_mode.png"), plot = p11, width = 10, height = 6, dpi = 300)
message("Saved: 11_shipping_time_by_mode.png")

message("All 11 visualizations successfully generated at 300 DPI!")
