# Week 2 Internship Project: Data Visualization and Insight Communication Using R

**Internship Track:** Data Analytics & Science  
**Task:** Week 2 – Data Visualization and Insight Communication using R  
**Primary Technology:** R (v4.6.1), `ggplot2`, `dplyr`, `readr`, `lubridate`, `scales`, `forcats`, `patchwork`, `tidyr`  
**Dataset:** Kaggle Superstore Sales Dataset (9,994 rows, 21 columns)  
**Report Output:** Microsoft Word DOCX (`report/Week2_Superstore_Data_Visualization_Report.docx`)  

---

## 1. Project Overview & Objective

The primary objective of this project is to apply modern statistical data visualization and analytical communication principles in R to transform retail enterprise data into actionable business intelligence. Rather than simply generating decorative charts, each visualization is engineered to address a specific commercial research question, ground its findings in empirical statistical methods, and communicate strategic insights in accessible language.

The analysis investigates:
1. **Sales Performance & Portfolio Diversity:** Top-line revenue distribution across categories and customer segments.
2. **Profitability Dynamics & Asymmetries:** The acute divergence between sales volume and bottom-line margin return.
3. **The "Furniture Deficit":** Granular sub-category root-cause analysis isolating loss drivers in Tables and Bookcases.
4. **Promotional Discounting Behavior:** Non-linear margin decay and empirical quantification of the "20% Discount Cliff."
5. **Macro-Geographic Variations:** Regional profit collapse in Central versus operational margin excellence in the West.
6. **Temporal Trajectories & Seasonality:** 48-month chronological growth trends and recurring Q4 holiday surges.
7. **Logistics Fulfillment Adherence:** Operational compliance across logistics tiers (Same Day to Standard Class).
8. **Forensic Outlier Auditing:** Rigorous differentiation between data entry corruption (0 errors) and legitimate commercial transactions.

---

## 2. Dataset Profile & Quality Summary

* **Dataset Source:** [Kaggle Superstore Dataset](https://www.kaggle.com/datasets/vivek468/superstore-dataset-final)
* **Raw File Location:** `data/raw/Superstore.csv`
* **Cleaned File Location:** `data/processed/superstore_clean.csv` & `superstore_clean.rds`
* **Observational Unit:** Individual line-item SKU within a commercial customer order.
* **Row Count:** 9,994 transactions.
* **Column Count:** 21 raw columns (expanded to 28 columns with engineered features).
* **Date Horizon:** January 4, 2011 to December 31, 2014 (48 continuous calendar months).
* **Missing Values:** Exactly 0 missing values (100% complete across all 21 fields).
* **Exact Duplicate Rows:** Exactly 0 duplicate rows.
* **Unique Orders:** 5,009 unique Order IDs representing 9,994 line items.
* **Date Parsing Resolution:** The raw data encodes dates strictly in `DD-MM-YYYY` format. Enforcing proper parsing via `lubridate::dmy()` eliminates all chronological anomalies, confirming 100% logical integrity (`Ship Date >= Order Date`).

---

## 3. Directory Structure

```text
Week2_Superstore_R_Visualization/
├── data/
│   ├── raw/
│   │   └── Superstore.csv                    # Immutable raw Kaggle dataset (9,994 rows)
│   └── processed/
│       ├── superstore_clean.csv              # Processed analytical dataset (28 columns)
│       └── superstore_clean.rds              # Serialized native R binary dataset
├── R/
│   ├── 01_setup.R                            # Package loading, environment options, ggplot2 theme
│   ├── 02_import_data.R                      # Ingestion, schema validation, diagnostic summary
│   ├── 03_data_quality.R                     # Completeness scan, deduplication, date integrity audit
│   ├── 04_data_cleaning.R                    # Feature engineering, date parsing, margin calculation
│   ├── 05_descriptive_analysis.R             # Methods 1-5: stats, frequencies, grouped summaries, correlations
│   ├── 06_visualizations.R                   # Generation of 11 publication-grade charts (300 DPI)
│   ├── 07_outlier_analysis.R                 # Method 6: IQR boundaries, extreme transaction forensics
│   └── 08_export_results.R                   # Master pipeline orchestration and asset validation
├── visualizations/
│   ├── 01_sales_by_category.png              # Horizontal bar chart of gross revenue by category
│   ├── 02_profit_by_category.png             # Comparative bar chart of net profit and margin %
│   ├── 03_sales_distribution.png             # Log10 histogram showing positive skew and median
│   ├── 04_profit_distribution.png            # Diverging histogram centered on $0 breakeven
│   ├── 05_monthly_sales_trend.png            # Chronological 48-month trend with LOESS curve
│   ├── 06_discount_vs_profit.png             # Bivariate scatter showing the 20% discount cliff
│   ├── 07_sales_vs_profit.png                # Bivariate scatter displaying widening risk dispersion
│   ├── 08_profit_by_subcategory.png          # Diverging horizontal bar chart across 17 sub-categories
│   ├── 09_profit_boxplot_category.png        # Non-parametric box plots of profit distributions
│   ├── 10_regional_performance.png           # Dual-metric grouped bar chart across US regions
│   └── 11_shipping_time_by_mode.png          # Fulfillment duration box plot across shipping tiers
├── outputs/
│   ├── tables/
│   │   ├── table1_dataset_structure.csv      # High-level dataset metadata and metrics
│   │   ├── data_quality_summary.csv          # Audit table for all 21 raw variables
│   │   ├── table2_numerical_summary.csv      # Parametric & non-parametric summary statistics
│   │   ├── table3_sales_profit_by_category.csv # Aggregated metrics by product category
│   │   ├── table4_regional_performance.csv   # Aggregated metrics by US geographic region
│   │   ├── table5_segment_performance.csv    # Aggregated metrics by customer segment
│   │   ├── table6_subcategory_profitability.csv # Complete profitability ranking for 17 sub-categories
│   │   ├── table7_outliers_summary.csv       # IQR outlier limits and anomaly counts
│   │   ├── correlation_matrix_pearson.csv    # Linear correlation coefficients
│   │   ├── correlation_matrix_spearman.csv   # Monotonic rank correlation coefficients
│   │   ├── monthly_trend_summary.csv         # 48-month chronological revenue and profit metrics
│   │   ├── yearly_trend_summary.csv          # 4-year annual growth and margin summary
│   │   ├── shipping_mode_summary.csv         # Logistics SLA adherence statistics
│   │   └── top_extreme_transactions.csv      # Forensic audit of top sales, profits, and losses
│   └── summaries/
│       └── project_execution_summary.txt     # Complete pipeline validation and runtime log
├── report/
│   └── Week2_Superstore_Data_Visualization_Report.docx # Comprehensive 26-section Word report
├── generate_doc_report.py                    # Programmatic DOCX generator with styling & figures
├── Week2_Superstore_R_Visualization.Rproj     # RStudio project configuration file
└── README.md                                 # Project documentation
```

---

## 4. Key Analytical Findings

1. **Volume vs. Margin Asymmetry:** Technology accounts for 36.4% of enterprise sales ($836.2K) and 50.8% of profit ($145.5K) at a 17.4% margin. Conversely, Furniture captures 32.3% of enterprise sales ($742.0K) but delivers only 6.4% of profit ($18.5K) at a 2.49% margin.
2. **The "Furniture Deficit" Identified:** Sub-category decomposition reveals that Chairs (+$26.6K) and Furnishings (+$13.1K) are profitable, but their earnings are wiped out by massive losses in Tables (-$17,725.48 net deficit) and Bookcases (-$3,472.56 net deficit).
3. **The 20% Discount Cliff:** While transactions discounted between 0% and 20% are consistently profitable, discounts exceeding 20% trigger immediate margin collapse (Spearman rank correlation $r_s = -0.5434$).
4. **Regional Margin Disparities:** The West region generates $725.5K in sales and $108.4K in profit (14.94% margin) with disciplined discounting (10.93% average discount). The Central region suffers severe margin erosion, earning only $39.7K profit on $501.2K sales (7.92% margin) due to aggressive discounting (24.03% average discount).
5. **High-Value Deal Volatility:** The variance of profit expands dramatically with sales revenue. Transactions over $5,000 diverge into extreme profits (+$8,400 in Copiers at 0% discount) or catastrophic commercial deficits (-$6,600 in 3D Printers at 70% discount).
6. **Logistics SLA Compliance:** Fulfillment operations adhere strictly to service-level agreements: Same Day fulfills in 0.04 days on average (max 1 day), while Standard Class fulfills in 5.01 days on average (max 7 days).

---

## 5. Execution & Reproducibility Instructions

### Prerequisites
* R (>= 4.3.0, verified on R 4.6.1 Windows ucrt)
* Python (>= 3.9, with `python-docx` installed for DOCX report compilation)

### Automated Pipeline Execution
To execute the complete R analytical workflow from beginning to end:
```bash
# In PowerShell or Terminal within the project directory:
Rscript R/08_export_results.R
```
This single command sequentially sources:
1. `R/01_setup.R` – Verifies directories and loads required libraries.
2. `R/02_import_data.R` – Ingests raw data and validates schema.
3. `R/03_data_quality.R` – Audits completeness, duplicates, and date format.
4. `R/04_data_cleaning.R` – Engineers derived features and serializes clean data.
5. `R/05_descriptive_analysis.R` – Generates all descriptive tables and correlation matrices.
6. `R/06_visualizations.R` – Exports all 11 charts at 300 DPI (10 × 6 inches).
7. `R/07_outlier_analysis.R` – Executes non-parametric IQR outlier checks and record forensics.

### Compiling the Word Report
To rebuild the Microsoft Word DOCX document with all embedded figures and tables:
```bash
python generate_doc_report.py
```
Output document generated: `report/Week2_Superstore_Data_Visualization_Report.docx` (~1.33 MB).
