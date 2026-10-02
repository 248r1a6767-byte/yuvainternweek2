# ==============================================================================
# Script: scripts/run_all.R
# Purpose: Master orchestration runner for Week 2 Data Visualization Internship Project.
#          Executes the complete data import, cleaning, descriptive analysis,
#          visualization generation (14 figures at 300 DPI), outlier forensics,
#          and automated asset validation.
# Author: Data Analyst & Visualization Specialist
# Date: 2026-09-28 (Updated 2026-10-02)
# ==============================================================================

cat("\n")
cat("==============================================================================\n")
cat("          STARTING FULL WEEK 2 R VISUALIZATION PIPELINE RUNNER               \n")
cat("==============================================================================\n\n")

start_pipeline_time <- Sys.time()

# Ensure working directory is project root
if (!dir.exists("R")) {
  if (dir.exists("../R")) {
    setwd("..")
  } else {
    stop("Please run this script from the Week2_Superstore_R_Visualization project root.")
  }
}

# Sequential execution
pipeline_scripts <- c(
  "R/01_setup.R",
  "R/02_import_data.R",
  "R/03_data_quality.R",
  "R/04_data_cleaning.R",
  "R/05_descriptive_analysis.R",
  "R/06_visualizations.R",
  "R/07_outlier_analysis.R",
  "R/08_export_results.R"
)

for (s in pipeline_scripts) {
  cat(sprintf("[PIPELINE STEP] Sourcing: %s ...\n", s))
  step_start <- Sys.time()
  source(s, echo = FALSE)
  step_elapsed <- round(as.numeric(difftime(Sys.time(), step_start, units = "secs")), 2)
  cat(sprintf("[PIPELINE STEP] Completed %s in %.2f seconds.\n\n", s, step_elapsed))
}

total_elapsed <- round(as.numeric(difftime(Sys.time(), start_pipeline_time, units = "secs")), 2)
cat("==============================================================================\n")
cat(sprintf("   MASTER PIPELINE COMPLETED SUCCESSFULLY IN %.2f SECONDS              \n", total_elapsed))
cat("==============================================================================\n\n")
