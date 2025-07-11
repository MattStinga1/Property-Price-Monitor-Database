# Load necessary libraries
library(readxl)
library(dplyr)
library(stringr)

# 🧹 Clear environment, plots, and console
rm(list = ls())
if (dev.cur() > 1) dev.off()
cat("\014")  # Clear console (only works in RStudio)

# 📂 File path to your Excel workbook
file_path <- "T:/Home/act_dpt/COMMON/Property/Price_Monitoring/2025 Accounts/2025_06_Property Rate Monitor_New .xlsx"

# 🔍 Pattern to identify valid sheet names (e.g., "Jun 2025", "Jun 2025 - BI PD")
valid_sheet_regex <- "^(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec|January|February|March|April|May|June|July|August|September|October|November|December)\\s\\d{4}(\\s-\\s(BI PD|Exposure Curve))?$"

# 📋 Get all sheet names and filter valid ones
all_sheets <- excel_sheets(file_path)
valid_sheets <- all_sheets[str_detect(all_sheets, valid_sheet_regex)]

# 📥 Read each valid sheet, skipping first 2 rows, and convert all columns to character
all_data <- lapply(valid_sheets, function(sheet) {
  read_excel(file_path, sheet = sheet, skip = 2) %>%
    mutate(across(everything(), as.character)) %>%
    mutate(source_sheet = sheet)  # Optional: track source sheet
})

# 🔗 Combine all data
combined_data <- bind_rows(all_data)

# 👀 View the result
View(combined_data)
