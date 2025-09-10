library(readxl)
library(dplyr)
library(writexl)

prepare_data <- function(file_path, sheet_number = 2, noise_rows = 2:10) {
  
  # 1. Read the entire sheet without headers
  df_full <- read_excel(file_path, sheet = sheet_number, col_names = FALSE)
  
  # 2. Use row 1 as column names
  colnames(df_full) <- as.character(df_full[1, ])
  
  # 3. Remove row 1 (header) + noise rows 2–10
  df <- df_full[-c(1, noise_rows), ]
  
  # 4. Reset row numbers
  rownames(df) <- NULL
  
  # 5. Rename first column to "Date"
  colnames(df)[1] <- "Date"
  
  # 6. Convert Excel numeric dates to proper Date
  df$Date <- as.Date(as.numeric(df$Date), origin = "1899-12-30")
  
  return(df)
}

# -----------------------------
# Step 1: Load and prepare all three datasets
# -----------------------------
state_data <- prepare_data(file.choose())      # choose state file
duration_data <- prepare_data(file.choose())   # choose duration file
country_data <- prepare_data(file.choose())    # choose country file

merged_data <- Reduce(function(x, y) merge(x, y, by = "Date", all = TRUE),
                      list(state_data, duration_data, country_data))

write_xlsx(merged_data, "merged_output_shorttermarrivals.xlsx")
