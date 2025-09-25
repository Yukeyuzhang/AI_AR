library(readxl)
library(dplyr)
library(tidyr)
library(writexl)

# Read data
file_path <- "E:/Train.xlsx"
df <- read_excel(file_path)

# Select allergen data
allergen_cols <- 2:13  # Columns containing allergen information
symptom_cols <- 17:21  # Columns containing symptom information
discomfort_col <- 22   # Column containing discomfort information
total_score_col <- 23  # Column containing total score information

# 1. Analysis of main allergen types
no_allergy_count <- df %>%
  filter(rowSums(select(., all_of(allergen_cols))) == 0) %>%
  nrow()

total_patients <- nrow(df)
no_allergy_ratio <- no_allergy_count / total_patients

allergen_distribution <- df %>%
  select(all_of(allergen_cols)) %>%
  pivot_longer(cols = everything(), names_to = "Allergen", values_to = "Severity") %>%
  group_by(Allergen, Severity) %>%
  summarise(Count = n(), .groups = "drop") %>%
  mutate(Percentage = Count / total_patients * 100)

# 2. Analysis of multiple allergy cases
multi_allergy_count <- df %>%
  filter(rowSums(select(., all_of(allergen_cols)) > 0) >= 2) %>%
  nrow()

multi_allergy_ratio <- multi_allergy_count / total_patients

multi_allergy_severity <- df %>%
  filter(rowSums(select(., all_of(allergen_cols)) > 0) >= 2) %>%
  select(all_of(allergen_cols)) %>%
  pivot_longer(cols = everything(), names_to = "Allergen", values_to = "Severity") %>%
  group_by(Severity) %>%
  summarise(Count = n(), .groups = "drop") %>%
  mutate(Percentage = Count / sum(Count) * 100)

# 3. Distribution of single allergen severity
single_allergy_df <- df %>%
  rowwise() %>%
  mutate(Single_Allergen = sum(c_across(all_of(allergen_cols)) > 0) == 1) %>%
  filter(Single_Allergen) %>%
  select(all_of(allergen_cols))

single_allergy_distribution <- single_allergy_df %>%
  pivot_longer(cols = everything(), names_to = "Allergen", values_to = "Severity") %>%
  filter(Severity > 0) %>%
  group_by(Allergen, Severity) %>%
  summarise(Count = n(), .groups = "drop") %>%
  mutate(Percentage = Count / sum(Count) * 100)

# Store results in a list
results <- list(
  "No_Allergy_Ratio" = data.frame(Metric = "No Allergy Ratio", Value = no_allergy_ratio),
  "Allergen_Distribution" = allergen_distribution,
  "Multi_Allergy_Ratio" = data.frame(Metric = "Multi Allergy Ratio", Value = multi_allergy_ratio),
  "Multi_Allergy_Severity" = multi_allergy_severity,
  "Single_Allergy_Distribution" = single_allergy_distribution
)

# Export to Excel
output_path <- "E:/Allergy_corr.xlsx"
write_xlsx(results, output_path)