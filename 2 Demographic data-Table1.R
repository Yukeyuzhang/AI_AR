library(readxl)
library(dplyr)

# Read data
file_path <- "E:/Train.xlsx"
file_path <- "E:/Test.xlsx"
df <- read_excel(file_path)

# Numeric columns to be analyzed
columns_to_test <- c("Age", 
                     "Nasal itching", "Nasal congestion", "Sneezing", 
                     "Clear rhinorrhea", "Ocular symptoms", "Overall discomfort", 
                     "Total VAS score")

# Force conversion of these columns to numeric type
df <- df %>%
  mutate(across(all_of(columns_to_test), as.numeric))

# Filter for Diagnosis values 0 (NAR) and 1 (AR)
df_filtered <- df %>% filter(Diagnosis %in% c(0, 1))

# Calculate median (25th%, 75th%)
describe_stat <- function(x) {
  x <- na.omit(x)  # Remove NA values
  if (length(x) == 0) return("NA (NA, NA)")  # Handle empty data
  q <- quantile(x, probs = c(0.25, 0.5, 0.75), na.rm = TRUE)
  return(sprintf("%.1f (%.1f, %.1f)", q[2], q[1], q[3]))  # Keep 1 decimal place
}

# Initialize results table
results <- data.frame(Variable = columns_to_test, Median_IQR_AR = NA, Median_IQR_NAR = NA, P_value = NA)

# Calculate statistical descriptions and P values
for (col in columns_to_test) {
  AR_values <- df_filtered %>% filter(Diagnosis == 1) %>% pull(!!sym(col)) %>% na.omit()
  NAR_values <- df_filtered %>% filter(Diagnosis == 0) %>% pull(!!sym(col)) %>% na.omit()
  
  if (length(AR_values) > 0 && length(NAR_values) > 0) {
    results$Median_IQR_AR[results$Variable == col] <- describe_stat(AR_values)
    results$Median_IQR_NAR[results$Variable == col] <- describe_stat(NAR_values)
    
    # Calculate P value using Mann-Whitney U test (Wilcoxon test)
    test_result <- wilcox.test(AR_values, NAR_values, exact = FALSE)
    results$P_value[results$Variable == col] <- sprintf("%.4f", test_result$p.value)
  } else {
    results$Median_IQR_AR[results$Variable == col] <- "NA (NA, NA)"
    results$Median_IQR_NAR[results$Variable == col] <- "NA (NA, NA)"
    results$P_value[results$Variable == col] <- "NA"
  }
}

# Output results
print(results)

# Save results as CSV file
write.csv(results, "E:/Table1.csv", row.names = FALSE)