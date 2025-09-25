# Load necessary R packages
library(readxl)
library(dplyr)
library(openxlsx)  # For saving Excel files

# Read data
file_path <- "E:/Data.xlsx"
df <- read_excel(file_path)

# Ensure Diagnosis is numeric
df <- df %>% mutate(Diagnosis = as.numeric(Diagnosis))

# Stratified division by Diagnosis
set.seed(6)  # Set random seed for reproducibility
train_index_0 <- sample(which(df$Diagnosis == 0), size = floor(sum(df$Diagnosis == 0) * 0.8))
train_index_1 <- sample(which(df$Diagnosis == 1), size = floor(sum(df$Diagnosis == 1) * 0.8))

# Get training and test set data
train_indices <- c(train_index_0, train_index_1)
train_set <- df[train_indices, ]
test_set <- df[-train_indices, ]

# Save training and test sets as Excel files
train_file <- "E:/Train.xlsx"
test_file <- "E:/Test.xlsx"

write.xlsx(train_set, train_file, rowNames = FALSE)
write.xlsx(test_set, test_file, rowNames = FALSE)

cat("Train:", train_file, "\n")
cat("Test:", test_file, "\n")