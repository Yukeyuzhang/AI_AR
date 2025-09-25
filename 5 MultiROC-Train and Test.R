library(readxl)
library(MASS)
library(pROC)
library(ggplot2)
library(openxlsx)
library(ROCR)

# Read training set data
train_file_path <- "E:/Train.xlsx"
train_data <- read_excel(train_file_path)
train_data$Diagnosis <- as.factor(train_data$Diagnosis)  # Convert Diagnosis to factor type for classification

# Extract independent variables (columns 17-22) and dependent variable from training set
X_train <- train_data[, 17:22]
# Convert all independent variables to numeric type (handle potential character-type numeric data)
X_train <- data.frame(lapply(X_train, function(x) as.numeric(as.character(x))))
# Combine independent variables and dependent variable into a single dataframe for modeling
train_data_model <- data.frame(X_train, Diagnosis = train_data$Diagnosis)

# Calculate class weights (to handle class imbalance)
weight_0 <- sum(train_data_model$Diagnosis == 1) / nrow(train_data_model)
weight_1 <- sum(train_data_model$Diagnosis == 0) / nrow(train_data_model)
# Assign weights: samples of class 1 get weight_0, samples of class 0 get weight_1
weights <- ifelse(train_data_model$Diagnosis == 1, weight_0, weight_1)

# Train logistic regression model
full_model <- glm(Diagnosis ~ ., data = train_data_model, 
                  family = binomial(link = "logit"),  # Binomial family for binary classification, logit link function
                  weights = weights)  # Apply calculated class weights
full_model$coefficients  # View model coefficients

# Predict probabilities for training set (outputs probability of belonging to class 1)
train_probabilities <- predict(full_model, type = "response")
# Combine actual Diagnosis labels and predicted probabilities into a dataframe
train_results <- data.frame(Diagnosis = train_data_model$Diagnosis, Probabilities = train_probabilities)
# Save training set prediction results to Excel file
write.xlsx(train_results, "E:/AI-AR/train_probabilities.xlsx", rowNames = FALSE)

# Calculate ROC curve for training set and generate 95% confidence interval (CI) for AUC
roc_train <- roc(train_results$Diagnosis, train_results$Probabilities)
auc_train <- auc(roc_train)  # Calculate AUC (Area Under ROC Curve) for training set
ci_train <- ci.auc(roc_train)  # Compute 95% CI for training set AUC

# Read test set data
test_file_path <- "E:/Test.xlsx"
test_data <- read_excel(test_file_path)
test_data$Diagnosis <- as.factor(test_data$Diagnosis)  # Convert Diagnosis to factor type

# Extract independent variables from test set (same columns as training set: 17-22)
X_test <- test_data[, 17:22]
# Convert test set independent variables to numeric type (consistent with training set processing)
X_test <- data.frame(lapply(X_test, function(x) as.numeric(as.character(x))))

# Predict probabilities for test set using the trained logistic regression model
test_probabilities <- predict(full_model, newdata = X_test, type = "response")
# Combine actual Diagnosis labels and predicted probabilities for test set
test_results <- data.frame(Diagnosis = test_data$Diagnosis, Probabilities = test_probabilities)
# Save test set prediction results to Excel file
write.xlsx(test_results, "E:/AI-AR/test_probabilities.xlsx", rowNames = FALSE)

# Calculate ROC curve for test set and generate 95% confidence interval for AUC
roc_test <- roc(test_results$Diagnosis, test_results$Probabilities)
auc_test <- auc(roc_test)  # Calculate test set AUC
ci_test <- ci.auc(roc_test)  # Compute 95% CI for test set AUC

# Plot ROC curves (training and test sets) and save as PDF
pdf("E:/AI-AR/ROC_Curve_Train_Test.pdf", width = 7.5, height = 7)  # Set PDF size
# Plot training set ROC curve (red, thick line)
plot(roc_train, col = "red", lwd = 2.5, main = "ROC Curve (Train & Test)")
# Add training set AUC and 95% CI as text annotation
text(0.8, 0.3, labels = paste0("Train AUC = ", round(auc_train, 3), " (", round(ci_train[1], 3), "-", round(ci_train[3], 3), ")"), col = "red", cex = 1.2)
# Add test set ROC curve (blue, thick line) to the same plot
plot(roc_test, col = "blue", lwd = 2.5, add = TRUE)
# Add test set AUC and 95% CI as text annotation
text(0.8, 0.2, labels = paste0("Test AUC = ", round(auc_test, 3), " (", round(ci_test[1], 3), "-", round(ci_test[3], 3), ")"), col = "blue", cex = 1.2)
# Add diagonal reference line (representing random guessing, gray dashed line)
abline(0, 1, col = "gray", lty = 4)
dev.off()  # Close PDF device to save the plot

# Calculate optimal cutoff threshold for training set (using Youden's J statistic)
best_cutoff_train <- coords(roc_train, "best", ret = "threshold", best.method = "youden")
print(paste("Train Cutoff:", round(best_cutoff_train, 3)))  # Print optimal cutoff (keep original Chinese label for clarity if needed)