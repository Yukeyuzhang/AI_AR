library(readxl)
library(MASS)
library(pROC)
library(ggplot2)
library(openxlsx)
library(ROCR)

# Read training set data
train_file_path <- "E:/Train.xlsx"
train_data <- read_excel(train_file_path)
train_data$Diagnosis <- as.factor(train_data$Diagnosis)  # Convert Diagnosis to factor for binary classification

# Extract independent variables (columns 17-23) and dependent variable from training set
X_train <- train_data[, 17:23]
# Convert all independent variables to numeric (handle potential character-formatted numeric data)
X_train <- data.frame(lapply(X_train, function(x) as.numeric(as.character(x))))
# Combine independent variables and dependent variable into a single dataframe for modeling
train_data_model <- data.frame(X_train, Diagnosis = train_data$Diagnosis)

# Calculate class weights (to address class imbalance in training data)
weight_0 <- sum(train_data_model$Diagnosis == 1) / nrow(train_data_model)
weight_1 <- sum(train_data_model$Diagnosis == 0) / nrow(train_data_model)
# Assign weights: samples with Diagnosis=1 get weight_0, samples with Diagnosis=0 get weight_1
weights <- ifelse(train_data_model$Diagnosis == 1, weight_0, weight_1)

# Initialize objects to store results for each independent variable (7 variables total)
colors <- c("#8C7D37", "#5757F9", "#DE8BF9", "#FF6000", "#05BE78", "#FF0000", "#000000") # 7 high-contrast colors
roc_train_list <- list()       # Store ROC objects for training set
auc_train_list <- list()       # Store AUC values for training set
cutoff_train_list <- list()    # Store optimal cutoff values for training set
coefficients_list <- list()    # Store model coefficients (slope and intercept)

# Train a separate logistic regression model for each independent variable
for (i in 1:7) {
  # Create formula: Diagnosis ~ current independent variable
  formula <- as.formula(paste("Diagnosis ~", colnames(X_train)[i]))
  # Train logistic regression model with class weights
  model <- glm(formula, data = train_data_model, family = binomial(link = "logit"), weights = weights)
  
  # Extract model coefficients (slope 'k' and intercept 'b')
  coefficients_list[[i]] <- coef(model)
  k <- coefficients_list[[i]][2]  # Slope: coefficient of the current independent variable
  b <- coefficients_list[[i]][1]  # Intercept: constant term of the model
  
  # Predict probabilities for the training set (probability of Diagnosis=1)
  train_probabilities <- predict(model, type = "response")
  # Combine actual Diagnosis labels and predicted probabilities
  train_results <- data.frame(Diagnosis = train_data_model$Diagnosis, Probabilities = train_probabilities)
  
  # Calculate ROC curve and optimal cutoff for the training set
  roc_train <- roc(train_results$Diagnosis, train_results$Probabilities)
  auc_train <- auc(roc_train)  # Compute AUC (Area Under ROC Curve)
  # Find optimal cutoff using Youden's J statistic (maximizes sensitivity + specificity - 1)
  best_cutoff <- coords(roc_train, "best", best.method = "youden", ret = c("threshold"))
  # Convert probability cutoff back to the original value of the independent variable
  # Derived from logit model: log(p/(1-p)) = b + k*x → solve for x when p = best_cutoff
  x_value <- (log(best_cutoff / (1 - best_cutoff)) - b) / k
  print(x_value)  # Print the optimal cutoff in original variable units
  
  # Store results in corresponding lists
  roc_train_list[[i]] <- roc_train
  auc_train_list[[i]] <- auc_train
  cutoff_train_list[[i]] <- x_value
}

# Plot and save ROC curves for all training set models (as PDF)
pdf("E:/AI-AR/ROC_Curve_Train.pdf", width = 7.5, height = 7)  # Set PDF dimensions
# Plot ROC curve for the first independent variable (initialize plot)
plot(roc_train_list[[1]], col = colors[1], lwd = 2.5, 
     main = "ROC Curve (Train)", 
     xlab = "False Positive Rate",  # X-axis: 1 - Specificity
     ylab = "True Positive Rate")   # Y-axis: Sensitivity
# Add ROC curves for the remaining 6 independent variables
for (i in 2:7) {
  plot(roc_train_list[[i]], col = colors[i], lwd = 2.5, add = TRUE)
}

# Adjust text positions to avoid overlap in annotations
y_offset <- 0.35  # Initial Y-position for the first annotation (higher to avoid bottom crowding)
y_gap <- 0.05    # Vertical gap between consecutive annotations
max_text_position <- 0.05  # Minimum Y-position to prevent text from going outside the plot

# Add annotations (AUC + 95% CI + optimal cutoff) for each model
for (i in 1:7) {
  # Calculate 95% confidence interval for AUC
  auc_ci <- ci.auc(roc_train_list[[i]])
  # Format AUC text (variable name + AUC + 95% CI)
  auc_text <- paste0(colnames(X_train)[i], " AUC = ", 
                     round(auc_train_list[[i]], 3), 
                     " (", round(auc_ci[1], 3), "-", round(auc_ci[3], 3), ")")
  # Format optimal cutoff text (in original variable units)
  cutoff_text <- paste0("Cutoff = ", round(cutoff_train_list[[i]], 3))
  
  # Calculate Y-position for current annotation
  text_position <- y_offset - (i - 1) * y_gap
  # Ensure text stays above the minimum Y-position
  if (text_position < max_text_position) {
    text_position <- max_text_position
  }
  
  # Add combined annotation to the plot (right-aligned at X=0.8)
  text(0.8, text_position, labels = paste(auc_text, cutoff_text), 
       col = colors[i], cex = 1.2, adj = 0)
}

# Add diagonal reference line (represents random guessing, gray dashed line)
abline(0, 1, col = "gray", lty = 2)
dev.off()  # Close PDF device to save the plot

# Read test set data
test_file_path <- "E:/Test.xlsx"
test_data <- read_excel(test_file_path)
test_data$Diagnosis <- as.factor(test_data$Diagnosis)  # Convert Diagnosis to factor

# Extract and preprocess independent variables from test set (match training set columns 17-23)
X_test <- test_data[, 17:23]
X_test <- data.frame(lapply(X_test, function(x) as.numeric(as.character(x))))  # Convert to numeric

# Initialize objects to store test set results
roc_test_list <- list()        # Store ROC objects for test set
auc_test_list <- list()        # Store AUC values for test set
ci_auc_test_list <- list()     # Store 95% CI for test set AUC

# Evaluate each trained model on the test set
for (i in 1:7) {
  # Recreate the formula for the current independent variable (consistent with training)
  formula <- as.formula(paste("Diagnosis ~", colnames(X_train)[i]))
  # Retrain the model (or use the previously trained one; retraining ensures consistency here)
  model <- glm(formula, data = train_data_model, family = binomial(link = "logit"), weights = weights)
  
  # Predict probabilities for the test set using the trained model
  test_probabilities <- predict(model, newdata = X_test, type = "response")
  
  # Calculate ROC curve, AUC, and 95% CI for the test set
  roc_test <- roc(test_data$Diagnosis, test_probabilities)
  auc_test <- auc(roc_test)
  ci_auc_test <- ci.auc(roc_test)  # Compute 95% CI for test set AUC
  
  # Store test set results
  roc_test_list[[i]] <- roc_test
  auc_test_list[[i]] <- auc_test
  ci_auc_test_list[[i]] <- ci_auc_test
}

# Plot and save ROC curves for all test set models (as PDF)
pdf("E:/AI-AR/ROC_Curve_Test.pdf", width = 7.5, height = 7)  # Set PDF dimensions
# Plot ROC curve for the first independent variable (initialize plot)
plot(roc_test_list[[1]], col = colors[1], lwd = 2.5, 
     main = "ROC Curve (Test)", 
     xlab = "False Positive Rate", 
     ylab = "True Positive Rate")
# Add ROC curves for the remaining 6 independent variables
for (i in 2:7) {
  plot(roc_test_list[[i]], col = colors[i], lwd = 2.5, add = TRUE)
}

# Adjust text positions (same logic as training set plot)
y_offset <- 0.35  
y_gap <- 0.05    
max_text_position <- 0.05  

# Add annotations (AUC + 95% CI) for each test set model
for (i in 1:7) {
  # Format AUC text with 95% CI
  auc_text <- paste0(colnames(X_test)[i], " AUC = ", 
                     round(auc_test_list[[i]], 3), 
                     " (", round(ci_auc_test_list[[i]][1], 3), "-", 
                     round(ci_auc_test_list[[i]][3], 3), ")")
  
  # Calculate Y-position and adjust if needed
  text_position <- y_offset - (i - 1) * y_gap
  if (text_position < max_text_position) {
    text_position <- max_text_position
  }
  
  # Add annotation to the plot
  text(0.8, text_position, labels = auc_text, col = colors[i], cex = 1.2, adj = 0)
}

# Add diagonal reference line for random guessing
abline(0, 1, col = "gray", lty = 2)
dev.off()  # Close PDF device to save the plot