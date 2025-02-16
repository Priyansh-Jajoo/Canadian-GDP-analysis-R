rm(list = ls())
library(openxlsx)
library(ggplot2)
library(dplyr)
library(tidyr)
library(forecast)
library(e1071)
library(randomForest)
library(scales)
library(broom)

# Read data
data <- read.xlsx("Canada_Real_GDP.xlsx", sheet = 1)

# Ensure required columns exist
required_columns <- c("date", "Value", "consumption_expenditure", "unemployment", "Inflation")
if (!all(required_columns %in% colnames(data))) {
  stop("Missing required columns in the dataset. Please check the input file.")
}

# 1. Data Preprocessing ----
# Convert date to proper format
data$date <- as.Date(as.numeric(data$date), origin = "1961-01-01")

# Calculate growth rates
data <- data %>%
  mutate(
    gdp_growth = (Value - lag(Value)) / lag(Value) * 100,
    consumption_growth = (consumption_expenditure - lag(consumption_expenditure)) / lag(consumption_expenditure) * 100
  ) %>%
  filter(!is.na(gdp_growth) & !is.na(consumption_growth))  # Remove rows with NA growth rates

# 2. Visualization Section ----
# GDP and Consumption Over Time
p1 <- ggplot(data, aes(x = date)) + 
  geom_line(aes(y = Value, color = "GDP")) + 
  geom_line(aes(y = consumption_expenditure, color = "Consumption")) +
  labs(x = "Date", y = "Millions of Chained 2012 CAD",
       title = "Real GDP and Consumption Trends",
       subtitle = "Canada (1960-2023)") +
  theme_minimal() +
  scale_color_manual(name = "Series", values = c("GDP" = "blue", "Consumption" = "red")) +
  scale_x_date(date_breaks = "5 years", date_labels = "%Y") +
  theme(axis.text.x = element_text(angle = 45))

# Unemployment and Inflation Over Time
p2 <- ggplot(data, aes(x = date)) + 
  geom_line(aes(y = unemployment, color = "Unemployment")) +
  geom_line(aes(y = Inflation, color = "Inflation")) +
  labs(x = "Date", y = "Rate (%)",
       title = "Unemployment and Inflation Trends") +
  theme_minimal() +
  scale_color_manual(name = "Series", values = c("Unemployment" = "purple", "Inflation" = "orange")) +
  scale_x_date(date_breaks = "5 years", date_labels = "%Y") +
  theme(axis.text.x = element_text(angle = 45))

# GDP vs Consumption Scatter
p3 <- ggplot(data, aes(x = Value, y = consumption_expenditure)) +
  geom_point(alpha = 0.5) +
  geom_smooth(method = "lm", se = TRUE) +
  labs(x = "Real GDP", y = "Real Consumption",
       title = "Consumption vs GDP Relationship",
       subtitle = "With 95% Confidence Interval") +
  theme_minimal()

# Phillips Curve (Unemployment vs Inflation)
p4 <- ggplot(data, aes(x = unemployment, y = Inflation)) +
  geom_point(alpha = 0.5) +
  geom_smooth(method = "loess", se = TRUE) +
  labs(x = "Unemployment Rate (%)", y = "Inflation Rate (%)",
       title = "Phillips Curve Analysis") +
  theme_minimal()

# 3. Statistical Analysis ----
# Descriptive statistics for all variables
stats_summary <- data %>%
  summarise(
    gdp_mean = mean(Value, na.rm = TRUE),
    gdp_sd = sd(Value, na.rm = TRUE),
    gdp_skew = skewness(Value, na.rm = TRUE),
    cons_mean = mean(consumption_expenditure, na.rm = TRUE),
    cons_sd = sd(consumption_expenditure, na.rm = TRUE),
    cons_skew = skewness(consumption_expenditure, na.rm = TRUE),
    unemp_mean = mean(unemployment, na.rm = TRUE),
    unemp_sd = sd(unemployment, na.rm = TRUE),
    inf_mean = mean(Inflation, na.rm = TRUE),
    inf_sd = sd(Inflation, na.rm = TRUE)
  )

# 4. Regression Analysis ----
# Basic consumption function
model1 <- lm(consumption_expenditure ~ Value, data = data)

# Multiple regression with all variables
model2 <- lm(consumption_expenditure ~ Value + unemployment + Inflation, data = data)

# Model summaries
summary(model1)
summary(model2)

# 5. Time Series Analysis ----
# GDP forecast
gdp_ts <- ts(data$Value, start = c(1960, 1), frequency = 1)
gdp_arima <- auto.arima(gdp_ts)
gdp_forecast <- forecast(gdp_arima, h = 5)  # 5-year forecast

# Consumption forecast
cons_ts <- ts(data$consumption_expenditure, start = c(1960, 1), frequency = 1)
cons_arima <- auto.arima(cons_ts)
cons_forecast <- forecast(cons_arima, h = 5)

# 6. Machine Learning ----
# Prepare data for Random Forest
rf_data <- data %>% 
  select(Value, consumption_expenditure, unemployment, Inflation) %>%
  na.omit()

# Random Forest model for consumption
rf_model <- randomForest(
  consumption_expenditure ~ ., 
  data = rf_data,
  ntree = 500,
  importance = TRUE
)

# Variable importance
importance(rf_model)

# 7. Economic Analysis ----
# Calculate marginal propensity to consume
mpc <- coef(model1)[2]

# Calculate Okun's Law coefficient (relationship between GDP growth and unemployment)
okun_model <- lm(unemployment ~ gdp_growth, data = data)
okun_coef <- coef(okun_model)[2]

# 8. Results Export ----
# Create results workbook
wb <- createWorkbook()

# Add sheets
addWorksheet(wb, "Model_Summary")
addWorksheet(wb, "Forecasts")
addWorksheet(wb, "Statistics")

# Write results
writeData(wb, "Model_Summary", tidy(model2))
writeData(wb, "Forecasts", as.data.frame(gdp_forecast))
writeData(wb, "Statistics", as.data.frame(stats_summary))

# Save workbook
saveWorkbook(wb, "economic_analysis_results.xlsx", overwrite = TRUE)

# 9. Plot Arrangement
install.packages("gridExtra")  
library(gridExtra)
library(ggplot2)
grid.arrange(p1, p2, p3, p4, ncol = 2)

# Print key findings
cat("\nKey Findings:\n")
cat("1. Marginal Propensity to Consume:", round(mpc, 3), "\n")
cat("2. Okun's Law Coefficient:", round(okun_coef, 3), "\n")
cat("3. R-squared (Multiple Regression):", round(summary(model2)$r.squared, 3), "\n")
cat("4. Most Important Predictor:", names(which.max(importance(rf_model)[, 1])), "\n")



