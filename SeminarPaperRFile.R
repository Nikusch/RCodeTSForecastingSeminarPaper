# --- LOAD LIBRARIES ---
## Core libraries for data manipulation and visualization
library(tidyverse)  # Comprehensive package including dplyr, ggplot2, and more
library(dplyr)      # Data manipulation (explicitly loaded for standalone use)
library(ggplot2)    # Visualization

## Libraries for data import
library(readxl)     # Reading Excel files

## Libraries for time series analysis
library(forecast)   # Forecasting models and tools
library(tseries)    # Time series and statistical tests
library(zoo)        # Working with irregular time series
library(lubridate)  # Date and time manipulation

## Libraries for statistical modeling and diagnostics
library(car)        # Companion to Applied Regression
library(lmtest)     # Diagnostic testing for linear models

## Libraries for enhanced plotting and aesthetics
library(gridExtra)  # Arranging multiple ggplot2 plots
library(extrafont)  # Custom fonts in plots

# --- LOAD & PREPARE DATA ---

## Define output and input file paths
output_dir <- "OUTPUT_DIR"  # Replace with desired output directory
file_path <- "INPUT_FILE_PATH"  # Replace with actual input file path

## Load sales data
sales_data <- read_excel(file_path) %>%
  mutate(
    MonthYear = as.Date(MonthYear),
    NewSales  = as.numeric(NewSales)
  ) %>%
  arrange(MonthYear)

## Create a time series object for sales data
start_year  <- year(min(sales_data$MonthYear))
start_month <- month(min(sales_data$MonthYear))
sales_ts    <- ts(sales_data$NewSales, start = c(start_year, start_month), frequency = 12)

## Load and process cryptocurrency data
crypto_file_path <- "INPUT_CRYPTO_FILE_PATH"  # Replace with crypto data file path
crypto_data <- read.csv(crypto_file_path) %>%
  mutate(
    Date      = as.Date(as.POSIXct(snapped_at / 1000, origin = "1970-01-01")),
    MarketCap = as.numeric(market_cap),
    MonthYear = as.Date(format(Date, "%Y-%m-01"))
  ) %>%
  group_by(MonthYear) %>%
  summarise(MarketCap = mean(MarketCap, na.rm = TRUE))

## Merge sales and cryptocurrency data
sales_data <- sales_data %>%
  left_join(crypto_data, by = "MonthYear") %>%
  mutate(MarketCap = ifelse(is.na(MarketCap), 0, MarketCap))

# --- INITIALIZE DATA FRAMES FOR STORING RESULTS ---

## Prepare data frames for storing metrics and forecasts
all_metrics   <- data.frame()  # Metrics for analysis
all_forecasts <- data.frame()  # Forecasted values

## Set forecasting horizon and future start date
h <- 12  # Forecast horizon in months
last_date <- max(sales_data$MonthYear)
future_start <- seq.Date(last_date, by = "month", length.out = h + 1)[2]  # Start of the forecast period

# --- HELPER FUNCTIONS ---

## 2A) Unified Metrics for Time-Series Models (ARIMA, SARIMA, TBATS, ETS, STL, etc.)
get_unified_metrics_ts <- function(model_obj, forecast_obj, actual_ts, model_name) {
  # Calculate AIC and BIC
  the_aic <- tryCatch(AIC(model_obj), error = function(e) NA)
  the_bic <- tryCatch(BIC(model_obj), error = function(e) NA)
  
  # Extract accuracy metrics from forecast package
  acc <- tryCatch(accuracy(forecast_obj), error = function(e) NA)
  if (is.matrix(acc)) {
    me   <- acc[1, "ME"]
    rmse <- acc[1, "RMSE"]
    mae  <- acc[1, "MAE"]
    mpe  <- acc[1, "MPE"]
    mape <- acc[1, "MAPE"]
    mase <- acc[1, "MASE"]
  } else {
    me <- rmse <- mae <- mpe <- mape <- mase <- NA
  }
  
  # Ljung-Box test for residual autocorrelation
  resid_vals <- tryCatch(residuals(model_obj), error = function(e) NA)
  if (!all(is.na(resid_vals))) {
    lb_test   <- Box.test(resid_vals, lag = 20, type = "Ljung-Box")
    lb_pvalue <- lb_test$p.value
  } else {
    lb_pvalue <- NA
  }
  
  # Regression-specific metrics (set to NA for TS models)
  r2 <- adj_r2 <- dw_pvalue <- NA
  
  # Create unified metrics data frame
  out <- data.frame(
    Model     = model_name,
    AIC       = the_aic,
    BIC       = the_bic,
    R2        = r2,
    Adj_R2    = adj_r2,
    ME        = me,
    RMSE      = rmse,
    MAE       = mae,
    MPE       = mpe,
    MAPE      = mape,
    MASE      = mase,
    LB_pvalue = lb_pvalue,
    DW_pvalue = dw_pvalue,
    stringsAsFactors = FALSE
  )
  return(out)
}

## 2B) Unified Metrics for Regression Models (Linear, including lagged)
get_unified_metrics_reg <- function(lm_obj, model_name, actual, fitted_vals) {
  # Calculate AIC and BIC
  the_aic <- AIC(lm_obj)
  the_bic <- BIC(lm_obj)
  
  # Extract R-squared and Adjusted R-squared
  lm_summary <- summary(lm_obj)
  the_r2     <- lm_summary$r.squared
  the_adj_r2 <- lm_summary$adj.r.squared
  
  # Residual-based metrics
  resid_vals <- actual - fitted_vals
  me         <- mean(resid_vals)
  rmse       <- sqrt(mean(resid_vals^2))
  mae        <- mean(abs(resid_vals))
  mpe        <- mean(resid_vals / actual) * 100
  mape       <- mean(abs(resid_vals / actual)) * 100
  mase       <- NA  # Not applicable for regression
  
  # Durbin-Watson test for autocorrelation
  dw_test    <- lmtest::dwtest(lm_obj)
  dw_pvalue  <- dw_test$p.value
  
  # Ljung-Box test (not typical for regression, set NA)
  lb_pvalue  <- NA
  
  # Create unified metrics data frame
  out <- data.frame(
    Model     = model_name,
    AIC       = the_aic,
    BIC       = the_bic,
    R2        = the_r2,
    Adj_R2    = the_adj_r2,
    ME        = me,
    RMSE      = rmse,
    MAE       = mae,
    MPE       = mpe,
    MAPE      = mape,
    MASE      = mase,
    LB_pvalue = lb_pvalue,
    DW_pvalue = dw_pvalue,
    stringsAsFactors = FALSE
  )
  return(out)
}

## 2C) Extract Forecast Table for Time-Series Models (80% & 95% Intervals)
get_forecast_table <- function(forecast_obj, model_name, start_date = NULL, freq = "month") {
  point_fore <- as.numeric(forecast_obj$mean)
  
  # Extract prediction intervals
  lower_80   <- as.numeric(forecast_obj$lower[, 1])
  lower_95   <- as.numeric(forecast_obj$lower[, 2])
  upper_80   <- as.numeric(forecast_obj$upper[, 1])
  upper_95   <- as.numeric(forecast_obj$upper[, 2])
  
  n_fore <- length(point_fore)
  
  # Generate forecast dates
  if (!is.null(start_date)) {
    date_seq <- seq.Date(from = as.Date(start_date), by = freq, length.out = n_fore)
  } else {
    date_seq <- 1:n_fore
  }
  
  # Create forecast data frame
  df_out <- data.frame(
    Model        = model_name,
    ForecastDate = date_seq,
    ForecastMean = point_fore,
    Lo80         = lower_80,
    Hi80         = upper_80,
    Lo95         = lower_95,
    Hi95         = upper_95,
    stringsAsFactors = FALSE
  )
  return(df_out)
}

## 2D) Helper Function for Time Series Analysis Plotting
create_appendix_plot_ts <- function(model_obj,
                                    forecast_obj,
                                    actual_ts,
                                    model_name = "MyModel",
                                    h          = 12) {
  
  # Ensure required libraries are loaded
  library(ggplot2)
  library(gridExtra)
  library(forecast)
  library(zoo)
  library(grid)
  library(dplyr)
  
  # Define output directory (generalized for user customization)
  output_dir <- "OUTPUT_DIR"  # Replace with desired output directory
  if (!dir.exists(output_dir)) {
    dir.create(output_dir, recursive = TRUE)
  }
  
  # Capture Model Summary
  model_summary_text <- capture.output(summary(model_obj))
  
  # -------------------------------
  # 1. Forecast Plot
  # -------------------------------
  p_forecast <- autoplot(forecast_obj) +
    ggtitle(paste("Forecast Plot -", model_name)) +
    xlab("Date") + 
    ylab("New Sales") +
    theme_minimal(base_size = 14) +
    theme(
      plot.title = element_text(hjust = 0.5, face = "bold"),
      axis.title = element_text(face = "bold")
    )
  
  forecast_plot_path <- file.path(output_dir, paste0("OutputsAppendix_", model_name, "_ForecastPlot.png"))
  ggsave(
    filename = forecast_plot_path,
    plot     = p_forecast,
    width    = 16,
    height   = 8,
    units    = "cm",
    dpi      = 300
  )
  
  # -------------------------------
  # 2. Forecast Table
  # -------------------------------
  df_fore_table <- data.frame(
    Date = as.Date(time(forecast_obj$mean)),
    Mean = as.numeric(forecast_obj$mean),
    Lo80 = as.numeric(forecast_obj$lower[,1]),
    Hi80 = as.numeric(forecast_obj$upper[,1]),
    Lo95 = as.numeric(forecast_obj$lower[,2]),
    Hi95 = as.numeric(forecast_obj$upper[,2])
  ) %>% mutate_if(is.numeric, round, digits = 2)
  
  table_grob <- gridExtra::tableGrob(
    df_fore_table,
    rows = NULL,
    theme = gridExtra::ttheme_default(
      core = list(fg_params = list(fontsize = 10)),
      colhead = list(fg_params = list(fontsize = 12, fontface = "bold"))
    )
  )
  
  forecast_table_path <- file.path(output_dir, paste0("OutputsAppendix_", model_name, "_ForecastTable.png"))
  ggsave(
    filename = forecast_table_path,
    plot     = ggplot() + 
      theme_void() + 
      annotation_custom(grid::grobTree(
        table_grob,
        vp = grid::viewport(width = 0.8, height = 0.8)
      ), xmin = -Inf, xmax = Inf, ymin = -Inf, ymax = Inf),
    width    = 20,
    height   = 10,
    units    = "cm",
    dpi      = 300
  )
  
  # -------------------------------
  # 3. Diagnostics
  # -------------------------------
  res <- residuals(model_obj)
  
  p_resid_ts <- autoplot(res) +
    ggtitle("Residual Time Series") +
    xlab("Date") +
    ylab("Residuals") +
    theme_minimal(base_size = 12) +
    theme(
      plot.title = element_text(hjust = 0.5, face = "bold"),
      axis.title = element_text(face = "bold"),
      axis.text.x = element_text(angle = 45, hjust = 1)
    )
  
  p_resid_hist <- ggplot(data.frame(res = res), aes(x = res)) +
    geom_histogram(color = "black", fill = "skyblue", bins = 20) +
    ggtitle("Histogram of Residuals") +
    xlab("Residual") +
    ylab("Count") +
    theme_minimal(base_size = 12) +
    theme(
      plot.title = element_text(hjust = 0.5, face = "bold"),
      axis.title = element_text(face = "bold"),
      axis.text.x = element_text(angle = 45, hjust = 1)
    )
  
  p_acf <- ggAcf(res) +
    ggtitle("ACF of Residuals") +
    theme_minimal(base_size = 12) +
    theme(
      plot.title = element_text(hjust = 0.5, face = "bold"),
      axis.title = element_text(face = "bold"),
      axis.text.x = element_text(angle = 45, hjust = 1)
    )
  
  diag_label <- paste0(
    "AIC = ", round(tryCatch(AIC(model_obj), error = function(e) NA), 1), " | ",
    "BIC = ", round(tryCatch(BIC(model_obj), error = function(e) NA), 1), " | ",
    "Ljung-Box p = ", round(tryCatch(Box.test(res, lag = 20, type = "Ljung-Box")$p.value, error = function(e) NA), 4)
  )
  
  diagnostics_full <- gridExtra::arrangeGrob(
    gridExtra::arrangeGrob(p_resid_ts, p_resid_hist, p_acf, ncol = 3),
    grid::textGrob(diag_label, gp = grid::gpar(fontsize = 12, fontface = "italic")),
    ncol = 1
  )
  
  diagnostics_path <- file.path(output_dir, paste0("OutputsAppendix_", model_name, "_Diagnostics.png"))
  ggsave(
    filename = diagnostics_path,
    plot     = diagnostics_full,
    width    = 16,
    height   = 8,
    units    = "cm",
    dpi      = 300
  )
  
  # -------------------------------
  # 4. Save Model Summary as TXT File
  # -------------------------------
  summary_path <- file.path(output_dir, paste0("OutputsAppendix_", model_name, "_ModelSummary.txt"))
  writeLines(model_summary_text, con = summary_path)
  
  # -------------------------------
  # 5. Return Components (Optional)
  # -------------------------------
  return(
    list(
      model_summary    = model_summary_text,
      forecast_table   = df_fore_table,
      forecast_plot    = p_forecast,
      diagnostics_plot = diagnostics_full
    )
  )
}

## 2E) Helper Function for Regression Models
create_appendix_plot_reg <- function(model_obj,
                                     forecast_results,
                                     actual_ts,
                                     model_name = "MyRegressionModel",
                                     h = 12) {
  
  # Generalize output directory for user customization
  output_dir <- "OUTPUT_DIR"  # Replace with your desired output directory
  if (!dir.exists(output_dir)) {
    dir.create(output_dir, recursive = TRUE)
  }
  
  # Capture model summary
  model_summary_text <- capture.output(summary(model_obj))
  
  # -------------------------------
  # 1. Forecast Plot with Actual Values
  # -------------------------------
  
  # Combine actual and forecasted data
  actual_data <- sales_data %>%
    select(MonthYear, NewSales) %>%
    rename(Date = MonthYear, Value = NewSales) %>%
    mutate(Type = "Actual")
  
  forecast_data <- forecast_results %>%
    select(ForecastDate, ForecastMean) %>%
    rename(Date = ForecastDate, Value = ForecastMean) %>%
    mutate(Type = "Forecast")
  
  combined_data <- bind_rows(actual_data, forecast_data)
  
  # Create the forecast plot
  p_forecast <- ggplot(data = combined_data, aes(x = Date, y = Value)) +
    geom_line(data = actual_data, aes(x = Date, y = Value), color = "black", size = 1) +
    geom_line(data = forecast_data, aes(x = Date, y = Value), color = "blue", size = 1) +
    geom_ribbon(
      data = forecast_results,
      aes(x = ForecastDate, ymin = Lo80, ymax = Hi80),
      fill = "lightblue", alpha = 0.5, inherit.aes = FALSE
    ) +
    geom_ribbon(
      data = forecast_results,
      aes(x = ForecastDate, ymin = Lo95, ymax = Hi95),
      fill = "blue", alpha = 0.2, inherit.aes = FALSE
    ) +
    ggtitle("Forecast") +
    xlab("Date") +
    ylab("New Sales") +
    theme_minimal(base_size = 12) +
    theme(
      plot.title = element_text(hjust = 0.5, face = "bold"),
      axis.title = element_text(face = "bold"),
      axis.text.x = element_text(angle = 45, hjust = 1),
      legend.position = "none"
    )
  
  # Save the forecast plot
  ggsave(
    filename = file.path(output_dir, paste0("OutputsAppendix_", model_name, "_Forecast.png")),
    plot     = p_forecast,
    width    = 16,
    height   = 8,
    units    = "cm",
    dpi      = 300
  )
  
  # -------------------------------
  # 2. Forecast Table
  # -------------------------------
  
  df_fore_table <- forecast_results %>%
    select(ForecastDate, ForecastMean, Lo80, Hi80, Lo95, Hi95) %>%
    mutate(across(everything(), round, digits = 2)) %>%
    rename(Date = ForecastDate, Mean = ForecastMean)
  
  table_grob <- gridExtra::tableGrob(
    df_fore_table,
    rows = NULL,
    theme = gridExtra::ttheme_default(
      core = list(fg_params = list(fontsize = 10)),
      colhead = list(fg_params = list(fontsize = 12, fontface = "bold"))
    )
  )
  
  ggsave(
    filename = file.path(output_dir, paste0("OutputsAppendix_", model_name, "_ForecastTable.png")),
    plot     = ggplot() +
      theme_void() +
      annotation_custom(grid::grobTree(
        table_grob,
        vp = grid::viewport(width = 0.8, height = 0.8)
      ), xmin = -Inf, xmax = Inf, ymin = -Inf, ymax = Inf),
    width    = 20,
    height   = 10,
    units    = "cm",
    dpi      = 300
  )
  
  # -------------------------------
  # 3. Diagnostics
  # -------------------------------
  
  res <- ts(residuals(model_obj), start = start(actual_ts), frequency = frequency(actual_ts))
  
  p_resid_ts <- autoplot(res) +
    ggtitle("Residual Time Series") +
    xlab("Date") +
    ylab("Residuals") +
    theme_minimal(base_size = 12) +
    theme(
      plot.title = element_text(hjust = 0.5, face = "bold"),
      axis.title = element_text(face = "bold"),
      axis.text.x = element_text(angle = 45, hjust = 1)
    )
  
  p_resid_hist <- ggplot(data.frame(res = as.numeric(res)), aes(x = res)) +
    geom_histogram(color = "black", fill = "skyblue", bins = 20) +
    ggtitle("Histogram of Residuals") +
    xlab("Residual") +
    ylab("Count") +
    theme_minimal(base_size = 12) +
    theme(
      plot.title = element_text(hjust = 0.5, face = "bold"),
      axis.title = element_text(face = "bold"),
      axis.text.x = element_text(angle = 45, hjust = 1)
    )
  
  p_acf <- ggAcf(res) +
    ggtitle("ACF of Residuals") +
    theme_minimal(base_size = 12) +
    theme(
      plot.title = element_text(hjust = 0.5, face = "bold"),
      axis.title = element_text(face = "bold"),
      axis.text.x = element_text(angle = 45, hjust = 1)
    )
  
  diag_label <- paste0(
    "AIC = ", round(tryCatch(AIC(model_obj), error = function(e) NA), 1), " | ",
    "BIC = ", round(tryCatch(BIC(model_obj), error = function(e) NA), 1), " | ",
    "DW = ", round(tryCatch(lmtest::dwtest(model_obj)$statistic, error = function(e) NA), 4), " | ",
    "R² = ", round(summary(model_obj)$r.squared, 4), " | ",
    "Adj R² = ", round(summary(model_obj)$adj.r.squared, 4)
  )
  
  diagnostics_full <- gridExtra::arrangeGrob(
    gridExtra::arrangeGrob(p_resid_ts, p_resid_hist, p_acf, ncol = 3),
    grid::textGrob(diag_label, gp = grid::gpar(fontsize = 12, fontface = "italic")),
    ncol = 1
  )
  
  ggsave(
    filename = file.path(output_dir, paste0("OutputsAppendix_", model_name, "_Diagnostics.png")),
    plot     = diagnostics_full,
    width    = 16,
    height   = 8,
    units    = "cm",
    dpi      = 300
  )
  
  # -------------------------------
  # 4. Save Model Summary as TXT File
  # -------------------------------
  
  writeLines(
    model_summary_text,
    con = file.path(output_dir, paste0("OutputsAppendix_", model_name, "_ModelSummary.txt"))
  )
  
  # -------------------------------
  # 5. Return Components (Optional)
  # -------------------------------
  
  return(
    list(
      model_summary = model_summary_text,
      forecast_plot = p_forecast,
      diagnostics_plot = diagnostics_full
    )
  )
}

# --- TIME-SERIES MODELS ---

## 5A) Simple Exponential Smoothing
ses_model <- ses(sales_ts, h = h)
ses_forecast <- forecast(ses_model, h = h, level = c(80, 95))
ses_metrics <- get_unified_metrics_ts(ses_model, ses_forecast, sales_ts, "SES")
all_metrics <- rbind(all_metrics, ses_metrics)
ses_fc_table <- get_forecast_table(ses_forecast, "SES", future_start)
all_forecasts <- rbind(all_forecasts, ses_fc_table)

ses_appendix <- create_appendix_plot_ts(
  model_obj    = ses_model,
  forecast_obj = ses_forecast,
  actual_ts    = sales_ts,
  model_name   = "SES",
  h            = h
)

## 5B) Holt's Linear Trend Method
holt_model <- holt(sales_ts, h = h)
holt_forecast <- forecast(holt_model, h = h, level = c(80, 95))
holt_metrics <- get_unified_metrics_ts(holt_model, holt_forecast, sales_ts, "Holt")
all_metrics <- rbind(all_metrics, holt_metrics)
holt_fc_table <- get_forecast_table(holt_forecast, "Holt", future_start)
all_forecasts <- rbind(all_forecasts, holt_fc_table)

holt_appendix <- create_appendix_plot_ts(
  model_obj    = holt_model,
  forecast_obj = holt_forecast,
  actual_ts    = sales_ts,
  model_name   = "Holt's Linear Trend",
  h            = h
)

## 5C) Holt-Winters' Additive
hw_model_add <- hw(sales_ts, seasonal = "additive", h = h)
hw_add_forecast <- forecast(hw_model_add, h = h, level = c(80, 95))
hw_add_metrics <- get_unified_metrics_ts(hw_model_add, hw_add_forecast, sales_ts, "HW_Additive")
all_metrics <- rbind(all_metrics, hw_add_metrics)
hw_add_fc_table <- get_forecast_table(hw_add_forecast, "HW_Additive", future_start)
all_forecasts <- rbind(all_forecasts, hw_add_fc_table)

hw_add_appendix <- create_appendix_plot_ts(
  model_obj    = hw_model_add,
  forecast_obj = hw_add_forecast,
  actual_ts    = sales_ts,
  model_name   = "Holt-Winters Additive",
  h            = h
)

## 5D) Holt-Winters' Multiplicative
hw_model_mult <- hw(sales_ts, seasonal = "multiplicative", h = h)
hw_mult_forecast <- forecast(hw_model_mult, h = h, level = c(80, 95))
hw_mult_metrics <- get_unified_metrics_ts(hw_model_mult, hw_mult_forecast, sales_ts, "HW_Multiplicative")
all_metrics <- rbind(all_metrics, hw_mult_metrics)
hw_mult_fc_table <- get_forecast_table(hw_mult_forecast, "HW_Multiplicative", future_start)
all_forecasts <- rbind(all_forecasts, hw_mult_fc_table)

hw_mult_appendix <- create_appendix_plot_ts(
  model_obj    = hw_model_mult,
  forecast_obj = hw_mult_forecast,
  actual_ts    = sales_ts,
  model_name   = "Holt-Winters Multiplicative",
  h            = h
)

## 5E) ARIMA (Non-seasonal)
arima_model    <- auto.arima(sales_ts, seasonal=FALSE)
arima_forecast <- forecast(arima_model, h=h, level=c(80,95))
arima_metrics  <- get_unified_metrics_ts(arima_model, arima_forecast, sales_ts, "ARIMA_NonSeasonal")
all_metrics    <- rbind(all_metrics, arima_metrics)
arima_fc_table <- get_forecast_table(arima_forecast, "ARIMA_NonSeasonal", future_start)
all_forecasts  <- rbind(all_forecasts, arima_fc_table)

arima_appendix <- create_appendix_plot_ts(
  model_obj    = arima_model,
  forecast_obj = arima_forecast,
  actual_ts    = sales_ts,
  model_name   = "ARIMA_NonSeasonal",
  h            = h
)

## 5F) SARIMA
sarima_model    <- auto.arima(sales_ts, seasonal=TRUE)
sarima_forecast <- forecast(sarima_model, h=h, level=c(80,95))
sarima_metrics  <- get_unified_metrics_ts(sarima_model, sarima_forecast, sales_ts, "SARIMA")
all_metrics     <- rbind(all_metrics, sarima_metrics)
sarima_fc_table <- get_forecast_table(sarima_forecast, "SARIMA", future_start)
all_forecasts   <- rbind(all_forecasts, sarima_fc_table)

sarima_appendix <- create_appendix_plot_ts(
  model_obj    = sarima_model,
  forecast_obj = sarima_forecast,
  actual_ts    = sales_ts,
  model_name   = "SARIMA",
  h            = h
)

## 5G) SARIMAX (with Crypto)
xreg_marketcap <- matrix(sales_data$MarketCap, ncol = 1)
sarimax_model <- auto.arima(sales_ts, xreg = xreg_marketcap, seasonal = TRUE)

# Use the last value of MarketCap for future xreg
last_cap_value <- tail(sales_data$MarketCap, 1)
future_xreg <- matrix(rep(last_cap_value, h), ncol = 1)

sarimax_forecast <- forecast(sarimax_model, xreg = future_xreg, h = h)
sarimax_metrics <- get_unified_metrics_ts(sarimax_model, sarimax_forecast, sales_ts, "SARIMAX_Crypto")
all_metrics <- rbind(all_metrics, sarimax_metrics)

sarimax_fc_table <- get_forecast_table(sarimax_forecast, "SARIMAX_Crypto", future_start)
all_forecasts <- rbind(all_forecasts, sarimax_fc_table)

sarimax_appendix <- create_appendix_plot_ts(
  model_obj    = sarimax_model,
  forecast_obj = sarimax_forecast,
  actual_ts    = sales_ts,
  model_name   = "SARIMAX_Crypto",
  h            = h
)

## 5H) TBATS
tbats_model <- tbats(sales_ts)
tbats_forecast <- forecast(tbats_model, h = h)

tbats_metrics <- get_unified_metrics_ts(tbats_model, tbats_forecast, sales_ts, "TBATS")
all_metrics <- rbind(all_metrics, tbats_metrics)

tbats_fc_table <- get_forecast_table(tbats_forecast, "TBATS", future_start)
all_forecasts <- rbind(all_forecasts, tbats_fc_table)

tbats_appendix <- create_appendix_plot_ts(
  model_obj    = tbats_model,
  forecast_obj = tbats_forecast,
  actual_ts    = sales_ts,
  model_name   = "TBATS",
  h            = h
)

## 5I) ETS (ZZZ, AAA, MMM)

# ETS ZZZ
ets_ZZZ <- ets(sales_ts, model = "ZZZ")
fcast_ZZZ <- forecast(ets_ZZZ, h = h)
zzz_metrics <- get_unified_metrics_ts(ets_ZZZ, fcast_ZZZ, sales_ts, "ETS_ZZZ")
all_metrics <- rbind(all_metrics, zzz_metrics)

zzz_fc_table <- get_forecast_table(fcast_ZZZ, "ETS_ZZZ", future_start)
all_forecasts <- rbind(all_forecasts, zzz_fc_table)

ets_ZZZ_appendix <- create_appendix_plot_ts(
  model_obj    = ets_ZZZ,
  forecast_obj = fcast_ZZZ,
  actual_ts    = sales_ts,
  model_name   = "ETS_ZZZ",
  h            = h
)

# ETS AAA
ets_AAA <- ets(sales_ts, model = "AAA")
fcast_AAA <- forecast(ets_AAA, h = h)
aaa_metrics <- get_unified_metrics_ts(ets_AAA, fcast_AAA, sales_ts, "ETS_AAA")
all_metrics <- rbind(all_metrics, aaa_metrics)

aaa_fc_table <- get_forecast_table(fcast_AAA, "ETS_AAA", future_start)
all_forecasts <- rbind(all_forecasts, aaa_fc_table)

ets_AAA_appendix <- create_appendix_plot_ts(
  model_obj    = ets_AAA,
  forecast_obj = fcast_AAA,
  actual_ts    = sales_ts,
  model_name   = "ETS_AAA",
  h            = h
)

# ETS MMM
ets_MMM <- tryCatch(ets(sales_ts, model = "MMM"), error = function(e) NULL, warning = function(w) NULL)
if (!is.null(ets_MMM)) {
  fcast_MMM <- forecast(ets_MMM, h = h)
  mmm_metrics <- get_unified_metrics_ts(ets_MMM, fcast_MMM, sales_ts, "ETS_MMM")
  all_metrics <- rbind(all_metrics, mmm_metrics)
  
  mmm_fc_table <- get_forecast_table(fcast_MMM, "ETS_MMM", future_start)
  all_forecasts <- rbind(all_forecasts, mmm_fc_table)
  
  ets_MMM_appendix <- create_appendix_plot_ts(
    model_obj    = ets_MMM,
    forecast_obj = fcast_MMM,
    actual_ts    = sales_ts,
    model_name   = "ETS_MMM",
    h            = h
  )
}

## 5J) STL + stlf() (ETS and ARIMA)

# STL + ETS
stlf_ets_model <- stlf(sales_ts, method = "ets", h = h)
stlf_ets_metrics <- get_unified_metrics_ts(stlf_ets_model$model, stlf_ets_model, sales_ts, "STL_ETS")
all_metrics <- rbind(all_metrics, stlf_ets_metrics)

stlf_ets_table <- get_forecast_table(stlf_ets_model, "STL_ETS", future_start)
all_forecasts <- rbind(all_forecasts, stlf_ets_table)

stlf_ets_appendix <- create_appendix_plot_ts(
  model_obj    = stlf_ets_model$model,
  forecast_obj = stlf_ets_model,
  actual_ts    = sales_ts,
  model_name   = "STL_ETS",
  h            = h
)

# STL + ARIMA
stlf_arima_model <- stlf(sales_ts, method = "arima", h = h)
stlf_arima_metrics <- get_unified_metrics_ts(stlf_arima_model$model, stlf_arima_model, sales_ts, "STL_ARIMA")
all_metrics <- rbind(all_metrics, stlf_arima_metrics)

stlf_arima_table <- get_forecast_table(stlf_arima_model, "STL_ARIMA", future_start)
all_forecasts <- rbind(all_forecasts, stlf_arima_table)

stlf_arima_appendix <- create_appendix_plot_ts(
  model_obj    = stlf_arima_model$model,
  forecast_obj = stlf_arima_model,
  actual_ts    = sales_ts,
  model_name   = "STL_ARIMA",
  h            = h
)

# --- HARMONIC REGRESSION (4 VARIANTS ACROSS K=1..6) ---

## Helper function for harmonic regression
fit_harmonic_regression <- function(ts_data, fourier_terms, external_vars=NULL) {
  if (is.null(external_vars)) {
    xreg <- fourier_terms
  } else {
    xreg <- cbind(fourier_terms, as.matrix(external_vars))
  }
  model <- auto.arima(ts_data, xreg = xreg, seasonal = FALSE)
  return(model)
}

## Add CEO effect as an external variable
sales_data <- sales_data %>%
  mutate(CEO_Effect = ifelse(MonthYear >= as.Date("2023-09-20"), 1, 0))

## Loop over K=1..6 for harmonic regression variants
for (Kval in 1:6) {
  fourier_in  <- fourier(sales_ts, K=Kval)
  fourier_out <- fourierf(sales_ts, K=Kval, h=h)
  
  xreg_crypto     <- as.matrix(sales_data$MarketCap)
  xreg_ceo        <- as.matrix(sales_data$CEO_Effect)
  xreg_ceo_crypto <- as.matrix(sales_data[, c("CEO_Effect", "MarketCap")])
  
  # A) Basic Harmonic Regression
  m_basic <- fit_harmonic_regression(sales_ts, fourier_in, NULL)
  fc_basic <- forecast(m_basic, xreg=fourier_out, h=h, level=c(80,95))
  mbasic_metrics <- get_unified_metrics_ts(m_basic, fc_basic, sales_ts, paste0("HarmBasic_K",Kval))
  all_metrics <- rbind(all_metrics, mbasic_metrics)
  basic_fc_table <- get_forecast_table(fc_basic, paste0("HarmBasic_K",Kval), future_start)
  all_forecasts  <- rbind(all_forecasts, basic_fc_table)
  
  basic_appendix <- create_appendix_plot_ts(
    model_obj    = m_basic,
    forecast_obj = fc_basic,
    actual_ts    = sales_ts,
    model_name   = paste0("HarmBasic_K", Kval),
    h            = h
  )
  
  # B) Harmonic Regression with Crypto
  m_crypto <- fit_harmonic_regression(sales_ts, fourier_in, xreg_crypto)
  last_crypto_val <- tail(xreg_crypto,1)
  future_crypto   <- matrix(rep(last_crypto_val,h), ncol=1)
  future_xreg_B   <- cbind(fourier_out, future_crypto)
  fc_crypto <- forecast(m_crypto, xreg=future_xreg_B, h=h, level=c(80,95))
  mcrypto_metrics <- get_unified_metrics_ts(m_crypto, fc_crypto, sales_ts, paste0("HarmCrypto_K",Kval))
  all_metrics <- rbind(all_metrics, mcrypto_metrics)
  crypto_fc_table <- get_forecast_table(fc_crypto, paste0("HarmCrypto_K",Kval), future_start)
  all_forecasts <- rbind(all_forecasts, crypto_fc_table)
  
  crypto_appendix <- create_appendix_plot_ts(
    model_obj    = m_crypto,
    forecast_obj = fc_crypto,
    actual_ts    = sales_ts,
    model_name   = paste0("HarmCrypto_K", Kval),
    h            = h
  )
  
  # C) Harmonic Regression with CEO Effect
  m_ceo <- fit_harmonic_regression(sales_ts, fourier_in, xreg_ceo)
  last_ceo_val <- tail(xreg_ceo,1)
  future_ceo   <- matrix(rep(last_ceo_val,h), ncol=1)
  fc_ceo_xreg  <- cbind(fourier_out, future_ceo)
  fc_ceo <- forecast(m_ceo, xreg=fc_ceo_xreg, h=h, level=c(80,95))
  mceo_metrics <- get_unified_metrics_ts(m_ceo, fc_ceo, sales_ts, paste0("HarmCEO_K",Kval))
  all_metrics <- rbind(all_metrics, mceo_metrics)
  ceo_fc_table <- get_forecast_table(fc_ceo, paste0("HarmCEO_K",Kval), future_start)
  all_forecasts <- rbind(all_forecasts, ceo_fc_table)
  
  ceo_appendix <- create_appendix_plot_ts(
    model_obj    = m_ceo,
    forecast_obj = fc_ceo,
    actual_ts    = sales_ts,
    model_name   = paste0("HarmCEO_K", Kval),
    h            = h
  )
  
  # D) Harmonic Regression with Both Variables
  m_both <- fit_harmonic_regression(sales_ts, fourier_in, xreg_ceo_crypto)
  last_ceo_val    <- tail(xreg_ceo,1)
  last_crypto_val <- tail(xreg_crypto,1)
  future_ceo_vals    <- matrix(rep(last_ceo_val,h), ncol=1)
  future_crypto_vals <- matrix(rep(last_crypto_val,h), ncol=1)
  future_xreg_both   <- cbind(fourier_out, future_ceo_vals, future_crypto_vals)
  fc_both <- forecast(m_both, xreg=future_xreg_both, h=h, level=c(80,95))
  mboth_metrics <- get_unified_metrics_ts(m_both, fc_both, sales_ts, paste0("HarmBoth_K",Kval))
  all_metrics <- rbind(all_metrics, mboth_metrics)
  both_fc_table <- get_forecast_table(fc_both, paste0("HarmBoth_K",Kval), future_start)
  all_forecasts <- rbind(all_forecasts, both_fc_table)
  
  both_appendix <- create_appendix_plot_ts(
    model_obj    = m_both,
    forecast_obj = fc_both,
    actual_ts    = sales_ts,
    model_name   = paste0("HarmBoth_K", Kval),
    h            = h
  )
}

# --- REGRESSION MODELS (NON-LAGGED) + FORECASTS ---

## Prepare data for regression models
sales_data <- sales_data %>%
  mutate(
    Month       = factor(format(MonthYear, "%m")),
    Trend       = as.numeric(MonthYear - min(MonthYear)),
    BlackFriday = ifelse(format(MonthYear, "%m") == "11", 1, 0)
  )

## Build future data for regression
future_dates_reg <- seq.Date(last_date, by = "month", length.out = h + 1)[-1]
month_levels <- levels(sales_data$Month)
reg_future_df <- data.frame(
  MonthYear   = future_dates_reg,
  Month       = factor(format(future_dates_reg, "%m"), levels = month_levels),
  Trend       = as.numeric(future_dates_reg - min(sales_data$MonthYear)),
  BlackFriday = ifelse(format(future_dates_reg, "%m") == "11", 1, 0)
)
reg_future_df$MarketCap <- tail(sales_data$MarketCap, 1)

## Define model formulas
model_definitions <- list(
  list(name = "Reg_DummyOnly", formula = NewSales ~ Month),
  list(name = "Reg_Trend_CEO", formula = NewSales ~ Month + Trend + CEO_Effect),
  list(name = "Reg_Interaction", formula = NewSales ~ Month + Trend + CEO_Effect + Trend_CEO),
  list(name = "Reg_Full", formula = NewSales ~ Month + Trend + CEO_Effect + Trend_CEO + MarketCap + BlackFriday),
  list(name = "Reg_Nonlinear", formula = NewSales ~ Month + Trend + Trend2 + CEO_Effect + Trend_CEO + MarketCap),
  list(name = "Reg_Simplified", formula = NewSales ~ Month + Trend + MarketCap)
)

## Adjust data for regression models
sales_data <- sales_data %>%
  mutate(
    CEO_Effect = ifelse(MonthYear >= as.Date("2023-09-20"), 1, 0),
    Trend_CEO = Trend * CEO_Effect,
    Trend2 = Trend^2
  )
reg_future_df <- reg_future_df %>%
  mutate(
    CEO_Effect = ifelse(MonthYear >= as.Date("2023-09-20"), 1, 0),
    Trend_CEO = Trend * CEO_Effect,
    Trend2 = Trend^2
  )

## Loop through regression models
for (model_def in model_definitions) {
  model_name <- model_def$name
  model_formula <- model_def$formula
  
  # Fit model
  model_obj <- lm(model_formula, data = sales_data)
  
  # Generate forecast
  fc <- forecast_regression_model(model_obj, model_name, reg_future_df,
                                  start_date = min(reg_future_df$MonthYear))
  all_forecasts <- rbind(all_forecasts, fc)
  
  # Generate metrics
  metrics <- get_unified_metrics_reg(model_obj, model_name,
                                     actual = sales_data$NewSales,
                                     fitted_vals = fitted(model_obj))
  all_metrics <- rbind(all_metrics, metrics)
  
  # Generate plots and diagnostics
  create_appendix_plot_reg(
    model_obj = model_obj,
    forecast_results = fc,
    actual_ts = sales_data$NewSales,
    model_name = model_name,
    h = h
  )
}

# --- LAGGED REGRESSION FORECAST ---

## Prepare lagged training data
lag_data <- sales_data %>%
  arrange(MonthYear) %>%
  mutate(Lag1 = dplyr::lag(NewSales)) %>%
  drop_na()

## Fit the final lagged regression model
final_lag_model <- lm(NewSales ~ Month + Trend + Lag1, data = lag_data)

## Multi-step forecasting function for lagged regression
forecast_lagged_regression_final <- function(lm_obj, model_name, future_data, start_date, 
                                             h, last_sales_value) {
  results_list <- list()
  
  for (i in seq_len(h)) {
    # Build row i
    row_i <- future_data[i, , drop = FALSE]
    if (i == 1) {
      row_i$Lag1 <- last_sales_value
    } else {
      row_i$Lag1 <- results_list[[i - 1]]$ForecastMean
    }
    
    # Generate forecasts with prediction intervals
    preds_80 <- predict(lm_obj, newdata = row_i, interval = "prediction", level = 0.80)
    preds_95 <- predict(lm_obj, newdata = row_i, interval = "prediction", level = 0.95)
    
    # Extract forecast components
    date_val <- seq.Date(from = as.Date(start_date), by = "month", length.out = h)[i]
    results_list[[i]] <- data.frame(
      Model        = model_name,
      ForecastDate = date_val,
      ForecastMean = preds_80[, "fit"],
      Lo80         = preds_80[, "lwr"],
      Hi80         = preds_80[, "upr"],
      Lo95         = preds_95[, "lwr"],
      Hi95         = preds_95[, "upr"],
      stringsAsFactors = FALSE
    )
  }
  
  df_final <- do.call(rbind, results_list)
  return(df_final)
}

## Prepare future data for lagged regression
future_data_lag <- reg_future_df %>%
  dplyr::select(MonthYear, Month, Trend) # Include required predictors
if ("MarketCap" %in% colnames(lag_data)) {
  future_data_lag$MarketCap <- tail(lag_data$MarketCap, 1)
}
future_data_lag$Lag1 <- NA

## Validate future data and extract start date
if (all(is.na(future_data_lag$MonthYear))) {
  stop("future_data_lag$MonthYear contains only NA values. Check reg_future_df preparation.")
}
start_date <- min(future_data_lag$MonthYear, na.rm = TRUE)

## Last known sales value
last_sales_value <- tail(lag_data$NewSales, 1)

## Multi-step forecast
lag_fc <- forecast_lagged_regression_final(
  lm_obj         = final_lag_model,
  model_name     = "Reg_AR1_Final_FC",
  future_data    = future_data_lag,
  start_date     = start_date,
  h              = h,
  last_sales_value = last_sales_value
)

## Combine forecasts
all_forecasts <- rbind(all_forecasts, lag_fc)

## Generate metrics for lagged model
metrics_lagfin <- get_unified_metrics_reg(final_lag_model, "Reg_AR1_Final",
                                          actual = lag_data$NewSales,
                                          fitted_vals = fitted(final_lag_model))
all_metrics <- rbind(all_metrics, metrics_lagfin)

## Generate plots and diagnostics
create_appendix_plot_reg(
  model_obj = final_lag_model,
  forecast_results = lag_fc,
  actual_ts = ts(lag_data$NewSales, start = start(lag_data$MonthYear), frequency = 12),
  model_name = "Reg_AR1_Final",
  h = h
)

# --- SIMPLE LAGGED REGRESSION FORECAST (STATIC Lag1) ---

## Prepare lagged training data
lag_data <- sales_data %>%
  arrange(MonthYear) %>%
  mutate(Lag1 = dplyr::lag(NewSales)) %>%  # Lag1 is purely historical
  drop_na()  # Ensure no NA rows in the training set

## Fit the lagged regression model
final_lag_model_static <- lm(NewSales ~ Month + Trend + Lag1, data = lag_data)

## Forecasting function with STATIC Lag1
forecast_lagged_regression_static <- function(lm_obj, model_name, future_data, start_date, 
                                              h, last_sales_value) {
  results_list <- list()
  
  # Initialize Lag1 for forecasts with last observed sales value
  future_data$Lag1 <- last_sales_value
  
  for (i in seq_len(h)) {
    # Build row i (future data)
    row_i <- future_data[i, , drop = FALSE]
    
    # Generate forecasts with prediction intervals
    preds_80 <- predict(lm_obj, newdata = row_i, interval = "prediction", level = 0.80)
    preds_95 <- predict(lm_obj, newdata = row_i, interval = "prediction", level = 0.95)
    
    # Extract forecast components
    date_val <- seq.Date(from = as.Date(start_date), by = "month", length.out = h)[i]
    mean_fc <- preds_80[, "fit"]
    lo80    <- preds_80[, "lwr"]
    hi80    <- preds_80[, "upr"]
    lo95    <- preds_95[, "lwr"]
    hi95    <- preds_95[, "upr"]
    
    # Store results
    results_list[[i]] <- data.frame(
      Model        = model_name,
      ForecastDate = date_val,
      ForecastMean = mean_fc,
      Lo80         = lo80,
      Hi80         = hi80,
      Lo95         = lo95,
      Hi95         = hi95,
      stringsAsFactors = FALSE
    )
  }
  
  # Combine results into a single data frame
  df_final <- do.call(rbind, results_list)
  return(df_final)
}

## Prepare future data for lagged regression
future_data_lag <- reg_future_df %>%
  dplyr::select(MonthYear, Month, Trend)  # Include all required predictors

## Validate and extract start date
if (all(is.na(future_data_lag$MonthYear))) {
  stop("future_data_lag$MonthYear contains only NA values. Check reg_future_df preparation.")
}
start_date <- min(future_data_lag$MonthYear, na.rm = TRUE)

## Last known sales value (static Lag1 initialization)
last_sales_value <- tail(lag_data$NewSales, 1)

## Multi-step forecast
lag_fc_static <- forecast_lagged_regression_static(
  lm_obj         = final_lag_model_static,
  model_name     = "Reg_AR1_Final_Static",
  future_data    = future_data_lag,
  start_date     = start_date,
  h              = h,
  last_sales_value = last_sales_value
)

## Combine forecasts
all_forecasts <- rbind(all_forecasts, lag_fc_static)

## Generate metrics for the simple lagged model
metrics_lag_static <- get_unified_metrics_reg(final_lag_model_static, "Reg_AR1_Final_Static",
                                              actual = lag_data$NewSales,
                                              fitted_vals = fitted(final_lag_model_static))
all_metrics <- rbind(all_metrics, metrics_lag_static)

## Generate plots and diagnostics
create_appendix_plot_reg(
  model_obj = final_lag_model_static,
  forecast_results = lag_fc_static,
  actual_ts = ts(lag_data$NewSales, start = start(lag_data$MonthYear), frequency = 12),
  model_name = "Reg_AR1_Final_Static",
  h = h
)

# --- WRITE OUTPUT TO CSV ---

## Save performance metrics to CSV
write.csv(all_metrics, "model_performance_metrics.csv", row.names = FALSE)

## Save forecasts to CSV
write.csv(all_forecasts, "model_forecasts.csv", row.names = FALSE)

## Completion message
message("All done! Check 'model_performance_metrics.csv' and 'model_forecasts.csv' for results.")
