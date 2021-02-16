library(readr)
library(tools)
library(plyr)
library(dplyr)
library(ggplot2)
library(corrplot)
library(scales)
library(ggrepel)
library(lubridate)
library(openxlsx)


# Load CSVs
# --------------------------------------------------

sessions_raw = read_csv("./data/DataAnalyst_Ecom_data_sessionCounts.csv")
cart_raw = read_csv("./data/DataAnalyst_Ecom_data_addsToCart.csv")


# Clean Data
# --------------------------------------------------

# Check data types
summary(sessions_raw)
summary(cart_raw)

# Change dim_date to date object
sessions_raw$dim_date = as.Date(sessions_raw$dim_date, format = "%m/%d/%y")

# Change column names
names(sessions_raw) = c("Browser", "Device Category", "Date", "Sessions", "Transactions", "QTY")
sessions_raw$"Device Category" = toTitleCase(sessions_raw$"Device Category")


# Month*Device Aggregation
# --------------------------------------------------

# Subset sessions_raw
df_MD = sessions_raw[c("Device Category", "Date", "Sessions", "Transactions", "QTY")]

# Change date to month
df_MD = df_MD %>% mutate(Date = floor_date(as_date(Date)))

# Group by month
df_MD = sessions_raw %>%
  group_by(df_MD$"Device Category", Date) %>%
  summarize(
    "Total Sessions" = sum(Sessions), 
    "Total Transactions" = sum(Transactions), 
    "Total QTY" = sum(QTY)
  )

# Order by date
df_MD = df_MD[order(df_MD$"Date"), ]


# Add ECR column
df_MD$ECR = df_MD$total_transactions/df_MD$total_sessions
df_MD$ECR = format(round(df_MD$ECR,4), nsmall = 4)

df_MD



wb = createWorkbook("wb_Sessions_Car")
addWorksheet(wb, "Sessions", gridLines = TRUE)
setColWidths(wb, "Sessions", cols = 1:6, widths = 15)
writeData(wb, sheet = "Sessions", df_MD, colNames = TRUE, rowNames = FALSE)
saveWorkbook(wb, "R_Challenge.xlsx", overwrite = TRUE)

