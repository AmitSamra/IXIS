library(readr)
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
View(sessions_raw)

cart_raw = read_csv("./data/DataAnalyst_Ecom_data_addsToCart.csv")
View(cart_raw)


# Clean Data
# --------------------------------------------------

# Check data types
summary(sessions_raw)
summary(cart_raw)

# Change dim_date to date object
sessions_raw$dim_date = as.Date(sessions_raw$dim_date, format = "%m/%d/%y")


# Month*Device Aggregation
# --------------------------------------------------

# Subset sessions_raw
df_MD = sessions_raw[c("dim_deviceCategory", "dim_date", "sessions", "transactions", "QTY")]

# Group by month
df_MD = sessions_raw %>%
  group_by(month=floor_date(dim_date, "year/month")) %>%
  summarize(
    total_sessions = sum(sessions), 
    total_transactions = sum(transactions), 
    total_QTY = sum(QTY)
    )

# Add ECR column
df_MD$ECR = df_MD$total_transactions/df_MD$total_sessions
df_MD$ECR = format(round(df_MD$ECR,4), nsmall = 4)

df_MD


wb = createWorkbook("wb_Sessions_Car")
addWorksheet(wb, "Sessions", gridLines = TRUE)
setColWidths(wb, "Sessions", cols = 2:6, widths = 15)
writeData(wb, sheet = "Sessions", df_MD, colNames = TRUE, rowNames = FALSE)
saveWorkbook(wb, "R_Challenge.xlsx", overwrite = TRUE)

