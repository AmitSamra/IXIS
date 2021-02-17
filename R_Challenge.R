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

# ----- sessions_raw -----
# Change sessions_raw$dim_date to date object
sessions_raw$dim_date = as.Date(sessions_raw$dim_date, format = "%m/%d/%y")

# Create year and month columns for sessions_raw
sessions_raw$Year = format(as.Date(sessions_raw$dim_date, format = "%m/%d/%Y"),"%Y")
sessions_raw$Month = format(as.Date(sessions_raw$dim_date, format = "%m/%d/%Y"),"%m")
sessions_raw$Date = as.Date(paste(sessions_raw$Year, sessions_raw$Month, 01), "%Y %m %d")

# Delete unnecessary columns in sessions_raw
sessions_raw = within(sessions_raw, rm(dim_date, Year, Month))

# Move sessions_raw$Date to first position
sessions_raw = sessions_raw %>%
  select(Date, everything())

# Change column names in sessions_raw
names(sessions_raw) = c("Date", "Browser", "Device Category", "Sessions", "Transactions", "QTY")
sessions_raw$`Device Category` = toTitleCase(sessions_raw$"Device Category")

# ----- cart_raw -----
# Create date column in cart_raw
cart_raw$Date = as.Date(paste(cart_raw$dim_year, cart_raw$dim_month, 01), "%Y %m %d")

# Delete cart_raw$dim_year and cart_raw$dim_month
cart_raw = within(cart_raw, rm(dim_year, dim_month))

# Move cart_raw$Date to first position
cart_raw = cart_raw %>%
  select(Date, everything())

# Change column names in cart_raw
names(cart_raw) = c("Date", "Adds To Cart")


# Sheet 1: Month*Device Aggregation
# --------------------------------------------------

# Subset sessions_raw
df_MD = sessions_raw[c("Date", "Device Category", "Sessions", "Transactions", "QTY")]

# Group by month
df_MD = df_MD %>%
  group_by(Date, `Device Category`) %>%
  summarize(
    "Total Sessions" = sum(Sessions), 
    "Total Transactions" = sum(Transactions), 
    "Total QTY" = sum(QTY)
  )

# Order by date
df_MD = df_MD[order(df_MD$"Date"), ]

# Add ECR column
df_MD$ECR = df_MD$`Total Transactions`/df_MD$`Total Sessions`
df_MD$ECR = format(round(df_MD$ECR,4), nsmall = 4)


# Sheet 2: Month over Month
# --------------------------------------------------

# Subset cart_raw
df_MM = cart_raw[c("Date", "Adds To Cart")]

# In order to combine df_MD and df_MM, we must regroup df_MD only by date
df_MD2 = sessions_raw[c("Date", "Sessions", "Transactions", "QTY")]
df_MD2 = df_MD2 %>%
  group_by(Date) %>%
  summarize(
    "Total Sessions" = sum(Sessions), 
    "Total Transactions" = sum(Transactions), 
    "Total QTY" = sum(QTY)
  )

# Merge df_MM and df_MD2
df_MD_MM = merge(df_MD2, df_MM)

# Calculate differences
df_MD_MM$`Sessions Change` = df_MD_MM$`Total Sessions` - lag(df_MD_MM$`Total Sessions`)
df_MD_MM$`Sessions % Change` = df_MD_MM$`Total Sessions` / lag(df_MD_MM$`Total Sessions`) -1
df_MD_MM$`Transactions Change` = df_MD_MM$`Total Transactions` - lag(df_MD_MM$`Total Transactions`)
df_MD_MM$`Transactions % Change` = df_MD_MM$`Total Transactions` / lag(df_MD_MM$`Total Transactions`) -1
df_MD_MM$`QTY Change` = df_MD_MM$`Total QTY` - lag(df_MD_MM$`Total QTY`)
df_MD_MM$`QTY % Change` = df_MD_MM$`Total QTY` / lag(df_MD_MM$`Total QTY`) -1
df_MD_MM$`Cart Change` = df_MD_MM$`Adds To Cart` - lag(df_MD_MM$`Adds To Cart`)
df_MD_MM$`Cart % Change` = df_MD_MM$`Adds To Cart` / lag(df_MD_MM$`Adds To Cart`) -1

# Round decimals
df_MD_MM$`Sessions % Change` = format(round(df_MD_MM$`Sessions % Change`,2), nsmall = 2)
df_MD_MM$`Transactions % Change` = format(round(df_MD_MM$`Transactions % Change`,2), nsmall = 2)
df_MD_MM$`QTY % Change` = format(round(df_MD_MM$`QTY % Change`,2), nsmall = 2)
df_MD_MM$`Cart % Change` = format(round(df_MD_MM$`Cart % Change`,2), nsmall = 2)


# Export XLSX Workbook
# --------------------------------------------------

# Create workbook
wb = createWorkbook("wb_Sessions_Car")

# Add Sheet 1
addWorksheet(wb, "Sessions", gridLines = TRUE)
setColWidths(wb, "Sessions", cols = 1:6, widths = 15)
writeData(wb, sheet = "Sessions", df_MD, colNames = TRUE, rowNames = FALSE)

# Add Sheet 2
addWorksheet(wb, "Cart", gridLines = TRUE)
setColWidths(wb, "Cart", cols = 1:13, widths = 15)
writeData(wb, sheet = "Cart", df_MD_MM, colNames = TRUE, rowNames = FALSE)

# Save Workbook
saveWorkbook(wb, "R_Challenge.xlsx", overwrite = TRUE)



# Visualizations
# --------------------------------------------------


df_MD
df_MD_MM


df_MD %>%
  ggplot( aes(x = Date, y = df_MD$`ECR`, group = df_MD$`Device Category`, color = df_MD$`Device Category` ) ) + 
  geom_line() +
  xlab('Date') +
  ylab('ECR') +
  ggtitle('ECR') +
  scale_x_date(breaks = "1 month", date_labels = "%y %b") +
  scale_color_manual(values=c("red2", "steelblue", "green4"))




