library(readr)
library(tools)
library(plyr)
library(dplyr)
library(ggplot2)
library(corrplot)
library(scales)
library(ggrepel)
library(tidyr)
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
names(cart_raw) = c("Date", "Add To Carts")


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
df_MD$ECR = as.numeric(df_MD$ECR)


# Sheet 2: Month over Month
# --------------------------------------------------

# Subset cart_raw
df_MM = cart_raw[c("Date", "Add To Carts")]

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
df_MD_MM$`Cart Change` = df_MD_MM$`Add To Carts` - lag(df_MD_MM$`Add To Carts`)
df_MD_MM$`Cart % Change` = df_MD_MM$`Add To Carts` / lag(df_MD_MM$`Add To Carts`) -1

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

# Plot Total Sessions
df_MD %>%
  ggplot( aes(x = Date, y = df_MD$`Total Sessions`, group = df_MD$`Device Category`, color = df_MD$`Device Category` ) ) + 
  geom_line() +
  labs(x='Date', y='Total Sessions', color='Device Category', title='Total Sessions') +
  scale_x_date(breaks = "1 month", date_labels = "%y %b") +
  scale_y_continuous(breaks=scales::breaks_extended(n=10), labels=comma)  +
  scale_color_manual(values=c("red2", "steelblue", "green4"))

ggsave('sessions.png', device='png', path='img')


# Plot ECR
df_MD %>%
  ggplot( aes(x = Date, y = df_MD$`ECR`, group = df_MD$`Device Category`, color = df_MD$`Device Category` ) ) + 
  geom_line() +
  labs(x='Date', y='ECR', color='Device Category', title='ECR') +
  scale_x_date(breaks = "1 month", date_labels = "%y %b") +
  scale_color_manual(values=c("red2", "steelblue", "green4")) +
  geom_text(hjust=0, vjust=-1, size=3, aes(label=round(df_MD$`ECR`, digits = 3))) +
  scale_y_continuous(breaks=scales::breaks_extended(n=10), labels=comma)

ggsave('ECR.png', device='png', path='img')

# Plot Sessions & Transactions
# Rename columns for binding
df_S = df_MD_MM[c("Date", "Total Sessions")]
names(df_S) = c("Date", "Total")
df_T = df_MD_MM[c("Date", "Total Transactions")]
names(df_T) = c("Date", "Total")
df_C = df_MD_MM[c("Date", "Add To Carts")]
names(df_C) = c("Date", "Total")

# Bind dataframes
df_S_T = dplyr::bind_rows(df_S, df_T, df_C, .id='id')

df_S_T %>%
  ggplot(aes(x=`Date`, y=`Total`, fill=id)) +
  geom_bar(stat='identity') +
  scale_y_continuous(breaks=scales::breaks_extended(n=10), labels=comma) +
  geom_text(hjust=.5, vjust=-3, size=3, aes(label=comma(`Total`))) +
  scale_fill_manual(labels=c('Sessions', 'Add to Carts', 'Transactions'), values=c('aquamarine4', 'aquamarine3', 'lightblue3')) +
  labs(title='Sessions, Add to Carts, Transactions', x='Date', y='Total', fill='Key') +
  scale_x_date(breaks = "1 month", date_labels = "%y %b")

ggsave('sessions_carts_transactions.png', device='png', path='img')

# Plot Transactions & Total QTY
# Bind dataframes
df_Trans = df_MD_MM[c("Date", "Total Transactions")]
names(df_Trans) = c("Date", "Total")
df_QTY = df_MD_MM[c("Date", "Total QTY")]
names(df_QTY) = c("Date", "Total")
df_Trans_QTY = dplyr::bind_rows(df_Trans, df_QTY, .id='id')
  
df_Trans_QTY %>%
  ggplot(aes(x=`Date`, y=`Total`, fill=id))+
  geom_bar(stat="identity", position="dodge") +
  scale_y_continuous(breaks=scales::breaks_extended(n=10), labels=comma) +
  scale_x_date(breaks = "1 month", date_labels = "%y %b") +
  labs(x='Date', y='Total', color='Key', title='Transactions & QTY', fill='Key') +
  scale_fill_manual(labels=c('Transactions', 'QTY'), values=c('aquamarine4', 'lightblue3'))

ggsave('transactions_QTY.png', device='png', path='img')
