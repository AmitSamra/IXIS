library(readr)
library(plyr)
library(dplyr)
library(ggplot2)
library(corrplot)
library(scales)
library(ggrepel)
library(openxlsx)

# Load CSVs
# --------------------------------------------------
df_Sessions = read_csv("./data/DataAnalyst_Ecom_data_sessionCounts.csv")
View(df_Sessions)

df_Cart = read_csv("./data/DataAnalyst_Ecom_data_addsToCart.csv")
View(df_Cart)

# Clean Data
# --------------------------------------------------

# Check data types
summary(df_Sessions)

# Change dim_date to date object
df_Sessions$dim_date = as.Date(df_Sessions$dim_date, format = "%m/%d/%Y")

# Month*Device Aggregation

df_MD = df_Session










wb = createWorkbook("wb_Sessions_Car")
addWorksheet(wb, "Sessions")
saveWorkbook(wb, "R_Challenge.xlsx", overwrite = TRUE)

