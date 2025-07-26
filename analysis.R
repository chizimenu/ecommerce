# ======================================================================
# Title: eCommerce Sales Analysis Script
# Author: Charles Ofuasia
# Purpose: Analyze sales trends, identify top-selling products,
#          peak/low sales months, and customer distribution by state.
# ======================================================================

# ====================
# Load Required Libraries
# ====================
library(tidyverse) # Data wrangling and plotting
library(lubridate) # Date parsing
library(openxlsx) # Excel export

# ====================================
# Create Output folder to store results
# ====================================
dir.create("analysis_output", showWarnings = FALSE)

# ====================
# Loading Dataset
# ===================
df_raw <- read_csv("eCommerce.csv")

# ===============================================
# Initial Data Inspection to understand structure
# ===============================================
cat("Data Preview:\n")
print(head(df_raw, 5))

cat("\nColumn Names:\n")
print(colnames(df_raw))

# =================================
# Data Cleaning & Transformation
# ===========================

# Conversion of Order_Date to Date format (assuming it's in dmy format)
# and Total_Sales to numeric after removing $ signs.
df <- df_raw %>%
    mutate(
        Order_Date = dmy(Order_Date), # Parse dates
        Total_Sales = as.numeric(gsub("\\$", "", Total_Sales)), # Remove $ signs
        Month = floor_date(Order_Date, "month"), # Round to first of month
        Month_Label = format(Month, "%b %Y"), # Label format: Jan 2021
        Year = year(Order_Date)
    )

# Drop rows with key missing values
df <- df %>%
    filter(!is.na(Order_Date), !is.na(Product), !is.na(Total_Sales), !is.na(State_Code))

# Ensure Month_Label is in correct order
df <- df %>%
    arrange(Month) %>%
    mutate(Month_Label = factor(Month_Label, levels = unique(Month_Label)))

# ====================
# Missing Value Summary
# ====================
missing_summary <- df_raw %>%
    summarise(
        missing_dates = sum(is.na(Order_Date)),
        missing_products = sum(is.na(Product)),
        missing_sales = sum(is.na(Total_Sales)),
        missing_states = sum(is.na(State_Code))
    )

write_csv(missing_summary, "analysis_output/missing_summary.csv")

# ====================
# Monthly Sales Summary
# ====================
monthly_sales <- df %>%
    group_by(Month_Label) %>%
    summarise(Monthly_Sales = sum(Total_Sales), .groups = "drop")

# ====================================
# Identify Peak and Lowest Selling Months
# ===================================
peak_month <- monthly_sales %>% slice_max(Monthly_Sales, n = 1)
lowest_month <- monthly_sales %>% slice_min(Monthly_Sales, n = 1)

# ====================
# Identify Top-Selling Product
# ====================
top_product <- df %>%
    group_by(Product) %>%
    summarise(Total_Sales = sum(Total_Sales), .groups = "drop") %>%
    arrange(desc(Total_Sales)) %>%
    slice(1)

# ====================
# Customer Count by State
# ====================
state_customer_count <- df %>%
    count(State_Code, name = "Customer_Count") %>%
    arrange(desc(Customer_Count))

most_customers <- state_customer_count %>% slice(1)
fewest_customers <- state_customer_count %>% slice(n())

# ====================
# Export Summary Files
# ====================
write_csv(peak_month, "analysis_output/peak_month.csv")
write_csv(lowest_month, "analysis_output/lowest_month.csv")
write_csv(top_product, "analysis_output/top_product.csv")
write_csv(state_customer_count, "analysis_output/customer_count_by_state.csv")
write_csv(most_customers, "analysis_output/state_most_customers.csv")
write_csv(fewest_customers, "analysis_output/state_fewest_customers.csv")

# ====================
# Export Summary as Text Report
# ====================
sink("analysis_output/ecommerce_summary.txt")

cat("📈 Peak Selling Month:\n")
print(peak_month)

cat("\n📉 Lowest Selling Month:\n")
print(lowest_month)

cat("\n🏆 Highest Selling Product:\n")
print(top_product)

cat("\n📍 State with Most Customers:\n")
print(most_customers)

cat("\n📍 State with Fewest Customers:\n")
print(fewest_customers)

sink()

# ====================
# Export to Excel Workbook
# ====================
wb <- createWorkbook()

addWorksheet(wb, "Peak Month")
writeData(wb, "Peak Month", peak_month)

addWorksheet(wb, "Lowest Month")
writeData(wb, "Lowest Month", lowest_month)

addWorksheet(wb, "Top Product")
writeData(wb, "Top Product", top_product)

addWorksheet(wb, "Customer by State")
writeData(wb, "Customer by State", state_customer_count)

saveWorkbook(wb, "analysis_output/ecommerce_report.xlsx", overwrite = TRUE)

# ====================
# Plot: Monthly Sales Trend
# ====================
sales_plot <- ggplot(monthly_sales, aes(x = Month_Label, y = Monthly_Sales)) +
    geom_line(color = "steelblue", group = 1, linewidth = 1.2) +
    geom_point(color = "darkred", size = 3) +
    labs(
        title = "Monthly Sales Trend",
        subtitle = "Total sales grouped by order month",
        x = "Month",
        y = "Total Sales"
    ) +
    theme_minimal() +
    theme(axis.text.x = element_text(angle = 45, hjust = 1))

# Display plot
print(sales_plot)

# Save to PNG
ggsave("analysis_output/monthly_sales_trend.png", sales_plot, width = 10, height = 6)

# ====================
# Plot: Top 5 Selling Products
# ====================
top5_products <- df %>%
    group_by(Product) %>%
    summarise(Total_Sales = sum(Total_Sales), .groups = "drop") %>%
    slice_max(Total_Sales, n = 5)

product_plot <- ggplot(top5_products, aes(x = reorder(Product, Total_Sales), y = Total_Sales)) +
    geom_col(fill = "darkgreen") +
    coord_flip() +
    labs(
        title = "Top 5 Selling Products",
        x = "Product",
        y = "Total Sales"
    ) +
    theme_minimal()

# Save plot
ggsave("analysis_output/top5_products.png", product_plot, width = 8, height = 5)

# ====================
# End of Script
# ====================
cat("✅ Analysis complete! Files saved to 'analysis_output/' folder.\n")
