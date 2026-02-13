============================================================================
COMPLETE EXAMPLE: Corporate Quarterly Report - R Data & Charts
============================================================================
This creates realistic business data and professional ggplot2 charts
Theme: Q4 2024 Business Performance Report
============================================================================

library(dplyr)
library(ggplot2)
library(tidyr)
library(scales)

cat("=" * 80, "\n")
cat("CORPORATE QUARTERLY REPORT - DATA GENERATION\n")
cat("=" * 80, "\n\n")

============================================================================
1. CREATE DUMMY BUSINESS DATA
============================================================================

cat("1. Creating business performance data...\n")

Summary Metrics (for Executive Summary slide)
summary_metrics <- data.frame(
  Metric = c("Quarterly Revenue", "Net Profit", "Customer Satisfaction", "Market Share"),
  Value = c(15.2, 3.8, 87, 23.5),
  Unit = c("M", "M", "%", "%"),
  Change = c(12.5, 8.3, 2.1, 1.8),
  stringsAsFactors = FALSE
)

print(summary_metrics)
cat("\n")

Monthly Sales Data (for trend chart)
monthly_sales <- data.frame(
  Month = factor(c("Jan", "Feb", "Mar", "Apr", "May", "Jun", 
                   "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"),
                levels = c("Jan", "Feb", "Mar", "Apr", "May", "Jun",
                          "Jul", "Aug", "Sep", "Oct", "Nov", "Dec")),
  Revenue = c(4.2, 4.5, 4.8, 5.1, 5.3, 5.6, 5.8, 6.1, 6.4, 6.8, 7.2, 7.6),
  Target = c(4.0, 4.2, 4.5, 4.8, 5.0, 5.3, 5.5, 5.8, 6.0, 6.5, 7.0, 7.5),
  Expenses = c(3.1, 3.2, 3.3, 3.5, 3.6, 3.8, 3.9, 4.1, 4.2, 4.4, 4.6, 4.8)
)

cat("Monthly sales data created (12 months)\n\n")

Product Performance (for comparison chart)
product_performance <- data.frame(
  Product = c("Product A", "Product B", "Product C", "Product D", "Product E"),
  Q3_Sales = c(2.8, 3.2, 1.9, 2.5, 1.6),
  Q4_Sales = c(3.5, 3.8, 2.2, 2.9, 2.1),
  Growth = c(25, 18.8, 15.8, 16, 31.3)
)

print(product_performance)
cat("\n")

Regional Distribution (for pie chart)
regional_sales <- data.frame(
  Region = c("North America", "Europe", "Asia Pacific", "Latin America", "Middle East"),
  Sales = c(6.5, 4.8, 2.9, 0.7, 0.3),
  Percentage = c(43.3, 32.0, 19.3, 4.7, 2.0)
)

cat("Regional sales data created\n\n")

Customer Segments (for stacked bar)
customer_segments <- data.frame(
  Month = rep(c("Oct", "Nov", "Dec"), each = 3),
  Segment = rep(c("Enterprise", "SMB", "Startup"), 3),
  Revenue = c(
    2.8, 1.5, 0.5,  # Oct
    3.1, 1.7, 0.6,  # Nov
    3.5, 1.9, 0.7   # Dec
  )
)

cat("Customer segment data created\n\n")

Department Budget (for horizontal bar)
department_budget <- data.frame(
  Department = c("Sales & Marketing", "Engineering", "Operations", 
                 "Customer Success", "HR & Admin"),
  Budget = c(4.5, 3.8, 2.2, 1.5, 0.8),
  Actual = c(4.2, 3.9, 2.0, 1.4, 0.7)
) %>%
  arrange(desc(Budget))

print(department_budget)
cat("\n")

============================================================================
2. CREATE OUTPUT DIRECTORY
============================================================================

cat("2. Creating output directories...\n")
system("mkdir -p /dbfs/FileStore/example_charts", ignore.stdout = TRUE)
system("mkdir -p /dbfs/FileStore/example_data", ignore.stdout = TRUE)
cat("   ✓ Directories created\n\n")

============================================================================
3. DEFINE CORPORATE COLOR PALETTE
============================================================================

Professional corporate color scheme
corporate_colors <- list(
  primary = "#1f4788",      # Deep Blue
  secondary = "#71a5de",    # Light Blue
  accent1 = "#f39c12",      # Orange
  accent2 = "#27ae60",      # Green
  accent3 = "#e74c3c",      # Red
  neutral = "#95a5a6",      # Gray
  success = "#27ae60",      # Green
  warning = "#f39c12",      # Orange
  danger = "#e74c3c"        # Red
)

palette_main <- c(corporate_colors$primary, corporate_colors$secondary, 
                  corporate_colors$accent1, corporate_colors$accent2, 
                  corporate_colors$accent3)

cat("3. Color palette defined\n")
cat("   Primary: ", corporate_colors$primary, "\n")
cat("   Secondary: ", corporate_colors$secondary, "\n\n")

============================================================================
4. CREATE PROFESSIONAL GGPLOT CHARTS
============================================================================

cat("4. Creating professional charts...\n\n")

--------------------------------------------------------------------------
CHART 1: Revenue vs Target - Line Chart
--------------------------------------------------------------------------
cat("   a) Creating Revenue vs Target line chart...\n")

Reshape data for plotting
monthly_long <- monthly_sales %>%
  select(Month, Revenue, Target) %>%
  pivot_longer(cols = c(Revenue, Target), 
               names_to = "Type", 
               values_to = "Amount")

chart1_revenue_trend <- ggplot(monthly_long, aes(x = Month, y = Amount, 
                                                  color = Type, group = Type)) +
  geom_line(size = 2) +
  geom_point(size = 4) +
  scale_color_manual(values = c("Revenue" = corporate_colors$primary, 
                                "Target" = corporate_colors$accent1)) +
  scale_y_continuous(labels = dollar_format(prefix = "$", suffix = "M"),
                     breaks = seq(0, 8, 1),
                     limits = c(0, 8)) +
  labs(
    title = "Monthly Revenue vs Target",
    subtitle = "2024 Performance",
    x = "Month",
    y = "Amount (Millions)",
    color = "Metric"
  ) +
  theme_minimal(base_size = 14) +
  theme(
    plot.title = element_text(size = 20, face = "bold", color = corporate_colors$primary),
    plot.subtitle = element_text(size = 14, color = "#666666"),
    axis.title = element_text(size = 13, face = "bold"),
    axis.text = element_text(size = 11),
    legend.position = "top",
    legend.title = element_text(size = 12, face = "bold"),
    legend.text = element_text(size = 11),
    panel.grid.minor = element_blank(),
    panel.grid.major = element_line(color = "#e0e0e0", size = 0.3),
    plot.margin = margin(20, 20, 20, 20)
  )

ggsave("/dbfs/FileStore/example_charts/01_revenue_trend.png", 
       chart1_revenue_trend,
       width = 10, height = 6, dpi = 150, bg = "white")
cat("      ✓ Saved: 01_revenue_trend.png\n")

--------------------------------------------------------------------------
CHART 2: Product Performance - Grouped Bar Chart
--------------------------------------------------------------------------
cat("   b) Creating Product Performance comparison...\n")

Reshape for grouped bar
product_long <- product_performance %>%
  select(Product, Q3_Sales, Q4_Sales) %>%
  pivot_longer(cols = c(Q3_Sales, Q4_Sales),
               names_to = "Quarter",
               values_to = "Sales") %>%
  mutate(Quarter = recode(Quarter, "Q3_Sales" = "Q3 2024", "Q4_Sales" = "Q4 2024"))

chart2_product_compare <- ggplot(product_long, 
                                  aes(x = reorder(Product, Sales), y = Sales, fill = Quarter)) +
  geom_bar(stat = "identity", position = "dodge", width = 0.7) +
  geom_text(aes(label = paste0("$", Sales, "M")), 
            position = position_dodge(width = 0.7),
            vjust = -0.5, size = 3.5, fontface = "bold") +
  scale_fill_manual(values = c("Q3 2024" = corporate_colors$neutral,
                               "Q4 2024" = corporate_colors$primary)) +
  scale_y_continuous(labels = dollar_format(prefix = "$", suffix = "M"),
                     limits = c(0, 4.5),
                     expand = expansion(mult = c(0, 0.1))) +
  labs(
    title = "Product Performance Comparison",
    subtitle = "Q3 vs Q4 2024",
    x = "Product",
    y = "Sales (Millions)",
    fill = "Quarter"
  ) +
  theme_minimal(base_size = 14) +
  theme(
    plot.title = element_text(size = 20, face = "bold", color = corporate_colors$primary),
    plot.subtitle = element_text(size = 14, color = "#666666"),
    axis.title = element_text(size = 13, face = "bold"),
    axis.text = element_text(size = 11),
    legend.position = "top",
    legend.title = element_text(size = 12, face = "bold"),
    panel.grid.major.x = element_blank(),
    panel.grid.minor = element_blank(),
    panel.grid.major.y = element_line(color = "#e0e0e0", size = 0.3),
    plot.margin = margin(20, 20, 20, 20)
  )

ggsave("/dbfs/FileStore/example_charts/02_product_performance.png",
       chart2_product_compare,
       width = 10, height = 6, dpi = 150, bg = "white")
cat("      ✓ Saved: 02_product_performance.png\n")

--------------------------------------------------------------------------
CHART 3: Regional Distribution - Pie Chart
--------------------------------------------------------------------------
cat("   c) Creating Regional Distribution pie chart...\n")

chart3_regional <- ggplot(regional_sales, 
                          aes(x = "", y = Sales, fill = reorder(Region, -Sales))) +
  geom_bar(stat = "identity", width = 1, color = "white", size = 2) +
  coord_polar("y", start = 0) +
  geom_text(aes(label = paste0(Region, "\n", Percentage, "%")),
            position = position_stack(vjust = 0.5),
            size = 4, fontface = "bold", color = "white") +
  scale_fill_manual(values = palette_main) +
  labs(title = "Revenue Distribution by Region",
       subtitle = "Q4 2024") +
  theme_void(base_size = 14) +
  theme(
    plot.title = element_text(size = 20, face = "bold", 
                             color = corporate_colors$primary, hjust = 0.5),
    plot.subtitle = element_text(size = 14, color = "#666666", hjust = 0.5),
    legend.position = "none",
    plot.margin = margin(20, 20, 20, 20)
  )

ggsave("/dbfs/FileStore/example_charts/03_regional_distribution.png",
       chart3_regional,
       width = 10, height = 7, dpi = 150, bg = "white")
cat("      ✓ Saved: 03_regional_distribution.png\n")

--------------------------------------------------------------------------
CHART 4: Customer Segments - Stacked Bar Chart
--------------------------------------------------------------------------
cat("   d) Creating Customer Segments stacked chart...\n")

chart4_segments <- ggplot(customer_segments, 
                          aes(x = Month, y = Revenue, fill = Segment)) +
  geom_bar(stat = "identity", position = "stack", width = 0.6) +
  geom_text(aes(label = paste0("$", Revenue, "M")),
            position = position_stack(vjust = 0.5),
            size = 3.5, fontface = "bold", color = "white") +
  scale_fill_manual(values = c("Enterprise" = corporate_colors$primary,
                               "SMB" = corporate_colors$secondary,
                               "Startup" = corporate_colors$accent1)) +
  scale_y_continuous(labels = dollar_format(prefix = "$", suffix = "M"),
                     expand = expansion(mult = c(0, 0.1))) +
  labs(
    title = "Revenue by Customer Segment",
    subtitle = "Q4 2024 Monthly Breakdown",
    x = "Month",
    y = "Revenue (Millions)",
    fill = "Customer Segment"
  ) +
  theme_minimal(base_size = 14) +
  theme(
    plot.title = element_text(size = 20, face = "bold", color = corporate_colors$primary),
    plot.subtitle = element_text(size = 14, color = "#666666"),
    axis.title = element_text(size = 13, face = "bold"),
    axis.text = element_text(size = 11),
    legend.position = "top",
    legend.title = element_text(size = 12, face = "bold"),
    panel.grid.major.x = element_blank(),
    panel.grid.minor = element_blank(),
    plot.margin = margin(20, 20, 20, 20)
  )

ggsave("/dbfs/FileStore/example_charts/04_customer_segments.png",
       chart4_segments,
       width = 10, height = 6, dpi = 150, bg = "white")
cat("      ✓ Saved: 04_customer_segments.png\n")

--------------------------------------------------------------------------
CHART 5: Department Budget - Horizontal Bar Chart
--------------------------------------------------------------------------
cat("   e) Creating Department Budget horizontal bar...\n")

Reshape for grouped horizontal bars
dept_long <- department_budget %>%
  pivot_longer(cols = c(Budget, Actual),
               names_to = "Type",
               values_to = "Amount")

chart5_budget <- ggplot(dept_long, 
                        aes(x = Amount, y = reorder(Department, Amount), fill = Type)) +
  geom_bar(stat = "identity", position = "dodge", width = 0.7) +
  geom_text(aes(label = paste0("$", Amount, "M")),
            position = position_dodge(width = 0.7),
            hjust = -0.1, size = 3.5, fontface = "bold") +
  scale_fill_manual(values = c("Budget" = corporate_colors$neutral,
                               "Actual" = corporate_colors$success)) +
  scale_x_continuous(labels = dollar_format(prefix = "$", suffix = "M"),
                     limits = c(0, 5.5),
                     expand = expansion(mult = c(0, 0.1))) +
  labs(
    title = "Department Budget vs Actual Spend",
    subtitle = "Q4 2024",
    x = "Amount (Millions)",
    y = "Department",
    fill = "Type"
  ) +
  theme_minimal(base_size = 14) +
  theme(
    plot.title = element_text(size = 20, face = "bold", color = corporate_colors$primary),
    plot.subtitle = element_text(size = 14, color = "#666666"),
    axis.title = element_text(size = 13, face = "bold"),
    axis.text = element_text(size = 11),
    legend.position = "top",
    legend.title = element_text(size = 12, face = "bold"),
    panel.grid.major.y = element_blank(),
    panel.grid.minor = element_blank(),
    plot.margin = margin(20, 20, 20, 20)
  )

ggsave("/dbfs/FileStore/example_charts/05_department_budget.png",
       chart5_budget,
       width = 10, height = 6, dpi = 150, bg = "white")
cat("      ✓ Saved: 05_department_budget.png\n")

--------------------------------------------------------------------------
CHART 6: Growth Indicators - Lollipop Chart
--------------------------------------------------------------------------
cat("   f) Creating Product Growth lollipop chart...\n")

chart6_growth <- ggplot(product_performance %>% arrange(Growth), 
                        aes(x = reorder(Product, Growth), y = Growth)) +
  geom_segment(aes(x = Product, xend = Product, y = 0, yend = Growth),
               color = corporate_colors$primary, size = 2) +
  geom_point(color = corporate_colors$accent2, size = 8) +
  geom_text(aes(label = paste0("+", Growth, "%")),
            hjust = -0.3, size = 4, fontface = "bold",
            color = corporate_colors$primary) +
  scale_y_continuous(labels = percent_format(scale = 1),
                     limits = c(0, 40),
                     expand = expansion(mult = c(0, 0.1))) +
  coord_flip() +
  labs(
    title = "Product Growth Rate",
    subtitle = "Q4 vs Q3 2024",
    x = "Product",
    y = "Growth Rate (%)"
  ) +
  theme_minimal(base_size = 14) +
  theme(
    plot.title = element_text(size = 20, face = "bold", color = corporate_colors$primary),
    plot.subtitle = element_text(size = 14, color = "#666666"),
    axis.title = element_text(size = 13, face = "bold"),
    axis.text = element_text(size = 11),
    panel.grid.major.y = element_blank(),
    panel.grid.minor = element_blank(),
    plot.margin = margin(20, 40, 20, 20)
  )

ggsave("/dbfs/FileStore/example_charts/06_growth_rates.png",
       chart6_growth,
       width = 10, height = 6, dpi = 150, bg = "white")
cat("      ✓ Saved: 06_growth_rates.png\n")

============================================================================
5. SAVE DATA AS CSV FOR PYTHON ACCESS
============================================================================

cat("\n5. Saving data files...\n")

write.csv(summary_metrics, "/dbfs/FileStore/example_data/summary_metrics.csv", row.names = FALSE)
write.csv(monthly_sales, "/dbfs/FileStore/example_data/monthly_sales.csv", row.names = FALSE)
write.csv(product_performance, "/dbfs/FileStore/example_data/product_performance.csv", row.names = FALSE)
write.csv(regional_sales, "/dbfs/FileStore/example_data/regional_sales.csv", row.names = FALSE)
write.csv(customer_segments, "/dbfs/FileStore/example_data/customer_segments.csv", row.names = FALSE)
write.csv(department_budget, "/dbfs/FileStore/example_data/department_budget.csv", row.names = FALSE)

cat("   ✓ All data files saved\n\n")

============================================================================
6. SUMMARY
============================================================================

cat("=" * 80, "\n")
cat("R SCRIPT COMPLETE!\n")
cat("=" * 80, "\n\n")

cat("✅ Charts Created (6):\n")
cat("   1. Revenue vs Target (line chart)\n")
cat("   2. Product Performance (grouped bar)\n")
cat("   3. Regional Distribution (pie chart)\n")
cat("   4. Customer Segments (stacked bar)\n")
cat("   5. Department Budget (horizontal bar)\n")
cat("   6. Product Growth (lollipop chart)\n\n")

cat("✅ Data Files Created (6):\n")
cat("   - summary_metrics.csv\n")
cat("   - monthly_sales.csv\n")
cat("   - product_performance.csv\n")
cat("   - regional_sales.csv\n")
cat("   - customer_segments.csv\n")
cat("   - department_budget.csv\n\n")

cat("📂 Output Locations:\n")
cat("   Charts: /dbfs/FileStore/example_charts/\n")
cat("   Data:   /dbfs/FileStore/example_data/\n\n")

cat("▶️  Next Step: Run the Python script to create PowerPoint presentation\n")
cat("=" * 80, "\n")
