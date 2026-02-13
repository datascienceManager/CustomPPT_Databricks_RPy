# Complete Example: Corporate Quarterly 

## 🎯 Overview

This is a **complete, working example** that demonstrates:
- ✅ Creating realistic business data in R
- ✅ Generating professional ggplot2 charts
- ✅ Building a PowerPoint presentation with Python
- ✅ Positioning charts precisely in slides
- ✅ Works with OR without a custom template

**Theme:** Q4 2024 Business Performance Report

---

## 📁 Files Included

| File | Purpose | Language |
|------|---------|----------|
| `complete_example_r_data_charts.R` | Generate data & charts | R |
| `complete_example_python_ppt.py` | Create PowerPoint | Python |
| `create_sample_template.py` | Generate sample template | Python |
| `COMPLETE_EXAMPLE_README.md` | This file | Markdown |

---

## 🚀 Quick Start (3 Steps)

### Step 1: Run R Script (Creates Data & Charts)

```r
# In Databricks R cell:
%r
source("/Workspace/path/to/complete_example_r_data_charts.R")
```

**What it creates:**
- ✅ 6 professional charts (PNG files)
- ✅ 6 CSV data files
- ✅ Corporate color scheme
- ✅ Realistic business metrics

**Output locations:**
- Charts: `/dbfs/FileStore/example_charts/`
- Data: `/dbfs/FileStore/example_data/`

### Step 2: Run Python Script (Creates PowerPoint)

```python
# In Databricks Python cell:
%python
exec(open("/Workspace/path/to/complete_example_python_ppt.py").read())
```

**What it creates:**
- ✅ 10-slide professional presentation
- ✅ Title slide
- ✅ Executive summary with metric cards
- ✅ 6 chart slides
- ✅ Data table
- ✅ Closing slide

**Output location:**
- Presentation: `/dbfs/FileStore/presentations/Q4_2024_Business_Report.pptx`

### Step 3: Download & Review

Download from:
```
/FileStore/presentations/Q4_2024_Business_Report.pptx
```

**Done! You now have a complete professional presentation!**

---

## 📊 What You'll Get

### Slide 1: Title Slide
**"Q4 2024 BUSINESS PERFORMANCE"**
- Professional branded design
- Company logo placeholder
- Date and subtitle

### Slide 2: Executive Summary
**4 Metric Cards:**
- 📊 Quarterly Revenue: $15.2M (↑12.5%)
- 💰 Net Profit: $3.8M (↑8.3%)
- ⭐ Customer Satisfaction: 87% (↑2.1%)
- 📈 Market Share: 23.5% (↑1.8%)

### Slide 3: Revenue Performance
**Line chart** showing:
- Monthly revenue vs target
- 12 months of data
- Clear trend visualization

### Slide 4: Product Performance
**Grouped bar chart** comparing:
- 5 products (A, B, C, D, E)
- Q3 vs Q4 performance
- Growth indicators

### Slide 5: Regional Distribution
**Pie chart** showing:
- North America: 43.3%
- Europe: 32.0%
- Asia Pacific: 19.3%
- Latin America: 4.7%
- Middle East: 2.0%

### Slide 6: Customer Segments
**Stacked bar chart** with:
- Enterprise, SMB, Startup segments
- Monthly breakdown (Oct, Nov, Dec)
- Revenue by segment

### Slide 7: Department Budget
**Horizontal bar chart** showing:
- Budget vs Actual spend
- 5 departments
- Variance analysis

### Slide 8: Growth Rates
**Lollipop chart** displaying:
- Product growth rates
- Q4 vs Q3 comparison
- Visual emphasis on high performers

### Slide 9: Detailed Metrics
**Data table** with:
- Product names
- Q3 and Q4 sales
- Growth percentages
- Professional formatting

### Slide 10: Closing
**Thank You slide** with:
- Questions & Discussion
- Professional branded design

---

## 🎨 The Data (Realistic Business Metrics)

### Summary Metrics
```
Quarterly Revenue:        $15.2M  (↑12.5%)
Net Profit:               $3.8M   (↑8.3%)
Customer Satisfaction:    87%     (↑2.1%)
Market Share:             23.5%   (↑1.8%)
```

### Monthly Sales Trend
```
Jan: $4.2M → Dec: $7.6M
Consistent growth throughout the year
Target achievement: 101% average
```

### Product Performance
```
Product A: $2.8M → $3.5M (+25%)
Product B: $3.2M → $3.8M (+18.8%)
Product C: $1.9M → $2.2M (+15.8%)
Product D: $2.5M → $2.9M (+16%)
Product E: $1.6M → $2.1M (+31.3%)
```

### Regional Sales
```
North America: $6.5M (43.3%)
Europe:        $4.8M (32.0%)
Asia Pacific:  $2.9M (19.3%)
Latin America: $0.7M (4.7%)
Middle East:   $0.3M (2.0%)
```

---

## 🎨 Chart Specifications

All charts use a professional corporate color scheme:

| Color | Hex Code | Usage |
|-------|----------|-------|
| Deep Blue | #1f4788 | Primary (bars, lines, headers) |
| Light Blue | #71a5de | Secondary (accents, highlights) |
| Orange | #f39c12 | Accent (targets, warnings) |
| Green | #27ae60 | Success (positive growth) |
| Red | #e74c3c | Danger (negative indicators) |

**Chart Dimensions:**
- Width: 10 inches
- Height: 6 inches
- DPI: 150 (high quality)
- Background: White
- Font: ggplot2 default (clean, professional)

---

## 📐 Positioning Details

### Standard Positioning (Works Without Template)

```python
# Title bar
Position: (0, 0)
Size: 10" × 0.8"

# Chart
Position: (0.5", 1.3")
Size: 9" wide (height auto-scales)

# Metric cards
2×2 Grid:
- Top-left: (0.5", 1.3")
- Top-right: (5.25", 1.3")
- Bottom-left: (0.5", 2.9")
- Bottom-right: (5.25", 2.9")
Card size: 4.25" × 1.3"
```

### With Template Positioning

The script automatically:
1. Detects template placeholders
2. Uses placeholder dimensions
3. Positions charts in content areas
4. Falls back to standard positioning if needed

---

## 🛠️ Customization Guide

### Change Colors

**In R script** (line ~70):
```r
corporate_colors <- list(
  primary = "#YOUR_COLOR",      # Your primary brand color
  secondary = "#YOUR_COLOR",    # Your secondary color
  accent1 = "#YOUR_COLOR",      # Accent colors
  # ... etc
)
```

**In Python script** (line ~30):
```python
class CorporateColors:
    PRIMARY = RGBColor(31, 71, 136)     # Your RGB values
    SECONDARY = RGBColor(113, 165, 222)
    # ... etc
```

### Change Data

**Edit the R script** (lines ~20-80):
```r
# Summary Metrics
summary_metrics <- data.frame(
  Metric = c("Your Metric 1", "Your Metric 2", ...),
  Value = c(YOUR_VALUES),
  # ... etc
)

# Monthly Sales
monthly_sales <- data.frame(
  Month = c(...),
  Revenue = c(YOUR_REVENUE_DATA),
  # ... etc
)
```

### Add More Charts

**In R:**
```r
# Create your chart
my_new_chart <- ggplot(your_data, aes(...)) +
  geom_...() +
  theme_minimal() +
  labs(title = "Your Chart Title")

# Save it
ggsave("/dbfs/FileStore/example_charts/07_my_chart.png",
       my_new_chart,
       width = 10, height = 6, dpi = 150, bg = "white")
```

**In Python:**
```python
# Add a slide with your chart
add_content_slide_with_chart(
    prs,
    "YOUR CHART TITLE",
    f"{CHART_DIR}/07_my_chart.png"
)
```

### Use Your Template

**Option 1: Use the sample template generator**
```python
%python
exec(open("/path/to/create_sample_template.py").read())
# Download, customize in PowerPoint, re-upload
```

**Option 2: Upload your existing template**
1. Upload `your_template.pptx` to `/dbfs/FileStore/templates/`
2. In Python script, update line ~20:
   ```python
   TEMPLATE_PATH = "/dbfs/FileStore/templates/your_template.pptx"
   ```
3. Run the script!

---

## 🎓 Learning from This Example

### R Best Practices Demonstrated

1. **Professional ggplot2 styling**
   - Consistent color scheme
   - Clear titles and labels
   - Proper axis formatting
   - Clean themes

2. **Data organization**
   - Tidy data format
   - Meaningful variable names
   - Proper data types

3. **Chart export**
   - High DPI (150)
   - Consistent dimensions
   - White backgrounds
   - PNG format (universal)

### Python Best Practices Demonstrated

1. **Modular code**
   - Helper functions for common tasks
   - Reusable slide templates
   - Configuration at top

2. **Error handling**
   - File existence checks
   - Try-except blocks
   - Fallback positioning

3. **Professional design**
   - Consistent spacing
   - Corporate colors
   - Clear hierarchy
   - Proper text formatting

---

## 📋 Checklist

Before running:
- [ ] R libraries installed (dplyr, ggplot2, tidyr, scales)
- [ ] Python libraries installed (python-pptx, pandas)
- [ ] Databricks notebook ready
- [ ] Output directories accessible

After running:
- [ ] 6 chart files created
- [ ] 6 data CSV files created
- [ ] PowerPoint file generated
- [ ] Charts appear in correct positions
- [ ] All slides formatted properly

---

## 🐛 Troubleshooting

### Issue: "Charts not found"
**Solution:** Make sure R script completed successfully. Check:
```r
list.files("/dbfs/FileStore/example_charts/")
```

### Issue: "Data files not found"
**Solution:** Verify data was saved:
```r
list.files("/dbfs/FileStore/example_data/")
```

### Issue: "Charts are positioned incorrectly"
**Solution:** Adjust positioning in Python script (lines ~200-300):
```python
# Current:
slide.shapes.add_picture(chart_path, Inches(0.5), Inches(1.3), width=Inches(9))

# Adjust:
slide.shapes.add_picture(chart_path, 
                        Inches(YOUR_LEFT), 
                        Inches(YOUR_TOP), 
                        width=Inches(YOUR_WIDTH))
```

### Issue: "Template not working"
**Solution:** Set `USE_TEMPLATE = False` to create without template:
```python
USE_TEMPLATE = False  # Line ~20
```

---

## 🚀 Next Steps

1. **Customize the data** to match your actual business metrics
2. **Adjust colors** to match your brand
3. **Add your logo** to the template
4. **Create additional charts** for your specific needs
5. **Modify positioning** to fit your template perfectly

---

## 📞 Support

If you need help:
1. Check the troubleshooting section
2. Review positioning guide files
3. Examine the code comments
4. Test with sample data first

---

## 📄 File Structure

```
project/
├── R Scripts/
│   └── complete_example_r_data_charts.R
│
├── Python Scripts/
│   ├── complete_example_python_ppt.py
│   └── create_sample_template.py
│
├── Output (Auto-generated)/
│   ├── Charts/
│   │   ├── 01_revenue_trend.png
│   │   ├── 02_product_performance.png
│   │   ├── 03_regional_distribution.png
│   │   ├── 04_customer_segments.png
│   │   ├── 05_department_budget.png
│   │   └── 06_growth_rates.png
│   │
│   ├── Data/
│   │   ├── summary_metrics.csv
│   │   ├── monthly_sales.csv
│   │   ├── product_performance.csv
│   │   ├── regional_sales.csv
│   │   ├── customer_segments.csv
│   │   └── department_budget.csv
│   │
│   └── Presentations/
│       └── Q4_2024_Business_Report.pptx
│
└── Documentation/
    └── COMPLETE_EXAMPLE_README.md
```

---

**Version:** 1.0  
**Last Updated:** February 2025  
**Status:** Production Ready ✅

---

**This is a complete, working example. Just run it and you'll have a professional presentation in minutes!** 🎉
