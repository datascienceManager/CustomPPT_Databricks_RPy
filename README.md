# Sports Viewing Analytics - Using Custom PowerPoint Templates in Databricks

## 🎯 Overview

This guide shows you how to use **YOUR PERSONAL PowerPoint template** with the R + Python Databricks workflow. Your company branding, colors, fonts, and slide layouts will be preserved!

---

## 📋 Table of Contents

- [Quick Start](#quick-start)
- [Step-by-Step Setup](#step-by-step-setup)
- [Understanding Template Layouts](#understanding-template-layouts)
- [Customizing for Your Template](#customizing-for-your-template)
- [Troubleshooting](#troubleshooting)
- [Advanced Customization](#advanced-customization)

---

## 🚀 Quick Start

### Prerequisites

1. **Your PowerPoint Template** (.pptx file)
2. **R Script**: `databricks_r_data_generator.R` (unchanged)
3. **Python Script**: `databricks_python_ppt_builder_with_template.py` (NEW)

### 3-Step Setup

```python
# Step 1: Upload your template to Databricks
# Go to: Data → Add Data → Upload File
# Upload to: /FileStore/templates/your_template.pptx

# Step 2: Run R cell
%r
source("/Workspace/path/to/databricks_r_data_generator.R")

# Step 3: Run Python cell
%python
# Update TEMPLATE_PATH in the script first!
exec(open("/Workspace/path/to/databricks_python_ppt_builder_with_template.py").read())
```

---

## 📝 Step-by-Step Setup

### Step 1: Upload Your Template

#### Method A: Using Databricks UI

1. In Databricks workspace, click **"Data"** in left sidebar
2. Click **"Add Data"** or **"Create Table"** 
3. Click **"Upload File"**
4. Select your template: `company_template.pptx`
5. Upload to: `/FileStore/templates/`
6. Your template path will be: `/dbfs/FileStore/templates/company_template.pptx`

#### Method B: Using Databricks CLI

```bash
databricks fs cp company_template.pptx dbfs:/FileStore/templates/company_template.pptx
```

#### Method C: Using Python in Notebook

```python
# Upload file from your local machine
dbutils.fs.cp("file:/tmp/company_template.pptx", 
              "dbfs:/FileStore/templates/company_template.pptx")
```

### Step 2: Update Template Path in Script

Open `databricks_python_ppt_builder_with_template.py` and find this line (around line 28):

```python
# CHANGE THIS LINE:
TEMPLATE_PATH = "/dbfs/FileStore/templates/your_template.pptx"

# TO YOUR ACTUAL PATH:
TEMPLATE_PATH = "/dbfs/FileStore/templates/company_template.pptx"
```

### Step 3: Run the Scripts

**R Cell:**
```r
%r
source("/Workspace/Users/your.email@company.com/databricks_r_data_generator.R")
```

**Python Cell:**
```python
%python
exec(open("/Workspace/Users/your.email@company.com/databricks_python_ppt_builder_with_template.py").read())
```

---

## 🎨 Understanding Template Layouts

PowerPoint templates have different **slide layouts** (Title Slide, Title + Content, Blank, etc.). Each layout has **placeholders** where content goes.

### Finding Your Template's Layouts

Add this code to your Python script to see available layouts:

```python
from pptx import Presentation

prs = Presentation("/dbfs/FileStore/templates/company_template.pptx")

print("Available Slide Layouts:")
for idx, layout in enumerate(prs.slide_layouts):
    print(f"\nLayout {idx}: {layout.name}")
    print(f"  Placeholders:")
    for pidx, placeholder in enumerate(layout.placeholders):
        print(f"    [{pidx}] {placeholder.name}")
```

**Example Output:**
```
Layout 0: Title Slide
  Placeholders:
    [0] Title
    [1] Subtitle

Layout 1: Title and Content
  Placeholders:
    [0] Title
    [1] Content

Layout 2: Section Header
  Placeholders:
    [0] Title
    [1] Text

Layout 5: Title Only
  Placeholders:
    [0] Title

Layout 6: Blank
  Placeholders: (none)
```

### Understanding the Output

- **Layout Index** (0, 1, 2...): Reference number for each layout
- **Layout Name**: Human-readable name
- **Placeholder Index** [0], [1]...: Position where you add content
- **Placeholder Name**: What content goes there (Title, Content, etc.)

---

## ⚙️ Customizing for Your Template

### Step 1: Map Your Layouts

Based on the output above, update these constants in the script (around line 237):

```python
# CUSTOMIZE THESE FOR YOUR TEMPLATE:
LAYOUT_TITLE = 0           # Your title slide layout
LAYOUT_SECTION = 2         # Your section header layout
LAYOUT_TITLE_CONTENT = 1   # Your title + content layout
LAYOUT_TITLE_ONLY = 5      # Your title only layout (for charts)
LAYOUT_BLANK = 6           # Your blank layout
```

**Example: If your template has these layouts:**
- Layout 0: Company Title Slide → Set `LAYOUT_TITLE = 0`
- Layout 3: Content with Bullets → Set `LAYOUT_TITLE_CONTENT = 3`
- Layout 8: Chart Slide → Set `LAYOUT_TITLE_ONLY = 8`

### Step 2: Configure Placeholder Indices

Different templates have different placeholder arrangements. Update these sections:

#### Title Slide (around line 256):

```python
# EXAMPLE 1: Standard template
add_content_to_placeholder(slide1, 0, "YOUR TITLE", font_size=44, bold=True)
add_content_to_placeholder(slide1, 1, "Your Subtitle", font_size=18)

# EXAMPLE 2: If your template has different placeholders
add_content_to_placeholder(slide1, 2, "YOUR TITLE", font_size=44, bold=True)  # Title is placeholder 2
add_content_to_placeholder(slide1, 3, "Your Subtitle", font_size=18)         # Subtitle is placeholder 3
```

#### Content Slides (around line 280):

```python
# EXAMPLE 1: Standard template
add_content_to_placeholder(slide2, 0, "SLIDE TITLE", font_size=36, bold=True)  # Title
add_content_to_placeholder(slide2, 1, "Content here", font_size=16)           # Content

# EXAMPLE 2: If your template uses different indices
add_content_to_placeholder(slide2, 1, "SLIDE TITLE", font_size=36, bold=True)  # Title is placeholder 1
add_content_to_placeholder(slide2, 2, "Content here", font_size=16)           # Content is placeholder 2
```

### Step 3: Adjust Image Positions

Charts need to fit your template's content area. Adjust these values (around line 310):

```python
# BEFORE (default positioning):
add_image_to_slide(slide4, "/dbfs/FileStore/charts/sports_pie.png", 
                   left=1, top=1.5, width=8)

# AFTER (adjusted for your template):
add_image_to_slide(slide4, "/dbfs/FileStore/charts/sports_pie.png", 
                   left=0.5,   # Move left/right
                   top=2.0,    # Move up/down
                   width=9)    # Make wider/narrower
```

**Tips for positioning:**
- **left**: Distance from left edge (inches)
- **top**: Distance from top edge (inches)
- **width**: Image width (height scales proportionally)
- Standard slide is 10" wide × 5.625" tall (16:9)

---

## 🎨 Real-World Example: Corporate Template

Let's say your company template has:

```
Layout 0: Company Title (with logo)
  [0] Title
  [10] Subtitle

Layout 3: Content Slide
  [0] Title Bar
  [15] Content Area

Layout 7: Full Width Chart
  [0] Title
```

### Configure the Script:

```python
# 1. Update layout constants
LAYOUT_TITLE = 0
LAYOUT_TITLE_CONTENT = 3
LAYOUT_TITLE_ONLY = 7

# 2. Update title slide
add_content_to_placeholder(slide1, 0, "SPORTS ANALYTICS", font_size=44, bold=True)
add_content_to_placeholder(slide1, 10, "2025 Report", font_size=18)  # Placeholder 10!

# 3. Update content slides
add_content_to_placeholder(slide2, 0, "EXECUTIVE SUMMARY", font_size=36, bold=True)
add_content_to_placeholder(slide2, 15, metrics_text, font_size=16)   # Placeholder 15!

# 4. Adjust chart positioning for your template
add_image_to_slide(slide4, "/dbfs/FileStore/charts/sports_pie.png", 
                   left=0.3, top=1.8, width=9.4)  # Full width for your layout
```

---

## 🐛 Troubleshooting

### Issue 1: "Template not found"

**Error Message:**
```
⚠ Template not found at: /dbfs/FileStore/templates/your_template.pptx
```

**Solutions:**
1. Verify file was uploaded:
   ```python
   dbutils.fs.ls("/FileStore/templates/")
   ```

2. Check exact filename (case-sensitive!):
   ```python
   # Wrong: company_Template.pptx
   # Right: company_template.pptx
   ```

3. Verify path in script matches uploaded location

### Issue 2: "Placeholder not found"

**Error Message:**
```
Warning: Placeholder 1 not found
```

**Solution:**
Run the layout inspection code to find correct placeholder indices:

```python
from pptx import Presentation

prs = Presentation("/dbfs/FileStore/templates/company_template.pptx")
layout = prs.slide_layouts[1]  # Check your layout index

print(f"Layout: {layout.name}")
for idx, p in enumerate(layout.placeholders):
    print(f"  [{idx}] {p.name}")
```

Then update your script with the correct indices.

### Issue 3: Text doesn't fit in placeholder

**Solution 1: Reduce font size**
```python
add_content_to_placeholder(slide2, 1, content, font_size=14)  # Smaller font
```

**Solution 2: Use manual text boxes**
```python
# Remove placeholder restriction
text_box = slide2.shapes.add_textbox(
    Inches(1), Inches(2), Inches(8), Inches(3)  # Custom position & size
)
text_frame = text_box.text_frame
text_frame.text = "Your content here"
text_frame.word_wrap = True
```

### Issue 4: Chart overlaps template elements

**Solution:**
Adjust chart position and size:

```python
# BEFORE:
add_image_to_slide(slide, chart_path, left=1, top=1.5, width=8)

# AFTER: Move down and make smaller
add_image_to_slide(slide, chart_path, left=1.5, top=2.0, width=7)
```

### Issue 5: Template colors don't match data

Your template might use specific colors. Extract them:

```python
from pptx import Presentation

prs = Presentation("/dbfs/FileStore/templates/company_template.pptx")
slide = prs.slides[0]  # First slide

# Check colors used in template
for shape in slide.shapes:
    if hasattr(shape, 'fill'):
        print(f"Shape: {shape.name}")
        if shape.fill.type == 1:  # Solid fill
            rgb = shape.fill.fore_color.rgb
            print(f"  Color: RGB({rgb[0]}, {rgb[1]}, {rgb[2]})")
            print(f"  Hex: #{rgb[0]:02X}{rgb[1]:02X}{rgb[2]:02X}")
```

Then update R charts to match:

```r
# Update R script chart colors to match template
chart_colors <- c('#FF6B35', '#004E89', '#1B9AAA')  # Your company colors
```

---

## 🔥 Advanced Customization

### 1. Preserve Template Master Slides

Your template's master slides define the overall design. They're automatically preserved!

```python
# The script preserves:
# - Master slide designs
# - Theme colors
# - Theme fonts
# - Background graphics
# - Footer/header settings
```

### 2. Add Company Logo to Every Slide

```python
LOGO_PATH = "/dbfs/FileStore/images/company_logo.png"

def add_logo_to_slide(slide):
    """Add company logo to bottom right of every slide"""
    if os.path.exists(LOGO_PATH):
        slide.shapes.add_picture(
            LOGO_PATH,
            Inches(8.5),  # Right side
            Inches(5),    # Bottom
            width=Inches(1)
        )

# Use it after creating each slide:
slide2 = prs.slides.add_slide(get_layout(LAYOUT_TITLE_CONTENT))
add_logo_to_slide(slide2)
```

### 3. Use Template Theme Colors in Charts

Extract theme colors from template and use in R:

```python
from pptx import Presentation
from pptx.enum.dml import MSO_THEME_COLOR

prs = Presentation("/dbfs/FileStore/templates/company_template.pptx")

# Get theme colors
theme = prs.slide_master.theme
color_scheme = theme.color_scheme

print("Template Theme Colors:")
print(f"Accent 1: {color_scheme.accent1.rgb}")
print(f"Accent 2: {color_scheme.accent2.rgb}")
# ... etc
```

Save to CSV and read in R:

```r
# Read template colors
theme_colors <- read.csv("/dbfs/FileStore/config/template_colors.csv")
chart_colors <- theme_colors$hex_color

# Use in charts
ggplot(...) + scale_fill_manual(values = chart_colors)
```

### 4. Dynamic Slide Selection

Choose layouts based on content:

```python
def get_best_layout_for_content(content_type):
    """Select layout based on content type"""
    layout_map = {
        'title': LAYOUT_TITLE,
        'text_heavy': LAYOUT_TITLE_CONTENT,
        'chart': LAYOUT_TITLE_ONLY,
        'two_column': LAYOUT_TWO_COLUMN,
        'blank': LAYOUT_BLANK
    }
    return get_layout(layout_map.get(content_type, LAYOUT_TITLE_CONTENT))

# Use it:
slide = prs.slides.add_slide(get_best_layout_for_content('chart'))
```

### 5. Template Validation

Add validation before processing:

```python
def validate_template(template_path):
    """Validate template has required layouts"""
    try:
        prs = Presentation(template_path)
        
        required_layouts = {
            'title': 0,
            'content': 1,
            'blank': 6
        }
        
        if len(prs.slide_layouts) < max(required_layouts.values()):
            print("⚠ Warning: Template may not have all required layouts")
            return False
        
        print("✓ Template validation passed")
        return True
        
    except Exception as e:
        print(f"✗ Template validation failed: {e}")
        return False

# Use before creating presentation:
if validate_template(TEMPLATE_PATH):
    # Proceed with presentation creation
    pass
```

---

## 📊 Complete Example Workflow

### Your Template Structure:
```
company_template.pptx
├── Layout 0: Company Title Slide
│   ├── [0] Title
│   └── [1] Subtitle
├── Layout 2: Content Slide
│   ├── [0] Title
│   └── [1] Content
└── Layout 5: Chart Slide
    └── [0] Title
```

### Complete Configuration:

```python
# ============================================================================
# CONFIGURATION FOR YOUR TEMPLATE
# ============================================================================

TEMPLATE_PATH = "/dbfs/FileStore/templates/company_template.pptx"

# Layout indices
LAYOUT_TITLE = 0
LAYOUT_TITLE_CONTENT = 2
LAYOUT_TITLE_ONLY = 5
LAYOUT_BLANK = 6

# ============================================================================
# SLIDE CREATION
# ============================================================================

# Title Slide
slide1 = prs.slides.add_slide(get_layout(LAYOUT_TITLE))
add_content_to_placeholder(slide1, 0, "SPORTS ANALYTICS REPORT", 
                          font_size=44, bold=True)
add_content_to_placeholder(slide1, 1, "Q4 2025 Analysis", 
                          font_size=18)

# Content Slide
slide2 = prs.slides.add_slide(get_layout(LAYOUT_TITLE_CONTENT))
add_content_to_placeholder(slide2, 0, "EXECUTIVE SUMMARY", 
                          font_size=36, bold=True)
add_content_to_placeholder(slide2, 1, metrics_text, 
                          font_size=16)

# Chart Slide
slide3 = prs.slides.add_slide(get_layout(LAYOUT_TITLE_ONLY))
add_content_to_placeholder(slide3, 0, "VIEWERSHIP TRENDS", 
                          font_size=36, bold=True)
add_image_to_slide(slide3, "/dbfs/FileStore/charts/trends.png",
                  left=0.5, top=1.8, width=9)
```

---

## ✅ Checklist

Before running your script:

- [ ] Template uploaded to `/dbfs/FileStore/templates/`
- [ ] `TEMPLATE_PATH` updated in script
- [ ] Layout indices configured for your template
- [ ] Placeholder indices verified
- [ ] Chart positions adjusted for template
- [ ] R script generates required charts
- [ ] Test run completed successfully

---

## 📞 Common Template Scenarios

### Scenario 1: Corporate Template with Logo

**Features:**
- Company logo on every slide
- Specific brand colors
- Footer with company name

**Setup:**
```python
LOGO_PATH = "/dbfs/FileStore/images/logo.png"

def add_branded_slide(layout_idx, title):
    slide = prs.slides.add_slide(get_layout(layout_idx))
    add_content_to_placeholder(slide, 0, title, font_size=36, bold=True)
    # Logo is already in template master!
    return slide
```

### Scenario 2: University/Academic Template

**Features:**
- Department header
- Citation footer
- Specific fonts

**Setup:**
```python
# Template already has these, just use layouts
slide = prs.slides.add_slide(get_layout(LAYOUT_TITLE_CONTENT))
# Department header and footer automatically appear!
```

### Scenario 3: Client Presentation Template

**Features:**
- Client logo
- Custom color scheme
- Specific slide numbering

**Setup:**
```python
# Everything preserved from template
# Just add your content to placeholders
```

---

## 🎓 Learning Resources

- [python-pptx Documentation](https://python-pptx.readthedocs.io/)
- [Working with Slide Layouts](https://python-pptx.readthedocs.io/en/latest/user/slides.html)
- [Understanding Placeholders](https://python-pptx.readthedocs.io/en/latest/user/placeholders-understanding.html)

---

## 📄 File Checklist

✅ **Files You Need:**
1. `databricks_r_data_generator.R` - Creates data and charts (unchanged)
2. `databricks_python_ppt_builder_with_template.py` - Builds presentation with template (NEW)
3. `your_company_template.pptx` - Your PowerPoint template
4. `DATABRICKS_TEMPLATE_README.md` - This file

---

**Version:** 1.0  
**Last Updated:** February 2025  
**Compatible with:** Databricks Runtime 11.0+, PowerPoint 2016+
