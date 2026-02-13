# Troubleshooting Guide: Positioning Charts in Custom PowerPoint Templates

## 🎯 Common Issues & Solutions

---

## Issue 1: "I don't know where my chart is appearing"

### Solution: Add a border to see the chart

```python
# Temporarily add this to see where your chart lands
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor

# After adding your picture:
pic = slide.shapes.add_picture(chart_path, Inches(1), Inches(2), width=Inches(8))

# Add a temporary border to see it
pic.line.color.rgb = RGBColor(255, 0, 0)  # Red border
pic.line.width = Pt(3)
```

---

## Issue 2: "Chart overlaps with template elements"

### Diagnosis: Find what's in your template

```python
from pptx import Presentation

prs = Presentation("/dbfs/FileStore/templates/your_template.pptx")
slide = prs.slides[0]  # First slide in template

print("Template elements:")
for idx, shape in enumerate(slide.shapes):
    print(f"\n{idx}. {shape.name}")
    if hasattr(shape, 'left'):
        print(f"   Position: left={shape.left.inches:.2f}\", top={shape.top.inches:.2f}\"")
        print(f"   Size: {shape.width.inches:.2f}\" × {shape.height.inches:.2f}\"")
```

### Solution: Adjust chart position to avoid overlaps

```python
# If template has logo at top-right (8", 0", 1.5" × 0.5")
# Make sure chart doesn't go past 8" horizontally

slide.shapes.add_picture(
    chart_path,
    Inches(1.0),      # Start 1" from left
    Inches(1.5),      # Start below logo
    width=Inches(6.5)  # End at 7.5" (before 8" logo starts)
)
```

---

## Issue 3: "Chart is too small/too large"

### Diagnosis: Check your chart's aspect ratio

```r
# In R, when creating the chart:
ggsave("/dbfs/FileStore/charts/my_chart.png", my_chart,
       width = 10, height = 6, dpi = 150)

# Aspect ratio = width/height = 10/6 = 1.67
```

### Solution: Match PowerPoint size to aspect ratio

```python
# If R chart is 10×6 (aspect 1.67)
chart_width = 8.0  # inches in PowerPoint
chart_height = chart_width / 1.67  # = 4.8 inches

# PowerPoint auto-scales height, but you can specify both:
slide.shapes.add_picture(
    chart_path,
    Inches(1.0),
    Inches(1.5),
    width=Inches(chart_width),
    height=Inches(chart_height)
)
```

---

## Issue 4: "I don't know my template's placeholder indices"

### Solution: Run this diagnostic

```python
from pptx import Presentation

TEMPLATE_PATH = "/dbfs/FileStore/templates/your_template.pptx"
prs = Presentation(TEMPLATE_PATH)

# Check each layout
for layout_idx, layout in enumerate(prs.slide_layouts):
    print(f"\n========================================")
    print(f"LAYOUT {layout_idx}: {layout.name}")
    print(f"========================================")
    
    for ph_idx, placeholder in enumerate(layout.placeholders):
        print(f"  Placeholder [{ph_idx}]: {placeholder.name}")
        print(f"     Type: {placeholder.placeholder_format.type}")
        print(f"     Position: ({placeholder.left.inches:.2f}\", {placeholder.top.inches:.2f}\")")
        print(f"     Size: {placeholder.width.inches:.2f}\" × {placeholder.height.inches:.2f}\"")
        print()
```

**Save this output and use it to update your script!**

---

## Issue 5: "Text doesn't appear where I want it"

### Solution: Use manual text boxes with exact coordinates

```python
# Instead of using placeholders:
text_box = slide.shapes.add_textbox(
    Inches(1.0),    # Left position
    Inches(0.5),    # Top position
    Inches(8.0),    # Width
    Inches(0.6)     # Height
)

text_frame = text_box.text_frame
text_frame.text = "Your Title Here"
text_frame.word_wrap = True  # Wrap long text
text_frame.margin_left = Inches(0.1)  # Internal padding

# Format the text
para = text_frame.paragraphs[0]
para.font.size = Pt(32)
para.font.bold = True
para.alignment = PP_ALIGN.CENTER
```

---

## Issue 6: "Chart appears behind template elements"

### Solution: Change z-order (bring to front)

```python
# Method 1: Add chart last (it will be on top)
# First add all template content, then add chart

# Method 2: Access and reorder (advanced)
# Charts added later appear on top automatically
```

---

## Issue 7: "Different slides need different positions"

### Solution: Create a positioning configuration

```python
# Configuration dictionary for each slide type
POSITIONS = {
    'title_slide': {
        'chart': {'left': 2.0, 'top': 2.5, 'width': 6.0}
    },
    'content_slide': {
        'chart': {'left': 1.0, 'top': 1.8, 'width': 8.0}
    },
    'comparison_slide': {
        'chart1': {'left': 0.5, 'top': 1.5, 'width': 4.5},
        'chart2': {'left': 5.0, 'top': 1.5, 'width': 4.5}
    }
}

# Use it:
def add_chart(slide, chart_path, slide_type, chart_name='chart'):
    pos = POSITIONS[slide_type][chart_name]
    slide.shapes.add_picture(
        chart_path,
        Inches(pos['left']),
        Inches(pos['top']),
        width=Inches(pos['width'])
    )

# Example:
add_chart(slide, "/dbfs/FileStore/charts/pie.png", 'content_slide')
```

---

## Issue 8: "How do I know if my coordinates are right?"

### Solution: Create a test slide with measurements

```python
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor

prs = Presentation()
slide = prs.slides.add_slide(prs.slide_layouts[6])  # Blank

# Draw grid
for i in range(11):  # 0-10 inches
    # Vertical line
    line = slide.shapes.add_connector(1, Inches(i), Inches(0), Inches(i), Inches(5.625))
    line.line.color.rgb = RGBColor(200, 200, 200)
    
    # Label
    label = slide.shapes.add_textbox(Inches(i), Inches(0.1), Inches(0.5), Inches(0.3))
    label.text = f"{i}\""
    label.text_frame.paragraphs[0].font.size = Pt(10)

for i in range(6):  # 0-5.625 inches
    # Horizontal line
    line = slide.shapes.add_connector(1, Inches(0), Inches(i), Inches(10), Inches(i))
    line.line.color.rgb = RGBColor(200, 200, 200)
    
    # Label
    label = slide.shapes.add_textbox(Inches(0.1), Inches(i), Inches(0.5), Inches(0.3))
    label.text = f"{i}\""
    label.text_frame.paragraphs[0].font.size = Pt(10)

# Now add your chart and see exactly where it lands
slide.shapes.add_picture(
    "/dbfs/FileStore/charts/test.png",
    Inches(1.0),
    Inches(1.5),
    width=Inches(8.0)
)

prs.save("/dbfs/FileStore/presentations/measurement_test.pptx")
```

---

## Issue 9: "R chart doesn't fit in PowerPoint"

### Solution: Match R dimensions to PowerPoint

#### In R (create chart):
```r
# Create chart with PowerPoint-friendly dimensions
ggsave(
    "/dbfs/FileStore/charts/my_chart.png",
    my_chart,
    width = 8,      # Match PowerPoint width
    height = 4.8,   # Match PowerPoint height (width/1.67)
    dpi = 150,      # Good resolution for PowerPoint
    bg = "white"    # White background
)
```

#### In Python (add to slide):
```python
# Use exact same size
slide.shapes.add_picture(
    "/dbfs/FileStore/charts/my_chart.png",
    Inches(1.0),
    Inches(1.5),
    width=Inches(8.0)  # Same as R width
)
```

---

## Issue 10: "Template has weird placeholder positions"

### Solution: Ignore placeholders, use manual positioning

```python
def add_slide_with_manual_positioning(prs, layout_idx, title, chart_path):
    """
    Bypass template placeholders entirely
    """
    slide = prs.slides.add_slide(prs.slide_layouts[layout_idx])
    
    # Remove all placeholders
    for shape in slide.shapes:
        if shape.is_placeholder:
            sp = shape.element
            sp.getparent().remove(sp)
    
    # Add title manually at fixed position
    title_box = slide.shapes.add_textbox(
        Inches(0.5), Inches(0.3),
        Inches(9), Inches(0.6)
    )
    title_box.text_frame.text = title
    title_box.text_frame.paragraphs[0].font.size = Pt(32)
    title_box.text_frame.paragraphs[0].font.bold = True
    
    # Add chart at fixed position
    slide.shapes.add_picture(
        chart_path,
        Inches(1.0),
        Inches(1.2),
        width=Inches(8.0)
    )
    
    return slide
```

---

## Best Practices Checklist

✅ **Before positioning:**
- [ ] Run template analysis to find placeholder indices
- [ ] Check R chart dimensions (width, height, aspect ratio)
- [ ] Note any fixed elements in template (logos, headers)

✅ **While positioning:**
- [ ] Start with one slide type, get it perfect
- [ ] Use consistent margins (e.g., 0.5" or 1.0")
- [ ] Test on actual PowerPoint (not just preview)

✅ **For debugging:**
- [ ] Add temporary borders to see chart boundaries
- [ ] Create grid test slide to verify positions
- [ ] Print placeholder info for each layout

✅ **Production tips:**
- [ ] Store positions in a config dictionary
- [ ] Use helper functions for common patterns
- [ ] Document which layout indices you're using

---

## Quick Command Reference

### Find placeholder positions:
```python
for i, ph in enumerate(slide.placeholders):
    print(f"[{i}] {ph.name}: ({ph.left.inches:.2f}\", {ph.top.inches:.2f}\")")
```

### Add chart below title:
```python
title = slide.placeholders[0]
chart_top = title.top.inches + title.height.inches + 0.3
slide.shapes.add_picture(path, Inches(1), Inches(chart_top), width=Inches(8))
```

### Center chart horizontally:
```python
chart_width = 8.0
left = (10 - chart_width) / 2
slide.shapes.add_picture(path, Inches(left), Inches(2), width=Inches(chart_width))
```

### Fill placeholder area:
```python
ph = slide.placeholders[1]
slide.shapes.add_picture(path, ph.left, ph.top, width=ph.width)
```

---

## Need More Help?

1. **Run the diagnostic script**: `ppt_positioning_guide.py`
2. **Try example patterns**: `ppt_positioning_patterns.py`
3. **Check template analysis output** for exact coordinates
4. **Create test slides** with grid overlay to verify positions

---

**Remember:** Every template is different. There's no universal solution, but these tools will help you find the right positions for YOUR specific template!
