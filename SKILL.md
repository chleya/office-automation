---
name: office-automation
description: Create, edit, and manipulate Office documents (Word, Excel, PowerPoint) and WPS-compatible files.
homepage: https://github.com/chleya/office-automation
metadata: {"clawdbot":{"emoji":"📊","requires":{"bins":["python"],"python":["python-docx","openpyxl","python-pptx","PyPDF2","pandas"]},"install":[{"id":"pip","kind":"pip","package":"office-automation","label":"Install office-automation dependencies"}]}}
---

# Office Automation Skill

Automate creation, editing, and manipulation of Office documents (Word, Excel, PowerPoint) and WPS-compatible files.

## Features

### 📝 Word Document Processing
- Create new documents from templates
- Add/edit text with formatting
- Insert tables, images, and hyperlinks
- Apply styles and themes
- Generate reports and letters

### 📊 Excel Spreadsheet Processing
- Create and populate worksheets
- Perform calculations and formulas
- Generate charts and graphs
- Data analysis and filtering
- Export to various formats

### 🎯 PowerPoint Presentation Processing
- Create presentations from templates
- Add slides with content
- Insert images, charts, and media
- Apply transitions and animations
- Generate slide decks automatically

### 🔄 Format Conversion
- Convert between Office formats
- Export to PDF
- Batch processing of multiple files
- Template-based generation

### 🏢 WPS Compatibility
- Work with standard Office formats (.docx, .xlsx, .pptx)
- WPS can open all generated files
- Support for WPS-specific features when available

## Quick Start

### Basic Usage

```python
from office_automation import OfficeAutomation

# Initialize
office = OfficeAutomation()

# Create a Word document
doc = office.create_word_document("Report.docx")
doc.add_heading("Monthly Report", level=1)
doc.add_paragraph("This is the monthly performance report.")
doc.save()

# Create an Excel spreadsheet
sheet = office.create_excel_workbook("Data.xlsx")
sheet.add_worksheet("Sales")
sheet.set_cell("A1", "Month")
sheet.set_cell("B1", "Revenue")
sheet.save()

# Create a PowerPoint presentation
pres = office.create_presentation("Presentation.pptx")
pres.add_slide(title="Project Overview", content="Key project details")
pres.save()
```

### Advanced Examples

```python
# Generate a complete report
report = office.generate_report(
    title="Q4 Financial Report",
    data={"revenue": [100, 150, 200, 250], "expenses": [80, 90, 100, 110]},
    output_formats=["docx", "pdf", "pptx"]
)

# Process multiple files
processor = office.BatchProcessor()
processor.process_folder(
    input_folder="input/",
    output_folder="output/",
    operation="convert_to_pdf"
)

# Use templates
template = office.load_template("company_report_template.docx")
filled = template.fill({
    "company_name": "Acme Corp",
    "report_date": "2024-01-28",
    "data": sales_data
})
filled.save("final_report.docx")
```

## Installation

### Dependencies

```bash
# Core dependencies
pip install python-docx openpyxl python-pptx PyPDF2 pandas

# Optional dependencies for advanced features
pip install pillow  # Image processing
pip install matplotlib  # Chart generation
pip install reportlab  # PDF generation
```

### For Clawdbot Integration

1. Copy this skill folder to your Clawdbot skills directory:
   ```
   C:\Users\Administrator\AppData\Roaming\npm\node_modules\clawdbot\skills\
   ```

2. Or use symbolic link:
   ```bash
   mklink /D "C:\Users\Administrator\AppData\Roaming\npm\node_modules\clawdbot\skills\office-automation" "F:\skill\office-automation"
   ```

## API Reference

### OfficeAutomation Class

```python
class OfficeAutomation:
    def create_word_document(filename: str) -> WordDocument
    def create_excel_workbook(filename: str) -> ExcelWorkbook  
    def create_presentation(filename: str) -> Presentation
    
    def read_document(filename: str) -> Document
    def convert_format(input_file: str, output_format: str) -> str
    def batch_process(files: List[str], operation: str) -> List[str]
    
    # Template system
    def load_template(template_path: str) -> Template
    def register_template(name: str, template_path: str)
    def generate_from_template(template_name: str, data: dict) -> Document
```

### WordDocument Class

```python
class WordDocument:
    def add_heading(text: str, level: int = 1)
    def add_paragraph(text: str, style: str = None)
    def add_table(data: List[List[str]], headers: List[str] = None)
    def add_image(image_path: str, width: int = None, height: int = None)
    def apply_style(style_name: str)
    def save(filename: str = None)
```

### ExcelWorkbook Class

```python
class ExcelWorkbook:
    def add_worksheet(name: str) -> Worksheet
    def set_cell(cell: str, value: any, formula: str = None)
    def add_chart(data_range: str, chart_type: str = "line")
    def apply_formatting(range: str, format: dict)
    def calculate_formulas()
    def save(filename: str = None)
```

### Presentation Class

```python
class Presentation:
    def add_slide(layout: str = "title_and_content", **kwargs) -> Slide
    def add_image(slide_index: int, image_path: str, position: tuple = None)
    def add_chart(slide_index: int, data: dict, chart_type: str = "bar")
    def apply_transition(slide_index: int, transition: str)
    def save(filename: str = None)
```

## Examples

See the `examples/` directory for complete working examples:

1. `create_report.py` - Generate a complete business report
2. `process_spreadsheet.py` - Analyze and visualize data
3. `generate_presentation.py` - Create a presentation from data
4. `batch_converter.py` - Convert multiple files between formats

## Configuration

Create a `config.yaml` file for custom settings:

```yaml
office_automation:
  default_templates:
    report: "templates/report_template.docx"
    invoice: "templates/invoice_template.xlsx"
    presentation: "templates/presentation_template.pptx"
  
  output_formats:
    preferred: "docx"
    fallback: "pdf"
  
  wps_compatibility:
    enable: true
    check_wps_installed: true
  
  batch_processing:
    max_workers: 4
    timeout_seconds: 300
```

## WPS Specific Features

If WPS Office is detected, additional features may be available:

```python
# Check WPS availability
if office.wps_available:
    # Use WPS-specific features
    office.wps_convert_to_wps_format(input_file, output_file)
    office.wps_apply_wps_template(template_name)
```

## Error Handling

The skill includes comprehensive error handling:

```python
try:
    document = office.create_word_document("report.docx")
    # ... operations
except DocumentCreationError as e:
    print(f"Failed to create document: {e}")
except FormatError as e:
    print(f"Format error: {e}")
except WPSNotAvailableError as e:
    print(f"WPS not available, using standard Office formats: {e}")
```

## Performance Tips

1. **Batch Processing**: Use `batch_process()` for multiple files
2. **Template Caching**: Templates are cached for faster access
3. **Lazy Loading**: Documents are loaded only when needed
4. **Memory Management**: Large files are processed in chunks

## Contributing

Contributions are welcome! Please see `CONTRIBUTING.md` for guidelines.

## License

MIT License - see `LICENSE` file for details.

## Support

For issues and questions:
- GitHub Issues: https://github.com/chleya/office-automation/issues
- Email: [your-email@example.com]
- Documentation: https://github.com/chleya/office-automation/wiki