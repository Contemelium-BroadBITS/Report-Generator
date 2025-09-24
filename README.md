# Report Generator

A dynamic report generation system that creates professional reports from YAML configurations and DOCX templates with embedded charts, tables, and text.

## 1. Installation and Setup

### Prerequisites

- Python 3.8 or higher
- Pandoc (optional, for enhanced PDF conversion)

### Installation Steps

1. Clone the repository or download the source code:

```bash
git clone https://github.com/Contemelium-BroadBITS/Report-Generator
cd Report-Generator
```

2. Install required dependencies:

```bash
pip install -r requirements.txt
```

3. Setup environment (optional but recommended):

```bash
# Create virtual environment
python -m venv venv

# Activate virtual environment
# On Windows:
venv\Scripts\activate
# On macOS/Linux:
source venv/bin/activate

# Install dependencies
pip install -r requirements.txt
```

4. Configure authentication:
   
   Edit the `OpExpertOperations.py` file to update credentials or set environment variables:

```python
# Using environment variables (recommended)
export OPEXPERT_USERNAME='your_username'
export OPEXPERT_PASSWORD='your_password'
export OPEXPERT_CRM_URL='your_crm_url'
```

5. Ensure proper folder structure:

```
project/
├── templates/     # Contains DOCX templates
├── reports/       # Generated reports will be saved here
├── *.yaml         # YAML configuration files
└── *.py           # Python source code
```

### Running the Report Generator

```bash
python ReportGenerator.py
```

## 2. How to Structure the YAML File

The YAML configuration defines the content and structure of the report. Each YAML file consists of a sequence of steps, where each step defines a component of the report.

### Basic Structure

```yaml
- step: 0  # Configuration step (always first)
  type: config
  cover_template_id: 'template_id'
  endcard_template_id: 'template_id'
  report_title: 'Your Report Title'
  report_description: 'Description of the report'
  default_palette: crimson_red  # Optional, default color palette
  
- step: 1  # Content step
  type: chart
  template_id: 'template_id'
  charts:
    - type: pie
      title: 'Chart Title'
      description: 'Chart Description'
      data: 'data_integration_id'
      column: 'column_name'  # Optional
```

### Available Step Types

1. **config**: Initial configuration for the report
   - Required at step 0
   - Defines global settings, cover page, end card, and default styles

2. **chart**: Add charts to the report
   - Can include multiple charts of different types
   - Each chart requires a title, description, and data source

3. **table**: Add tables to the report
   - Requires a data integration ID
   - Can specify which columns to include

4. **caption**: Add text captions
   - Simple text blocks with title and description

### Chart Types and Configuration

The following chart types are supported:

1. **pie**: Circular chart showing proportions
   ```yaml
   type: pie
   title: 'Pie Chart Title'
   description: 'Description'
   data: 'integration_id'
   column: 'column_name'  # Column to categorize data
   ```

2. **donut**: Similar to pie chart but with a hole in the center
   ```yaml
   type: donut
   title: 'Donut Chart Title'
   description: 'Description'
   data: 'integration_id'
   column: 'column_name'  # Column to categorize data
   ```

3. **bar**: Vertical bar chart
   ```yaml
   type: bar
   title: 'Bar Chart Title'
   description: 'Description'
   data: 'integration_id'
   abscissa: 'x_axis_column'  # X-axis data
   ordinate: 'y_axis_column'  # Y-axis data
   ```

4. **line**: Line chart showing trends over time or categories
   ```yaml
   type: line
   title: 'Line Chart Title'
   description: 'Description'
   data: 'integration_id'
   abscissa: 'x_axis_column'  # X-axis data
   ordinate: 'y_axis_column'  # Y-axis data
   ```

### Color Palettes

Available color palettes:
- crimson_red (default)
- ocean_blue
- violet
- emerald_green
- sunset_orange

You can specify a default palette for the entire report or override it for individual charts.

### Example YAML Configuration

```yaml
- step: 0
  type: config
  template_type: 'dynamic'
  cover_template_id: 'template-cover-id'
  endcard_template_id: 'template-endcard-id'
  report_title: 'Quarterly Sales Report'
  report_description: 'Analysis of Q3 2025 sales performance'
  default_palette: ocean_blue

- step: 1
  type: chart
  template_id: 'template-charts-id'
  charts:
    - type: pie
      title: 'Regional Sales Distribution'
      description: 'Sales distribution across different regions'
      data: 'sales-data-integration-id'
      column: 'region'
    - type: line
      title: 'Monthly Sales Trend'
      description: 'Sales trend over the past 12 months'
      data: 'monthly-sales-id'
      abscissa: 'month'
      ordinate: 'revenue'

- step: 2
  type: table
  template_id: 'template-table-id'
  title: 'Top Performing Products'
  description: 'Products with the highest revenue in Q3 2025'
  data: 'product-data-id'
  columns: ['product_name', 'units_sold', 'revenue', 'growth']
```

## 3. How to Structure the Template

Templates are Microsoft Word documents (.docx) with special placeholders that are replaced with dynamic content during report generation.

### Template Basics

1. Create a standard Word document (.docx)
2. Insert placeholders using the {{ variable_name }} syntax
3. For charts and tables, create a table cell with a specific placeholder format

### Common Placeholders

- `{{ report_title }}`: The title of the report (from YAML config)
- `{{ report_description }}`: The description of the report (from YAML config)
- `{{ current_date }}`: The date when the report is generated

### Chart Placeholders

For charts, create a table cell with a placeholder in this format:
```
{{ chart1_image }}
```

Where `chart1` corresponds to the first chart in your YAML configuration, `chart2` for the second, and so on.

### Table Structure

For tables, create a placeholder where you want the table to appear:
```
{{ table1 }}
```

### Template Organization

Create separate templates for different sections of your report:
1. **Cover template**: First page with title and basic information
2. **Content templates**: Various sections containing charts, tables, and text
3. **Endcard template**: Final page with conclusion or contact information

### Template IDs

Each template must have a unique ID, specified in the YAML configuration. Template files should be named with this ID and placed in the `templates/` directory.

Example:
- File name: `templates/c7f6af3c-09e0-1e14-b788-68b4ee963318.docx`
- Referenced in YAML as: `cover_template_id: 'c7f6af3c-09e0-1e14-b788-68b4ee963318'`

### Styling and Formatting

- Style and format your template as desired in Word
- CHarts will pick the closest color palette of the image's placeholder in the word template
- Charts will automatically default to the color palette specified in the YAML configuration
- The generator will respect your template's fonts, margins, and styles

### Tips for Effective Templates

1. Use consistent formatting throughout your templates
2. Create adequate space for charts and tables
3. Include headers and footers for professional appearance
4. Consider page breaks between sections
5. Use Word styles for consistent formatting

### Example Template Structure

A typical report might use these templates:
1. Cover page template with title, logo, and date
2. Executive summary template with key findings
3. Chart section templates with placeholders for different chart types
4. Table section templates for detailed data
5. Conclusion template with final remarks

---

## Support and Troubleshooting

For issues or questions, please contact the development team.

Common issues:
- Ensure all required Python packages are installed
- Check that template IDs in YAML match the actual template files
- Verify data integration IDs are correct and accessible
- For PDF conversion issues, ensure proper installation of required system packages