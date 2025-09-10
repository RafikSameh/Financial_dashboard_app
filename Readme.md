# Cash Flow Forecast Dashboard

A comprehensive PyQt6-based desktop application for analyzing and visualizing cash flow data from Excel files. This dashboard provides interactive charts, KPI tracking, and detailed data analysis for financial forecasting.

## Features

### Core Functionality
- **Excel Data Import**: Load cash flow data from Excel files with automatic data processing
- **Interactive Dashboard**: Real-time visualization of cash flow metrics and trends
- **Multiple Chart Types**: Waterfall charts, bar charts, pie charts, and time series analysis
- **KPI Tracking**: Key performance indicators with automated calculations
- **PDF Report Export**: Generate comprehensive reports with charts and data tables

### Advanced Analytics
- **Operating Cash Flow Analysis**: Detailed breakdown of inflows and outflows
- **Multi-Country Support**: Separate analysis for Egypt (EG) and KSA operations
- **Time Period Selection**: Interactive range sliders for custom date filtering
- **Category-based Filtering**: Dynamic filtering by cash flow categories
- **Trend Analysis**: Month-over-month cash flow visualization

## Technology Stack

- **Python 3.8+**
- **PyQt6**: Modern GUI framework
- **Pandas**: Data manipulation and analysis
- **Plotly**: Interactive data visualization
- **ReportLab**: PDF report generation
- **NumPy**: Numerical computing

## Installation

### Prerequisites
```bash
pip install pandas numpy plotly PyQt6 reportlab openpyxl
```

### Optional Dependencies
For enhanced chart export capabilities:
```bash
pip install kaleido  # For static image export
```

## File Structure

```
cash-flow-dashboard/
├── Data_Handler.py          # Core data processing class
├── utils.py                 # Plotting utilities and chart generation
├── pyqt_dash3.py           # Main PyQt6 application
└── README.md               # This file
```

## Usage

### Running the Application
```bash
python pyqt_dash3.py
```

### Loading Data
1. Click "Load Excel File" button
2. Select your Excel file containing cash flow data
3. The application expects a sheet named "Cash Flow - User Input"
4. Data will be automatically processed and charts will update

### Expected Excel Format
The Excel file should contain:
- Sheet: "Cash Flow - User Input" 
- Columns: Category, Country, Item, Cash Flow Type, followed by monthly data
- Specific rows for totals and balance calculations
- Categories numbered 1-4 for different cash flow types

### Navigation
- **Dashboard Tab**: Overview with KPIs and main charts
- **Cash Flow Data Tab**: Raw data table view
- **Operating Cash Inflow Tab**: Interactive pie charts with filtering
- **Operating Cash Outflow Tab**: Detailed outflow analysis

## Key Components

### Data_Handler Class
Handles Excel file processing and data preparation:
- Automatic sheet parsing and cleaning
- Cash flow categorization (Operating In/Out, Investing, Financing)
- KPI calculations and totals aggregation
- Multi-country data separation

### plot_diagrams Class
Generates interactive Plotly visualizations:
- **Waterfall Chart**: Cash movement summary
- **Monthly Cash Flow**: Stacked bar charts with trend lines
- **Operating Diagram**: Multi-view analysis with country filtering

### PyQt6 Dashboard
Main application interface featuring:
- Modern, responsive design
- Tabbed interface for different data views
- Interactive controls (sliders, dropdowns, date pickers)
- Real-time chart updates
- Export capabilities

## Interactive Features

### Operating Cash Inflow Analysis
- **Range Slider**: Select custom time periods
- **Category Filtering**: View all categories or drill down to specific items
- **Dual Visualization**: Pie chart for distribution, bar chart for trends
- **Real-time Updates**: Charts update automatically with filter changes

### Export Capabilities
- **PDF Reports**: Comprehensive reports with KPIs, charts, and data tables
- **Chart Images**: Individual chart export for presentations
- **Data Tables**: Formatted tables included in reports

## Data Processing Pipeline

1. **Excel Import**: Read and parse Excel file structure
2. **Data Cleaning**: Remove empty rows, standardize column names
3. **Categorization**: Classify cash flows by type and country
4. **Aggregation**: Calculate totals, subtotals, and KPIs
5. **Visualization**: Generate interactive charts and tables
6. **Export**: Create formatted PDF reports

## Customization

### Adding New Charts
Extend the `plot_diagrams` class:
```python
def new_chart_method(self, data):
    fig = go.Figure()
    # Add your chart logic
    return fig
```

### Modifying KPIs
Update the `update_kpis` method in the dashboard:
```python
def update_kpis(self):
    calc_dict = self.data_handler.calculations_dict
    # Add new KPI calculations
```

### Custom Styling
Modify the `apply_theme` method to change the application appearance:
```python
def apply_theme(self):
    self.setStyleSheet("""
        /* Your custom CSS styles */
    """)
```

## Error Handling

The application includes comprehensive error handling:
- Invalid Excel file formats
- Missing required sheets or columns
- Data type conversion errors
- Chart rendering failures
- PDF export issues

Error messages are displayed via Qt message boxes with detailed information.

## Performance Considerations

- **Background Loading**: Excel processing occurs in separate threads
- **Efficient Data Structures**: Pandas DataFrames for optimal performance
- **Chart Caching**: Plotly figures cached to avoid regeneration
- **Memory Management**: Proper widget cleanup and resource management

## Troubleshooting

### Common Issues

**"No data available" error:**
- Verify Excel file contains "Cash Flow - User Input" sheet
- Check that data starts from the expected row structure
- Ensure monthly columns contain valid date formats

**Charts not displaying:**
- Install required dependencies (plotly, kaleido)
- Check console output for JavaScript/WebEngine errors
- Verify data contains numeric values for visualization

**PDF export fails:**
- Ensure write permissions to target directory
- Install reportlab dependency
- Check available disk space

**Slow performance:**
- Large Excel files may take time to process
- Consider filtering date ranges for better performance
- Close unused tabs to free memory

## Contributing

When contributing to this project:
1. Follow PEP 8 style guidelines
2. Add docstrings to new methods
3. Test with various Excel file formats
4. Update this README for new features

## License

This project is provided as-is for internal use. Ensure compliance with all third-party library licenses.

## Support

For technical issues or feature requests, please:
1. Check the troubleshooting section above
2. Review console output for error details
3. Verify Excel file format compatibility
4. Test with sample data first