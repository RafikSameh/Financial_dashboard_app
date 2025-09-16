# integrated_pyqt_dashboard.py - Complete version with PyQt Operating Inflow Pie Chart
import sys
import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import plotly.graph_objs as go
import plotly.io as pio
import json
import numbers

from PyQt6.QtWidgets import (QApplication, QMainWindow, QVBoxLayout, QHBoxLayout, 
                            QWidget, QPushButton, QLabel, QFileDialog, QTableWidget, 
                            QTableWidgetItem, QTabWidget, QGridLayout, QFrame,
                            QDateEdit, QComboBox, QMessageBox, QProgressBar,
                            QScrollArea, QSplitter, QSlider)
from PyQt6.QtCore import Qt, QDate, QThread, pyqtSignal, QTimer
from PyQt6.QtGui import QFont, QPalette, QColor, QPixmap, QIcon, QPainter, QBrush, QPen
from PyQt6.QtWebEngineWidgets import QWebEngineView
from plotly.subplots import make_subplots
from PyQt6.QtWidgets import *
from PyQt6.QtCore import *
import io
import base64
from datetime import datetime
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter, A4
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image, PageBreak
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from reportlab.pdfgen import canvas
import plotly.io as pio
import tempfile
import os
import warnings
warnings.filterwarnings("ignore")


# Import your custom classes
from Data_Handler import Data_Handler
from utils import plot_diagrams

class RangeSlider(QWidget):
    """Custom range slider widget for date selection"""
    valueChanged = pyqtSignal(tuple)
    
    def __init__(self, minimum=0, maximum=100, parent=None):
        super().__init__(parent)
        self.minimum = minimum
        self.maximum = maximum
        self.low = minimum
        self.high = maximum
        self._dragging_low = False
        self._dragging_high = False
        
        self.setFixedHeight(40)
        self.setMouseTracking(True)
        self.setStyleSheet("""
            QWidget {
                background-color: #f8f9fa;
                border: 1px solid #dee2e6;
                border-radius: 5px;
            }
        """)
        
    def setValue(self, low, high):
        """Set the range values"""
        self.low = max(self.minimum, min(low, self.maximum))
        self.high = max(self.minimum, min(high, self.maximum))
        if self.low > self.high:
            self.low, self.high = self.high, self.low
        self.update()
        
    def value(self):
        """Get current range values"""
        return (int(self.low), int(self.high))
    
    def paintEvent(self, event):
        """Custom paint event for the range slider"""
        painter = QPainter(self)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)
        
        # Track
        track_rect = self.rect().adjusted(10, 15, -10, -15)
        painter.fillRect(track_rect, QBrush(Qt.GlobalColor.lightGray))
        
        # Selected range
        if self.maximum > self.minimum:
            start_pos = 10 + (self.low - self.minimum) / (self.maximum - self.minimum) * (self.width() - 20)
            end_pos = 10 + (self.high - self.minimum) / (self.maximum - self.minimum) * (self.width() - 20)
            
            selected_rect = track_rect.adjusted(int(start_pos - 10), 0, int(end_pos - self.width() + 10), 0)
            painter.fillRect(selected_rect, QBrush(Qt.GlobalColor.blue))
            
            # Handles
            painter.setBrush(QBrush(Qt.GlobalColor.darkBlue))
            painter.drawEllipse(int(start_pos - 5), 10, 10, 10)
            painter.drawEllipse(int(end_pos - 5), 10, 10, 10)
    
    def mousePressEvent(self, event):
        """Handle mouse press for dragging handles"""
        if event.button() == Qt.MouseButton.LeftButton and self.maximum > self.minimum:
            pos = event.position().x()
            
            start_pos = 10 + (self.low - self.minimum) / (self.maximum - self.minimum) * (self.width() - 20)
            end_pos = 10 + (self.high - self.minimum) / (self.maximum - self.minimum) * (self.width() - 20)
            
            # Check which handle is closer
            if abs(pos - start_pos) < abs(pos - end_pos):
                self._dragging_low = True
            else:
                self._dragging_high = True
    
    def mouseMoveEvent(self, event):
        """Handle mouse move for dragging"""
        if (self._dragging_low or self._dragging_high) and self.maximum > self.minimum:
            pos = event.position().x()
            value = self.minimum + (pos - 10) / (self.width() - 20) * (self.maximum - self.minimum)
            value = max(self.minimum, min(value, self.maximum))
            
            if self._dragging_low:
                self.low = value
            elif self._dragging_high:
                self.high = value
                
            self.setValue(self.low, self.high)
            self.valueChanged.emit(self.value())
    
    def mouseReleaseEvent(self, event):
        """Handle mouse release"""
        self._dragging_low = False
        self._dragging_high = False

class OperatingInflowPieWidget(QWidget):
    """Widget for operating cash inflow pie chart with controls"""
    chart_updated = pyqtSignal(object)  # Emits the figure
    
    def __init__(self, data_handler=None, parent=None):
        super().__init__(parent)
        self.data_handler = data_handler
        self.df_long = None
        self.months = []
    
        self.setup_ui()
        self.setup_data()
        
    def setup_ui(self):
        """Setup the user interface"""
        layout = QVBoxLayout(self)
        
        # Title
        title_label = QLabel("Operating Cash Inflow Analysis")
        title_label.setFont(QFont("Arial", 14, QFont.Weight.Bold))
        title_label.setStyleSheet("padding: 10px; background-color: #f8f9fa; border-radius: 5px;")
        layout.addWidget(title_label)
        
        # Controls frame
        controls_frame = QFrame()
        controls_frame.setFrameStyle(QFrame.Shape.Box)
        controls_frame.setStyleSheet("""
            QFrame {
                background-color: #ffffff;
                border: 1px solid #dee2e6;
                border-radius: 8px;
                padding: 15px;
                margin: 5px;
            }
        """)
        controls_layout = QVBoxLayout(controls_frame)
        
        # Period selection
        period_layout = QVBoxLayout()
        period_label = QLabel("Select Period:")
        period_label.setFont(QFont("Arial", 10, QFont.Weight.Bold))
        period_layout.addWidget(period_label)
        
        # Range slider
        self.period_slider = RangeSlider()
        period_layout.addWidget(self.period_slider)
        
        # Period labels
        period_labels_layout = QHBoxLayout()
        self.period_start_label = QLabel("Start")
        self.period_end_label = QLabel("End")
        self.period_start_label.setStyleSheet("font-size: 10px; color: #6c757d;")
        self.period_end_label.setStyleSheet("font-size: 10px; color: #6c757d;")
        
        period_labels_layout.addWidget(self.period_start_label)
        period_labels_layout.addStretch()
        period_labels_layout.addWidget(self.period_end_label)
        period_layout.addLayout(period_labels_layout)
        
        controls_layout.addLayout(period_layout)
        
        # Category selection
        category_layout = QHBoxLayout()
        category_label = QLabel("Category:")
        category_label.setFont(QFont("Arial", 10, QFont.Weight.Bold))
        
        self.category_selector = QComboBox()
        self.category_selector.setStyleSheet("""
            QComboBox {
                padding: 5px;
                border: 1px solid #dee2e6;
                border-radius: 3px;
                background-color: white;
                min-width: 150px;
            }
        """)
        
        category_layout.addWidget(category_label)
        category_layout.addWidget(self.category_selector)
        category_layout.addStretch()
        
        controls_layout.addLayout(category_layout)
        
        # Update button
        self.update_btn = QPushButton("Update Chart")
        self.update_btn.setStyleSheet("""
            QPushButton {
                background-color: #0066cc;
                color: white;
                border: none;
                padding: 8px 16px;
                border-radius: 4px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #0056b3;
            }
            QPushButton:disabled {
                background-color: #6c757d;
            }
        """)
        self.update_btn.clicked.connect(self.update_chart)
        controls_layout.addWidget(self.update_btn)
        
        layout.addWidget(controls_frame)
        
        # Connect signals
        self.period_slider.valueChanged.connect(self.on_period_changed)
        self.category_selector.currentTextChanged.connect(self.update_chart)
        
        # Auto-update timer (debounce)
        self.update_timer = QTimer()
        self.update_timer.setSingleShot(True)
        self.update_timer.timeout.connect(self.update_chart)
    
    def setup_data(self):
        """Prepare data and initialize controls"""
        if not self.data_handler or not hasattr(self.data_handler, 'operating_cf_in') or self.data_handler.operating_cf_in is None:
            self.show_no_data()
            return
        
        try:
            # Prepare data
            df = self.data_handler.operating_cf_in.copy()
            
            # Remove the last row if it's a summary row
            if len(df) > 0:
                df = df[:len(df)-1]
            
            # Melt the dataframe
            self.df_long = df.melt(
                id_vars=["Category", "Item"], 
                var_name="Month", 
                value_name="Value"
            )
            
            # Convert Month to datetime
            self.df_long["Month"] = pd.to_datetime(self.df_long["Month"], errors='coerce')
            
            # Remove rows with invalid dates
            self.df_long = self.df_long.dropna(subset=['Month'])
            
            # Get sorted months
            self.months = sorted(self.df_long["Month"].unique())
            
            if not self.months:
                self.show_no_data()
                return
            
            # Setup period slider
            self.period_slider.minimum = 0
            self.period_slider.maximum = len(self.months) - 1
            self.period_slider.setValue(0, len(self.months) - 1)
            
            # Update period labels
            self.update_period_labels()
            
            # Setup category dropdown
            categories = ["All"] + sorted(self.df_long["Category"].astype(str).unique().tolist())
            self.category_selector.clear()
            self.category_selector.addItems(categories)
            
            # Enable controls
            self.update_btn.setEnabled(True)
            
            # Initial chart update
            self.update_chart()
            
        except Exception as e:
            print(f"Error setting up data: {e}")
            self.show_no_data()
    
    def show_no_data(self):
        """Show empty state when no data is available"""
        empty_fig = go.Figure()
        empty_fig.add_annotation(
            text="No operating cash inflow data available",
            x=0.5, y=0.5,
            showarrow=False,
            font={"size":16, "color":"gray"}
        )
        empty_fig.update_layout(
            xaxis={"showgrid":False, "showticklabels":False, "zeroline":False},
            yaxis={"showgrid":False, "showticklabels":False, "zeroline":False},
            plot_bgcolor='white'
        )
        self.chart_updated.emit(empty_fig)
        
        # Disable controls
        self.update_btn.setEnabled(False)
    
    def on_period_changed(self, value):
        """Handle period slider changes with debouncing"""
        self.update_period_labels()
        # Debounce the update to avoid too many rapid updates
        self.update_timer.stop()
        self.update_timer.start(300)  # 300ms delay
    
    def update_period_labels(self):
        """Update period range labels"""
        if self.months:
            start_idx, end_idx = self.period_slider.value()
            start_idx = max(0, min(start_idx, len(self.months) - 1))
            end_idx = max(0, min(end_idx, len(self.months) - 1))
            
            start_month = self.months[start_idx].strftime('%Y-%m')
            end_month = self.months[end_idx].strftime('%Y-%m')
            
            self.period_start_label.setText(start_month)
            self.period_end_label.setText(end_month)
    
    def update_chart(self):
        """Update the chart based on current selections"""
        if self.df_long is None or self.df_long.empty or not self.months:
            self.show_no_data()
            return
        
        try:
            # Get selected range
            start_idx, end_idx = self.period_slider.value()
            start_idx = max(0, min(start_idx, len(self.months) - 1))
            end_idx = max(0, min(end_idx, len(self.months) - 1))
            
            start_date = self.months[start_idx]
            end_date = self.months[end_idx]
            
            # Filter data by date range
            mask = (self.df_long["Month"] >= start_date) & (self.df_long["Month"] <= end_date)
            selected = self.df_long[mask].copy()
            
            # Apply category filter
            category = self.category_selector.currentText()
            if category and category != "All":
                selected = selected[selected["Category"] == category]
            
            if selected.empty:
                self.show_no_data()
                return
            
            # Prepare pie data
            if category == "All":
                pie_data = selected.groupby("Category")["Value"].sum().reset_index()
                pie_labels = pie_data["Category"]
                chart_title = "Distribution by Category"
            else:
                pie_data = selected.groupby("Item")["Value"].sum().reset_index()
                pie_labels = pie_data["Item"]
                chart_title = f"Distribution by Item - {category}"
            
            # Filter out zero values
            pie_data = pie_data[pie_data["Value"] > 0]
            pie_labels = pie_data["Category"] if category == "All" else pie_data["Item"]
            
            # Prepare bar data
            if category == "All":
                bar_data = selected.groupby(["Month", "Category"])["Value"].sum().reset_index()
                bar_groups = bar_data["Category"].unique()
                group_col = "Category"
            else:
                bar_data = selected.groupby(["Month", "Item"])["Value"].sum().reset_index()
                bar_groups = bar_data["Item"].unique()
                group_col = "Item"
            
            # Create figure with subplots
            fig = make_subplots(
                rows=1, cols=2,
                specs=[[{"type": "domain"}, {"type": "xy"}]],
                subplot_titles=(chart_title, "Trend Over Time"),
                column_widths=[0.4, 0.6]
            )
            
            # Add pie chart
            if not pie_data.empty:
                fig.add_trace(go.Pie(
                    labels=pie_labels,
                    values=pie_data["Value"],
                    textinfo="label+percent",
                    name="Distribution",
                    hovertemplate="<b>%{label}</b><br>Value: %{value:,.0f}<br>Percentage: %{percent}<extra></extra>"
                ), row=1, col=1)
            
            # Add bar chart
            colors = ['#1f77b4', '#ff7f0e', '#2ca02c', '#d62728', '#9467bd', '#8c564b', '#e377c2', '#7f7f7f', '#bcbd22', '#17becf']
            for i, grp in enumerate(bar_groups):
                grp_data = bar_data[bar_data[group_col] == grp]
                if not grp_data.empty:
                    fig.add_trace(go.Bar(
                        x=grp_data["Month"],
                        y=grp_data["Value"] / 1000,  # Scale to thousands
                        name=grp,
                        legendgroup=grp,
                        marker_color=colors[i % len(colors)],
                        hovertemplate="<b>%{fullData.name}</b><br>Date: %{x}<br>Value: %{y:.1f}K<extra></extra>"
                    ), row=1, col=2)
            
            # Update layout
            period_text = f"{start_date.strftime('%Y-%m')} to {end_date.strftime('%Y-%m')}"
            fig.update_layout(
                title=f"Operating Cash Inflow Analysis - {period_text}",
                barmode="stack",
                showlegend=True,
                autosize=True,
                height=500,
                plot_bgcolor='white',
                paper_bgcolor='white',
                font=dict(size=12),
                legend=dict(
                    orientation="v",
                    yanchor="top",
                    y=1,
                    xanchor="left",
                    x=1.02
                )
            )
            
            # Update axes
            fig.update_xaxes(title_text="Month", row=1, col=2)
            fig.update_yaxes(title_text="Value (Thousands)", row=1, col=2)
            
            # Emit the figure
            self.chart_updated.emit(fig)
            
        except Exception as e:
            print(f"Error updating chart: {e}")
            self.show_no_data()

class DataLoader(QThread):
    """Background thread for loading Excel data"""
    data_loaded = pyqtSignal(object, object)  # Data_Handler, plot_diagrams
    error_occurred = pyqtSignal(str)
    
    def __init__(self, file_path):
        super().__init__()
        self.file_path = file_path
    
    def run(self):
        try:
            # Initialize your custom classes
            data_handler = Data_Handler(self.file_path)
            plot_handler = plot_diagrams()
            
            self.data_loaded.emit(data_handler, plot_handler)
            
        except Exception as e:
            self.error_occurred.emit(str(e))

class ChartWidget(QWidget):
    """Custom widget for displaying Plotly charts"""
    def __init__(self, parent=None):
        super().__init__(parent)
        self.layout = QVBoxLayout()
        self.setLayout(self.layout)
    
    def clear_chart(self):
        """Remove existing chart"""
        for i in reversed(range(self.layout.count())):
            widget = self.layout.itemAt(i).widget()
            if widget:
                widget.setParent(None)
    
    def display_plotly_figure(self, fig):
        """Display a Plotly figure in the widget"""
        self.clear_chart()
        
        # Convert Plotly figure to HTML
        html = pio.to_html(fig, include_plotlyjs='cdn', full_html=False)
        web_view = QWebEngineView()
        web_view.setHtml(html)
        self.layout.addWidget(web_view)

class KPIWidget(QFrame):
    """Custom widget for displaying KPIs"""
    def __init__(self, title, value, change=None, parent=None):
        super().__init__(parent)
        self.setFrameStyle(QFrame.Shape.Box)
        self.setStyleSheet("""
            QFrame {
                background-color: #f8f9fa;
                border: 1px solid #dee2e6;
                border-radius: 8px;
                padding: 15px;
                margin: 5px;
            }
        """)
        
        layout = QVBoxLayout()
        
        # Title
        title_label = QLabel(title)
        title_label.setFont(QFont("Arial", 10))
        title_label.setStyleSheet("color: #6c757d; font-weight: bold;")
        title_label.setWordWrap(True)
        
        # Value
        value_label = QLabel(value)
        value_label.setFont(QFont("Arial", 14, QFont.Weight.Bold))
        value_label.setStyleSheet("color: #212529;")
        value_label.setWordWrap(True)
        
        layout.addWidget(title_label)
        layout.addWidget(value_label)
        
        # Change (optional)
        if change:
            change_label = QLabel(f"{change}%")
            change_label.setFont(QFont("Arial", 12))
            if float(change.replace('%', '').replace('+', '')) >= 0:
                change_label.setStyleSheet("color: #28a745;")
            else:
                change_label.setStyleSheet("color: #dc3545;")
            layout.addWidget(change_label)
        
        self.setLayout(layout)

class CashFlowDashboard(QMainWindow):
    """Main dashboard window integrated with custom classes"""
    def __init__(self):
        super().__init__()
        self.data_handler = None
        self.plot_handler = None
        self.file_path = None
        self.init_ui()
        self.apply_theme()

    
    def init_ui(self):
        """Initialize the user interface"""
        self.setWindowTitle("Cash Flow Forecast Dashboard")
        self.setGeometry(100, 100, 1600, 1000)
    
        # Central widget
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
    
        # Main layout for central widget
        central_layout = QVBoxLayout(central_widget)
    
        # Header
        self.create_header(central_layout)
    
        # Control panel
        self.create_controls(central_layout)
    
        # Scrollable content area
        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)
    
        # Content widget inside scroll area
        content_widget = QWidget()
        self.scroll_layout = QVBoxLayout(content_widget)
    
        # Add dashboard / charts / other content inside scroll_layout
        self.create_content_area(self.scroll_layout)
    
        scroll_area.setWidget(content_widget)
    
        # Add scroll area to central layout
        central_layout.addWidget(scroll_area)
    
        # Status bar
        self.statusBar().showMessage("Ready - Load Excel file to begin")

    
    def create_header(self, layout):
        """Create header section"""
        header_widget = QWidget()
        header_widget.setStyleSheet("""
            QWidget {
                background-color: #2E3440;
                color: white;
                border-radius: 10px;
                padding: 20px;
            }
        """)
        
        header_layout = QHBoxLayout()
        
        title_label = QLabel("Cash Flow Forecast Dashboard")
        title_label.setFont(QFont("Arial", 18, QFont.Weight.Bold))
        title_label.setStyleSheet("color: white;")
        
        # Add current date
        date_label = QLabel(f"As of: {datetime.now().strftime('%Y-%m-%d')}")
        date_label.setFont(QFont("Arial", 12))
        date_label.setStyleSheet("color: #d8dee9;")
        
        header_layout.addWidget(title_label)
        header_layout.addStretch()
        header_layout.addWidget(date_label)
        
        header_widget.setLayout(header_layout)
        layout.addWidget(header_widget)
    
    def create_controls(self, layout):
        """Create control panel"""
        controls_widget = QWidget()
        controls_layout = QHBoxLayout()
        
        # Load Excel button
        self.load_btn = QPushButton("Load Excel File")
        self.load_btn.setStyleSheet("""
            QPushButton {
                background-color: #0066cc;
                color: white;
                border: none;
                padding: 10px 20px;
                border-radius: 5px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #0056b3;
            }
        """)
        self.load_btn.clicked.connect(self.load_excel_file)
        
        # Refresh button
        self.refresh_btn = QPushButton("Refresh Charts")
        self.refresh_btn.setEnabled(False)
        self.refresh_btn.clicked.connect(self.refresh_dashboard)
        
        # Export button
        self.export_btn = QPushButton("Export Report")
        self.export_btn.setEnabled(False)
        self.export_btn.clicked.connect(self.export_report)
        
        controls_layout.addWidget(self.load_btn)
        controls_layout.addWidget(self.refresh_btn)
        controls_layout.addWidget(self.export_btn)
        controls_layout.addStretch()
        
        controls_widget.setLayout(controls_layout)
        layout.addWidget(controls_widget)
    
    def create_content_area(self, layout):
        """Create main content area"""
        # Create tab widget
        self.tab_widget = QTabWidget()
        
        # Dashboard tab
        self.create_dashboard_tab()
        
        # Data tabs
        self.create_data_tabs()
        
        layout.addWidget(self.tab_widget)
    
    def create_dashboard_tab(self):
        """Create main dashboard tab with enhanced KPI layout and scrollable content"""
        # Create scroll area for the entire dashboard
        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)
        scroll_area.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        scroll_area.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)

        # Create the main content widget that will be scrollable
        dashboard_widget = QWidget()
        dashboard_layout = QVBoxLayout(dashboard_widget)
        dashboard_layout.setContentsMargins(15, 15, 15, 15)
        dashboard_layout.setSpacing(20)

        # Enhanced KPI Section at the top - Fixed position
        kpi_section = QFrame()
        kpi_section.setFrameStyle(QFrame.Shape.Box)
        kpi_section.setStyleSheet("""
            QFrame {
                background-color: #f8f9fa;
                border: 2px solid #dee2e6;
                border-radius: 10px;
                padding: 20px;
                margin-bottom: 10px;
            }
        """)

        kpi_main_layout = QVBoxLayout(kpi_section)

        # KPI Title
        kpi_title = QLabel("Key Performance Indicators & Insights")
        kpi_title.setFont(QFont("Arial", 16, QFont.Weight.Bold))
        kpi_title.setStyleSheet("color: #2E3440; margin-bottom: 15px;")
        kpi_title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        kpi_main_layout.addWidget(kpi_title)

        # Enhanced KPI Grid Layout with more rows to accommodate insights
        self.kpi_frame = QFrame()
        self.kpi_layout = QGridLayout(self.kpi_frame)
        self.kpi_layout.setSpacing(15)

        # Set column stretch factors for better layout
        self.kpi_layout.setColumnStretch(0, 1)
        self.kpi_layout.setColumnStretch(1, 1)
        self.kpi_layout.setColumnStretch(2, 1)

        kpi_main_layout.addWidget(self.kpi_frame)

        # Add KPI section to main layout
        dashboard_layout.addWidget(kpi_section)

        # Charts Section Title
        charts_title = QLabel("Financial Charts & Analysis")
        charts_title.setFont(QFont("Arial", 14, QFont.Weight.Bold))
        charts_title.setStyleSheet("color: #0066cc; padding: 10px 0;")
        dashboard_layout.addWidget(charts_title)

        # Charts section with better organization
        charts_container = QFrame()
        charts_container.setFrameStyle(QFrame.Shape.Box)
        charts_container.setStyleSheet("""
            QFrame {
                background-color: white;
                border: 1px solid #dee2e6;
                border-radius: 8px;
                padding: 15px;
            }
        """)
        charts_layout = QVBoxLayout(charts_container)

        # Top row charts (side by side)
        top_charts_frame = QFrame()
        top_charts_layout = QHBoxLayout(top_charts_frame)

        # Waterfall Chart
        waterfall_container = QFrame()
        waterfall_container.setFrameStyle(QFrame.Shape.StyledPanel)
        waterfall_container.setStyleSheet("border: 1px solid #ccc; border-radius: 5px;")
        waterfall_layout = QVBoxLayout(waterfall_container)

        waterfall_title = QLabel("Cash Movement Analysis")
        waterfall_title.setFont(QFont("Arial", 12, QFont.Weight.Bold))
        waterfall_title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        waterfall_title.setStyleSheet("padding: 8px; background-color: #e9ecef; border-radius: 3px;")

        self.waterfall_chart = ChartWidget()
        self.waterfall_chart.setMinimumHeight(400)

        waterfall_layout.addWidget(waterfall_title)
        waterfall_layout.addWidget(self.waterfall_chart)

        # Monthly Chart
        monthly_container = QFrame()
        monthly_container.setFrameStyle(QFrame.Shape.StyledPanel)
        monthly_container.setStyleSheet("border: 1px solid #ccc; border-radius: 5px;")
        monthly_layout = QVBoxLayout(monthly_container)

        monthly_title = QLabel("Monthly Cash Flow Trend")
        monthly_title.setFont(QFont("Arial", 12, QFont.Weight.Bold))
        monthly_title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        monthly_title.setStyleSheet("padding: 8px; background-color: #e9ecef; border-radius: 3px;")

        self.monthly_chart = ChartWidget()
        self.monthly_chart.setMinimumHeight(400)

        monthly_layout.addWidget(monthly_title)
        monthly_layout.addWidget(self.monthly_chart)

        # Add charts to top row
        top_charts_layout.addWidget(waterfall_container)
        top_charts_layout.addWidget(monthly_container)

        # Operating Cash Flow Chart (full width)
        operating_container = QFrame()
        operating_container.setFrameStyle(QFrame.Shape.StyledPanel)
        operating_container.setStyleSheet("border: 1px solid #ccc; border-radius: 5px; margin-top: 15px;")
        operating_layout = QVBoxLayout(operating_container)

        operating_title = QLabel("Operating Cash Flow Analysis")
        operating_title.setFont(QFont("Arial", 12, QFont.Weight.Bold))
        operating_title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        operating_title.setStyleSheet("padding: 8px; background-color: #e9ecef; border-radius: 3px;")

        self.operating_cash_chart = ChartWidget()
        self.operating_cash_chart.setMinimumHeight(500)

        operating_layout.addWidget(operating_title)
        operating_layout.addWidget(self.operating_cash_chart)

        # Add all chart components to charts container
        charts_layout.addWidget(top_charts_frame)
        charts_layout.addWidget(operating_container)

        # Add charts container to main layout
        dashboard_layout.addWidget(charts_container)

        # Add some bottom spacing
        dashboard_layout.addStretch(0)

        # Set the scrollable widget
        scroll_area.setWidget(dashboard_widget)

        # Add the scroll area to the tab
        self.tab_widget.addTab(scroll_area, "Dashboard")

    def create_data_tabs(self):
        """Create data view tabs"""
        # Cash Flow Data tab
        self.create_cash_flow_data_tab()
        
        # Operating Cash Inflow tab
        self.create_operating_inflow_tab()
        
        # Operating Cash Outflow tab  
        self.create_operating_outflow_tab()
    
    def create_cash_flow_data_tab(self):
        """Create cash flow data tab"""
        data_widget = QWidget()
        data_layout = QVBoxLayout()
        
        self.cash_flow_table = QTableWidget()
        data_layout.addWidget(self.cash_flow_table)
        
        data_widget.setLayout(data_layout)
        self.tab_widget.addTab(data_widget, "Cash Flow Data")
    
    def create_operating_inflow_tab(self):
        """Create operating cash inflow tab with interactive pie chart - scrollable version"""
        # Create scroll area as the main container
        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)  # Allow content to resize with the scroll area
        scroll_area.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        scroll_area.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)

        # Create the main content widget that will be scrollable
        inflow_widget = QWidget()
        inflow_layout = QVBoxLayout(inflow_widget)

        # Add some padding for better appearance
        inflow_layout.setContentsMargins(10, 10, 10, 10)
        inflow_layout.setSpacing(15)

        # Create splitter for controls and chart
        splitter = QSplitter(Qt.Orientation.Vertical)

        # Create the interactive pie chart widget
        self.inflow_pie_widget = OperatingInflowPieWidget()

        # Create chart display widget with minimum size
        self.inflow_pie_chart = ChartWidget()
        self.inflow_pie_chart.setMinimumHeight(600)  # Ensure chart has minimum height

        # Connect the chart update signal
        self.inflow_pie_widget.chart_updated.connect(self.inflow_pie_chart.display_plotly_figure)

        # Add widgets to splitter
        splitter.addWidget(self.inflow_pie_widget)
        splitter.addWidget(self.inflow_pie_chart)

        # Set stretch factors
        splitter.setStretchFactor(0, 0)  # Controls don't stretch
        splitter.setStretchFactor(1, 1)  # Chart stretches

        # Set minimum height for the splitter to ensure good visibility
        splitter.setMinimumHeight(1000)

        inflow_layout.addWidget(splitter)

        # Add section separator
        separator = QFrame()
        separator.setFrameShape(QFrame.Shape.HLine)
        separator.setFrameShadow(QFrame.Shadow.Sunken)
        separator.setStyleSheet("QFrame { color: #dee2e6; }")
        inflow_layout.addWidget(separator)

        # Add table section with title
        table_title = QLabel("Operating Cash Inflow Data Table")
        table_title.setFont(QFont("Arial", 12, QFont.Weight.Bold))
        table_title.setStyleSheet("padding: 10px; background-color: #f8f9fa; border-radius: 5px; margin-bottom: 5px;")
        inflow_layout.addWidget(table_title)

        # Add table with minimum size
        self.operating_inflow_table = QTableWidget()
        self.operating_inflow_table.setMinimumHeight(300)  # Ensure table has minimum height
        self.operating_inflow_table.setAlternatingRowColors(True)  # Better table appearance
        self.operating_inflow_table.setStyleSheet("""
            QTableWidget {
                gridline-color: #dee2e6;
                background-color: white;
                alternate-background-color: #f8f9fa;
            }
            QHeaderView::section {
                background-color: #e9ecef;
                padding: 8px;
                border: 1px solid #dee2e6;
                font-weight: bold;
            }
        """)
        inflow_layout.addWidget(self.operating_inflow_table)

        # Add some bottom spacing
        inflow_layout.addStretch(0)

        # Set the scrollable widget
        scroll_area.setWidget(inflow_widget)

        # Add the scroll area to the tab
        self.tab_widget.addTab(scroll_area, "Operating Cash Inflow")
    
    def create_operating_outflow_tab(self):
        """Create operating cash outflow tab with scroll area"""
        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)
        scroll_area.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        scroll_area.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        
        outflow_widget = QWidget()
        outflow_layout = QVBoxLayout()
        
        # Add title
        title_label = QLabel("Operating Cash Outflow Data")
        title_label.setFont(QFont("Arial", 14, QFont.Weight.Bold))
        title_label.setStyleSheet("padding: 10px; background-color: #f8f9fa;")
        outflow_layout.addWidget(title_label)
        
        self.operating_outflow_table = QTableWidget()
        outflow_layout.addWidget(self.operating_outflow_table)
        
        outflow_widget.setLayout(outflow_layout)
        scroll_area.setWidget(outflow_widget)
        self.tab_widget.addTab(scroll_area, "Operating Cash Outflow")
    
    def apply_theme(self):
        """Apply modern theme to the application"""
        self.setStyleSheet("""
            QMainWindow {
                background-color: #ffffff;
            }
            QTabWidget::pane {
                border: 1px solid #dee2e6;
                border-radius: 5px;
            }
            QTabBar::tab {
                background-color: #f8f9fa;
                border: 1px solid #dee2e6;
                padding: 8px 12px;
                margin-right: 2px;
            }
            QTabBar::tab:selected {
                background-color: #0066cc;
                color: white;
            }
            QPushButton {
                background-color: #f8f9fa;
                border: 1px solid #dee2e6;
                padding: 8px 16px;
                border-radius: 4px;
            }
            QPushButton:hover {
                background-color: #e9ecef;
            }
        """)
    
    def load_excel_file(self):
        """Load Excel file dialog"""
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Select Excel File",
            "",
            "Excel Files (*.xlsx *.xls)"
        )
        
        if file_path:
            self.statusBar().showMessage("Loading data...")
            self.file_path = file_path
            # Create and start data loader thread
            self.data_loader = DataLoader(file_path)
            self.data_loader.data_loaded.connect(self.on_data_loaded)
            self.data_loader.error_occurred.connect(self.on_error)
            self.data_loader.start()
    
    def on_data_loaded(self, data_handler, plot_handler):
        """Handle successful data loading with enhanced KPI initialization"""
        self.data_handler = data_handler
        self.plot_handler = plot_handler

        # Update the pie widget with new data
        if hasattr(self, 'inflow_pie_widget'):
            self.inflow_pie_widget.data_handler = data_handler
            self.inflow_pie_widget.setup_data()

        # Initialize enhanced KPI system first
        self.update_dashboard()
        self.update_data_tables()

        self.refresh_btn.setEnabled(True)
        self.export_btn.setEnabled(True)

        self.statusBar().showMessage("Data loaded successfully - Enhanced KPI insights available")

        # Show message about KPI insights if totals data is available
        if (hasattr(data_handler, 'totals') and 
            data_handler.totals is not None and 
            not data_handler.totals.empty):
            self.statusBar().showMessage("Data loaded successfully - Interactive KPI insights with time controls are now active")
        else:
            self.statusBar().showMessage("Data loaded successfully - Basic KPI display active (no time series data found)")

    def on_error(self, error_msg):
        """Handle data loading errors"""
        QMessageBox.critical(self, "Error Loading Data", f"Failed to load Excel file:\n{error_msg}")
        self.statusBar().showMessage("Error loading data")
    
    def update_dashboard(self):
        """Update dashboard with real data"""
        if self.data_handler and self.plot_handler:
            # Update KPIs first (they're now at the top)
            self.update_kpis()
            
            # Update charts using your plot_diagrams class
            try:
                if hasattr(self.plot_handler, 'waterfall_cash_movement_fig'):
                    self.waterfall_chart.display_plotly_figure(self.plot_handler.waterfall_cash_movement_fig)
                if hasattr(self.plot_handler, 'monthly_cash_flow_fig'):
                    self.monthly_chart.display_plotly_figure(self.plot_handler.monthly_cash_flow_fig)
                if hasattr(self.plot_handler, 'operating_cash_flow_diagram_fig'):
                    self.operating_cash_chart.display_plotly_figure(self.plot_handler.operating_cash_flow_diagram_fig)
                    
            except Exception as e:
                self.statusBar().showMessage(f"Error updating charts: {str(e)}")
                
            # Remove the KPI insights from scroll_layout since KPIs are now at the top
            if hasattr(self, "kpi_insights_widget") and self.kpi_insights_widget:
                # This removes the complex KPI insights that were added to scroll_layout
                if self.kpi_insights_widget.parent():
                    self.kpi_insights_widget.parent().layout().removeWidget(self.kpi_insights_widget)
                self.kpi_insights_widget.deleteLater()
                self.kpi_insights_widget = None
    
    def update_kpis(self):
        """Update KPI widgets with real data - Enhanced version with insights integration"""
        # Clear existing KPIs
        for i in reversed(range(self.kpi_layout.count())):
            widget = self.kpi_layout.itemAt(i).widget()
            if widget:
                widget.setParent(None)

        if not self.data_handler:
            error_kpi = self.create_enhanced_kpi_widget(
                "No Data", 
                "Please load Excel file to view KPIs", 
                "⚠️", 
                "#dc3545"
            )
            self.kpi_layout.addWidget(error_kpi, 0, 0)
            return

        try:
            # Initialize KPI insights functionality
            self.setup_kpi_insights()

        except Exception as e:
            error_kpi = self.create_enhanced_kpi_widget(
                "Error", 
                f"Failed to load KPIs: {str(e)[:50]}...", 
                "⚠️", 
                "#dc3545"
            )
            self.kpi_layout.addWidget(error_kpi, 0, 0)

    def setup_kpi_insights(self):
        """Setup KPI insights with time period controls integrated into main dashboard"""

        # Check if we have totals data for advanced insights
        has_totals = (hasattr(self.data_handler, 'totals') and 
                      self.data_handler.totals is not None and 
                      not self.data_handler.totals.empty)

        if has_totals:
            self.setup_advanced_kpi_insights()
        else:
            self.setup_basic_kpi_display()

    def setup_basic_kpi_display(self):
        """Setup basic KPI display when totals data is not available"""
        calc_dict = getattr(self.data_handler, 'calculations_dict', {})

        # Define basic KPI data
        basic_kpis = [
            {
                "title": "Beginning Cash Balance", 
                "value": f"{calc_dict.get('Cash Beginning Balance', 0):,.0f} SAR",
                "icon": "💰",
                "color": "#17a2b8"
            },
            {
                "title": "Total Operating Inflow", 
                "value": f"{calc_dict.get('Total Operating Cash Inflow', 0):,.0f} SAR",
                "icon": "📈", 
                "color": "#28a745"
            },
            {
                "title": "Total Operating Outflow", 
                "value": f"{calc_dict.get('Total Operating Cash Outflow', 0):,.0f} SAR",
                "icon": "📉",
                "color": "#dc3545"
            },
            {
                "title": "Net Cash Flow", 
                "value": f"{calc_dict.get('Total Operating Cash Inflow', 0) - calc_dict.get('Total Operating Cash Outflow', 0):,.0f} SAR",
                "icon": "⚖️",
                "color": "#6f42c1"
            },
            {
                "title": "Ending Cash Balance", 
                "value": f"{calc_dict.get('Cash Ending Balance', 0):,.0f} SAR",
                "icon": "🏦",
                "color": "#fd7e14"
            }
        ]

        # Create basic KPI widgets
        for i, kpi in enumerate(basic_kpis):
            kpi_widget = self.create_enhanced_kpi_widget(
                kpi["title"], 
                kpi["value"], 
                kpi["icon"], 
                kpi["color"]
            )
            row = i // 3
            col = i % 3
            self.kpi_layout.addWidget(kpi_widget, row, col)

    def setup_advanced_kpi_insights(self):
        """Setup advanced KPI insights with time period controls"""
        try:
            df = self.data_handler.totals.copy()

            # Identify date and non-date columns
            non_date_cols = []
            for col in df.columns:
                col_lower = str(col).lower()
                if any(keyword in col_lower for keyword in ['country', 'cash flow', 'category', 'item', 'type']):
                    non_date_cols.append(col)

            if not non_date_cols:
                non_date_cols = ['Country', 'Cash Flow Type', 'Category', 'Item']

            # Get and parse date columns
            date_columns = [col for col in df.columns if col not in non_date_cols]
            date_objects, successful_cols = self.parse_date_columns(date_columns)

            if not date_objects:
                # Fall back to basic display if no valid dates
                self.setup_basic_kpi_display()
                return

            # Store data for updates
            self.kpi_df = df
            self.kpi_date_columns = successful_cols
            self.kpi_date_objects = sorted(date_objects)
            self.kpi_non_date_cols = non_date_cols

            # Create time period controls in KPI section
            self.create_kpi_time_controls()

            # Initial KPI calculation and display
            self.update_advanced_kpis()

        except Exception as e:
            print(f"Error setting up advanced KPI insights: {e}")
            self.setup_basic_kpi_display()

    def parse_date_columns(self, date_columns):
        """Parse date columns and return valid date objects and column names"""
        date_objects = []
        successful_cols = []

        for col in date_columns:
            try:
                col_str = str(col)
                date_obj = None

                # Try different parsing strategies
                if ' ' in col_str:
                    date_str = col_str.split(' ')[0]
                    try:
                        date_obj = datetime.strptime(date_str, '%Y-%m-%d')
                    except ValueError:
                        pass

                if date_obj is None:
                    try:
                        date_obj = pd.to_datetime(col_str).to_pydatetime()
                    except:
                        pass

                if date_obj is None:
                    date_formats = ['%Y-%m-%d', '%Y/%m/%d', '%m/%d/%Y', '%d/%m/%Y', '%Y-%m', '%Y%m%d']
                    for fmt in date_formats:
                        try:
                            date_obj = datetime.strptime(col_str, fmt)
                            break
                        except ValueError:
                            continue

                if date_obj is None:
                    import re
                    date_match = re.search(r'(\d{4})-(\d{2})-(\d{2})', col_str)
                    if date_match:
                        try:
                            date_obj = datetime(int(date_match.group(1)), 
                                              int(date_match.group(2)), 
                                              int(date_match.group(3)))
                        except ValueError:
                            pass

                if date_obj:
                    date_objects.append(date_obj)
                    successful_cols.append(col)

            except Exception:
                continue

        return date_objects, successful_cols
    

    def create_kpi_time_controls(self):
        """Create time period control with RangeSlider in the KPI section"""

        # Create a control widget that spans the full width
        control_widget = QFrame()
        control_widget.setFrameStyle(QFrame.Shape.Box)
        control_widget.setStyleSheet("""
            QFrame {
                background-color: #e3f2fd;
                border: 2px solid #1976d2;
                border-radius: 8px;
                padding: 15px;
                margin: 5px;
            }
        """)

        control_layout = QVBoxLayout(control_widget)

        # Title for controls
        control_title = QLabel("📅 Time Period Selection")
        control_title.setFont(QFont("Arial", 12, QFont.Weight.Bold))
        control_title.setStyleSheet("color: #1976d2; margin-bottom: 10px;")
        control_title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        control_layout.addWidget(control_title)

        # Labels layout
        labels_layout = QHBoxLayout()
        self.kpi_start_label = QLabel(self.kpi_date_objects[0].strftime('%Y-%m-%d'))
        self.kpi_start_label.setStyleSheet("background-color: white; padding: 4px; border-radius: 3px; font-size: 10px;")

        self.kpi_end_label = QLabel(self.kpi_date_objects[-1].strftime('%Y-%m-%d'))
        self.kpi_end_label.setStyleSheet("background-color: white; padding: 4px; border-radius: 3px; font-size: 10px;")

        labels_layout.addWidget(QLabel("Start:"))
        labels_layout.addWidget(self.kpi_start_label)
        labels_layout.addStretch()
        labels_layout.addWidget(QLabel("End:"))
        labels_layout.addWidget(self.kpi_end_label)

        control_layout.addLayout(labels_layout)

        # ✅ Use custom RangeSlider instead of two QSliders
        self.kpi_range_slider = RangeSlider(0, len(self.kpi_date_objects) - 1)
        self.kpi_range_slider.setValue(0, len(self.kpi_date_objects) - 1)
        control_layout.addWidget(self.kpi_range_slider)

        # Add control widget to KPI layout (spans full width)
        self.kpi_layout.addWidget(control_widget, 0, 0, 1, 3)  # Row 0, spans 3 columns

        # Connect signal
        self.kpi_range_slider.valueChanged.connect(self.on_kpi_period_changed)


    def on_kpi_period_changed(self, values):
        """Update labels when range slider is moved"""
        start_idx, end_idx = values
        self.kpi_start_label.setText(self.kpi_date_objects[start_idx].strftime('%Y-%m-%d'))
        self.kpi_end_label.setText(self.kpi_date_objects[end_idx].strftime('%Y-%m-%d'))
        
        # 🔹 Call your KPI update logic here
        self.update_advanced_kpis()

    def update_advanced_kpis(self):
        """Update KPIs based on selected time period"""
        try:
            # ✅ Get start & end indices from RangeSlider
            start_idx, end_idx = self.kpi_range_slider.value()

            # Ensure indices are valid
            if start_idx > end_idx:
                start_idx, end_idx = end_idx, start_idx

            # Get selected date range
            selected_date_cols = self.kpi_date_columns[start_idx:end_idx + 1]
            selected_dates = self.kpi_date_objects[start_idx:end_idx + 1]

            # Calculate KPIs for selected period
            kpi_insights = self.calculate_period_kpis(selected_date_cols, selected_dates)

            # Clear existing KPI widgets (except control widget in row 0)
            for row in range(1, self.kpi_layout.rowCount()):
                for col in range(self.kpi_layout.columnCount()):
                    item = self.kpi_layout.itemAtPosition(row, col)
                    if item:
                        widget = item.widget()
                        if widget:
                            widget.setParent(None)

            # Create and display KPI widgets
            self.create_advanced_kpi_widgets(kpi_insights, len(selected_date_cols))

        except Exception as e:
            print(f"Error updating advanced KPIs: {e}")

    
    def calculate_period_kpis(self, selected_cols, selected_dates):
        """Calculate KPIs for the selected time period"""
        kpis = {}

        for _, row in self.kpi_df.iterrows():
            item_name = row.get('Item', 'Unknown')
            if pd.notna(item_name):
                # Get values for selected period
                period_values = []
                for col in selected_cols:
                    try:
                        val = pd.to_numeric(row[col], errors='coerce')
                        period_values.append(val if pd.notna(val) else 0)
                    except:
                        period_values.append(0)

                # Calculate metrics
                total = sum(period_values)
                avg_per_period = total / len(period_values) if period_values else 0
                max_value = max(period_values) if period_values else 0
                min_value = min(period_values) if period_values else 0

                # Trend calculation
                if len(period_values) >= 2:
                    trend_change = period_values[-1] - period_values[0]
                    trend_pct = (trend_change / abs(period_values[0]) * 100) if period_values[0] != 0 else 0
                else:
                    trend_change = 0
                    trend_pct = 0

                kpis[item_name] = {
                    'total': total,
                    'average': avg_per_period,
                    'max': max_value,
                    'min': min_value,
                    'trend_change': trend_change,
                    'trend_pct': trend_pct,
                    'latest': period_values[-1] if period_values else 0,
                    'first': period_values[0] if period_values else 0
                }

        return kpis

    def create_advanced_kpi_widgets(self, kpi_insights, num_periods):
        """Create and display advanced KPI widgets in organized sections"""

        # Calculate aggregate metrics
        total_inflow = sum([m['total'] for name, m in kpi_insights.items() if 'Inflow' in name and m['total'] > 0])
        total_outflow = sum([abs(m['total']) for name, m in kpi_insights.items() if 'Outflow' in name and m['total'] < 0])
        net_cash_flow = total_inflow - total_outflow

        # PRIMARY KPIs (Row 1)
        primary_kpis = [
            {
                "title": "Net Cash Flow", 
                "value": f"{net_cash_flow:,.0f} SAR",
                "trend": f"{((net_cash_flow / total_inflow * 100) if total_inflow > 0 else 0):+.1f}%" if total_inflow > 0 else "0.0%",
                "icon": "⚖️",
                "color": "#28a745" if net_cash_flow > 0 else "#dc3545"
            },
            {
                "title": "Total Inflow", 
                "value": f"{total_inflow:,.0f} SAR",
                "trend": f"+{((total_inflow / (total_inflow + total_outflow)) * 100):,.1f}%" if (total_inflow + total_outflow) > 0 else "0.0%",
                "icon": "📈", 
                "color": "#28a745"
            },
            {
                "title": "Total Outflow", 
                "value": f"{total_outflow:,.0f} SAR",
                "trend": f"-{((total_outflow / (total_inflow + total_outflow)) * 100):,.1f}%" if (total_inflow + total_outflow) > 0 else "0.0%",
                "icon": "📉",
                "color": "#dc3545"
            }
        ]

        for i, kpi in enumerate(primary_kpis):
            kpi_widget = self.create_enhanced_kpi_widget_with_trend(
                kpi["title"], 
                kpi["value"],
                kpi["trend"],
                kpi["icon"], 
                kpi["color"]
            )
            self.kpi_layout.addWidget(kpi_widget, 1, i)

        # PERFORMANCE KPIs (Row 2) - Top performers and trends
        performance_items = [(name, m['trend_pct']) for name, m in kpi_insights.items() 
                            if abs(m['total']) > 0 and m['trend_pct'] != 0]
        performance_items.sort(key=lambda x: x[1], reverse=True)

        performance_kpis = []
        if performance_items:
            # Best performer
            best_item, best_trend = performance_items[0]
            performance_kpis.append({
                "title": "Best Performer", 
                "value": best_item[:15] + "..." if len(best_item) > 15 else best_item,
                "trend": f"{best_trend:+.1f}%",
                "icon": "🏆",
                "color": "#ffc107"
            })

            # Worst performer
            if len(performance_items) > 1:
                worst_item, worst_trend = performance_items[-1]
                performance_kpis.append({
                    "title": "Needs Attention", 
                    "value": worst_item[:15] + "..." if len(worst_item) > 15 else worst_item,
                    "trend": f"{worst_trend:+.1f}%",
                    "icon": "⚠️",
                    "color": "#dc3545"
                })
            else:
                # If only one item, show stability indicator
                performance_kpis.append({
                    "title": "Status", 
                    "value": "Single Item",
                    "trend": "0.0%",
                    "icon": "📊",
                    "color": "#6c757d"
                })

            # Growth summary
            growing_items = len([x for x in performance_items if x[1] > 0])
            declining_items = len([x for x in performance_items if x[1] < 0])
            stable_items = len(kpi_insights) - len(performance_items)

            growth_trend = f"+{growing_items}" if growing_items > declining_items else f"-{declining_items}" if declining_items > growing_items else "0"
            performance_kpis.append({
                "title": "Growth Summary", 
                "value": f"{growing_items}↑ {declining_items}↓ {stable_items}→",
                "trend": f"{growth_trend} net",
                "icon": "📊",
                "color": "#17a2b8"
            })
        else:
            # No trend data available
            performance_kpis = [
                {
                    "title": "Performance", 
                    "value": "No Trends",
                    "trend": "0.0%",
                    "icon": "📊",
                    "color": "#6c757d"
                },
                {
                    "title": "Analysis", 
                    "value": "Insufficient Data",
                    "trend": "0.0%",
                    "icon": "⚠️",
                    "color": "#ffc107"
                },
                {
                    "title": "Summary", 
                    "value": f"{len(kpi_insights)} Items",
                    "trend": "0.0%",
                    "icon": "📈",
                    "color": "#17a2b8"
                }
            ]

        for i, kpi in enumerate(performance_kpis):
            kpi_widget = self.create_enhanced_kpi_widget_with_trend(
                kpi["title"], 
                kpi["value"],
                kpi["trend"],
                kpi["icon"], 
                kpi["color"]
            )
            self.kpi_layout.addWidget(kpi_widget, 2, i)

        # ANALYSIS KPIs (Row 3) - Period analysis
        avg_monthly_flow = net_cash_flow / num_periods if num_periods > 0 else 0
        cash_flow_ratio = total_inflow / total_outflow if total_outflow > 0 else float('inf')

        analysis_kpis = [
            {
                "title": "Analysis Period", 
                "value": f"{num_periods} Month{'s' if num_periods != 1 else ''}",
                "trend": f"{((num_periods / len(self.kpi_date_objects)) * 100):,.1f}%" if hasattr(self, 'kpi_date_objects') and self.kpi_date_objects else "100.0%",
                "icon": "📅",
                "color": "#6f42c1"
            },
            {
                "title": "Avg Monthly Flow", 
                "value": f"{avg_monthly_flow:,.0f} SAR",
                "trend": f"{((avg_monthly_flow / net_cash_flow * 100) if net_cash_flow != 0 else 0):+.1f}%",
                "icon": "📊",
                "color": "#fd7e14"
            },
            {
                "title": "Cash Flow Ratio", 
                "value": f"{cash_flow_ratio:,.2f}" if cash_flow_ratio != float('inf') else "∞",
                "trend": f"{((cash_flow_ratio - 1) * 100):+.1f}%" if cash_flow_ratio != float('inf') else "+100.0%",
                "icon": "🎯",
                "color": "#17a2b8"
            }
        ]

        for i, kpi in enumerate(analysis_kpis):
            kpi_widget = self.create_enhanced_kpi_widget_with_trend(
                kpi["title"], 
                kpi["value"],
                kpi["trend"],
                kpi["icon"], 
                kpi["color"]
            )
            self.kpi_layout.addWidget(kpi_widget, 3, i)
    def create_enhanced_kpi_widget_with_trend(self, title, value, trend, icon, color):
        """Create enhanced KPI widget with drop shadow and clean style"""
    
        # Frame container
        kpi_widget = QFrame()
        kpi_widget.setFrameStyle(QFrame.Shape.NoFrame)
        kpi_widget.setStyleSheet(f"""
            QFrame {{
                background-color: white;
                border: 1px solid #dee2e6;
                border-radius: 12px;
                padding: 10px;
                margin: 8px;
            }}
            QLabel {{
                color: #212529;
            }}
        """)
    
        # ✅ Add drop shadow
        shadow = QGraphicsDropShadowEffect()
        shadow.setBlurRadius(15)
        shadow.setXOffset(2)
        shadow.setYOffset(2)
        shadow.setColor(QColor(0, 0, 0, 50))
        kpi_widget.setGraphicsEffect(shadow)
    
        layout = QVBoxLayout(kpi_widget)
        layout.setSpacing(6)
    
        # 🔹 Title (small & subtle)
        title_label = QLabel(title)
        title_label.setFont(QFont("Arial", 9, QFont.Weight.Bold))
        title_label.setStyleSheet("color: #6c757d;")  # grey like your image
        title_label.setWordWrap(True)
        title_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
    
        # 🔹 Value (bold & centered)
        value_label = QLabel(value)
        value_label.setFont(QFont("Arial", 13, QFont.Weight.Bold))
        value_label.setStyleSheet("color: #000;")
        value_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
    
        # 🔹 Trend (optional, smaller font)
        trend_label = QLabel(trend if trend else "")
        trend_label.setFont(QFont("Arial", 9))
        trend_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        if trend and "%" in str(trend):
            if "+" in trend:
                trend_label.setStyleSheet("color: #28a745; font-weight: bold;")  # green
            elif "-" in trend:
                trend_label.setStyleSheet("color: #dc3545; font-weight: bold;")  # red
            else:
                trend_label.setStyleSheet("color: #6c757d;")  # neutral grey
    
        # Add widgets
        layout.addWidget(title_label)
        layout.addWidget(value_label)
        if trend:
            layout.addWidget(trend_label)
    
        return kpi_widget

    def create_operating_inflow_pie(self):
        """
        PyQt compatible version of create_operating_inflow_pie for the CashFlowDashboard class.
        This replaces the original Jupyter widgets version and returns a simple figure for dashboard display.
        """
        if not self.data_handler or not hasattr(self.data_handler, 'operating_cf_in') or self.data_handler.operating_cf_in is None:
            # Return empty figure
            empty_fig = go.Figure()
            empty_fig.add_annotation(
                text="No operating cash inflow data available",
                x=0.5, y=0.5,
                showarrow=False,
                font=dict(size=16, color="gray")
            )
            empty_fig.update_layout(
                xaxis=dict(showgrid=False, showticklabels=False, zeroline=False),
                yaxis=dict(showgrid=False, showticklabels=False, zeroline=False),
                plot_bgcolor='white'
            )
            return empty_fig
        
        try:
            # Use default settings for initial display
            df = self.data_handler.operating_cf_in.copy()
            
            # Remove the last row if it's a summary row
            if len(df) > 0:
                df = df[:len(df)-1]
            
            # Melt the dataframe
            df_long = df.melt(id_vars=["Category", "Item"], var_name="Month", value_name="Value")
            df_long["Month"] = pd.to_datetime(df_long["Month"], errors='coerce')
            df_long = df_long.dropna(subset=['Month'])
            
            if df_long.empty:
                return go.Figure()
            
            # Use all data for initial display
            selected = df_long.copy()
            
            # Create pie chart by category (default view)
            pie_data = selected.groupby("Category")["Value"].sum().reset_index()
            pie_data = pie_data[pie_data["Value"] > 0]
            
            # Create bar chart data
            bar_data = selected.groupby(["Month", "Category"])["Value"].sum().reset_index()
            bar_groups = bar_data["Category"].unique()
            
            # Create figure
            fig = make_subplots(
                rows=1, cols=2,
                specs=[[{"type": "domain"}, {"type": "xy"}]],
                subplot_titles=("Distribution by Category", "Trend Over Time"),
                column_widths=[0.4, 0.6]
            )
            
            # Add pie chart
            if not pie_data.empty:
                fig.add_trace(go.Pie(
                    labels=pie_data["Category"],
                    values=pie_data["Value"],
                    textinfo="label+percent",
                    name="Distribution"
                ), row=1, col=1)
            
            # Add bar chart
            colors = ['#1f77b4', '#ff7f0e', '#2ca02c', '#d62728', '#9467bd']
            for i, grp in enumerate(bar_groups):
                grp_data = bar_data[bar_data["Category"] == grp]
                if not grp_data.empty:
                    fig.add_trace(go.Bar(
                        x=grp_data["Month"],
                        y=grp_data["Value"] / 1000,
                        name=grp,
                        marker_color=colors[i % len(colors)]
                    ), row=1, col=2)
            
            # Update layout
            fig.update_layout(
                title="Operating Cash Inflow Analysis",
                barmode="stack",
                showlegend=True,
                autosize=True,
                height=500,
                plot_bgcolor='white'
            )
            
            fig.update_xaxes(title_text="Month", row=1, col=2)
            fig.update_yaxes(title_text="Value (Thousands)", row=1, col=2)
            
            return fig
            
        except Exception as e:
            print(f"Error creating operating inflow pie chart: {e}")
            return go.Figure()
    
    def update_data_tables(self):
        """Update all data tables"""
        if self.data_handler:
            # Update main cash flow table
            if hasattr(self.data_handler, 'CashFlow') and self.data_handler.CashFlow is not None:
                self.populate_table(self.cash_flow_table, self.data_handler.CashFlow)
            
            # Update operating cash inflow table
            if hasattr(self.data_handler, 'operating_cf_in') and self.data_handler.operating_cf_in is not None:
                self.populate_table(self.operating_inflow_table, self.data_handler.operating_cf_in)
            
            # Update operating cash outflow table
            if hasattr(self.data_handler, 'operating_cf_out') and self.data_handler.operating_cf_out is not None:
                self.populate_table(self.operating_outflow_table, self.data_handler.operating_cf_out)
    
    def populate_table(self, table_widget, dataframe):
        """Populate a table widget with dataframe data"""
        if dataframe is None or dataframe.empty:
            return
        
        try:
            table_widget.setRowCount(len(dataframe))
            table_widget.setColumnCount(len(dataframe.columns))
            table_widget.setHorizontalHeaderLabels([str(col) for col in dataframe.columns])
            
            for i in range(len(dataframe)):
                for j, col in enumerate(dataframe.columns):
                    value = dataframe.iloc[i, j]
                    if pd.isna(value):
                        item = QTableWidgetItem("")
                    else:
                        item = QTableWidgetItem(str(value))
                    table_widget.setItem(i, j, item)
                    
        except Exception as e:
            print(f"Error populating table: {e}")
    
    def refresh_dashboard(self):
        """Refresh dashboard data"""
        if self.data_handler and self.plot_handler:
            self.data_loader = DataLoader(self.file_path)
            self.data_loader.data_loaded.connect(self.on_data_loaded)
            self.data_loader.error_occurred.connect(self.on_error)
            self.data_loader.start()
            self.update_dashboard()
            
            # Refresh the pie widget
            if hasattr(self, 'inflow_pie_widget'):
                self.inflow_pie_widget.update_chart()
                
            self.statusBar().showMessage("Dashboard refreshed")

    def export_report(self):
        """Export comprehensive dashboard report to PDF"""
        if not self.data_handler:
            QMessageBox.warning(self, "No Data", "Please load data before exporting report.")
            return

        file_path, _ = QFileDialog.getSaveFileName(
            self,
            "Save Report",
            f"cash_flow_report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf",
            "PDF Files (*.pdf)"
        )

        if file_path:
            try:
                # Show progress bar
                progress = QProgressBar()
                progress.setWindowTitle("Generating Report...")
                progress.setRange(0, 100)
                progress.show()

                # Create PDF document
                doc = SimpleDocTemplate(
                    file_path,
                    pagesize=A4,
                    rightMargin=72,
                    leftMargin=72,
                    topMargin=72,
                    bottomMargin=18
                )

                elements = []

                # Define styles
                styles = getSampleStyleSheet()
                title_style = ParagraphStyle(
                    'CustomTitle',
                    parent=styles['Heading1'],
                    fontSize=24,
                    spaceAfter=30,
                    alignment=1,
                    textColor=colors.HexColor('#2E3440')
                )
                heading_style = ParagraphStyle(
                    'CustomHeading',
                    parent=styles['Heading2'],
                    fontSize=16,
                    spaceAfter=12,
                    textColor=colors.HexColor('#0066cc')
                )

                # Title page
                progress.setValue(10)
                elements.append(Paragraph("Cash Flow Forecast Dashboard Report", title_style))
                elements.append(Spacer(1, 20))
                elements.append(Paragraph(f"Generated on: {datetime.now().strftime('%B %d, %Y at %I:%M %p')}", styles['Normal']))
                elements.append(Spacer(1, 40))

                # === Executive Summary with KPIs ===
                progress.setValue(20)
                elements.append(Paragraph("Executive Summary", heading_style))

                if hasattr(self.data_handler, 'calculations_dict'):
                    calc_dict = self.data_handler.calculations_dict

                    kpi_data = [
                        ['Key Performance Indicator', 'Value (SAR)'],
                        ['Beginning Cash Balance', f"{calc_dict.get('Cash Beginning Balance', 0):,.0f}"],
                        ['Total Operating Cash Inflow', f"{calc_dict.get('Total Operating Cash Inflow', 0):,.0f}"],
                        ['Total Operating Cash Outflow', f"{calc_dict.get('Total Operating Cash Outflow', 0):,.0f}"],
                        ['Net Operating Cash Flow', f"{calc_dict.get('Total Operating Cash Inflow', 0) - calc_dict.get('Total Operating Cash Outflow', 0):,.0f}"],
                        ['Ending Cash Balance', f"{calc_dict.get('Cash Ending Balance', 0):,.0f}"]
                    ]

                    kpi_table = Table(kpi_data, colWidths=[3*inch, 2*inch])
                    kpi_table.setStyle(TableStyle([
                        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#f8f9fa')),
                        ('TEXTCOLOR', (0, 0), (-1, 0), colors.black),
                        ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
                        ('ALIGN', (1, 1), (1, -1), 'RIGHT'),
                        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                        ('FONTSIZE', (0, 0), (-1, 0), 12),
                        ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
                        ('BACKGROUND', (0, 1), (-1, -1), colors.white),
                        ('GRID', (0, 0), (-1, -1), 1, colors.black),
                        ('FONTSIZE', (0, 1), (-1, -1), 10),
                    ]))
                    elements.append(kpi_table)
                    elements.append(Spacer(1, 30))

                # === Advanced KPI Section (from selected period) ===
                progress.setValue(30)
                elements.append(Paragraph("Advanced KPIs (Selected Period)", heading_style))

                if hasattr(self, 'kpi_range_slider'):
                    start_idx, end_idx = self.kpi_range_slider.value()
                    selected_date_cols = self.kpi_date_columns[start_idx:end_idx + 1]
                    selected_dates = self.kpi_date_objects[start_idx:end_idx + 1]

                    period_kpis = self.calculate_period_kpis(selected_date_cols, selected_dates)

                    adv_kpi_data = [['Metric', 'Value']]
                    for key, val in period_kpis.items():
                        if isinstance(val, numbers.Number):
                            # ✅ format numeric values nicely
                            adv_kpi_data.append([key, f"{val:,.2f}"])
                        elif isinstance(val, dict):
                            # ✅ expand dict values into multiple rows
                            adv_kpi_data.append([key, ""])  # main KPI name as a header row
                            for sub_key, sub_val in val.items():
                                if isinstance(sub_val, numbers.Number):
                                    adv_kpi_data.append([f"   ↳ {sub_key}", f"{sub_val:,.2f}"])
                                else:
                                    adv_kpi_data.append([f"   ↳ {sub_key}", str(sub_val)])
                        elif isinstance(val, (list, tuple)):
                            # ✅ expand lists/tuples as separate rows
                            adv_kpi_data.append([key, ""])
                            for i, item in enumerate(val, 1):
                                if isinstance(item, numbers.Number):
                                    adv_kpi_data.append([f"   [{i}]", f"{item:,.2f}"])
                                else:
                                    adv_kpi_data.append([f"   [{i}]", str(item)])
                        else:
                            # ✅ fallback: just convert to string
                            adv_kpi_data.append([key, str(val)])

                    adv_kpi_table = Table(adv_kpi_data, colWidths=[3*inch, 2*inch])
                    adv_kpi_table.setStyle(TableStyle([
                        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#e8f0fe')),
                        ('TEXTCOLOR', (0, 0), (-1, 0), colors.black),
                        ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
                        ('ALIGN', (1, 1), (1, -1), 'RIGHT'),
                        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                        ('FONTSIZE', (0, 0), (-1, 0), 11),
                        ('BOTTOMPADDING', (0, 0), (-1, 0), 10),
                        ('BACKGROUND', (0, 1), (-1, -1), colors.white),
                        ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
                        ('FONTSIZE', (0, 1), (-1, -1), 9),
                    ]))
                    elements.append(adv_kpi_table)
                    elements.append(Spacer(1, 30))

                # === Charts Section ===
                progress.setValue(50)
                elements.append(Paragraph("Charts and Visualizations", heading_style))
                elements.append(Spacer(1, 12))

                chart_images = []

                if hasattr(self.plot_handler, 'waterfall_cash_movement_fig'):
                    waterfall_img = self.plotly_to_image(self.plot_handler.waterfall_cash_movement_fig)
                    if waterfall_img:
                        chart_images.append(("Cash Movement Waterfall Chart", waterfall_img))

                if hasattr(self.plot_handler, 'monthly_cash_flow_fig'):
                    monthly_img = self.plotly_to_image(self.plot_handler.monthly_cash_flow_fig)
                    if monthly_img:
                        chart_images.append(("Monthly Cash Flow Trend", monthly_img))

                if hasattr(self.plot_handler, 'operating_cash_flow_diagram_fig'):
                    operating_img = self.plotly_to_image(self.plot_handler.operating_cash_flow_diagram_fig)
                    if operating_img:
                        chart_images.append(("Operating Cash Flow Analysis", operating_img))

                # Add inflow pie
                if hasattr(self, 'data_handler') and self.data_handler:
                    pie_fig = self.create_operating_inflow_pie()
                    pie_img = self.plotly_to_image(pie_fig)
                    if pie_img:
                        chart_images.append(("Operating Cash Inflow Distribution", pie_img))

                # Loop through and insert charts
                for chart_title, img_data in chart_images:
                    elements.append(Paragraph(chart_title, styles['Heading3']))
                    elements.append(Spacer(1, 6))
                    img = Image(io.BytesIO(img_data), width=6*inch, height=4*inch)
                    elements.append(img)
                    elements.append(Spacer(1, 20))

                # === Data Tables ===
                progress.setValue(80)
                elements.append(PageBreak())
                elements.append(Paragraph("Data Tables", heading_style))
                elements.append(Spacer(1, 12))

                if hasattr(self.data_handler, 'operating_cf_in') and self.data_handler.operating_cf_in is not None:
                    elements.append(Paragraph("Operating Cash Inflow Data", styles['Heading3']))
                    elements.append(Spacer(1, 6))

                    df = self.data_handler.operating_cf_in.head(20)
                    data = [df.columns.tolist()] + [[str(cell) if not pd.isna(cell) else "" for cell in row.tolist()] for _, row in df.iterrows()]
                    table = Table(data, repeatRows=1)
                    table.setStyle(TableStyle([
                        ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
                        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                        ('FONTSIZE', (0, 0), (-1, 0), 8),
                        ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
                        ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
                        ('GRID', (0, 0), (-1, -1), 1, colors.black),
                        ('FONTSIZE', (0, 1), (-1, -1), 6),
                    ]))
                    elements.append(table)
                    elements.append(Spacer(1, 20))

                # Footer
                progress.setValue(90)
                elements.append(Spacer(1, 30))
                elements.append(Paragraph("Report Notes:", styles['Heading3']))
                elements.append(Paragraph("• This report was generated automatically from the Cash Flow Dashboard application.", styles['Normal']))
                elements.append(Paragraph("• All figures are in Saudi Arabian Riyals (SAR) unless otherwise specified.", styles['Normal']))
                elements.append(Paragraph("• For detailed analysis, please refer to the interactive dashboard.", styles['Normal']))

                # Build PDF
                doc.build(elements)
                progress.setValue(100)
                progress.close()

                QMessageBox.information(
                    self,
                    "Export Successful",
                    f"Report has been successfully exported to:\n{file_path}\n\n"
                    f"File size: {os.path.getsize(file_path) / 1024:.1f} KB"
                )
                self.statusBar().showMessage(f"Report exported successfully to {os.path.basename(file_path)}")

            except Exception as e:
                if 'progress' in locals():
                    progress.close()
                QMessageBox.critical(
                    self,
                    "Export Error",
                    f"Failed to export report:\n{str(e)}\n\n"
                    "Please ensure you have write permissions to the selected location."
                )
                self.statusBar().showMessage("Report export failed")

    def plotly_to_image(self, fig, format='png', width=800, height=600):
        """Convert Plotly figure to image bytes"""
        try:
            # Export to image bytes
            img_bytes = pio.to_image(
                fig, 
                format=format, 
                width=width, 
                height=height,
                engine="kaleido"  # Use kaleido engine for better compatibility
            )
            return img_bytes
        except Exception as e:
            print(f"Error converting plot to image: {e}")
            return None


def main():
    """Main application entry point"""
    app = QApplication(sys.argv)
    
    # Set application properties
    app.setApplicationName("Cash Flow Dashboard")
    app.setApplicationVersion("1.0")
    app.setOrganizationName("Your Company")
    app.setWindowIcon(QIcon("analysis.ico"))
    
    # Create and show main window
    dashboard = CashFlowDashboard()
    dashboard.show()
    
    sys.exit(app.exec())

if __name__ == "__main__":
    main()