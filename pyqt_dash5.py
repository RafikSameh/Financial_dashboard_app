# integrated_pyqt_dashboard.py - Complete version with PyQt Operating Inflow Pie Chart
import sys
import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import plotly.graph_objs as go
import plotly.io as pio
import json

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
            font=dict(size=16, color="gray")
        )
        empty_fig.update_layout(
            xaxis=dict(showgrid=False, showticklabels=False, zeroline=False),
            yaxis=dict(showgrid=False, showticklabels=False, zeroline=False),
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
        """Create main dashboard tab with KPIs at top and scrollable content"""
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

        # KPI Section at the top - Fixed position
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
        kpi_title = QLabel("Key Performance Indicators")
        kpi_title.setFont(QFont("Arial", 16, QFont.Weight.Bold))
        kpi_title.setStyleSheet("color: #2E3440; margin-bottom: 15px;")
        kpi_title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        kpi_main_layout.addWidget(kpi_title)

        # KPI Grid Layout
        self.kpi_frame = QFrame()
        self.kpi_layout = QGridLayout(self.kpi_frame)
        self.kpi_layout.setSpacing(15)
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
        """Handle successful data loading"""
        self.data_handler = data_handler
        self.plot_handler = plot_handler
        
        # Update the pie widget with new data
        if hasattr(self, 'inflow_pie_widget'):
            self.inflow_pie_widget.data_handler = data_handler
            self.inflow_pie_widget.setup_data()
        
        self.update_dashboard()
        self.update_data_tables()
        
        self.refresh_btn.setEnabled(True)
        self.export_btn.setEnabled(True)
        
        self.statusBar().showMessage("Data loaded successfully")
    
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
        """Update KPI widgets with real data - improved version"""
        # Clear existing KPIs
        for i in reversed(range(self.kpi_layout.count())):
            widget = self.kpi_layout.itemAt(i).widget()
            if widget:
                widget.setParent(None)

        if self.data_handler:
            try:
                calc_dict = self.data_handler.calculations_dict

                # Define KPI data with better formatting
                kpis = [
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

                # Create enhanced KPI widgets
                for i, kpi in enumerate(kpis):
                    kpi_widget = self.create_enhanced_kpi_widget(
                        kpi["title"], 
                        kpi["value"], 
                        kpi["icon"], 
                        kpi["color"]
                    )

                    # Arrange in 2 rows if more than 3 KPIs
                    row = i // 3
                    col = i % 3
                    self.kpi_layout.addWidget(kpi_widget, row, col)

            except Exception as e:
                error_kpi = self.create_enhanced_kpi_widget(
                    "Error", 
                    f"Failed to load KPIs: {str(e)[:50]}...", 
                    "⚠️", 
                    "#dc3545"
                )
                self.kpi_layout.addWidget(error_kpi, 0, 0)


    def create_enhanced_kpi_widget(self, title, value, icon, color):
        """Create enhanced KPI widget with better styling"""
        kpi_widget = QFrame()
        kpi_widget.setFrameStyle(QFrame.Shape.Box)
        kpi_widget.setStyleSheet(f"""
            QFrame {{
                background-color: white;
                border: 2px solid {color};
                border-radius: 10px;
                padding: 15px;
                margin: 5px;
                min-height: 100px;
                max-height: 120px;
            }}
            QFrame:hover {{
                background-color: #f8f9fa;
                border-color: {color};
                box-shadow: 0 4px 8px rgba(0,0,0,0.1);
            }}
        """)

        layout = QVBoxLayout(kpi_widget)
        layout.setSpacing(5)

        # Icon and Title row
        header_layout = QHBoxLayout()

        icon_label = QLabel(icon)
        icon_label.setFont(QFont("Arial", 20))
        icon_label.setFixedSize(30, 30)
        icon_label.setAlignment(Qt.AlignmentFlag.AlignCenter)

        title_label = QLabel(title)
        title_label.setFont(QFont("Arial", 10, QFont.Weight.Bold))
        title_label.setStyleSheet(f"color: {color}; font-weight: bold;")
        title_label.setWordWrap(True)

        header_layout.addWidget(icon_label)
        header_layout.addWidget(title_label)
        header_layout.addStretch()

        # Value
        value_label = QLabel(value)
        value_label.setFont(QFont("Arial", 12, QFont.Weight.Bold))
        value_label.setStyleSheet("color: #212529;")
        value_label.setWordWrap(True)
        value_label.setAlignment(Qt.AlignmentFlag.AlignCenter)

        layout.addLayout(header_layout)
        layout.addWidget(value_label)
        layout.addStretch()

        return kpi_widget

    def create_enhanced_kpi_widget_with_trend(self, title, value, trend, icon, color):
        """Create enhanced KPI widget with trend indicator"""
        kpi_widget = QFrame()
        kpi_widget.setFrameStyle(QFrame.Shape.Box)
        kpi_widget.setStyleSheet(f"""
            QFrame {{
                background-color: white;
                border: 2px solid {color};
                border-radius: 10px;
                padding: 15px;
                margin: 5px;
                min-height: 110px;
                max-height: 130px;
            }}
            QFrame:hover {{
                background-color: #f8f9fa;
                border-color: {color};
                box-shadow: 0 4px 8px rgba(0,0,0,0.1);
            }}
        """)

        layout = QVBoxLayout(kpi_widget)
        layout.setSpacing(3)

        # Icon and Title row
        header_layout = QHBoxLayout()

        icon_label = QLabel(icon)
        icon_label.setFont(QFont("Arial", 18))
        icon_label.setFixedSize(25, 25)
        icon_label.setAlignment(Qt.AlignmentFlag.AlignCenter)

        title_label = QLabel(title)
        title_label.setFont(QFont("Arial", 9, QFont.Weight.Bold))
        title_label.setStyleSheet(f"color: {color}; font-weight: bold;")
        title_label.setWordWrap(True)

        header_layout.addWidget(icon_label)
        header_layout.addWidget(title_label)
        header_layout.addStretch()

        # Value
        value_label = QLabel(value)
        value_label.setFont(QFont("Arial", 11, QFont.Weight.Bold))
        value_label.setStyleSheet("color: #212529;")
        value_label.setWordWrap(True)
        value_label.setAlignment(Qt.AlignmentFlag.AlignCenter)

        # Trend (if provided)
        if trend and trend.strip():
            trend_label = QLabel(trend)
            trend_label.setFont(QFont("Arial", 10))
            trend_value = float(trend.replace('%', '').replace('+', ''))
            if trend_value > 0:
                trend_label.setStyleSheet("color: #28a745; font-weight: bold;")
                trend_text = f"↑ {trend}"
            elif trend_value < 0:
                trend_label.setStyleSheet("color: #dc3545; font-weight: bold;")
                trend_text = f"↓ {trend}"
            else:
                trend_label.setStyleSheet("color: #6c757d;")
                trend_text = f"→ {trend}"

            trend_label.setText(trend_text)
            trend_label.setAlignment(Qt.AlignmentFlag.AlignCenter)

            layout.addLayout(header_layout)
            layout.addWidget(value_label)
            layout.addWidget(trend_label)
        else:
            layout.addLayout(header_layout)
            layout.addWidget(value_label)
            layout.addStretch()

        return kpi_widget


    def preview_kpi_insights_with_slider(self):
        """Show KPI Insights inside the dashboard content area (scrollable)"""

        if not self.data_handler or not hasattr(self.data_handler, 'totals') or self.data_handler.totals is None:
            QMessageBox.warning(self, "Warning", "No totals data available for KPI insights!")
            return

        try:
            df = self.data_handler.totals.copy()

            # --- Identify date columns ---
            non_date_cols = [c for c in df.columns if any(k in str(c).lower()
                              for k in ['country', 'cash flow', 'category', 'item', 'type'])]
            if not non_date_cols:
                non_date_cols = ['Country', 'Cash Flow Type', 'Category', 'Item']
            date_columns = [c for c in df.columns if c not in non_date_cols]

            # Parse dates
            date_objects, successful_cols = [], []
            for col in date_columns:
                try:
                    date_obj = pd.to_datetime(str(col), errors="coerce")
                    if pd.notna(date_obj):
                        date_objects.append(date_obj.to_pydatetime())
                        successful_cols.append(col)
                except Exception:
                    pass
            date_columns = successful_cols
            if not date_objects:
                QMessageBox.warning(self, "Warning", "No valid date columns found in DataFrame!")
                return
            date_objects.sort()

            # --- Remove old KPI widget if exists ---
            if hasattr(self, "kpi_insights_widget") and self.kpi_insights_widget:
                self.scroll_layout.removeWidget(self.kpi_insights_widget)
                self.kpi_insights_widget.deleteLater()

            # --- Create KPI Insights widget ---
            self.kpi_insights_widget = QWidget()
            layout = QVBoxLayout(self.kpi_insights_widget)

            # Title
            title_label = QLabel("Cash Flow KPI Insights")
            title_label.setStyleSheet("font-size:18px;font-weight:bold;color:#2E86AB;")
            title_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
            layout.addWidget(title_label)

            # RangeSlider
            self.range_slider = RangeSlider(0, len(date_objects)-1)
            self.range_slider.setValue(0, len(date_objects)-1)
            self.range_label = QLabel(
                f"{date_objects[0].strftime('%Y-%m-%d')} → {date_objects[-1].strftime('%Y-%m-%d')}")
            self.range_label.setStyleSheet("background:#e3f2fd;padding:6px;border-radius:3px;")

            slider_box = QVBoxLayout()
            slider_box.addWidget(self.range_label)
            slider_box.addWidget(self.range_slider)
            layout.addLayout(slider_box)

            # KPI sections
            self.primary_kpi_layout = QGridLayout()
            self.secondary_kpi_layout = QGridLayout()
            self.performance_kpi_layout = QGridLayout()

            layout.addWidget(QLabel("Primary KPIs"))
            layout.addLayout(self.primary_kpi_layout)
            layout.addWidget(QLabel("Secondary KPIs"))
            layout.addLayout(self.secondary_kpi_layout)
            layout.addWidget(QLabel("Performance Indicators"))
            layout.addLayout(self.performance_kpi_layout)

            # Add to dashboard scrollable layout
            self.scroll_layout.addWidget(self.kpi_insights_widget)

            # --- Update logic ---
            def update_kpi_insights(values):
                start_idx, end_idx = values
                self.range_label.setText(
                    f"{date_objects[start_idx].strftime('%Y-%m-%d')} → {date_objects[end_idx].strftime('%Y-%m-%d')}"
                )
                selected_date_cols = date_columns[start_idx:end_idx+1]
                selected_dates = date_objects[start_idx:end_idx+1]

                kpi_insights = calculate_period_kpis(df, selected_date_cols, selected_dates)
                update_all_kpi_widgets(kpi_insights, len(selected_date_cols))

            def calculate_period_kpis(df, selected_cols, selected_dates):
                kpis = {}
                for _, row in df.iterrows():
                    item = row.get("Item", "Unknown")
                    vals = pd.to_numeric(row[selected_cols], errors="coerce").fillna(0).tolist()
                    if vals:
                        total = sum(vals)
                        kpis[item] = {
                            "total": total,
                            "average": total/len(vals),
                            "trend_pct": ((vals[-1]-vals[0])/vals[0]*100 if vals[0] else 0),
                        }
                return kpis

            def clear_layout(layout):
                for i in reversed(range(layout.count())):
                    w = layout.itemAt(i).widget()
                    if w:
                        w.setParent(None)

            def update_all_kpi_widgets(kpi_insights, num_periods):
                clear_layout(self.primary_kpi_layout)
                clear_layout(self.secondary_kpi_layout)
                clear_layout(self.performance_kpi_layout)

                inflow = sum(v["total"] for k,v in kpi_insights.items() if "Inflow" in k)
                outflow = sum(abs(v["total"]) for k,v in kpi_insights.items() if "Outflow" in k)
                net = inflow - outflow

                self.primary_kpi_layout.addWidget(QLabel(f"Net Cash Flow: {net:,.0f}"), 0, 0)
                self.secondary_kpi_layout.addWidget(QLabel(f"Periods: {num_periods}"), 0, 0)

                for i,(item,metrics) in enumerate(kpi_insights.items()):
                    self.performance_kpi_layout.addWidget(QLabel(f"{item}: {metrics['trend_pct']:+.1f}%"), i//3, i%3)

            # Connect slider
            self.range_slider.valueChanged.connect(update_kpi_insights)

            # Initial update
            update_kpi_insights(self.range_slider.value())

        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to create KPI insights: {str(e)}")

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

                # Container for the 'Flowable' objects
                elements = []

                # Define styles
                styles = getSampleStyleSheet()
                title_style = ParagraphStyle(
                    'CustomTitle',
                    parent=styles['Heading1'],
                    fontSize=24,
                    spaceAfter=30,
                    alignment=1,  # Center alignment
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

                # Executive Summary with KPIs
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

                progress.setValue(40)

                # Charts section
                elements.append(Paragraph("Charts and Visualizations", heading_style))
                elements.append(Spacer(1, 12))

                # Export charts as images and add to PDF
                chart_images = []

                # Waterfall chart
                if hasattr(self.plot_handler, 'waterfall_cash_movement_fig'):
                    waterfall_img = self.plotly_to_image(self.plot_handler.waterfall_cash_movement_fig)
                    if waterfall_img:
                        chart_images.append(("Cash Movement Waterfall Chart", waterfall_img))

                progress.setValue(50)

                # Monthly cash flow chart
                if hasattr(self.plot_handler, 'monthly_cash_flow_fig'):
                    monthly_img = self.plotly_to_image(self.plot_handler.monthly_cash_flow_fig)
                    if monthly_img:
                        chart_images.append(("Monthly Cash Flow Trend", monthly_img))

                progress.setValue(60)

                # Operating cash flow diagram
                if hasattr(self.plot_handler, 'operating_cash_flow_diagram_fig'):
                    operating_img = self.plotly_to_image(self.plot_handler.operating_cash_flow_diagram_fig)
                    if operating_img:
                        chart_images.append(("Operating Cash Flow Analysis", operating_img))

                progress.setValue(70)

                # Operating inflow pie chart
                if hasattr(self, 'data_handler') and self.data_handler:
                    pie_fig = self.create_operating_inflow_pie()
                    pie_img = self.plotly_to_image(pie_fig)
                    if pie_img:
                        chart_images.append(("Operating Cash Inflow Distribution", pie_img))

                # Add charts to PDF
                for chart_title, img_data in chart_images:
                    elements.append(Paragraph(chart_title, styles['Heading3']))
                    elements.append(Spacer(1, 6))

                    # Create image from bytes
                    img = Image(io.BytesIO(img_data), width=6*inch, height=4*inch)
                    elements.append(img)
                    elements.append(Spacer(1, 20))

                progress.setValue(80)

                # Data tables section
                elements.append(PageBreak())
                elements.append(Paragraph("Data Tables", heading_style))
                elements.append(Spacer(1, 12))

                # Operating Cash Inflow Table
                if hasattr(self.data_handler, 'operating_cf_in') and self.data_handler.operating_cf_in is not None:
                    elements.append(Paragraph("Operating Cash Inflow Data", styles['Heading3']))
                    elements.append(Spacer(1, 6))

                    # Convert dataframe to table (limited rows for PDF)
                    df = self.data_handler.operating_cf_in.head(20)  # Limit to first 20 rows
                    data = [df.columns.tolist()]
                    for index, row in df.iterrows():
                        data.append([str(cell) if not pd.isna(cell) else "" for cell in row.tolist()])

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

                progress.setValue(90)

                # Footer information
                elements.append(Spacer(1, 30))
                elements.append(Paragraph("Report Notes:", styles['Heading3']))
                elements.append(Paragraph("• This report was generated automatically from the Cash Flow Dashboard application.", styles['Normal']))
                elements.append(Paragraph("• All figures are in Saudi Arabian Riyals (SAR) unless otherwise specified.", styles['Normal']))
                elements.append(Paragraph("• For detailed analysis, please refer to the interactive dashboard.", styles['Normal']))

                # Build PDF
                doc.build(elements)
                progress.setValue(100)
                progress.close()

                # Success message
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