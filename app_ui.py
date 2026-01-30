import os
import sys
import datetime
import threading
from PyQt5.QtWidgets import (QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, 
                            QTabWidget, QPushButton, QLabel, QFileDialog, 
                            QComboBox, QMessageBox, QGroupBox, QGridLayout,
                            QTableWidget, QTableWidgetItem, QHeaderView, QSplitter,
                            QTextEdit, QSpinBox, QDateEdit, QCheckBox, QFrame, QApplication,
                            QProgressDialog, QProgressBar, QStyleFactory)
from PyQt5.QtCore import Qt, QSize, QDate, pyqtSignal, QObject
from PyQt5.QtGui import QIcon, QFont, QColor, QPalette, QLinearGradient, QBrush
from PyQt5.QtWebEngineWidgets import QWebEngineView
import matplotlib.pyplot as plt
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.figure import Figure
import pandas as pd
import numpy as np
from sales_analyzer import SalesAnalyzer

# 设置matplotlib的样式
plt.style.use('seaborn-v0_8')

# 自定义颜色方案
COLOR_PRIMARY = "#3498db"       # 主色调（蓝色）
COLOR_SECONDARY = "#2ecc71"     # 次要色调（绿色）
COLOR_ACCENT = "#e74c3c"        # 强调色（红色）
COLOR_BACKGROUND = "#f9f9f9"    # 背景色（淡灰）
COLOR_TEXT = "#2c3e50"          # 文字颜色（深蓝灰）
COLOR_BUTTON = "#3498db"        # 按钮颜色
COLOR_HOVER = "#2980b9"         # 悬停颜色
COLOR_LIGHT_BG = "#ecf0f1"      # 浅色背景
COLOR_DARK_BG = "#34495e"       # 暗色背景

class WorkerSignals(QObject):
    """用于在线程间传递信号的类"""
    finished = pyqtSignal()
    error = pyqtSignal(str)
    progress = pyqtSignal(int, str)
    result = pyqtSignal(object)

class DataLoadWorker(threading.Thread):
    """处理数据加载的工作线程"""
    def __init__(self, analyzer, file_path):
        super().__init__()
        self.analyzer = analyzer
        self.file_path = file_path
        self.signals = WorkerSignals()
        self.daemon = True  # 设置为守护线程

    def run(self):
        try:
            # 发送进度信号
            self.signals.progress.emit(10, "正在读取文件...")
            
            # 获取文件扩展名
            file_ext = os.path.splitext(self.file_path)[1].lower()
            
            # 判断文件类型和大小
            file_size = os.path.getsize(self.file_path) / (1024 * 1024)  # 转换为MB
            is_large_file = file_size > 10  # 大于10MB视为大文件
            
            # 更新进度
            self.signals.progress.emit(20, f"正在加载数据到内存...({file_size:.1f} MB)")
            
            # 加载数据
            self.analyzer.load_data(self.file_path)
            
            # 更新进度
            self.signals.progress.emit(50, "正在处理日期数据...")
            
            # 确保日期列为日期类型
            if '日期' in self.analyzer.df.columns:
                self.analyzer.df['日期'] = pd.to_datetime(self.analyzer.df['日期'])
                
                # 如果没有年、月、季度列，则从日期列生成
                if '年' not in self.analyzer.df.columns:
                    self.analyzer.df['年'] = self.analyzer.df['日期'].dt.year
                if '月' not in self.analyzer.df.columns:
                    self.analyzer.df['月'] = self.analyzer.df['日期'].dt.month
                if '季度' not in self.analyzer.df.columns:
                    self.analyzer.df['季度'] = (self.analyzer.df['日期'].dt.month - 1) // 3 + 1
            
            # 完成并发送结果
            self.signals.progress.emit(90, "数据加载完成!")
            self.signals.result.emit(self.analyzer.df)
            self.signals.finished.emit()
            
        except Exception as e:
            self.signals.error.emit(str(e))

class MatplotlibCanvas(FigureCanvas):
    """Matplotlib画布类，用于在PyQt中显示图表"""
    def __init__(self, parent=None, width=5, height=4, dpi=100):
        self.fig = Figure(figsize=(width, height), dpi=dpi)
        self.axes = self.fig.add_subplot(111)
        
        # 设置图表样式
        self.fig.patch.set_facecolor(COLOR_LIGHT_BG)
        self.axes.set_facecolor(COLOR_LIGHT_BG)
        
        # 增强网格线样式
        self.axes.grid(True, linestyle='--', alpha=0.7, color="#bdc3c7")
        
        # 设置轴标签和标题的字体样式
        self.axes.xaxis.label.set_color(COLOR_TEXT)
        self.axes.yaxis.label.set_color(COLOR_TEXT)
        self.axes.title.set_color(COLOR_TEXT)
        
        super(MatplotlibCanvas, self).__init__(self.fig)
        self.setParent(parent)
        self.setMinimumSize(400, 300)
        
    def clear(self):
        """清除图表"""
        self.axes.clear()
        self.draw()
        
    def plot(self, *args, **kwargs):
        """绘制图表"""
        self.axes.plot(*args, **kwargs)
        self.draw()

class EcommerceAnalysisApp(QMainWindow):
    """电商销售分析应用主窗口"""
    def __init__(self):
        super().__init__()
        self.analyzer = SalesAnalyzer()
        self.setup_style()
        self.init_ui()
        
    def setup_style(self):
        """设置应用程序样式"""
        # 设置应用程序样式为Fusion，这样我们可以自定义调色板
        QApplication.setStyle(QStyleFactory.create('Fusion'))
        
        # 创建自定义调色板
        palette = QPalette()
        palette.setColor(QPalette.Window, QColor(COLOR_BACKGROUND))
        palette.setColor(QPalette.WindowText, QColor(COLOR_TEXT))
        palette.setColor(QPalette.Base, QColor(COLOR_LIGHT_BG))
        palette.setColor(QPalette.AlternateBase, QColor(COLOR_BACKGROUND))
        palette.setColor(QPalette.ToolTipBase, QColor(COLOR_DARK_BG))
        palette.setColor(QPalette.ToolTipText, QColor(COLOR_LIGHT_BG))
        palette.setColor(QPalette.Text, QColor(COLOR_TEXT))
        palette.setColor(QPalette.Button, QColor(COLOR_BUTTON))
        palette.setColor(QPalette.ButtonText, QColor("white"))
        palette.setColor(QPalette.BrightText, QColor("white"))
        palette.setColor(QPalette.Link, QColor(COLOR_PRIMARY))
        palette.setColor(QPalette.Highlight, QColor(COLOR_PRIMARY))
        palette.setColor(QPalette.HighlightedText, QColor("white"))
        
        # 应用调色板
        QApplication.setPalette(palette)
        
        # 设置全局样式表
        QApplication.instance().setStyleSheet("""
            QMainWindow {
                background-color: #f9f9f9;
            }
            
            QTabWidget::pane {
                border: 1px solid #bdc3c7;
                border-radius: 4px;
                padding: 2px;
                background-color: #f9f9f9;
            }
            
            QTabBar::tab {
                background-color: #ecf0f1;
                color: #2c3e50;
                min-width: 100px;
                min-height: 30px;
                padding: 5px 10px;
                margin-right: 2px;
                border: 1px solid #bdc3c7;
                border-top-left-radius: 4px;
                border-top-right-radius: 4px;
            }
            
            QTabBar::tab:selected {
                background-color: #3498db;
                color: white;
                border-bottom-color: #3498db;
            }
            
            QTabBar::tab:hover:!selected {
                background-color: #d0d9e0;
            }
            
            QPushButton {
                background-color: #3498db;
                color: white;
                border: none;
                border-radius: 4px;
                padding: 8px 16px;
                min-height: 30px;
                font-weight: bold;
            }
            
            QPushButton:hover {
                background-color: #2980b9;
            }
            
            QPushButton:pressed {
                background-color: #21618c;
            }
            
            QPushButton:disabled {
                background-color: #bdc3c7;
                color: #7f8c8d;
            }
            
            QComboBox {
                border: 1px solid #bdc3c7;
                border-radius: 4px;
                padding: 4px 10px;
                min-height: 25px;
                background-color: white;
            }
            
            QComboBox:hover {
                border: 1px solid #3498db;
            }
            
            QComboBox::drop-down {
                subcontrol-origin: padding;
                subcontrol-position: top right;
                width: 20px;
                border-left: 1px solid #bdc3c7;
                border-top-right-radius: 4px;
                border-bottom-right-radius: 4px;
            }
            
            QGroupBox {
                font-weight: bold;
                border: 1px solid #bdc3c7;
                border-radius: 4px;
                margin-top: 10px;
                padding-top: 15px;
            }
            
            QGroupBox::title {
                subcontrol-origin: margin;
                subcontrol-position: top left;
                padding: 0 5px;
                color: #2c3e50;
            }
            
            QTableWidget {
                gridline-color: #d4d4d4;
                border: 1px solid #bdc3c7;
                border-radius: 4px;
                selection-background-color: #3498db;
                selection-color: white;
            }
            
            QHeaderView::section {
                background-color: #ecf0f1;
                color: #2c3e50;
                padding: 5px;
                border: 1px solid #bdc3c7;
                font-weight: bold;
            }
            
            QTextEdit {
                border: 1px solid #bdc3c7;
                border-radius: 4px;
                padding: 5px;
                background-color: white;
            }
            
            QDateEdit, QSpinBox {
                border: 1px solid #bdc3c7;
                border-radius: 4px;
                padding: 4px 10px;
                background-color: white;
            }
            
            QCheckBox {
                spacing: 5px;
            }
            
            QCheckBox::indicator {
                width: 18px;
                height: 18px;
            }
            
            QCheckBox::indicator:checked {
                background-color: #3498db;
                border: 1px solid #2980b9;
                border-radius: 3px;
            }
            
            QStatusBar {
                background-color: #ecf0f1;
                color: #2c3e50;
                border-top: 1px solid #bdc3c7;
            }
            
            QSplitter::handle {
                background-color: #bdc3c7;
            }
            
            QSplitter::handle:horizontal {
                width: 2px;
            }
            
            QProgressBar {
                border: 1px solid #bdc3c7;
                border-radius: 4px;
                text-align: center;
                color: white;
                background-color: #ecf0f1;
            }
            
            QProgressBar::chunk {
                background-color: #3498db;
                width: 10px;
                margin: 0.5px;
            }
        """)
        
    def init_ui(self):
        """初始化用户界面"""
        self.setWindowTitle('基于大数据的电商平台商品销售趋势分析与决策软件')
        self.setGeometry(100, 100, 1280, 800)  # 增大窗口尺寸
        
        # 创建中央部件
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        
        # 主布局
        main_layout = QVBoxLayout(central_widget)
        main_layout.setContentsMargins(10, 10, 10, 10)  # 设置布局边距
        main_layout.setSpacing(10)  # 设置控件间距
        
        # 添加标题标签
        title_label = QLabel("基于大数据的电商平台商品销售趋势分析与决策软件")
        title_label.setAlignment(Qt.AlignCenter)
        title_label.setStyleSheet(f"""
            font-size: 24px;
            font-weight: bold;
            color: {COLOR_TEXT};
            margin: 10px 0;
            padding: 5px;
        """)
        main_layout.addWidget(title_label)
        
        # 创建主选项卡
        self.main_tabs = QTabWidget()
        self.main_tabs.setStyleSheet("""
            QTabWidget::tab-bar {
                alignment: center;
            }
        """)
        
        # 创建四个主要模块的选项卡
        self.data_management_tab = QWidget()
        self.data_analysis_tab = QWidget()
        self.trend_prediction_tab = QWidget()
        self.report_generation_tab = QWidget()
        
        # 设置四个主选项卡
        self.main_tabs.addTab(self.data_management_tab, "数据管理")
        self.main_tabs.addTab(self.data_analysis_tab, "数据分析")
        self.main_tabs.addTab(self.trend_prediction_tab, "趋势预测")
        self.main_tabs.addTab(self.report_generation_tab, "报告生成")
        
        # 初始化四个主要模块
        self.init_data_management_tab()
        self.init_data_analysis_tab()
        self.init_trend_prediction_tab()
        self.init_report_generation_tab()
        
        main_layout.addWidget(self.main_tabs)
        
        # 状态栏
        self.statusBar().setStyleSheet(f"""
            QStatusBar {{
                font-size: 12px;
                padding: 3px;
                background-color: {COLOR_LIGHT_BG};
                color: {COLOR_TEXT};
            }}
        """)
        self.statusBar().showMessage('就绪')
    
    def init_data_management_tab(self):
        """初始化数据管理选项卡"""
        layout = QVBoxLayout(self.data_management_tab)
        layout.setContentsMargins(15, 15, 15, 15)
        layout.setSpacing(15)
        
        # 数据加载区域
        data_group = QGroupBox("数据加载")
        data_group.setStyleSheet(f"""
            QGroupBox {{
                font-size: 14px;
                font-weight: bold;
                border: 2px solid {COLOR_PRIMARY};
                border-radius: 6px;
                margin-top: 12px;
                padding-top: 15px;
            }}
            QGroupBox::title {{
                subcontrol-origin: margin;
                subcontrol-position: top left;
                padding: 0 10px;
                color: {COLOR_PRIMARY};
                background-color: {COLOR_BACKGROUND};
            }}
        """)
        data_layout = QHBoxLayout()
        data_layout.setContentsMargins(10, 15, 10, 10)
        data_layout.setSpacing(10)
        
        self.load_btn = QPushButton("加载数据")
        self.load_btn.setIcon(QIcon.fromTheme("document-open"))
        self.load_btn.setMinimumHeight(40)
        self.load_btn.setStyleSheet(f"""
            QPushButton {{
                background-color: {COLOR_PRIMARY};
                font-size: 14px;
                min-width: 120px;
            }}
        """)
        self.load_btn.clicked.connect(self.load_data)
        data_layout.addWidget(self.load_btn)
        
        self.data_path_label = QLabel("未加载数据")
        self.data_path_label.setStyleSheet(f"""
            QLabel {{
                font-size: 13px;
                color: {COLOR_TEXT};
                padding: 5px 10px;
                background-color: white;
                border: 1px solid #ddd;
                border-radius: 4px;
            }}
        """)
        data_layout.addWidget(self.data_path_label)
        data_layout.addStretch()
        
        data_group.setLayout(data_layout)
        layout.addWidget(data_group)
        
        # 数据基本信息表格
        info_group = QGroupBox("数据基本信息")
        info_group.setStyleSheet(f"""
            QGroupBox {{
                font-size: 14px;
                font-weight: bold;
                border: 2px solid {COLOR_SECONDARY};
                border-radius: 6px;
                margin-top: 12px;
                padding-top: 15px;
            }}
            QGroupBox::title {{
                subcontrol-origin: margin;
                subcontrol-position: top left;
                padding: 0 10px;
                color: {COLOR_SECONDARY};
                background-color: {COLOR_BACKGROUND};
            }}
        """)
        info_layout = QVBoxLayout()
        info_layout.setContentsMargins(10, 15, 10, 10)
        
        self.data_info_table = QTableWidget()
        self.data_info_table.setColumnCount(2)
        self.data_info_table.setHorizontalHeaderLabels(["属性", "值"])
        self.data_info_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.data_info_table.setAlternatingRowColors(True)
        self.data_info_table.setStyleSheet("""
            QTableWidget {
                font-size: 12px;
            }
            QTableWidget::item {
                padding: 5px;
            }
        """)
        info_layout.addWidget(self.data_info_table)
        
        info_group.setLayout(info_layout)
        layout.addWidget(info_group)
        
        # 筛选控制区域
        filter_group = QGroupBox("数据筛选")
        filter_group.setStyleSheet(f"""
            QGroupBox {{
                font-size: 14px;
                font-weight: bold;
                border: 2px solid {COLOR_ACCENT};
                border-radius: 6px;
                margin-top: 12px;
                padding-top: 15px;
            }}
            QGroupBox::title {{
                subcontrol-origin: margin;
                subcontrol-position: top left;
                padding: 0 10px;
                color: {COLOR_ACCENT};
                background-color: {COLOR_BACKGROUND};
            }}
        """)
        filter_layout = QHBoxLayout()
        filter_layout.setContentsMargins(10, 15, 10, 10)
        filter_layout.setSpacing(10)
        
        # 年份筛选
        year_label = QLabel("年份:")
        year_label.setStyleSheet("font-weight: bold;")
        self.year_combo = QComboBox()
        self.year_combo.addItem("全部")
        self.year_combo.setMinimumWidth(100)
        self.year_combo.currentTextChanged.connect(self.update_full_data_view)
        
        # 季度筛选
        quarter_label = QLabel("季度:")
        quarter_label.setStyleSheet("font-weight: bold;")
        self.quarter_combo = QComboBox()
        self.quarter_combo.addItem("全部")
        self.quarter_combo.addItems(["1", "2", "3", "4"])
        self.quarter_combo.setMinimumWidth(80)
        self.quarter_combo.currentTextChanged.connect(self.update_full_data_view)
        
        # 月份筛选
        month_label = QLabel("月份:")
        month_label.setStyleSheet("font-weight: bold;")
        self.month_combo = QComboBox()
        self.month_combo.addItem("全部")
        self.month_combo.addItems([str(i) for i in range(1, 13)])
        self.month_combo.setMinimumWidth(80)
        self.month_combo.currentTextChanged.connect(self.update_full_data_view)
        
        # 商品类别筛选
        category_label = QLabel("商品类别:")
        category_label.setStyleSheet("font-weight: bold;")
        self.data_category_combo = QComboBox()
        self.data_category_combo.addItem("全部")
        self.data_category_combo.setMinimumWidth(120)
        self.data_category_combo.currentTextChanged.connect(self.update_full_data_view)
        
        # 添加筛选控件到布局
        filter_layout.addWidget(year_label)
        filter_layout.addWidget(self.year_combo)
        filter_layout.addWidget(quarter_label)
        filter_layout.addWidget(self.quarter_combo)
        filter_layout.addWidget(month_label)
        filter_layout.addWidget(self.month_combo)
        filter_layout.addWidget(category_label)
        filter_layout.addWidget(self.data_category_combo)
        filter_layout.addStretch()
        
        # 重置筛选按钮
        self.reset_filter_btn = QPushButton("重置筛选")
        self.reset_filter_btn.setIcon(QIcon.fromTheme("edit-clear"))
        self.reset_filter_btn.setStyleSheet(f"""
            QPushButton {{
                background-color: {COLOR_DARK_BG};
                min-width: 100px;
            }}
        """)
        self.reset_filter_btn.clicked.connect(self.reset_data_filters)
        filter_layout.addWidget(self.reset_filter_btn)
        
        filter_group.setLayout(filter_layout)
        layout.addWidget(filter_group)
        
        # 统计信息区域
        stats_layout = QHBoxLayout()
        stats_layout.setContentsMargins(5, 5, 5, 5)
        
        stats_frame = QFrame()
        stats_frame.setFrameShape(QFrame.StyledPanel)
        stats_frame.setStyleSheet(f"""
            QFrame {{
                background-color: {COLOR_LIGHT_BG};
                border-radius: 5px;
                border: 1px solid #ddd;
            }}
        """)
        stats_inner_layout = QHBoxLayout(stats_frame)
        
        self.filtered_count_label = QLabel("已筛选记录数: 0")
        self.filtered_count_label.setStyleSheet(f"""
            QLabel {{
                font-weight: bold;
                font-size: 13px;
                color: {COLOR_PRIMARY};
                padding: 3px;
            }}
        """)
        stats_inner_layout.addWidget(self.filtered_count_label)
        
        self.total_count_label = QLabel("总记录数: 0")
        self.total_count_label.setStyleSheet(f"""
            QLabel {{
                font-weight: bold;
                font-size: 13px;
                color: {COLOR_TEXT};
                padding: 3px;
            }}
        """)
        stats_inner_layout.addWidget(self.total_count_label)
        
        stats_inner_layout.addStretch()
        
        stats_layout.addWidget(stats_frame)
        layout.addLayout(stats_layout)
        
        # 完整数据表格
        self.full_data_table = QTableWidget()
        self.full_data_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)
        self.full_data_table.setAlternatingRowColors(True)  # 交替行颜色
        self.full_data_table.setStyleSheet("""
            QTableWidget {
                font-size: 12px;
                gridline-color: #d4d4d4;
            }
            QTableWidget::item {
                padding: 4px;
            }
            QTableWidget::item:selected {
                background-color: #3498db;
                color: white;
            }
        """)
        layout.addWidget(self.full_data_table)
    
    def init_data_analysis_tab(self):
        """初始化数据分析选项卡"""
        layout = QVBoxLayout(self.data_analysis_tab)
        layout.setContentsMargins(15, 15, 15, 15)
        layout.setSpacing(15)
        
        # 分析控制区域
        control_group = QGroupBox("分析控制")
        control_group.setStyleSheet(f"""
            QGroupBox {{
                font-size: 14px;
                font-weight: bold;
                border: 2px solid {COLOR_PRIMARY};
                border-radius: 6px;
                margin-top: 12px;
                padding-top: 15px;
            }}
            QGroupBox::title {{
                subcontrol-origin: margin;
                subcontrol-position: top left;
                padding: 0 10px;
                color: {COLOR_PRIMARY};
                background-color: {COLOR_BACKGROUND};
            }}
        """)
        control_layout = QHBoxLayout()
        control_layout.setContentsMargins(10, 15, 10, 10)
        control_layout.setSpacing(15)
        
        # 时间单位选择
        time_label = QLabel("时间单位:")
        time_label.setStyleSheet("font-weight: bold;")
        self.time_combo = QComboBox()
        self.time_combo.addItems(["日", "月", "季度", "年"])
        self.time_combo.setCurrentIndex(1)  # 默认选择"月"
        self.time_combo.setMinimumWidth(100)
        self.time_combo.currentTextChanged.connect(self.update_analysis)
        
        # 商品类别选择
        category_label = QLabel("商品类别:")
        category_label.setStyleSheet("font-weight: bold;")
        self.category_combo = QComboBox()
        self.category_combo.addItems(["全部"])
        self.category_combo.setMinimumWidth(150)
        self.category_combo.currentTextChanged.connect(self.update_analysis)
        
        # 地区级别选择
        region_label = QLabel("地区级别:")
        region_label.setStyleSheet("font-weight: bold;")
        self.region_combo = QComboBox()
        self.region_combo.addItems(["省份", "城市"])
        self.region_combo.setMinimumWidth(100)
        self.region_combo.currentTextChanged.connect(self.update_analysis)
        
        # 刷新分析按钮
        refresh_btn = QPushButton("刷新分析")
        refresh_btn.setIcon(QIcon.fromTheme("view-refresh"))
        refresh_btn.setStyleSheet(f"""
            QPushButton {{
                background-color: {COLOR_SECONDARY};
                min-width: 120px;
            }}
        """)
        refresh_btn.clicked.connect(self.update_analysis)
        
        control_layout.addWidget(time_label)
        control_layout.addWidget(self.time_combo)
        control_layout.addWidget(category_label)
        control_layout.addWidget(self.category_combo)
        control_layout.addWidget(region_label)
        control_layout.addWidget(self.region_combo)
        control_layout.addWidget(refresh_btn)
        control_layout.addStretch()
        
        control_group.setLayout(control_layout)
        layout.addWidget(control_group)
        
        # 分析结果选项卡
        self.analysis_tabs = QTabWidget()
        self.analysis_tabs.setStyleSheet(f"""
            QTabWidget::pane {{
                border: 1px solid #bdc3c7;
                border-radius: 5px;
                padding: 5px;
                background-color: white;
            }}
            
            QTabBar::tab {{
                background-color: #ecf0f1;
                color: {COLOR_TEXT};
                min-width: 120px;
                padding: 8px 15px;
                margin-right: 2px;
                border-top-left-radius: 4px;
                border-top-right-radius: 4px;
            }}
            
            QTabBar::tab:selected {{
                background-color: {COLOR_PRIMARY};
                color: white;
            }}
            
            QTabBar::tab:hover:!selected {{
                background-color: #d0d9e0;
            }}
        """)
        
        # 销售趋势分析选项卡
        self.trend_tab = QWidget()
        trend_layout = QVBoxLayout(self.trend_tab)
        trend_layout.setContentsMargins(10, 15, 10, 10)
        trend_layout.setSpacing(15)
        
        # 添加标题标签
        trend_title = QLabel("销售趋势分析")
        trend_title.setAlignment(Qt.AlignCenter)
        trend_title.setStyleSheet(f"""
            font-size: 16px;
            font-weight: bold;
            color: {COLOR_PRIMARY};
            margin-bottom: 5px;
        """)
        trend_layout.addWidget(trend_title)
        
        # 销售趋势图
        trend_frame = QFrame()
        trend_frame.setFrameShape(QFrame.StyledPanel)
        trend_frame.setStyleSheet("background-color: white; border-radius: 5px; border: 1px solid #ddd;")
        trend_inner_layout = QVBoxLayout(trend_frame)
        
        self.trend_canvas = MatplotlibCanvas(self.trend_tab)
        trend_inner_layout.addWidget(self.trend_canvas)
        
        trend_layout.addWidget(trend_frame)
        
        # 销售热图
        heatmap_frame = QFrame()
        heatmap_frame.setFrameShape(QFrame.StyledPanel)
        heatmap_frame.setStyleSheet("background-color: white; border-radius: 5px; border: 1px solid #ddd;")
        heatmap_inner_layout = QVBoxLayout(heatmap_frame)
        
        heatmap_title = QLabel("月度-类别销售热力图")
        heatmap_title.setAlignment(Qt.AlignCenter)
        heatmap_title.setStyleSheet(f"""
            font-size: 14px;
            font-weight: bold;
            color: {COLOR_TEXT};
        """)
        heatmap_inner_layout.addWidget(heatmap_title)
        
        self.heatmap_canvas = MatplotlibCanvas(self.trend_tab)
        heatmap_inner_layout.addWidget(self.heatmap_canvas)
        
        trend_layout.addWidget(heatmap_frame)
        
        self.analysis_tabs.addTab(self.trend_tab, "销售趋势")
        
        # 商品类别分析选项卡
        self.category_tab = QWidget()
        category_layout = QVBoxLayout(self.category_tab)
        category_layout.setContentsMargins(10, 15, 10, 10)
        category_layout.setSpacing(15)
        
        # 添加标题标签
        category_title = QLabel("商品类别分析")
        category_title.setAlignment(Qt.AlignCenter)
        category_title.setStyleSheet(f"""
            font-size: 16px;
            font-weight: bold;
            color: {COLOR_PRIMARY};
            margin-bottom: 5px;
        """)
        category_layout.addWidget(category_title)
        
        # 类别销售额对比图
        category_frame = QFrame()
        category_frame.setFrameShape(QFrame.StyledPanel)
        category_frame.setStyleSheet("background-color: white; border-radius: 5px; border: 1px solid #ddd;")
        category_inner_layout = QVBoxLayout(category_frame)
        
        self.category_canvas = MatplotlibCanvas(self.category_tab)
        category_inner_layout.addWidget(self.category_canvas)
        
        category_layout.addWidget(category_frame)
        
        # 热销商品表格
        products_frame = QFrame()
        products_frame.setFrameShape(QFrame.StyledPanel)
        products_frame.setStyleSheet("background-color: white; border-radius: 5px; border: 1px solid #ddd;")
        products_inner_layout = QVBoxLayout(products_frame)
        
        products_title = QLabel("热销商品TOP10")
        products_title.setAlignment(Qt.AlignCenter)
        products_title.setStyleSheet(f"""
            font-size: 14px;
            font-weight: bold;
            color: {COLOR_TEXT};
        """)
        products_inner_layout.addWidget(products_title)
        
        self.top_products_table = QTableWidget()
        self.top_products_table.setColumnCount(2)
        self.top_products_table.setHorizontalHeaderLabels(["商品名称", "销售额"])
        self.top_products_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.top_products_table.setAlternatingRowColors(True)
        self.top_products_table.setStyleSheet("""
            QTableWidget {
                font-size: 12px;
            }
            QTableWidget::item {
                padding: 5px;
            }
        """)
        products_inner_layout.addWidget(self.top_products_table)
        
        category_layout.addWidget(products_frame)
        
        self.analysis_tabs.addTab(self.category_tab, "商品类别")
        
        # 地区销售分析选项卡
        self.region_tab = QWidget()
        region_layout = QVBoxLayout(self.region_tab)
        region_layout.setContentsMargins(10, 15, 10, 10)
        region_layout.setSpacing(15)
        
        # 添加标题标签
        region_title = QLabel("地区销售分析")
        region_title.setAlignment(Qt.AlignCenter)
        region_title.setStyleSheet(f"""
            font-size: 16px;
            font-weight: bold;
            color: {COLOR_PRIMARY};
            margin-bottom: 5px;
        """)
        region_layout.addWidget(region_title)
        
        # 地区销售对比图
        region_frame = QFrame()
        region_frame.setFrameShape(QFrame.StyledPanel)
        region_frame.setStyleSheet("background-color: white; border-radius: 5px; border: 1px solid #ddd;")
        region_inner_layout = QVBoxLayout(region_frame)
        
        self.region_canvas = MatplotlibCanvas(self.region_tab)
        region_inner_layout.addWidget(self.region_canvas)
        
        region_layout.addWidget(region_frame)
        
        self.analysis_tabs.addTab(self.region_tab, "地区销售")
        
        # 客户分析选项卡
        self.customer_tab = QWidget()
        customer_layout = QVBoxLayout(self.customer_tab)
        customer_layout.setContentsMargins(10, 15, 10, 10)
        customer_layout.setSpacing(15)
        
        # 添加标题标签
        customer_title = QLabel("客户群体分析")
        customer_title.setAlignment(Qt.AlignCenter)
        customer_title.setStyleSheet(f"""
            font-size: 16px;
            font-weight: bold;
            color: {COLOR_PRIMARY};
            margin-bottom: 5px;
        """)
        customer_layout.addWidget(customer_title)
        
        # 客户群体分析表格
        customer_frame = QFrame()
        customer_frame.setFrameShape(QFrame.StyledPanel)
        customer_frame.setStyleSheet("background-color: white; border-radius: 5px; border: 1px solid #ddd;")
        customer_inner_layout = QVBoxLayout(customer_frame)
        
        self.customer_table = QTableWidget()
        self.customer_table.setColumnCount(3)
        self.customer_table.setHorizontalHeaderLabels(["客户群体", "客户数量", "平均消费额"])
        self.customer_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.customer_table.setAlternatingRowColors(True)
        self.customer_table.setStyleSheet("""
            QTableWidget {
                font-size: 12px;
            }
            QTableWidget::item {
                padding: 5px;
            }
        """)
        customer_inner_layout.addWidget(self.customer_table)
        
        customer_layout.addWidget(customer_frame)
        
        self.analysis_tabs.addTab(self.customer_tab, "客户分析")
        
        # 促销效果分析选项卡
        self.promotion_tab = QWidget()
        promotion_layout = QVBoxLayout(self.promotion_tab)
        promotion_layout.setContentsMargins(10, 15, 10, 10)
        promotion_layout.setSpacing(15)
        
        # 添加标题标签
        promotion_title = QLabel("促销效果分析")
        promotion_title.setAlignment(Qt.AlignCenter)
        promotion_title.setStyleSheet(f"""
            font-size: 16px;
            font-weight: bold;
            color: {COLOR_PRIMARY};
            margin-bottom: 5px;
        """)
        promotion_layout.addWidget(promotion_title)
        
        # 购物节影响分析图
        festival_frame = QFrame()
        festival_frame.setFrameShape(QFrame.StyledPanel)
        festival_frame.setStyleSheet("background-color: white; border-radius: 5px; border: 1px solid #ddd;")
        festival_inner_layout = QVBoxLayout(festival_frame)
        
        self.festival_canvas = MatplotlibCanvas(self.promotion_tab)
        festival_inner_layout.addWidget(self.festival_canvas)
        
        promotion_layout.addWidget(festival_frame)
        
        self.analysis_tabs.addTab(self.promotion_tab, "促销效果")
        
        layout.addWidget(self.analysis_tabs)
    
    def init_trend_prediction_tab(self):
        """初始化趋势预测选项卡"""
        layout = QVBoxLayout(self.trend_prediction_tab)
        layout.setContentsMargins(15, 15, 15, 15)
        layout.setSpacing(15)
        
        # 创建子选项卡
        self.prediction_tabs = QTabWidget()
        self.prediction_tabs.setStyleSheet(f"""
            QTabWidget::pane {{
                border: 1px solid #bdc3c7;
                border-radius: 5px;
                padding: 5px;
                background-color: white;
            }}
            
            QTabBar::tab {{
                background-color: #ecf0f1;
                color: {COLOR_TEXT};
                min-width: 120px;
                padding: 8px 15px;
                margin-right: 2px;
                border-top-left-radius: 4px;
                border-top-right-radius: 4px;
            }}
            
            QTabBar::tab:selected {{
                background-color: {COLOR_PRIMARY};
                color: white;
            }}
            
            QTabBar::tab:hover:!selected {{
                background-color: #d0d9e0;
            }}
        """)
        
        # 销售预测选项卡
        self.forecast_tab = QWidget()
        forecast_layout = QVBoxLayout(self.forecast_tab)
        forecast_layout.setContentsMargins(10, 15, 10, 10)
        forecast_layout.setSpacing(15)
        
        # 添加标题标签
        forecast_title = QLabel("销售预测分析")
        forecast_title.setAlignment(Qt.AlignCenter)
        forecast_title.setStyleSheet(f"""
            font-size: 16px;
            font-weight: bold;
            color: {COLOR_PRIMARY};
            margin-bottom: 5px;
        """)
        forecast_layout.addWidget(forecast_title)
        
        # 预测控制区域
        forecast_control_group = QGroupBox("预测控制")
        forecast_control_group.setStyleSheet(f"""
            QGroupBox {{
                font-size: 14px;
                font-weight: bold;
                border: 2px solid {COLOR_PRIMARY};
                border-radius: 6px;
                margin-top: 12px;
                padding-top: 15px;
            }}
            QGroupBox::title {{
                subcontrol-origin: margin;
                subcontrol-position: top left;
                padding: 0 10px;
                color: {COLOR_PRIMARY};
                background-color: {COLOR_BACKGROUND};
            }}
        """)
        forecast_control_layout = QHBoxLayout()
        forecast_control_layout.setContentsMargins(10, 15, 10, 10)
        forecast_control_layout.setSpacing(15)
        
        # 预测时间单位选择
        forecast_time_label = QLabel("时间单位:")
        forecast_time_label.setStyleSheet("font-weight: bold;")
        self.forecast_time_combo = QComboBox()
        self.forecast_time_combo.addItems(["日", "月", "季度"])
        self.forecast_time_combo.setCurrentIndex(1)  # 默认选择"月"
        self.forecast_time_combo.setMinimumWidth(100)
        
        # 预测方法选择
        forecast_method_label = QLabel("预测方法:")
        forecast_method_label.setStyleSheet("font-weight: bold;")
        self.forecast_method_combo = QComboBox()
        self.forecast_method_combo.addItems(["指数平滑", "线性回归"])
        self.forecast_method_combo.setMinimumWidth(100)
        
        # 尝试导入Prophet，如果可用则添加到预测方法中
        try:
            from prophet import Prophet
            self.forecast_method_combo.addItems(["Prophet"])
        except ImportError:
            pass
        
        # 预测周期设置
        forecast_periods_label = QLabel("预测周期数:")
        forecast_periods_label.setStyleSheet("font-weight: bold;")
        self.forecast_periods_spin = QSpinBox()
        self.forecast_periods_spin.setRange(1, 24)
        self.forecast_periods_spin.setValue(6)
        self.forecast_periods_spin.setMinimumWidth(70)
        self.forecast_periods_spin.setStyleSheet("""
            QSpinBox {
                padding: 4px;
                font-size: 13px;
            }
        """)
        
        # 预测按钮
        self.forecast_btn = QPushButton("生成预测")
        self.forecast_btn.setIcon(QIcon.fromTheme("system-run"))
        self.forecast_btn.setStyleSheet(f"""
            QPushButton {{
                background-color: {COLOR_SECONDARY};
                min-width: 120px;
                min-height: 35px;
                font-size: 13px;
            }}
        """)
        self.forecast_btn.clicked.connect(self.update_sales_forecast)
        
        # 添加控件到布局
        forecast_control_layout.addWidget(forecast_time_label)
        forecast_control_layout.addWidget(self.forecast_time_combo)
        forecast_control_layout.addWidget(forecast_method_label)
        forecast_control_layout.addWidget(self.forecast_method_combo)
        forecast_control_layout.addWidget(forecast_periods_label)
        forecast_control_layout.addWidget(self.forecast_periods_spin)
        forecast_control_layout.addWidget(self.forecast_btn)
        forecast_control_layout.addStretch()
        
        forecast_control_group.setLayout(forecast_control_layout)
        forecast_layout.addWidget(forecast_control_group)
        
        # 预测图表
        forecast_chart_frame = QFrame()
        forecast_chart_frame.setFrameShape(QFrame.StyledPanel)
        forecast_chart_frame.setStyleSheet("background-color: white; border-radius: 5px; border: 1px solid #ddd;")
        forecast_chart_layout = QVBoxLayout(forecast_chart_frame)
        
        forecast_chart_title = QLabel("销售预测趋势图")
        forecast_chart_title.setAlignment(Qt.AlignCenter)
        forecast_chart_title.setStyleSheet(f"""
            font-size: 14px;
            font-weight: bold;
            color: {COLOR_TEXT};
        """)
        forecast_chart_layout.addWidget(forecast_chart_title)
        
        self.forecast_canvas = MatplotlibCanvas(self.forecast_tab)
        forecast_chart_layout.addWidget(self.forecast_canvas)
        
        forecast_layout.addWidget(forecast_chart_frame)
        
        # 预测数据表格
        forecast_table_frame = QFrame()
        forecast_table_frame.setFrameShape(QFrame.StyledPanel)
        forecast_table_frame.setStyleSheet("background-color: white; border-radius: 5px; border: 1px solid #ddd;")
        forecast_table_layout = QVBoxLayout(forecast_table_frame)
        
        forecast_table_title = QLabel("销售预测数据明细")
        forecast_table_title.setAlignment(Qt.AlignCenter)
        forecast_table_title.setStyleSheet(f"""
            font-size: 14px;
            font-weight: bold;
            color: {COLOR_TEXT};
        """)
        forecast_table_layout.addWidget(forecast_table_title)
        
        self.forecast_table = QTableWidget()
        self.forecast_table.setColumnCount(4)
        self.forecast_table.setHorizontalHeaderLabels(["日期", "实际销售额", "预测销售额", "数据类型"])
        self.forecast_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.forecast_table.setAlternatingRowColors(True)
        self.forecast_table.setStyleSheet("""
            QTableWidget {
                font-size: 12px;
            }
            QTableWidget::item {
                padding: 5px;
            }
            QTableWidget::item[data-type="预测"] {
                background-color: rgba(231, 76, 60, 0.1);
            }
        """)
        forecast_table_layout.addWidget(self.forecast_table)
        
        forecast_layout.addWidget(forecast_table_frame)
        
        self.prediction_tabs.addTab(self.forecast_tab, "销售预测")
        
        # 决策支持选项卡
        self.decision_tab = QWidget()
        decision_layout = QVBoxLayout(self.decision_tab)
        decision_layout.setContentsMargins(10, 15, 10, 10)
        decision_layout.setSpacing(15)
        
        # 添加标题标签
        decision_title = QLabel("智能决策支持")
        decision_title.setAlignment(Qt.AlignCenter)
        decision_title.setStyleSheet(f"""
            font-size: 16px;
            font-weight: bold;
            color: {COLOR_PRIMARY};
            margin-bottom: 5px;
        """)
        decision_layout.addWidget(decision_title)
        
        # 决策控制区域
        decision_control_group = QGroupBox("决策支持控制")
        decision_control_group.setStyleSheet(f"""
            QGroupBox {{
                font-size: 14px;
                font-weight: bold;
                border: 2px solid {COLOR_PRIMARY};
                border-radius: 6px;
                margin-top: 12px;
                padding-top: 15px;
            }}
            QGroupBox::title {{
                subcontrol-origin: margin;
                subcontrol-position: top left;
                padding: 0 10px;
                color: {COLOR_PRIMARY};
                background-color: {COLOR_BACKGROUND};
            }}
        """)
        decision_control_layout = QHBoxLayout()
        decision_control_layout.setContentsMargins(10, 15, 10, 10)
        decision_control_layout.setSpacing(15)
        
        # 商品类别选择
        decision_category_label = QLabel("商品类别:")
        decision_category_label.setStyleSheet("font-weight: bold;")
        self.decision_category_combo = QComboBox()
        self.decision_category_combo.addItems(["全部"])
        self.decision_category_combo.setMinimumWidth(150)
        
        # 生成建议按钮
        self.generate_suggestions_btn = QPushButton("生成决策建议")
        self.generate_suggestions_btn.setIcon(QIcon.fromTheme("dialog-information"))
        self.generate_suggestions_btn.setStyleSheet(f"""
            QPushButton {{
                background-color: {COLOR_SECONDARY};
                min-width: 140px;
                min-height: 35px;
                font-size: 13px;
            }}
        """)
        self.generate_suggestions_btn.clicked.connect(self.update_decision_suggestions)
        
        # 添加控件到布局
        decision_control_layout.addWidget(decision_category_label)
        decision_control_layout.addWidget(self.decision_category_combo)
        decision_control_layout.addWidget(self.generate_suggestions_btn)
        decision_control_layout.addStretch()
        
        decision_control_group.setLayout(decision_control_layout)
        decision_layout.addWidget(decision_control_group)
        
        # 决策建议文本区域
        decision_text_frame = QFrame()
        decision_text_frame.setFrameShape(QFrame.StyledPanel)
        decision_text_frame.setStyleSheet("background-color: white; border-radius: 5px; border: 1px solid #ddd;")
        decision_text_layout = QVBoxLayout(decision_text_frame)
        
        decision_text_title = QLabel("决策建议")
        decision_text_title.setAlignment(Qt.AlignCenter)
        decision_text_title.setStyleSheet(f"""
            font-size: 14px;
            font-weight: bold;
            color: {COLOR_TEXT};
        """)
        decision_text_layout.addWidget(decision_text_title)
        
        self.suggestions_text = QTextEdit()
        self.suggestions_text.setReadOnly(True)
        self.suggestions_text.setStyleSheet(f"""
            QTextEdit {{
                font-size: 13px;
                line-height: 1.5;
                padding: 10px;
                border: none;
                background-color: {COLOR_LIGHT_BG};
                border-radius: 5px;
            }}
        """)
        decision_text_layout.addWidget(self.suggestions_text)
        
        decision_layout.addWidget(decision_text_frame)
        
        self.prediction_tabs.addTab(self.decision_tab, "决策支持")
        
        layout.addWidget(self.prediction_tabs)
    
    def init_report_generation_tab(self):
        """初始化报告生成选项卡"""
        layout = QVBoxLayout(self.report_generation_tab)
        layout.setContentsMargins(15, 15, 15, 15)
        layout.setSpacing(15)
        
        # 添加标题标签
        report_title = QLabel("报告生成中心")
        report_title.setAlignment(Qt.AlignCenter)
        report_title.setStyleSheet(f"""
            font-size: 16px;
            font-weight: bold;
            color: {COLOR_PRIMARY};
            margin-bottom: 5px;
        """)
        layout.addWidget(report_title)
        
        # 创建左右分栏布局
        splitter = QSplitter(Qt.Horizontal)
        splitter.setStyleSheet(f"""
            QSplitter::handle {{
                background-color: {COLOR_LIGHT_BG};
                width: 2px;
            }}
        """)
        
        # 左侧设置区域
        left_widget = QWidget()
        left_layout = QVBoxLayout(left_widget)
        left_layout.setContentsMargins(5, 5, 5, 5)
        left_layout.setSpacing(15)
        
        # 报告设置区域
        report_settings_group = QGroupBox("报告设置")
        report_settings_group.setStyleSheet(f"""
            QGroupBox {{
                font-size: 14px;
                font-weight: bold;
                border: 2px solid {COLOR_PRIMARY};
                border-radius: 6px;
                margin-top: 12px;
                padding-top: 15px;
            }}
            QGroupBox::title {{
                subcontrol-origin: margin;
                subcontrol-position: top left;
                padding: 0 10px;
                color: {COLOR_PRIMARY};
                background-color: {COLOR_BACKGROUND};
            }}
        """)
        settings_layout = QGridLayout()
        settings_layout.setContentsMargins(10, 15, 10, 10)
        settings_layout.setSpacing(10)
        
        # 报告标题
        title_label = QLabel("报告标题:")
        title_label.setStyleSheet("font-weight: bold;")
        self.report_title_edit = QTextEdit()
        self.report_title_edit.setMaximumHeight(60)
        self.report_title_edit.setText("电商平台销售分析报告")
        self.report_title_edit.setStyleSheet(f"""
            QTextEdit {{
                font-size: 13px;
                border: 1px solid #bdc3c7;
                border-radius: 4px;
                padding: 8px;
                background-color: white;
            }}
        """)
        settings_layout.addWidget(title_label, 0, 0)
        settings_layout.addWidget(self.report_title_edit, 0, 1)
        
        # 报告时间范围
        date_label = QLabel("报告时间范围:")
        date_label.setStyleSheet("font-weight: bold;")
        date_layout = QHBoxLayout()
        
        start_date_label = QLabel("开始日期:")
        self.start_date_edit = QDateEdit()
        self.start_date_edit.setCalendarPopup(True)
        self.start_date_edit.setDate(QDate.currentDate().addMonths(-6))
        self.start_date_edit.setStyleSheet("""
            QDateEdit {
                padding: 5px;
                border: 1px solid #bdc3c7;
                border-radius: 4px;
                background-color: white;
                min-height: 25px;
                min-width: 100px;
            }
        """)
        
        end_date_label = QLabel("结束日期:")
        self.end_date_edit = QDateEdit()
        self.end_date_edit.setCalendarPopup(True)
        self.end_date_edit.setDate(QDate.currentDate())
        self.end_date_edit.setStyleSheet("""
            QDateEdit {
                padding: 5px;
                border: 1px solid #bdc3c7;
                border-radius: 4px;
                background-color: white;
                min-height: 25px;
                min-width: 100px;
            }
        """)
        
        date_layout.addWidget(start_date_label)
        date_layout.addWidget(self.start_date_edit)
        date_layout.addWidget(end_date_label)
        date_layout.addWidget(self.end_date_edit)
        date_layout.addStretch()
        
        settings_layout.addWidget(date_label, 1, 0)
        settings_layout.addLayout(date_layout, 1, 1)
        
        # 报告内容选择
        content_label = QLabel("报告内容:")
        content_label.setStyleSheet("font-weight: bold;")
        
        # 创建一个漂亮的内容选择框架
        content_frame = QFrame()
        content_frame.setFrameShape(QFrame.StyledPanel)
        content_frame.setStyleSheet(f"""
            QFrame {{
                background-color: white;
                border-radius: 5px;
                border: 1px solid #ddd;
            }}
            QCheckBox {{
                font-size: 13px;
                padding: 5px;
            }}
            QCheckBox::indicator {{
                width: 18px;
                height: 18px;
                border: 1px solid #bdc3c7;
                border-radius: 3px;
            }}
            QCheckBox::indicator:checked {{
                background-color: {COLOR_PRIMARY};
                border: 1px solid {COLOR_PRIMARY};
                image: url(:/qt-project.org/styles/commonstyle/images/check-16.png);
            }}
        """)
        content_layout = QGridLayout(content_frame)
        content_layout.setContentsMargins(15, 10, 15, 10)
        content_layout.setSpacing(10)
        
        self.include_trend_check = QCheckBox("销售趋势分析")
        self.include_trend_check.setChecked(True)
        content_layout.addWidget(self.include_trend_check, 0, 0)
        
        self.include_category_check = QCheckBox("商品类别分析")
        self.include_category_check.setChecked(True)
        content_layout.addWidget(self.include_category_check, 0, 1)
        
        self.include_region_check = QCheckBox("地区销售分析")
        self.include_region_check.setChecked(True)
        content_layout.addWidget(self.include_region_check, 1, 0)
        
        self.include_customer_check = QCheckBox("客户分析")
        self.include_customer_check.setChecked(True)
        content_layout.addWidget(self.include_customer_check, 1, 1)
        
        self.include_promotion_check = QCheckBox("促销效果分析")
        self.include_promotion_check.setChecked(True)
        content_layout.addWidget(self.include_promotion_check, 2, 0)
        
        self.include_forecast_check = QCheckBox("销售预测")
        self.include_forecast_check.setChecked(True)
        content_layout.addWidget(self.include_forecast_check, 2, 1)
        
        self.include_decision_check = QCheckBox("决策建议")
        self.include_decision_check.setChecked(True)
        content_layout.addWidget(self.include_decision_check, 3, 0)
        
        settings_layout.addWidget(content_label, 2, 0)
        settings_layout.addWidget(content_frame, 2, 1)
        
        # 报告生成按钮
        self.generate_report_btn = QPushButton("生成报告")
        self.generate_report_btn.clicked.connect(self.generate_report)
        settings_layout.addWidget(self.generate_report_btn, 3, 0, 1, 2)
        
        report_settings_group.setLayout(settings_layout)
        left_layout.addWidget(report_settings_group)
        
        # 保存报告按钮
        save_layout = QHBoxLayout()
        save_layout.addStretch()
        
        self.save_report_btn = QPushButton("保存报告")
        self.save_report_btn.clicked.connect(self.save_report)
        self.save_report_btn.setMinimumWidth(120)
        save_layout.addWidget(self.save_report_btn)
        
        left_layout.addLayout(save_layout)
        left_layout.addStretch()
        
        # 右侧预览区域
        right_widget = QWidget()
        right_layout = QVBoxLayout(right_widget)
        
        # 报告预览标题
        preview_label = QLabel("报告预览")
        preview_label.setAlignment(Qt.AlignCenter)
        preview_label.setStyleSheet("font-size: 14px; font-weight: bold; margin-bottom: 10px;")
        right_layout.addWidget(preview_label)
        
        # 使用QWebEngineView进行HTML预览
        self.report_preview = QWebEngineView()
        self.report_preview.setMinimumWidth(600)  # 设置最小宽度
        self.report_preview.setMinimumHeight(700)  # 设置最小高度
        
        # 添加一些默认的HTML内容
        default_html = """
        <html>
        <body style="font-family: Arial, sans-serif; text-align: center; padding: 50px;">
            <h2>欢迎使用基于大数据的电商平台商品销售趋势分析与决策软件</h2>
            <p>点击"生成报告"按钮查看完整报告预览</p>
            <p style="color: #666;">报告将在此处显示</p>
        </body>
        </html>
        """
        self.report_preview.setHtml(default_html)
        
        right_layout.addWidget(self.report_preview)
        
        # 添加左右部件到分割器
        splitter.addWidget(left_widget)
        splitter.addWidget(right_widget)
        
        # 设置初始分割比例
        splitter.setSizes([300, 700])
        
        layout.addWidget(splitter)
    
    def update_data_overview(self):
        """更新数据概览信息"""
        if self.analyzer.df is None:
            return
            
        # 获取数据摘要
        summary = self.analyzer.get_data_summary()
        
        if summary:
            # 更新数据信息表格
            self.data_info_table.setRowCount(len(summary))
            row = 0
            for key, value in summary.items():
                self.data_info_table.setItem(row, 0, QTableWidgetItem(key))
                
                # 如果值是列表，将其转换为字符串
                if isinstance(value, list):
                    value_str = ", ".join(map(str, value))
                else:
                    value_str = str(value)
                    
                self.data_info_table.setItem(row, 1, QTableWidgetItem(value_str))
                row += 1
            
            # 更新筛选控件的选项
            self.update_filter_options()
            
            # 更新完整数据视图
            self.update_full_data_view()
            
    def update_filter_options(self):
        """更新筛选控件的选项"""
        if self.analyzer.df is None:
            return
            
        # 获取数据中的年份选项
        if '年' in self.analyzer.df.columns:
            years = sorted(self.analyzer.df['年'].unique())
        elif '日期' in self.analyzer.df.columns:
            years = sorted(self.analyzer.df['日期'].dt.year.unique())
        else:
            years = []
            
        # 更新年份下拉框
        current_year = self.year_combo.currentText()
        self.year_combo.clear()
        self.year_combo.addItem("全部")
        self.year_combo.addItems([str(int(year)) for year in years])
        
        # 尝试恢复之前的选择
        index = self.year_combo.findText(current_year)
        if index >= 0:
            self.year_combo.setCurrentIndex(index)
            
        # 获取数据中的商品类别选项
        if '商品类别' in self.analyzer.df.columns:
            categories = sorted(self.analyzer.df['商品类别'].unique())
        else:
            categories = []
            
        # 更新商品类别下拉框
        current_category = self.data_category_combo.currentText()
        self.data_category_combo.clear()
        self.data_category_combo.addItem("全部")
        self.data_category_combo.addItems(categories)
        
        # 尝试恢复之前的选择
        index = self.data_category_combo.findText(current_category)
        if index >= 0:
            self.data_category_combo.setCurrentIndex(index)
            
    def update_full_data_view(self):
        """更新完整数据视图"""
        if self.analyzer.df is None:
            return
            
        # 应用筛选条件
        filtered_df = self.get_filtered_data()
        
        # 更新统计信息
        self.filtered_count_label.setText(f"已筛选记录数: {len(filtered_df)}")
        self.total_count_label.setText(f"总记录数: {len(self.analyzer.df)}")
        
        # 更新完整数据表格
        self.full_data_table.setRowCount(len(filtered_df))
        self.full_data_table.setColumnCount(len(filtered_df.columns))
        self.full_data_table.setHorizontalHeaderLabels(filtered_df.columns)
        
        # 设置数据
        for i in range(len(filtered_df)):
            for j in range(len(filtered_df.columns)):
                value = filtered_df.iloc[i, j]
                
                # 格式化日期列
                if filtered_df.columns[j] == '日期':
                    if pd.notnull(value):
                        value = pd.to_datetime(value).strftime('%Y-%m-%d')
                
                self.full_data_table.setItem(i, j, QTableWidgetItem(str(value)))
        
    def get_filtered_data(self):
        """根据筛选条件获取过滤后的数据"""
        if self.analyzer.df is None:
            return pd.DataFrame()
            
        filtered_df = self.analyzer.df.copy()
        
        # 应用年份筛选
        year = self.year_combo.currentText()
        if year != "全部" and year.strip():  # 确保年份不是空字符串
            try:
                year_value = int(year)
                if '年' in filtered_df.columns:
                    filtered_df = filtered_df[filtered_df['年'] == year_value]
                elif '日期' in filtered_df.columns:
                    filtered_df = filtered_df[filtered_df['日期'].dt.year == year_value]
            except (ValueError, TypeError):
                # 忽略无效的年份值
                pass
                
        # 应用季度筛选
        quarter = self.quarter_combo.currentText()
        if quarter != "全部" and quarter.strip():  # 确保季度不是空字符串
            try:
                quarter_value = int(quarter)
                if '季度' in filtered_df.columns:
                    filtered_df = filtered_df[filtered_df['季度'] == quarter_value]
                elif '日期' in filtered_df.columns:
                    filtered_df = filtered_df[filtered_df['日期'].dt.quarter == quarter_value]
            except (ValueError, TypeError):
                # 忽略无效的季度值
                pass
                
        # 应用月份筛选
        month = self.month_combo.currentText()
        if month != "全部" and month.strip():  # 确保月份不是空字符串
            try:
                month_value = int(month)
                if '月' in filtered_df.columns:
                    filtered_df = filtered_df[filtered_df['月'] == month_value]
                elif '日期' in filtered_df.columns:
                    filtered_df = filtered_df[filtered_df['日期'].dt.month == month_value]
            except (ValueError, TypeError):
                # 忽略无效的月份值
                pass
                
        # 应用商品类别筛选
        category = self.data_category_combo.currentText()
        if category != "全部" and '商品类别' in filtered_df.columns:
            filtered_df = filtered_df[filtered_df['商品类别'] == category]
            
        return filtered_df
        
    def reset_data_filters(self):
        """重置所有筛选条件"""
        self.year_combo.setCurrentText("全部")
        self.quarter_combo.setCurrentText("全部")
        self.month_combo.setCurrentText("全部")
        self.data_category_combo.setCurrentText("全部")
        self.update_full_data_view()
        
    def load_data(self):
        """加载数据文件"""
        # 获取当前文件所在目录
        current_dir = os.path.dirname(os.path.abspath(__file__))
        # 构建data文件夹的路径
        data_dir = os.path.join(current_dir, 'data')
        
        # 如果data文件夹不存在，创建它
        if not os.path.exists(data_dir):
            os.makedirs(data_dir)
            QMessageBox.warning(self, "提示", "data文件夹不存在，已自动创建。请将数据文件放入data文件夹中。")
            return
            
        file_path, _ = QFileDialog.getOpenFileName(
            self, 
            "选择数据文件", 
            data_dir,  # 设置初始目录为data文件夹
            "CSV文件 (*.csv);;Excel文件 (*.xlsx *.xls)"
        )
        
        if file_path:
            try:
                # 创建进度对话框
                progress = QProgressDialog("正在加载数据...", "取消", 0, 100, self)
                progress.setWindowTitle("数据加载")
                progress.setWindowModality(Qt.WindowModal)
                progress.setMinimumDuration(0)  # 立即显示
                progress.setValue(0)
                
                # 设置进度条的样式
                progress.setStyleSheet(f"""
                    QProgressDialog {{
                        background-color: {COLOR_LIGHT_BG};
                        border-radius: 8px;
                        border: 1px solid #ddd;
                        min-width: 400px;
                        min-height: 120px;
                    }}
                    QProgressBar {{
                        border: 1px solid #bdc3c7;
                        border-radius: 5px;
                        text-align: center;
                        color: white;
                        background-color: {COLOR_LIGHT_BG};
                        height: 20px;
                    }}
                    QProgressBar::chunk {{
                        background-color: qlineargradient(spread:pad, x1:0, y1:0, x2:1, y2:0, 
                                            stop:0 {COLOR_PRIMARY}, stop:1 {COLOR_SECONDARY});
                        width: 10px;
                        margin: 0.5px;
                        border-radius: 3px;
                    }}
                    QLabel {{
                        font-size: 14px;
                        color: {COLOR_TEXT};
                        margin-bottom: 5px;
                    }}
                    QPushButton {{
                        background-color: {COLOR_PRIMARY};
                        color: white;
                        border: none;
                        border-radius: 4px;
                        padding: 6px 12px;
                        min-width: 80px;
                    }}
                """)
                
                progress.show()
                
                # 创建并启动工作线程
                self.worker = DataLoadWorker(self.analyzer, file_path)
                
                # 连接信号
                self.worker.signals.progress.connect(
                    lambda value, message: self._update_progress(progress, value, message))
                self.worker.signals.error.connect(
                    lambda error_msg: self._handle_load_error(progress, error_msg))
                self.worker.signals.finished.connect(
                    lambda: self._finalize_data_load(progress, file_path))
                self.worker.signals.result.connect(
                    lambda df: self._process_loaded_data(df, file_path))
                
                # 启动线程
                self.worker.start()
                
            except Exception as e:
                QMessageBox.critical(self, "错误", f"加载数据失败: {str(e)}")
    
    def _update_progress(self, progress_dialog, value, message):
        """更新进度对话框"""
        if progress_dialog and progress_dialog.isVisible():
            progress_dialog.setValue(value)
            progress_dialog.setLabelText(message)
    
    def _handle_load_error(self, progress_dialog, error_msg):
        """处理加载错误"""
        if progress_dialog and progress_dialog.isVisible():
            progress_dialog.close()
        QMessageBox.critical(self, "错误", f"加载数据失败: {error_msg}")
    
    def _finalize_data_load(self, progress_dialog, file_path):
        """完成数据加载"""
        if progress_dialog and progress_dialog.isVisible():
            progress_dialog.setValue(100)
            progress_dialog.close()
        
        file_size = os.path.getsize(file_path) / (1024 * 1024)  # 转换为MB
        self.statusBar().showMessage(f'已成功加载数据: {os.path.basename(file_path)} ({file_size:.1f} MB)')
        
        # 显示数据文件名
        self.data_path_label.setText(os.path.basename(file_path))
        
        # 切换到数据概览选项卡
        self.main_tabs.setCurrentIndex(0)
    
    def _process_loaded_data(self, df, file_path):
        """处理已加载的数据，更新UI组件"""
        try:
            # 更新报告生成选项卡中的日期范围
            if '日期' in self.analyzer.df.columns:
                min_date = self.analyzer.df['日期'].min()
                max_date = self.analyzer.df['日期'].max()
                
                # 更新报告日期范围
                self.start_date_edit.setDate(QDate(min_date.year, min_date.month, min_date.day))
                self.end_date_edit.setDate(QDate(max_date.year, max_date.month, max_date.day))
            
            # 更新商品类别下拉框
            if self.analyzer.df is not None:
                categories = ["全部"] + sorted(self.analyzer.df['商品类别'].unique().tolist())
                
                # 更新数据分析选项卡中的类别下拉框
                self.category_combo.clear()
                self.category_combo.addItems(categories)
                
                # 更新决策支持选项卡中的类别下拉框
                self.decision_category_combo.clear()
                self.decision_category_combo.addItems(categories)
            
            # 更新数据概览
            self.update_data_overview()
            
            # 更新分析结果
            self.update_analysis()
        
        except Exception as e:
            QMessageBox.warning(self, "警告", f"更新界面时出错: {str(e)}")
    
    def update_analysis(self):
        """更新所有分析结果"""
        if self.analyzer.df is None:
            QMessageBox.warning(self, "警告", "请先加载数据")
            return
            
        try:
            # 获取当前选择
            time_unit = self.time_combo.currentText()
            category = self.category_combo.currentText()
            if category == "全部":
                category = None
            region_level = self.region_combo.currentText()
            
            # 更新各个分析模块
            try:
                self.update_trend_analysis(time_unit, category)
            except Exception as e:
                QMessageBox.warning(self, "警告", f"更新销售趋势分析时出错: {str(e)}")
                
            try:
                self.update_category_analysis(category)
            except Exception as e:
                QMessageBox.warning(self, "警告", f"更新商品类别分析时出错: {str(e)}")
                
            try:
                self.update_region_analysis(region_level)
            except Exception as e:
                QMessageBox.warning(self, "警告", f"更新地区销售分析时出错: {str(e)}")
                
            try:
                self.update_customer_analysis()
            except Exception as e:
                QMessageBox.warning(self, "警告", f"更新客户分析时出错: {str(e)}")
                
            try:
                self.update_promotion_analysis()
            except Exception as e:
                QMessageBox.warning(self, "警告", f"更新促销效果分析时出错: {str(e)}")
                
        except Exception as e:
            QMessageBox.critical(self, "错误", f"更新分析结果时出错: {str(e)}\n\n请检查数据格式是否正确。")
    
    def update_trend_analysis(self, time_unit, category):
        """更新销售趋势分析"""
        # 清除现有图表
        self.trend_canvas.clear()
        self.heatmap_canvas.clear()
        
        try:
            # 绘制销售趋势图
            sales_data = self.analyzer.get_sales_by_time(time_unit, category)
            if sales_data is not None and len(sales_data) > 0:
                # 直接使用Matplotlib API绘图
                # 设置更现代的颜色和风格
                self.trend_canvas.axes.plot(sales_data['时间'], sales_data['总价'], marker='o', linewidth=2.5,
                                          color=COLOR_PRIMARY, markerfacecolor='white',
                                          markeredgecolor=COLOR_PRIMARY, markersize=8)
                
                # 添加区域阴影
                self.trend_canvas.axes.fill_between(sales_data['时间'], 0, sales_data['总价'], 
                                                  alpha=0.1, color=COLOR_PRIMARY)
                
                # 设置标题和标签
                self.trend_canvas.axes.set_title(f'按{time_unit}销售趋势', fontsize=14, fontweight='bold')
                self.trend_canvas.axes.set_xlabel('时间', fontsize=12)
                self.trend_canvas.axes.set_ylabel('销售额（元）', fontsize=12)
                self.trend_canvas.axes.grid(True, linestyle='--', alpha=0.7, color='#ddd')
                
                # 美化背景
                self.trend_canvas.axes.set_facecolor(COLOR_LIGHT_BG)
                
                # 美化坐标轴
                self.trend_canvas.axes.spines['top'].set_visible(False)
                self.trend_canvas.axes.spines['right'].set_visible(False)
                self.trend_canvas.axes.spines['left'].set_color('#ddd')
                self.trend_canvas.axes.spines['bottom'].set_color('#ddd')
                
                # 优化x轴标签显示
                if len(sales_data) > 12:
                    # 只显示部分标签
                    step = max(1, len(sales_data) // 12)  # 最多显示12个标签
                    labels = self.trend_canvas.axes.get_xticklabels()
                    for i, label in enumerate(labels):
                        if i % step != 0:
                            label.set_visible(False)
                
                # 旋转标签使其不重叠
                plt.setp(self.trend_canvas.axes.get_xticklabels(), rotation=45, ha='right')
                
                # 添加数值标签
                for i, (x, y) in enumerate(zip(sales_data['时间'], sales_data['总价'])):
                    if i % step == 0 or i == len(sales_data) - 1:  # 只对部分点添加标签
                        self.trend_canvas.axes.annotate(f'{y:,.0f}', 
                                                     xy=(x, y), 
                                                     xytext=(0, 10),
                                                     textcoords='offset points',
                                                     ha='center',
                                                     va='bottom',
                                                     fontsize=9,
                                                     color=COLOR_TEXT)
                
                self.trend_canvas.fig.tight_layout()
                self.trend_canvas.draw()
            
            # 绘制热图
            try:
                # 完全重新创建图表对象，确保没有残留
                self.heatmap_canvas.fig.clear()
                self.heatmap_canvas.axes = self.heatmap_canvas.fig.add_subplot(111)
                
                if self.analyzer.df is not None:
                    data = self.analyzer.df.copy()
                    if category:
                        data = data[data['商品类别'] == category]
                        
                    # 确保数据包含月份信息
                    if '月' not in data.columns and '日期' in data.columns:
                        data['月'] = data['日期'].dt.month
                        
                    # 创建月份-类别交叉表
                    pivot = pd.pivot_table(
                        data, 
                        values='总价', 
                        index='商品类别', 
                        columns='月',
                        aggfunc='sum'
                    ).fillna(0)
                    
                    if not pivot.empty:
                        # 使用更好看的颜色映射
                        cmap = plt.cm.get_cmap('viridis')
                        im = self.heatmap_canvas.axes.imshow(pivot.values, cmap=cmap, aspect='auto')
                        
                        # 添加标题
                        self.heatmap_canvas.axes.set_title('月份-商品类别销售热图', fontsize=14, fontweight='bold')
                        
                        # 设置坐标轴
                        self.heatmap_canvas.axes.set_yticks(range(len(pivot.index)))
                        self.heatmap_canvas.axes.set_yticklabels(pivot.index)
                        self.heatmap_canvas.axes.set_xticks(range(len(pivot.columns)))
                        self.heatmap_canvas.axes.set_xticklabels(pivot.columns)
                        
                        # 添加数值标签
                        for i in range(len(pivot.index)):
                            for j in range(len(pivot.columns)):
                                value = pivot.values[i, j]
                                text_color = 'white' if value > pivot.values.max() / 2 else 'black'
                                self.heatmap_canvas.axes.text(j, i, f'{value:,.0f}', 
                                                       ha='center', va='center', 
                                                       color=text_color, fontsize=8)
                        
                        # 添加新的colorbar
                        cbar = self.heatmap_canvas.fig.colorbar(im)
                        cbar.set_label('销售额（元）', fontsize=12)
                        
                        self.heatmap_canvas.fig.tight_layout()
                        self.heatmap_canvas.draw()
            except Exception as e:
                print(f"热图生成错误: {str(e)}")
                # 热图生成失败不影响整体功能
                pass
                
        except Exception as e:
            raise Exception(f"绘制销售趋势图出错: {str(e)}")
    
    def update_category_analysis(self, category):
        """更新商品类别分析"""
        # 清除现有图表
        self.category_canvas.clear()
        
        try:
            # 直接在这里实现类别销售额对比图，而不是使用analyzer的方法
            if self.analyzer.df is not None:
                data = self.analyzer.df.copy()
                if category:
                    data = data[data['商品类别'] == category]
                
                # 按商品类别分组计算销售额
                category_sales = data.groupby('商品类别')['总价'].sum().reset_index()
                category_sales = category_sales.sort_values('总价', ascending=False)
                
                if len(category_sales) > 0:
                    # 绘制横向条形图
                    bars = self.category_canvas.axes.barh(category_sales['商品类别'], category_sales['总价'])
                    self.category_canvas.axes.set_title('各商品类别销售额')
                    self.category_canvas.axes.set_xlabel('销售额（元）')
                    self.category_canvas.axes.set_ylabel('商品类别')
                    self.category_canvas.axes.grid(True, linestyle='--', alpha=0.7, axis='x')
                    
                    # 添加数据标签
                    for bar in bars:
                        width = bar.get_width()
                        self.category_canvas.axes.text(width + width*0.02, bar.get_y() + bar.get_height()/2, 
                                f'{width:,.0f}', ha='left', va='center')
                    
                    # 限制显示的类别数量以避免图表过拥挤
                    if len(category_sales) > 15:
                        self.category_canvas.axes.set_yticks(range(15))
                        self.category_canvas.axes.set_yticklabels(category_sales['商品类别'].iloc[:15])
                    
                    self.category_canvas.fig.tight_layout()
                    self.category_canvas.draw()
            
            # 更新热销商品表格
            top_products = self.analyzer.get_top_products(n=10, measure='销售额', category=category)
            if top_products is not None and len(top_products) > 0:
                self.top_products_table.setRowCount(len(top_products))
                for i, (_, row) in enumerate(top_products.iterrows()):
                    self.top_products_table.setItem(i, 0, QTableWidgetItem(row['商品名称']))
                    self.top_products_table.setItem(i, 1, QTableWidgetItem(f"{row['销售额']:,.2f}"))
        except Exception as e:
            raise Exception(f"更新商品类别分析出错: {str(e)}")
    
    def update_region_analysis(self, region_level):
        """更新地区销售分析"""
        # 清除现有图表
        self.region_canvas.clear()
        
        try:
            # 直接在这里实现地区销售对比图，而不是使用analyzer的方法
            if self.analyzer.df is not None:
                if region_level not in ['省份', '城市']:
                    raise ValueError(f"不支持的地区级别: {region_level}")
                
                # 检查数据是否包含指定的地区列
                if region_level not in self.analyzer.df.columns:
                    raise ValueError(f"数据中缺少{region_level}列")
                
                # 按地区分组计算销售额
                region_sales = self.analyzer.df.groupby(region_level)['总价'].sum().reset_index()
                region_sales = region_sales.sort_values('总价', ascending=False)
                
                if len(region_sales) > 0:
                    # 限制地区数量，只显示销售额前10名
                    top_regions = region_sales.head(10).copy()
                    
                    # 设置图表背景颜色
                    self.region_canvas.axes.set_facecolor(COLOR_LIGHT_BG)
                    
                    # 生成渐变颜色列表
                    cmap = plt.cm.get_cmap('Blues')
                    colors = [cmap(0.3 + 0.7 * i / len(top_regions)) for i in range(len(top_regions))]
                    
                    # 绘制横向条形图，使用自定义颜色
                    bars = self.region_canvas.axes.barh(top_regions[region_level], top_regions['总价'], 
                                                     color=colors, height=0.7, edgecolor='white', linewidth=1)
                    
                    # 设置标题和标签
                    self.region_canvas.axes.set_title(f'各{region_level}销售额对比（Top 10）', 
                                                   fontsize=14, fontweight='bold')
                    self.region_canvas.axes.set_xlabel('销售额（元）', fontsize=12)
                    self.region_canvas.axes.set_ylabel(region_level, fontsize=12)
                    
                    # 美化坐标轴
                    self.region_canvas.axes.spines['top'].set_visible(False)
                    self.region_canvas.axes.spines['right'].set_visible(False)
                    self.region_canvas.axes.spines['left'].set_color('#ddd')
                    self.region_canvas.axes.spines['bottom'].set_color('#ddd')
                    
                    # 添加网格线
                    self.region_canvas.axes.grid(True, linestyle='--', alpha=0.7, color='#ddd', axis='x')
                    
                    # 添加数据标签
                    for bar in bars:
                        width = bar.get_width()
                        self.region_canvas.axes.text(width + width*0.02, bar.get_y() + bar.get_height()/2, 
                                                  f'{width:,.0f}', ha='left', va='center', 
                                                  fontsize=10, fontweight='bold')
                    
                    # 增大字体大小，使标签更清晰可读
                    self.region_canvas.axes.tick_params(axis='y', labelsize=10)
                    
                    # 设置固定的图表大小，以便在报告中有足够空间显示
                    self.region_canvas.fig.set_size_inches(10, 8)
                    
                    # 添加顶部和底部的空白边距
                    self.region_canvas.axes.margins(y=0.1)
                    
                    # 调整布局，确保所有元素可见
                    self.region_canvas.fig.tight_layout()
                    self.region_canvas.draw()
        except Exception as e:
            raise Exception(f"更新地区销售分析出错: {str(e)}")
    
    def update_customer_analysis(self):
        """更新客户分析"""
        try:
            # 获取客户细分数据
            customer_data, cluster_features = self.analyzer.get_customer_segments()
            if customer_data is not None and cluster_features is not None:
                # 计算每个客户群体的统计信息
                segment_stats = customer_data.groupby('客户群体标签').agg({
                    '顾客ID': 'count',
                    '总消费额': 'mean'
                }).reset_index()
                
                # 更新客户群体分析表格
                self.customer_table.setRowCount(len(segment_stats))
                for i, (_, row) in enumerate(segment_stats.iterrows()):
                    self.customer_table.setItem(i, 0, QTableWidgetItem(row['客户群体标签']))
                    self.customer_table.setItem(i, 1, QTableWidgetItem(str(row['顾客ID'])))
                    self.customer_table.setItem(i, 2, QTableWidgetItem(f"{row['总消费额']:,.2f}"))
        except Exception as e:
            raise Exception(f"更新客户分析出错: {str(e)}")
    
    def update_promotion_analysis(self):
        """更新促销效果分析"""
        # 清除现有图表
        self.festival_canvas.clear()
        
        try:
            # 直接在这里实现购物节影响分析图，而不是使用analyzer的方法
            if self.analyzer.df is not None:
                # 检查数据中是否已有购物节列
                if '购物节' not in self.analyzer.df.columns:
                    # 创建购物节标记
                    data = self.analyzer.df.copy()
                    data['购物节'] = '普通日期'
                    
                    # 添加年、月、日列（如果不存在）
                    if '日期' in data.columns:
                        data['年'] = data['日期'].dt.year
                        data['月'] = data['日期'].dt.month
                        data['日'] = data['日期'].dt.day
                    
                    # 春节 (假设2022年春节在2月1日，2023年春节在1月22日)
                    spring_festival_2022 = (data['年'] == 2022) & (data['月'] == 2) & (data['日'] <= 7)
                    spring_festival_2023 = (data['年'] == 2023) & (data['月'] == 1) & (data['日'] >= 20) & (data['日'] <= 27)
                    
                    # 618购物节
                    festival_618 = (data['月'] == 6) & (data['日'] >= 10) & (data['日'] <= 20)
                    
                    # 双11购物节
                    festival_1111 = (data['月'] == 11) & (data['日'] >= 9) & (data['日'] <= 12)
                    
                    # 创建标记特殊日期的列
                    data.loc[spring_festival_2022 | spring_festival_2023, '购物节'] = '春节'
                    data.loc[festival_618, '购物节'] = '618购物节'
                    data.loc[festival_1111, '购物节'] = '双11购物节'
                else:
                    data = self.analyzer.df.copy()
                
                try:
                    # 特殊购物节的销售统计
                    festival_sales = data.groupby(['购物节']).agg({
                        '订单ID': 'count',
                        '总价': 'sum',
                        '折扣率': 'mean'
                    }).reset_index()
                    
                    if not festival_sales.empty:
                        # 绘制条形图
                        bars = self.festival_canvas.axes.bar(festival_sales['购物节'], festival_sales['总价'])
                        self.festival_canvas.axes.set_title('各购物节销售额对比')
                        self.festival_canvas.axes.set_xlabel('购物节')
                        self.festival_canvas.axes.set_ylabel('销售额（元）')
                        self.festival_canvas.axes.grid(True, linestyle='--', alpha=0.7, axis='y')
                        
                        # 添加折扣率标签
                        if '折扣率' in festival_sales.columns:
                            for i, (_, row) in enumerate(festival_sales.iterrows()):
                                self.festival_canvas.axes.text(i, row['总价'] * 0.5, 
                                            f'平均折扣: {row["折扣率"]:.2f}', 
                                            ha='center', va='center', color='white', fontweight='bold')
                        
                        self.festival_canvas.fig.tight_layout()
                        self.festival_canvas.draw()
                except Exception as e:
                    print(f"处理购物节数据时出错: {str(e)}")
                    # 错误不影响整体功能
        except Exception as e:
            raise Exception(f"更新促销效果分析出错: {str(e)}")
    
    def update_sales_forecast(self):
        """更新销售预测结果"""
        try:
            # 清除现有图表和表格
            self.forecast_canvas.clear()
            self.forecast_table.setRowCount(0)
            
            if self.analyzer.df is None:
                raise ValueError("请先加载数据")
                
            # 获取预测参数
            time_unit = self.forecast_time_combo.currentText()
            method = self.forecast_method_combo.currentText().lower()
            periods = self.forecast_periods_spin.value()
            
            # 获取商品类别
            category = self.category_combo.currentText()
            if category == "全部":
                category = None
                
            # 根据选择的预测方法调整参数
            if method == "指数平滑":
                method = "exponential_smoothing"
            elif method == "线性回归":
                method = "linear"
            elif method == "prophet":
                method = "prophet"
                
            # 生成预测结果
            forecast_data = self.analyzer.predict_sales(
                time_unit=time_unit,
                category=category,
                periods=periods,
                method=method
            )
            
            if forecast_data is not None and not forecast_data.empty:
                # 绘制预测图表
                self.plot_forecast(forecast_data)
                
                # 更新预测表格
                self.update_forecast_table(forecast_data)
                
                self.statusBar().showMessage(f'销售预测已完成：使用{method}方法，预测{periods}个{time_unit}')
                
        except Exception as e:
            QMessageBox.warning(self, "警告", f"生成销售预测时出错: {str(e)}")
            
    def plot_forecast(self, forecast_data):
        """绘制预测图表"""
        self.forecast_canvas.axes.clear()
        
        # 设置图表背景颜色
        self.forecast_canvas.axes.set_facecolor(COLOR_LIGHT_BG)
        
        # 提取数据
        history_data = forecast_data[forecast_data['数据类型'] == '历史']
        forecast = forecast_data[forecast_data['数据类型'] == '预测']
        
        # 绘制历史数据
        if not history_data.empty:
            self.forecast_canvas.axes.plot(
                history_data['日期'], 
                history_data['实际销售额'], 
                color=COLOR_PRIMARY, 
                marker='o', 
                label='历史销售额',
                linewidth=2.5,
                markerfacecolor='white',
                markeredgecolor=COLOR_PRIMARY,
                markersize=6
            )
            
            # 添加历史数据的阴影区域
            self.forecast_canvas.axes.fill_between(
                history_data['日期'],
                0, 
                history_data['实际销售额'],
                color=COLOR_PRIMARY,
                alpha=0.1
            )
            
        # 绘制预测数据
        if not forecast.empty:
            self.forecast_canvas.axes.plot(
                forecast['日期'], 
                forecast['预测销售额'], 
                color=COLOR_ACCENT, 
                marker='x', 
                label='预测销售额',
                linewidth=2.5,
                linestyle='--',
                markersize=6
            )
            
            # 如果有预测区间，绘制区间
            if '预测下限' in forecast.columns and '预测上限' in forecast.columns:
                self.forecast_canvas.axes.fill_between(
                    forecast['日期'],
                    forecast['预测下限'],
                    forecast['预测上限'],
                    color=COLOR_ACCENT,
                    alpha=0.2,
                    label='预测区间'
                )
                
        # 美化坐标轴
        self.forecast_canvas.axes.spines['top'].set_visible(False)
        self.forecast_canvas.axes.spines['right'].set_visible(False)
        self.forecast_canvas.axes.spines['left'].set_color('#ddd')
        self.forecast_canvas.axes.spines['bottom'].set_color('#ddd')
        
        # 设置图表属性
        self.forecast_canvas.axes.set_title('销售额预测分析', fontsize=14, fontweight='bold')
        self.forecast_canvas.axes.set_xlabel('日期', fontsize=12)
        self.forecast_canvas.axes.set_ylabel('销售额（元）', fontsize=12)
        self.forecast_canvas.axes.grid(True, linestyle='--', alpha=0.7, color='#ddd')
        
        # 美化图例
        self.forecast_canvas.axes.legend(
            loc='upper left',
            frameon=True,
            framealpha=0.9,
            facecolor='white',
            edgecolor='#ddd'
        )
        
        # 优化x轴标签
        if len(forecast_data) > 12:
            step = max(1, len(forecast_data) // 12)  # 最多显示12个标签
            labels = self.forecast_canvas.axes.get_xticklabels()
            for i, label in enumerate(labels):
                if i % step != 0:
                    label.set_visible(False)
                    
        # 旋转标签
        plt.setp(self.forecast_canvas.axes.get_xticklabels(), rotation=45, ha='right')
        
        # 添加分隔线，区分历史和预测
        if not history_data.empty and not forecast.empty:
            # 找到历史和预测的分界点
            last_history_date = history_data['日期'].iloc[-1]
            self.forecast_canvas.axes.axvline(
                x=last_history_date, 
                color='#777',
                linestyle='--',
                alpha=0.7
            )
            
            # 在分隔线上方添加"预测开始"文字
            y_max = self.forecast_canvas.axes.get_ylim()[1]
            self.forecast_canvas.axes.text(
                last_history_date, 
                y_max * 0.95, 
                '预测开始',
                ha='center',
                va='top',
                color='#777',
                fontsize=9,
                bbox=dict(boxstyle="round,pad=0.3", fc='white', ec="#ddd", alpha=0.8)
            )
        
        self.forecast_canvas.fig.tight_layout()
        self.forecast_canvas.draw()
        
    def update_forecast_table(self, forecast_data):
        self.forecast_canvas.fig.tight_layout()
        self.forecast_canvas.draw()
        
    def update_forecast_table(self, forecast_data):
        """更新预测数据表格"""
        # 清空表格
        self.forecast_table.setRowCount(0)
        
        # 设置行数
        self.forecast_table.setRowCount(len(forecast_data))
        
        # 添加数据到表格
        for i, (_, row) in enumerate(forecast_data.iterrows()):
            # 日期
            self.forecast_table.setItem(i, 0, QTableWidgetItem(str(row['日期'])))
            
            # 实际销售额
            if '实际销售额' in row and not pd.isna(row['实际销售额']):
                self.forecast_table.setItem(i, 1, QTableWidgetItem(f"{row['实际销售额']:,.2f}"))
            else:
                self.forecast_table.setItem(i, 1, QTableWidgetItem(""))
                
            # 预测销售额
            self.forecast_table.setItem(i, 2, QTableWidgetItem(f"{row['预测销售额']:,.2f}"))
            
            # 数据类型
            item = QTableWidgetItem(row['数据类型'])
            if row['数据类型'] == '预测':
                item.setBackground(QColor(255, 240, 240))  # 淡红色背景
            self.forecast_table.setItem(i, 3, item)
            
    def update_decision_suggestions(self):
        """更新决策建议"""
        try:
            # 清空建议文本
            self.suggestions_text.clear()
            
            if self.analyzer.df is None:
                raise ValueError("请先加载数据")
                
            # 获取商品类别
            category = self.decision_category_combo.currentText()
            if category == "全部":
                category = None
                
            # 生成决策建议
            suggestions = self.analyzer.generate_decision_suggestions(category)
            
            if suggestions:
                # 构建HTML格式的建议文本
                html = '<html><body style="font-family: Arial, sans-serif;">'
                html += f'<h2 style="color: #333;">电商平台决策建议{" - " + category if category else ""}</h2>'
                
                for suggestion in suggestions:
                    html += f'<div style="margin-bottom: 20px;">'
                    html += f'<h3 style="color: #0066cc;">{suggestion["类型"]}</h3>'
                    html += f'<p style="line-height: 1.5; color: #444;">{suggestion["建议"].replace("；", "<br>")}</p>'
                    html += '</div>'
                
                html += '</body></html>'
                
                # 设置文本
                self.suggestions_text.setHtml(html)
                
                self.statusBar().showMessage('决策建议已生成')
                
        except Exception as e:
            QMessageBox.warning(self, "警告", f"生成决策建议时出错: {str(e)}")
            self.suggestions_text.setPlainText(f"生成决策建议时出错: {str(e)}")
    
    def _path_to_url(self, path):
        """将本地文件路径转换为URL格式，确保在所有平台上正确显示"""
        # 确保路径是绝对路径
        abs_path = os.path.abspath(path)
        
        # 转换Windows路径分隔符
        url_path = abs_path.replace('\\', '/')
        
        # 添加file:// 前缀
        if not url_path.startswith('/'):
            url_path = '/' + url_path
        
        return f'file://{url_path}'
    
    def generate_report(self):
        """生成分析报告"""
        try:
            if self.analyzer.df is None:
                raise ValueError("请先加载数据")
                
            # 设置等待光标
            QApplication.setOverrideCursor(Qt.WaitCursor)
                
            # 获取报告标题和时间范围
            report_title = self.report_title_edit.toPlainText()
            start_date = self.start_date_edit.date().toString("yyyy-MM-dd")
            end_date = self.end_date_edit.date().toString("yyyy-MM-dd")
            
            # 创建HTML报告
            html = self.create_report_html(
                title=report_title,
                start_date=start_date,
                end_date=end_date
            )
            
            # 修改图片路径为file://格式
            for key, path in self.report_image_paths.items():
                # 使用辅助函数转换路径
                path_pattern = f'src="{path}"'
                file_url = f'src="{self._path_to_url(path)}"'
                html = html.replace(path_pattern, file_url)
            
            # 显示报告预览
            self.report_preview.setHtml(html)
            
            # 调试输出
            print("报告图片路径检查:")
            for key, path in self.report_image_paths.items():
                print(f"{key}原始路径: {path}")
                print(f"{key}URL路径: {self._path_to_url(path)}")
                print(f"文件是否存在: {os.path.exists(path)}")
            
            # 恢复正常光标
            QApplication.restoreOverrideCursor()
            
            self.statusBar().showMessage('报告已生成')
            
        except Exception as e:
            # 恢复正常光标
            QApplication.restoreOverrideCursor()
            import traceback
            error_details = traceback.format_exc()
            print(f"生成报告出错: {error_details}")
            QMessageBox.warning(self, "警告", f"生成报告时出错: {str(e)}")
    
    def create_report_html(self, title, start_date, end_date):
        """创建HTML格式的报告"""
        # 确保所有临时图片文件都保存在同一目录中
        current_dir = os.path.dirname(os.path.abspath(__file__))
        image_dir = current_dir
        os.makedirs(image_dir, exist_ok=True)
        
        # 预先创建图表文件列表，用于跟踪生成的所有图表
        image_paths = {}
        
        # 导入base64用于图片编码
        import base64
        import io
        
        # 定义函数将图表直接转换为base64编码的图片数据
        def fig_to_base64(canvas):
            if not hasattr(canvas, 'fig') or not hasattr(canvas, 'axes'):
                return None
            
            # 检查图表是否有内容
            has_content = False
            if hasattr(canvas.axes, 'get_lines') and canvas.axes.get_lines():
                has_content = True
            elif hasattr(canvas.axes, 'get_images') and canvas.axes.get_images():
                has_content = True
            elif hasattr(canvas.axes, 'patches') and canvas.axes.patches:
                has_content = True
            
            if not has_content:
                print("图表内容为空，跳过")
                return None
            
            try:
                # 将图表保存到内存buffer
                buf = io.BytesIO()
                canvas.fig.savefig(buf, format='png', dpi=150, bbox_inches='tight')
                buf.seek(0)
                
                # 转换为base64编码
                img_data = base64.b64encode(buf.getvalue()).decode('utf-8')
                return img_data
            except Exception as e:
                print(f"图表转换为base64出错: {str(e)}")
                return None
        
        # 报告头部和样式
        html = f'''
        <!DOCTYPE html>
        <html>
        <head>
            <meta charset="UTF-8">
            <title>{title}</title>
            <style>
                body {{ font-family: Arial, sans-serif; margin: 20px; color: #333; }}
                h1 {{ color: #0066cc; text-align: center; margin-bottom: 20px; font-size: 24px; }}
                h2 {{ color: #0066cc; margin-top: 30px; border-bottom: 1px solid #ddd; padding-bottom: 5px; font-size: 20px; }}
                h3 {{ color: #333; margin-top: 20px; font-size: 16px; }}
                .report-header {{ text-align: center; margin-bottom: 30px; }}
                .report-date {{ font-style: italic; color: #666; margin-bottom: 20px; text-align: center; }}
                .chart-container {{ text-align: center; margin: 20px 0; }}
                .chart {{ max-width: 100%; height: auto; border: 1px solid #ddd; }}
                table {{ border-collapse: collapse; width: 100%; margin: 20px 0; }}
                th, td {{ text-align: left; padding: 8px; border: 1px solid #ddd; }}
                th {{ background-color: #f2f2f2; }}
                tr:nth-child(even) {{ background-color: #f9f9f9; }}
                .suggestion {{ margin: 10px 0; padding: 10px; background-color: #f2f9ff; border-left: 3px solid #0066cc; }}
                
                /* 响应式设计，确保在小屏幕上也能良好显示 */
                @media screen and (max-width: 800px) {{
                    .chart {{ width: 100%; }}
                    table {{ font-size: 12px; }}
                }}
            </style>
        </head>
        <body>
            <div class="report-header">
                <h1>{title}</h1>
                <div class="report-date">报告时间范围: {start_date} 至 {end_date}</div>
                <div>生成时间: {datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")}</div>
            </div>
        '''
        
        # 数据概览部分
        summary = self.analyzer.get_data_summary()
        if summary and self.include_trend_check.isChecked():
            html += '''
            <h2>1. 数据概览</h2>
            <table>
                <tr><th>属性</th><th>值</th></tr>
            '''
            
            for key, value in summary.items():
                # 如果值是列表，将其转换为字符串
                if isinstance(value, list):
                    value_str = ", ".join(map(str, value))
                else:
                    value_str = str(value)
                
                html += f'<tr><td>{key}</td><td>{value_str}</td></tr>'
            
            html += '</table>'
        
        # 销售趋势分析部分
        if self.include_trend_check.isChecked():
            html += '<h2>2. 销售趋势分析</h2>'
            
            # 直接将趋势图转换为base64并嵌入
            if hasattr(self, 'trend_canvas'):
                trend_base64 = fig_to_base64(self.trend_canvas)
                if trend_base64:
                    image_paths['trend'] = 'base64_embedded'
                    html += f'''
                    <div class="chart-container">
                        <h3>销售趋势图</h3>
                        <img src="data:image/png;base64,{trend_base64}" class="chart" alt="销售趋势图">
                    </div>
                    '''
                else:
                    html += '<p>无法生成销售趋势图</p>'
            
            # 直接将热图转换为base64并嵌入
            if hasattr(self, 'heatmap_canvas'):
                heatmap_base64 = fig_to_base64(self.heatmap_canvas)
                if heatmap_base64:
                    image_paths['heatmap'] = 'base64_embedded'
                    html += f'''
                    <div class="chart-container">
                        <h3>销售热图</h3>
                        <img src="data:image/png;base64,{heatmap_base64}" class="chart" alt="销售热图">
                    </div>
                    '''
                else:
                    html += '<p>无法生成销售热图</p>'
        
        # 商品类别分析部分
        if self.include_category_check.isChecked():
            html += '<h2>3. 商品类别分析</h2>'
            
            # 直接将类别销售图转换为base64并嵌入
            if hasattr(self, 'category_canvas'):
                category_base64 = fig_to_base64(self.category_canvas)
                if category_base64:
                    image_paths['category'] = 'base64_embedded'
                    html += f'''
                    <div class="chart-container">
                        <h3>商品类别销售额对比</h3>
                        <img src="data:image/png;base64,{category_base64}" class="chart" alt="商品类别销售额对比">
                    </div>
                    '''
                else:
                    html += '<p>无法生成商品类别销售图</p>'
            
            # 添加热销商品表格
            if hasattr(self, 'top_products_table') and self.top_products_table.rowCount() > 0:
                html += '''
                <h3>热销商品TOP10</h3>
                <table>
                    <tr><th>商品名称</th><th>销售额</th></tr>
                '''
                
                for row in range(self.top_products_table.rowCount()):
                    product_name = self.top_products_table.item(row, 0).text()
                    product_sales = self.top_products_table.item(row, 1).text()
                    html += f'<tr><td>{product_name}</td><td>{product_sales}</td></tr>'
                
                html += '</table>'
        
        # 地区销售分析部分
        if self.include_region_check.isChecked():
            html += '<h2>4. 地区销售分析</h2>'
            
            # 直接将地区销售图转换为base64并嵌入
            if hasattr(self, 'region_canvas'):
                region_base64 = fig_to_base64(self.region_canvas)
                if region_base64:
                    image_paths['region'] = 'base64_embedded'
                    html += f'''
                    <div class="chart-container">
                        <h3>地区销售额对比</h3>
                        <img src="data:image/png;base64,{region_base64}" class="chart" alt="地区销售额对比">
                    </div>
                    '''
                else:
                    html += '<p>无法生成地区销售图</p>'
        
        # 客户分析部分
        if self.include_customer_check.isChecked():
            html += '<h2>5. 客户分析</h2>'
            
            # 添加客户群体分析表格
            if hasattr(self, 'customer_table') and self.customer_table.rowCount() > 0:
                html += '''
                <h3>客户群体分析</h3>
                <table>
                    <tr><th>客户群体</th><th>客户数量</th><th>平均消费额</th></tr>
                '''
                
                for row in range(self.customer_table.rowCount()):
                    segment = self.customer_table.item(row, 0).text()
                    count = self.customer_table.item(row, 1).text()
                    avg_amount = self.customer_table.item(row, 2).text()
                    html += f'<tr><td>{segment}</td><td>{count}</td><td>{avg_amount}</td></tr>'
                
                html += '</table>'
        
        # 促销效果分析部分
        if self.include_promotion_check.isChecked():
            html += '<h2>6. 促销效果分析</h2>'
            
            # 直接将促销效果图转换为base64并嵌入
            if hasattr(self, 'festival_canvas'):
                festival_base64 = fig_to_base64(self.festival_canvas)
                if festival_base64:
                    image_paths['festival'] = 'base64_embedded'
                    html += f'''
                    <div class="chart-container">
                        <h3>购物节销售额对比</h3>
                        <img src="data:image/png;base64,{festival_base64}" class="chart" alt="购物节销售额对比">
                    </div>
                    '''
                else:
                    html += '<p>无法生成购物节销售图</p>'
        
        # 销售预测部分
        if self.include_forecast_check.isChecked():
            html += '<h2>7. 销售预测</h2>'
            
            # 直接将销售预测图转换为base64并嵌入
            if hasattr(self, 'forecast_canvas'):
                forecast_base64 = fig_to_base64(self.forecast_canvas)
                if forecast_base64:
                    image_paths['forecast'] = 'base64_embedded'
                    html += f'''
                    <div class="chart-container">
                        <h3>销售预测图</h3>
                        <img src="data:image/png;base64,{forecast_base64}" class="chart" alt="销售预测图">
                    </div>
                    '''
                else:
                    html += '<p>无法生成销售预测图</p>'
            
            # 添加预测数据表格
            if hasattr(self, 'forecast_table') and self.forecast_table.rowCount() > 0:
                html += '''
                <h3>销售预测数据</h3>
                <table>
                    <tr><th>日期</th><th>实际销售额</th><th>预测销售额</th><th>数据类型</th></tr>
                '''
                
                for row in range(self.forecast_table.rowCount()):
                    if self.forecast_table.item(row, 0) and self.forecast_table.item(row, 3):
                        if self.forecast_table.item(row, 3).text() == "预测":  # 只显示预测部分
                            date = self.forecast_table.item(row, 0).text()
                            actual = self.forecast_table.item(row, 1).text() if self.forecast_table.item(row, 1) else "-"
                            forecast = self.forecast_table.item(row, 2).text() if self.forecast_table.item(row, 2) else "-"
                            data_type = self.forecast_table.item(row, 3).text()
                            html += f'<tr><td>{date}</td><td>{actual}</td><td>{forecast}</td><td>{data_type}</td></tr>'
                
                html += '</table>'
        
        # 决策建议部分
        if self.include_decision_check.isChecked():
            html += '<h2>8. 决策建议</h2>'
            
            # 获取决策建议
            try:
                category = None
                suggestions = self.analyzer.generate_decision_suggestions(category)
                
                if suggestions:
                    for suggestion in suggestions:
                        html += f'''
                        <div class="suggestion">
                            <h3>{suggestion["类型"]}</h3>
                            <p>{suggestion["建议"].replace("；", "<br>")}</p>
                        </div>
                        '''
            except Exception as e:
                html += f'<p>无法生成决策建议: {str(e)}</p>'
        
        # 报告结尾
        html += '''
            <div style="margin-top: 50px; text-align: center; color: #666; font-size: 12px;">
                <p>由基于大数据的电商平台商品销售趋势分析与决策软件自动生成</p>
            </div>
        </body>
        </html>
        '''
        
        # 保存图像路径以供后续使用
        self.report_image_paths = image_paths
        
        return html
    
    def save_report(self):
        """保存报告到HTML文件"""
        try:
            # 检查是否已生成报告
            if not hasattr(self, 'report_image_paths'):
                QMessageBox.warning(self, "警告", "请先生成报告再保存")
                return
                
            # 获取当前文件所在目录
            current_dir = os.path.dirname(os.path.abspath(__file__))
            # 构建报告保存路径
            report_dir = os.path.join(current_dir, 'reports')
            
            # 如果reports文件夹不存在，创建它
            if not os.path.exists(report_dir):
                os.makedirs(report_dir)
            
            # 获取当前日期时间作为文件名一部分
            now_str = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
            default_filename = f"销售分析报告_{now_str}.html"
            
            # 打开文件保存对话框
            file_path, _ = QFileDialog.getSaveFileName(
                self,
                "保存报告",
                os.path.join(report_dir, default_filename),
                "HTML文件 (*.html)"
            )
            
            if file_path:
                # 创建进度对话框
                progress = QProgressDialog("正在保存报告...", "取消", 0, 100, self)
                progress.setWindowTitle("保存报告")
                progress.setWindowModality(Qt.WindowModal)
                progress.setMinimumDuration(0)
                progress.setValue(0)
                progress.show()
                
                progress.setValue(50)
                progress.setLabelText("正在处理HTML内容...")
                
                # 获取当前预览中的HTML内容
                self.report_preview.page().toHtml(lambda content: self._process_and_save_html(file_path, content, None, progress))
                
        except Exception as e:
            import traceback
            error_details = traceback.format_exc()
            print(f"保存报告出错: {error_details}")
            QMessageBox.warning(self, "警告", f"保存报告时出错: {str(e)}")
        
    def _process_and_save_html(self, file_path, content, image_map, progress):
        """处理HTML内容并保存到文件"""
        try:
            progress.setValue(70)
            progress.setLabelText("正在更新图片路径...")
            
            # 创建报告文件夹
            report_name = os.path.splitext(os.path.basename(file_path))[0]
            report_folder = os.path.join(os.path.dirname(file_path), report_name + "_files")
            
            # 确保报告文件夹存在
            if not os.path.exists(report_folder):
                os.makedirs(report_folder)
            
            # 替换HTML中的图片路径
            for key, path in self.report_image_paths.items():
                # 查找图片标签
                url_path = self._path_to_url(path)  # 原始的file://路径
                
                # 计算相对路径
                rel_path = os.path.join(report_name + "_files", os.path.basename(path))
                
                # 替换路径
                content = content.replace(f'src="{url_path}"', f'src="{rel_path}"')
                
                # 复制图片文件到报告文件夹
                target_path = os.path.join(report_folder, os.path.basename(path))
                import shutil
                if os.path.exists(path):
                    try:
                        shutil.copy2(path, target_path)
                        print(f"已复制图片: {path} -> {target_path}")
                    except Exception as e:
                        print(f"复制图片出错: {path} -> {target_path}: {str(e)}")
                else:
                    print(f"警告: 源图片不存在: {path}")
            
            progress.setValue(90)
            progress.setLabelText("正在写入文件...")
            
            # 保存处理后的HTML内容
            with open(file_path, 'w', encoding='utf-8') as f:
                f.write(content)
                
            progress.setValue(100)
            QMessageBox.information(self, "成功", f"报告已成功保存至:\n{file_path}")
            self.statusBar().showMessage(f'报告已保存至: {file_path}')
            
        except Exception as e:
            import traceback
            error_details = traceback.format_exc()
            print(f"处理和保存HTML出错: {error_details}")
            QMessageBox.warning(self, "警告", f"保存报告时出错: {str(e)}")
        finally:
            if progress and progress.isVisible():
                progress.close()