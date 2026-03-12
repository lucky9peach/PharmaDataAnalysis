import os
import sys
import pandas as pd
import numpy as np
import platform

# 解决跨平台中文字体显示方块的问题
import matplotlib
if platform.system() == 'Darwin':
    matplotlib.rcParams['font.sans-serif'] = ['Arial Unicode MS', 'PingFang SC']
else:
    matplotlib.rcParams['font.sans-serif'] = ['Microsoft YaHei', 'SimHei']
matplotlib.rcParams['axes.unicode_minus'] = False

from PySide6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, 
                               QLabel, QPushButton, QComboBox, QFrame, QScrollArea, QSplitter, 
                               QMessageBox, QTableWidget, QTableWidgetItem, QHeaderView, QTabWidget,
                               QTreeWidget, QTreeWidgetItem, QRadioButton, QButtonGroup, QMenu,
                               QSpinBox, QDoubleSpinBox)
from PySide6.QtCore import Qt, QTimer
from PySide6.QtGui import QStandardItemModel, QStandardItem, QImage, QPainter, QClipboard
from matplotlib.backends.backend_qtagg import FigureCanvasQTAgg as FigureCanvas
import matplotlib.pyplot as plt

# 复用核心配置
try:
    import core_config
except ImportError:
    pass

# 抽取自 core_config 的欧洲经济区 + 英国
EEA_AND_UK_MARKETS = [
    "AUSTRIA", "BELGIUM", "BULGARIA", "CROATIA", "CYPRUS",
    "CZECH REPUBLIC", "DENMARK", "ESTONIA", "FINLAND", "FRANCE",
    "GERMANY", "GREECE", "HUNGARY", "ICELAND", "IRELAND",
    "ITALY", "LATVIA", "LIECHTENSTEIN", "LITHUANIA", "LUXEMBOURG",
    "MALTA", "NETHERLANDS", "NORWAY", "POLAND", "PORTUGAL",
    "ROMANIA", "SLOVAKIA", "SLOVENIA", "SPAIN", "SWEDEN",
    "UNITED KINGDOM"
]

# 通用带勾选库复用 (简化版)
from PySide6.QtGui import QStandardItemModel, QStandardItem
class CheckableComboBox(QComboBox):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.source_model = QStandardItemModel(self)
        self.source_model.dataChanged.connect(self.updateText)
        
        from PySide6.QtCore import QSortFilterProxyModel
        self.pFilterModel = QSortFilterProxyModel(self)
        self.pFilterModel.setFilterCaseSensitivity(Qt.CaseInsensitive)
        self.pFilterModel.setSourceModel(self.source_model)
        
        self.setModel(self.pFilterModel)
        self.view().pressed.connect(self.handleItemPressed)
        self._changed = False
        
        self.setEditable(True)
        self.lineEdit().setReadOnly(False)
        self.lineEdit().setPlaceholderText("可输入文本搜索或下拉勾选...")
        self.lineEdit().installEventFilter(self)
        self.lineEdit().textEdited.connect(self._on_text_edited)
        self.updateText()
        
    def _on_text_edited(self, text):
        self.pFilterModel.setFilterFixedString(text)
        self.showPopup()
        
    def addCheckableItem(self, text, checked=True):
        item = QStandardItem(text)
        item.setFlags(Qt.ItemIsUserCheckable | Qt.ItemIsEnabled)
        item.setData(Qt.Checked if checked else Qt.Unchecked, Qt.CheckStateRole)
        self.source_model.appendRow(item)
        
    def clear(self):
        self.source_model.clear()
        super().clear()
        self.updateText()
        
    def check_all(self):
        for i in range(self.source_model.rowCount()):
            item = self.source_model.item(i)
            if item and item.isEnabled():
                item.setCheckState(Qt.Checked)
        self._changed = True
        self.updateText()
        
    def uncheck_all(self):
        for i in range(self.source_model.rowCount()):
            item = self.source_model.item(i)
            if item and item.isEnabled():
                item.setCheckState(Qt.Unchecked)
        self._changed = True
        self.updateText()
        
    def eventFilter(self, widget, event):
        if widget == self.lineEdit() and event.type() == event.Type.MouseButtonRelease:
            self.showPopup()
            return True
        return super().eventFilter(widget, event)
        
    def handleItemPressed(self, index):
        source_index = self.pFilterModel.mapToSource(index)
        item = self.source_model.itemFromIndex(source_index)
        if item and item.isEnabled():
            if item.checkState() == Qt.Checked:
                item.setCheckState(Qt.Unchecked)
            else:
                item.setCheckState(Qt.Checked)
            self._changed = True
            
    def hidePopup(self):
        if not self._changed:
            super().hidePopup()
            self.pFilterModel.setFilterFixedString("")
            self.updateText()
        self._changed = False

    def updateText(self):
        if self.lineEdit().hasFocus():
            return
        checked = self.get_checked_items()
        if checked:
            self.lineEdit().setText(", ".join(checked))
        else:
            self.lineEdit().setText("全选 / 未选...")

    def get_checked_items(self):
        checkedItems = []
        for i in range(self.source_model.rowCount()):
            item = self.source_model.item(i)
            if item and item.checkState() == Qt.Checked:
                checkedItems.append(item.text())
        return checkedItems


class ForecastApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("销预测与成本分析系统 (Volume Forecast & Cost Analysis)")
        self.resize(1300, 850)
        
        self.raw_df = pd.DataFrame() # 当前加载的API数据
        self.filtered_df = pd.DataFrame() # 应用维度过滤后的数据
        self.cache_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "Cache", "API_Views")
        
        self.setup_ui()
        self.load_api_list()
        
    def setup_ui(self):
        self.setStyleSheet("""
            QMainWindow { background-color: #f7fafc; }
            QFrame#Sidebar { background-color: #ffffff; border-right: 1px solid #e2e8f0; }
            QLabel.Title { font-size: 16px; font-weight: bold; color: #2c5282; margin-bottom: 5px; }
            QLabel.Sub { font-size: 13px; font-weight: bold; color: #4a5568; margin-top: 10px; }
            QComboBox, QLineEdit { padding: 5px; border: 1px solid #cbd5e0; border-radius: 4px; }
            QPushButton.ActionBtn { background-color: #3182ce; color: white; font-weight: bold; padding: 8px; border-radius: 4px; }
            QPushButton.ActionBtn:hover { background-color: #2b6cb0; }
            QPushButton.ActionBtn:disabled { background-color: #a0aec0; }
        """)
        
        central = QWidget()
        self.setCentralWidget(central)
        layout = QHBoxLayout(central)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)
        
        # 侧边栏：步骤输入区
        self.sidebar = QFrame()
        self.sidebar.setObjectName("Sidebar")
        self.sidebar.setFixedWidth(320)
        side_layout = QVBoxLayout(self.sidebar)
        side_layout.setAlignment(Qt.AlignTop)
        side_layout.setContentsMargins(15, 20, 15, 20)
        
        title = QLabel("步骤导引")
        title.setProperty("class", "Title")
        side_layout.addWidget(title)
        
        # 1. 第一步：加载 API
        lbl_step1 = QLabel("1. 极速调取单品 (API)")
        lbl_step1.setProperty("class", "Sub")
        self.api_combo = QComboBox()
        self.api_combo.currentIndexChanged.connect(self.on_api_selected)
        side_layout.addWidget(lbl_step1)
        side_layout.addWidget(self.api_combo)
        
        # 2. 第二步：国家选择
        lbl_step2 = QLabel("2. 地域界定 (多选)")
        lbl_step2.setProperty("class", "Sub")
        self.country_combo = CheckableComboBox()
        
        btn_eea_row = QHBoxLayout()
        btn_eea_row.setContentsMargins(0,0,0,0)
        btn_eea_all = QPushButton("全选 EEA+UK")
        btn_eea_clr = QPushButton("清空")
        btn_style = "QPushButton { padding: 4px; font-size: 11px; background-color: #cbd5e0; border-radius: 3px; } QPushButton:hover { background-color: #a0aec0; }"
        btn_eea_all.setStyleSheet(btn_style)
        btn_eea_clr.setStyleSheet(btn_style)
        btn_eea_all.clicked.connect(self.country_combo.check_all)
        btn_eea_clr.clicked.connect(self.country_combo.uncheck_all)
        btn_eea_row.addWidget(self.country_combo)
        btn_eea_row.addWidget(btn_eea_all)
        btn_eea_row.addWidget(btn_eea_clr)
        
        side_layout.addWidget(lbl_step2)
        side_layout.addLayout(btn_eea_row)
        
        # 3. 第三步：剂型大类合并
        lbl_step3 = QLabel("3. 剂型锁定 (自动归类合并)")
        lbl_step3.setProperty("class", "Sub")
        self.dosage_combo = CheckableComboBox()
        side_layout.addWidget(lbl_step3)
        side_layout.addWidget(self.dosage_combo)
        
        # 执行分析按钮
        side_layout.addSpacing(20)
        self.btn_analyze = QPushButton("▶ 执行聚合测算")
        self.btn_analyze.setProperty("class", "ActionBtn")
        self.btn_analyze.clicked.connect(self.run_analysis)
        side_layout.addWidget(self.btn_analyze)
        
        side_layout.addStretch()
        layout.addWidget(self.sidebar)
        
        # 右侧：展示区
        self.main_area = QTabWidget()
        self.main_area.setStyleSheet("""
            QTabWidget::pane { border: 1px solid #e2e8f0; background: white; }
            QTabBar::tab { background: #edf2f7; border: 1px solid #cbd5e0; padding: 8px 20px; border-top-left-radius: 4px; border-top-right-radius: 4px; }
            QTabBar::tab:selected { background: #ffffff; border-bottom-color: #ffffff; font-weight: bold; color: #2b6cb0; }
        """)
        layout.addWidget(self.main_area)
        
        # ---- 标签页：综合产品评估面板 (Dashboard) ----
        self.tab_dashboard = QWidget()
        self.setup_dashboard_tab()
        self.main_area.addTab(self.tab_dashboard, "3. 综合评估地图 (Dashboard)")
        
        # ---- 标签页：规格与包装拆解分析 ----
        self.tab_pack = QWidget()
        self.setup_pack_tab()
        self.main_area.addTab(self.tab_pack, "4. 规格包装拆解 (Pack Size)")
        
        # ---- 标签页：长尾市场 (Others) 机会洞察 ----
        self.tab_others = QWidget()
        self.setup_others_tab()
        self.main_area.addTab(self.tab_others, "5. 长尾市场战略纵深 (Others Market)")
        
        # ---- 标签页：选品与收益智能推荐 ----
        self.tab_strategy = QWidget()
        self.setup_strategy_tab()
        self.main_area.addTab(self.tab_strategy, "6. 选品与收益智能推荐 (Strategy)")
        
    def setup_dashboard_tab(self):
        ly = QVBoxLayout(self.tab_dashboard)
        
        desc = QLabel("此面板对各类核心竞争性指标进行宏观汇总，涵盖容量、单价、复合增长率、核心竞争者等，高度贴合战略立项表格的评估需求。")
        desc.setWordWrap(True)
        desc.setStyleSheet("color: #718096; padding: 10px; background-color: #f6ad5511; border-radius: 5px;")
        ly.addWidget(desc)
        
        self.table_dashboard = QTableWidget()
        cols = [
            "国家 (Country)", "分子 (Molecule)", "剂型 (Formulation)", "规格组合 (Strength)",
            "总体市场容量/万单位", "总体市场金额/万美元", "预估单价/美元", "5年CAGR(量)",
            "对手数量", "≥10%份额厂家数", "主要对手", "主流规格", "API年用量/kg"
        ]
        self.table_dashboard.setColumnCount(len(cols))
        self.table_dashboard.setHorizontalHeaderLabels(cols)
        self.table_dashboard.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)
        self.table_dashboard.setEditTriggers(QTableWidget.NoEditTriggers)
        
        self.btn_export_dash = QPushButton("📥 导出评估地图至 Excel")
        self.btn_export_dash.setProperty("class", "ActionBtn")
        self.btn_export_dash.clicked.connect(self.export_dashboard)
        
        top_bar = QHBoxLayout()
        top_bar.addStretch()
        top_bar.addWidget(self.btn_export_dash)
        
        ly.addLayout(top_bar)
        ly.addWidget(self.table_dashboard)
        
    def export_dashboard(self):
        if self.table_dashboard.rowCount() == 0:
            QMessageBox.warning(self, "空数据", "当前没有可导出的数据！")
            return
            
        from PySide6.QtWidgets import QFileDialog
        path, _ = QFileDialog.getSaveFileName(self, "保存评估地图", "综合评估地图.xlsx", "Excel Files (*.xlsx)")
        if not path:
            return
            
        try:
            headers = [self.table_dashboard.horizontalHeaderItem(i).text() for i in range(self.table_dashboard.columnCount())]
            data = []
            for row in range(self.table_dashboard.rowCount()):
                row_data = []
                for col in range(self.table_dashboard.columnCount()):
                    item = self.table_dashboard.item(row, col)
                    row_data.append(item.text() if item else "")
                data.append(row_data)
                
            df = pd.DataFrame(data, columns=headers)
            df.to_excel(path, index=False)
            QMessageBox.information(self, "导出成功", f"文件已保存至：\n{path}")
        except Exception as e:
            QMessageBox.critical(self, "导出失败", f"导出时发生错误：\n{str(e)}")

    def setup_pack_tab(self):
        ly = QVBoxLayout(self.tab_pack)
        
        desc = QLabel("此表拆解了在您限定的【国家/剂型】下，不同【给药规格（Strength）】与【包装数量（Pack Size）】在不同维度的组合与占比情况。\n从上到下：上方为全盘的规格与包装拆分饼图/环形图可视化；下方从左至右依次为：A. 全盘规格与包装树状图； B. 点击不同国家查看其规格结构； C. 点击某国家内的公司查看其规格结构与量价。\n【图表支持右键复制】")
        desc.setWordWrap(True)
        desc.setStyleSheet("color: #718096; padding: 10px; background-color: #ebf8ff; border-radius: 5px;")
        ly.addWidget(desc)
        
        pack_vsplit = QSplitter(Qt.Vertical)
        
        # 上半部分: 可视化图表
        frame_chart = QFrame()
        l_chart = QVBoxLayout(frame_chart)
        self.fig_pack = plt.Figure(figsize=(10, 4), dpi=100)
        self.canvas_pack = FigureCanvas(self.fig_pack)
        self.canvas_pack.setContextMenuPolicy(Qt.CustomContextMenu)
        self.canvas_pack.customContextMenuRequested.connect(lambda pos: self.show_copy_menu(pos, self.canvas_pack))
        l_chart.addWidget(self.canvas_pack)
        pack_vsplit.addWidget(frame_chart)
        
        # 下半部分: 树状图列表
        main_split = QSplitter(Qt.Horizontal)
        
        # 1. 全盘架构 (总体)
        frame_global = QFrame()
        l_global = QVBoxLayout(frame_global)
        l_global.addWidget(QLabel("<b>[1] 选定范围整体大盘 (Global)</b>"))
        self.tree_global = QTreeWidget()
        self.tree_global.setHeaderLabels(["层级结构 (规格 -> 包装)", "销量 (Units)", "占比 (%)"])
        self.tree_global.header().setSectionResizeMode(0, QHeaderView.Stretch)
        l_global.addWidget(self.tree_global)
        main_split.addWidget(frame_global)
        
        # 2. 国家维度
        frame_country = QFrame()
        l_country = QVBoxLayout(frame_country)
        l_country.addWidget(QLabel("<b>[2] 按国家下钻 (Country)</b>"))
        
        country_split = QSplitter(Qt.Vertical)
        
        self.table_pack_countries = QTableWidget()
        self.table_pack_countries.setColumnCount(3)
        self.table_pack_countries.setHorizontalHeaderLabels(["国家", "总销量", "占比"])
        self.table_pack_countries.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.table_pack_countries.setSelectionBehavior(QTableWidget.SelectRows)
        self.table_pack_countries.setEditTriggers(QTableWidget.NoEditTriggers)
        self.table_pack_countries.cellClicked.connect(self.on_pack_country_clicked)
        country_split.addWidget(self.table_pack_countries)
        
        self.tree_country_detail = QTreeWidget()
        self.tree_country_detail.setHeaderLabels(["选中国家的层级结构", "销量", "占比"])
        self.tree_country_detail.header().setSectionResizeMode(0, QHeaderView.Stretch)
        country_split.addWidget(self.tree_country_detail)
        
        l_country.addWidget(country_split)
        main_split.addWidget(frame_country)
        
        # 3. 公司维度 (在指定国家下)
        frame_comp = QFrame()
        l_comp = QVBoxLayout(frame_comp)
        
        top_comp = QHBoxLayout()
        top_comp.addWidget(QLabel("<b>[3] 市场内公司下钻 (Company)</b>"))
        
        self.radio_comp_all = QRadioButton("全部公司")
        self.radio_comp_all.setChecked(True)
        self.radio_comp_others = QRadioButton("长尾 Others (>5名)")
        
        self.comp_group = QButtonGroup(self)
        self.comp_group.addButton(self.radio_comp_all)
        self.comp_group.addButton(self.radio_comp_others)
        self.comp_group.buttonClicked.connect(self.refresh_pack_companies)
        
        top_comp.addStretch()
        top_comp.addWidget(self.radio_comp_all)
        top_comp.addWidget(self.radio_comp_others)
        l_comp.addLayout(top_comp)
        
        comp_split = QSplitter(Qt.Vertical)
        
        self.table_pack_comps = QTableWidget()
        self.table_pack_comps.setColumnCount(3)
        self.table_pack_comps.setHorizontalHeaderLabels(["公司", "总销量", "占比"])
        self.table_pack_comps.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.table_pack_comps.setSelectionBehavior(QTableWidget.SelectRows)
        self.table_pack_comps.setEditTriggers(QTableWidget.NoEditTriggers)
        self.table_pack_comps.cellClicked.connect(self.on_pack_comp_clicked)
        comp_split.addWidget(self.table_pack_comps)
        
        self.tree_comp_detail = QTreeWidget()
        self.tree_comp_detail.setHeaderLabels(["选中公司的层级结构", "销量", "占比", "该组合均价"])
        self.tree_comp_detail.header().setSectionResizeMode(0, QHeaderView.Stretch)
        comp_split.addWidget(self.tree_comp_detail)
        
        l_comp.addWidget(comp_split)
        main_split.addWidget(frame_comp)
        
        main_split.setSizes([350, 350, 450])
        
        pack_vsplit.addWidget(main_split)
        pack_vsplit.setSizes([350, 400])
        ly.addWidget(pack_vsplit)
        
    def setup_others_tab(self):
        ly = QVBoxLayout(self.tab_others)
        
        desc = QLabel("在目标国家，由于前五大巨头（含原研或仿制头部）固若金汤，我们常通过争夺【Others】(排名五名开外的尾部门板) 的份额来切入。\n点击右侧国家列表中的国家，可在下方查看该国家 Others 内各公司的详细销售及主导品种情况。")
        desc.setWordWrap(True)
        desc.setStyleSheet("color: #718096; padding: 10px; background-color: #f0fff4; border-radius: 5px;")
        ly.addWidget(desc)
        
        # 上下分栏
        right_split = QSplitter(Qt.Vertical)
        
        # 表格容器: 国家级别
        self.table_others = QTableWidget()
        self.table_others.setColumnCount(5)
        self.table_others.setHorizontalHeaderLabels(["国家", "全境总销量", "Others总销量", "Others分红比", "Others内企业数"])
        self.table_others.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)
        self.table_others.setEditTriggers(QTableWidget.NoEditTriggers)
        self.table_others.setSelectionBehavior(QTableWidget.SelectRows)
        self.table_others.cellClicked.connect(self.on_country_clicked)
        right_split.addWidget(self.table_others)
        
        # 表格容器: Others企业详情
        self.table_others_detail = QTableWidget()
        self.table_others_detail.setColumnCount(6)
        self.table_others_detail.setHorizontalHeaderLabels(["公司", "长尾占比", "主力产品", "单品占比", "销售量", "预估单价"])
        self.table_others_detail.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.table_others_detail.setEditTriggers(QTableWidget.NoEditTriggers)
        right_split.addWidget(self.table_others_detail)
        
        ly.addWidget(right_split)

    def on_country_clicked(self, row, col):
        country_item = self.table_others.item(row, 0)
        if not country_item: return
        country = country_item.text()
        
        df = self.filtered_df
        if df.empty or 'market_region' not in df.columns or 'corporation_name' not in df.columns: return
        
        vdf = df[df['market_region'] == country].copy()
        if vdf.empty: return
        
        c_grp = vdf.groupby('corporation_name')['sales_volume_units'].sum().sort_values(ascending=False)
        if len(c_grp) <= 5:
            self.table_others_detail.setRowCount(0)
            return
            
        others_companies = c_grp.iloc[5:].index.tolist()
        others_df = vdf[vdf['corporation_name'].isin(others_companies)]
        others_vol_sum = c_grp.iloc[5:].sum()
        
        self.table_others_detail.setRowCount(0)
        
        for comp in others_companies:
            comp_df = others_df[others_df['corporation_name'] == comp]
            comp_vol = comp_df['sales_volume_units'].sum()
            if comp_vol <= 0: continue
            
            comp_pct = comp_vol / others_vol_sum if others_vol_sum > 0 else 0
            
            # Analyze main SKU
            if 'strength_raw' in comp_df.columns:
                sku_grp = comp_df.groupby('strength_raw')['sales_volume_units'].sum().sort_values(ascending=False)
                main_sku = sku_grp.index[0] if not sku_grp.empty else "未知"
                main_sku_vol = sku_grp.iloc[0] if not sku_grp.empty else 0
                main_sku_pct = main_sku_vol / comp_vol if comp_vol > 0 else 0
                
                sku_df = comp_df[comp_df['strength_raw'] == main_sku]
                if 'sales_value_usd' in sku_df.columns and main_sku_vol > 0:
                    est_price = sku_df['sales_value_usd'].sum() / main_sku_vol
                else:
                    est_price = 0
            else:
                main_sku = "未知"
                main_sku_pct = 0
                est_price = 0
                
            r = self.table_others_detail.rowCount()
            self.table_others_detail.insertRow(r)
            self.table_others_detail.setItem(r, 0, QTableWidgetItem(str(comp)))
            self.table_others_detail.setItem(r, 1, QTableWidgetItem(f"{comp_pct:.1%}"))
            self.table_others_detail.setItem(r, 2, QTableWidgetItem(str(main_sku)))
            self.table_others_detail.setItem(r, 3, QTableWidgetItem(f"{main_sku_pct:.1%}"))
            
            vol_item = QTableWidgetItem(f"{comp_vol:,.0f}")
            vol_item.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
            self.table_others_detail.setItem(r, 4, vol_item)
            
            price_item = QTableWidgetItem(f"${est_price:,.4f}")
            price_item.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
            self.table_others_detail.setItem(r, 5, price_item)
            
        self.table_others_detail.resizeColumnsToContents()
        
    def setup_strategy_tab(self):
        ly = QVBoxLayout(self.tab_strategy)
        
        desc = QLabel("此面板基于以上各维度数据为您生成【准入战略与收益预期】。\n由于产线包装机的限制（例如最大只能包装30片/盒等），系统将自动排除您无法生产的超大包装，并推荐各国剩余的最优组合；同时根据您预期的渗透率计算收益总池。")
        desc.setWordWrap(True)
        desc.setStyleSheet("color: #718096; padding: 10px; background-color: #fce88322; border-radius: 5px;")
        ly.addWidget(desc)
        
        # 控制台 (输入限制条件)
        ctrl_frame = QFrame()
        ctrl_layout = QHBoxLayout(ctrl_frame)
        
        ctrl_layout.addWidget(QLabel("<b>📦我方最大包装能力 (片/盒):</b>"))
        self.spin_max_pack = QSpinBox()
        self.spin_max_pack.setRange(1, 1000)
        self.spin_max_pack.setValue(30)
        self.spin_max_pack.setFixedWidth(60)
        ctrl_layout.addWidget(self.spin_max_pack)
        
        ctrl_layout.addSpacing(10)
        
        ctrl_layout.addWidget(QLabel("<b>💊最大规格制造限制 (mg):</b>"))
        self.spin_max_str = QSpinBox()
        self.spin_max_str.setRange(1, 5000)
        self.spin_max_str.setValue(1000) # 默认1000mg 基本都能做
        self.spin_max_str.setFixedWidth(70)
        ctrl_layout.addWidget(self.spin_max_str)
        
        ctrl_layout.addSpacing(10)
        
        ctrl_layout.addWidget(QLabel("<b>🎯 渗透率阶段预期 (%) Y1-Y3:</b>"))
        
        self.spin_pen_y1 = QDoubleSpinBox()
        self.spin_pen_y1.setRange(0.1, 100.0)
        self.spin_pen_y1.setValue(2.0)
        self.spin_pen_y1.setFixedWidth(60)
        ctrl_layout.addWidget(self.spin_pen_y1)
        
        self.spin_pen_y2 = QDoubleSpinBox()
        self.spin_pen_y2.setRange(0.1, 100.0)
        self.spin_pen_y2.setValue(5.0)
        self.spin_pen_y2.setFixedWidth(60)
        ctrl_layout.addWidget(self.spin_pen_y2)
        
        self.spin_pen_y3 = QDoubleSpinBox()
        self.spin_pen_y3.setRange(0.1, 100.0)
        self.spin_pen_y3.setValue(7.0)
        self.spin_pen_y3.setFixedWidth(60)
        ctrl_layout.addWidget(self.spin_pen_y3)
        
        ctrl_layout.addSpacing(20)
        btn_refresh_strat = QPushButton("⚡ 生成起量预测")
        btn_refresh_strat.setProperty("class", "ActionBtn")
        btn_refresh_strat.clicked.connect(self.render_strategy_analysis)
        ctrl_layout.addWidget(btn_refresh_strat)
        
        btn_copy_strat = QPushButton("📋 复制表格到Excel")
        btn_copy_strat.clicked.connect(self.copy_table_to_clipboard)
        ctrl_layout.addWidget(btn_copy_strat)
        
        ctrl_layout.addStretch()
        
        ly.addWidget(ctrl_frame)
        
        main_split = QSplitter(Qt.Vertical)
        
        # 上层：推荐表格
        self.table_strategy = QTableWidget()
        self.table_strategy.setColumnCount(8)
        self.table_strategy.setHorizontalHeaderLabels(["国家", "推荐主打规格", "可行畅销包装", "受限后可达总盘(Units)", "Y1量(Units)", "Y2量(Units)", "Y3量(Units)", "参考IMS单价(USD)"])
        self.table_strategy.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.table_strategy.setEditTriggers(QTableWidget.NoEditTriggers)
        main_split.addWidget(self.table_strategy)
        
        # 下层：收益图表
        frame_chart = QFrame()
        l_chart = QVBoxLayout(frame_chart)
        self.fig_strategy = plt.Figure(figsize=(10, 3), dpi=100)
        self.canvas_strategy = FigureCanvas(self.fig_strategy)
        self.canvas_strategy.setContextMenuPolicy(Qt.CustomContextMenu)
        self.canvas_strategy.customContextMenuRequested.connect(lambda pos: self.show_copy_menu(pos, self.canvas_strategy))
        l_chart.addWidget(self.canvas_strategy)
        main_split.addWidget(frame_chart)
        
        # 允许纵向拉伸表格，图表作为次要参考在最底端
        main_split.setSizes([600, 250])
        ly.addWidget(main_split)

    def load_api_list(self):
        from PySide6.QtWidgets import QCompleter
        self.api_combo.blockSignals(True)
        self.api_combo.clear()
        self.api_combo.addItem("--- 请选择 ---")
        
        self.api_combo.setEditable(True)
        self.api_combo.setInsertPolicy(QComboBox.NoInsert)
        if self.api_combo.completer():
            self.api_combo.completer().setCompletionMode(QCompleter.PopupCompletion)
            self.api_combo.completer().setFilterMode(Qt.MatchContains)
            
        api_names = []
        if os.path.exists(self.cache_dir):
            for f in os.listdir(self.cache_dir):
                if f.startswith("core_cache_") and f.endswith(".parquet"):
                    api_name = f.replace("core_cache_", "").replace(".parquet", "")
                    api_names.append(api_name)
                    
        api_names.sort(key=lambda x: x.upper())
        for name in api_names:
            self.api_combo.addItem(name)
            
        self.api_combo.blockSignals(False)

    def on_api_selected(self, idx):
        if idx <= 0: return
        api_name = self.api_combo.currentText()
        filepath = os.path.join(self.cache_dir, f"core_cache_{api_name}.parquet")
        
        try:
            self.raw_df = pd.read_parquet(filepath)
            
            if 'pack_size' not in self.raw_df.columns and 'sales_volume_units' in self.raw_df.columns and 'units_large' in self.raw_df.columns:
                self.raw_df['pack_size'] = np.where(self.raw_df['units_large'] > 0, 
                                                (self.raw_df['sales_volume_units'] / self.raw_df['units_large']).round(), 0).astype(int)
            
            # 1. 初始化国家 (自动剔除EEA以外国家)
            self.country_combo.clear()
            if 'market_region' in self.raw_df.columns:
                mkt_list = sorted(self.raw_df['market_region'].dropna().unique().tolist())
                for m in mkt_list:
                    if str(m).upper() in EEA_AND_UK_MARKETS:
                        self.country_combo.addCheckableItem(str(m), checked=True)
            
            # 2. 剂型智能合并机制 (Step 3 核心)
            self.dosage_combo.clear()
            if 'dosage_form' in self.raw_df.columns:
                # 模糊匹配聚类大组
                def categorize_form(x):
                    s = str(x).upper()
                    if '片' in s or 'TABLET' in s or 'TAB' in s: return '【片剂(含所有变种)】'
                    elif '胶囊' in s or 'CAPSULE' in s or 'CAP' in s: return '【胶囊剂(含所有变种)】'
                    elif '注射' in s or 'INJECTION' in s or 'VIAL' in s or 'AMP' in s: return '【注射剂】'
                    elif '口服液' in s or 'SYRUP' in s or 'SOLUTION' in s: return '【口服液/溶液】'
                    return '【其他剂型】'
                
                self.raw_df['dosage_form_grouped'] = self.raw_df['dosage_form'].apply(categorize_form)
                grp_dosages = sorted(self.raw_df['dosage_form_grouped'].dropna().unique().tolist())
                for d in grp_dosages:
                    self.dosage_combo.addCheckableItem(str(d), checked=True)
                    
            print(f"✅ 成功极速加载: {api_name} (行数:{len(self.raw_df)})")
            
        except Exception as e:
            QMessageBox.critical(self, "读取错误", f"无法加载缓存: {str(e)}")

    def run_analysis(self):
        if self.raw_df.empty: return
        
        # 1. 拦截基础条件过滤 (Filter Application)
        sel_countries = self.country_combo.get_checked_items()
        sel_dosages = self.dosage_combo.get_checked_items()
        
        df = self.raw_df.copy()
        if 'market_region' in df.columns and sel_countries:
            df = df[df['market_region'].isin(sel_countries)]
        if 'dosage_form_grouped' in df.columns and sel_dosages:
            df = df[df['dosage_form_grouped'].isin(sel_dosages)]
            
        if df.empty:
            QMessageBox.warning(self, "无数据", "在所选国家及剂型下，没有任何销售数据可用。")
            return
            
        self.filtered_df = df
        
        # 2. 更新视图
        self.render_dashboard()
        self.render_pack_analysis()
        self.render_others_analysis()
        self.render_strategy_analysis()
        
        QMessageBox.information(self, "执行完成", "聚合拆解数据测算完成！")
        
    def render_dashboard(self):
        df = self.filtered_df
        if df.empty: return
        
        req_cols = ['market_region', 'year', 'sales_volume_units', 'sales_value_usd', 'corporation_name', 'strength_raw', 'api_kg', 'api_name', 'dosage_form']
        missing = [c for c in req_cols if c not in df.columns]
        if missing: return
        
        # 确定最近的基准年份 (用于容量和对手计算)
        latest_year = df['year'].max()
        
        # 确定最早的年份 (用于CAGR)
        earliest_year = df['year'].min()
        year_diff = latest_year - earliest_year
        
        results = []
        
        for country, cdf in df.groupby('market_region'):
            # 基准年数据
            curr_df = cdf[cdf['year'] == latest_year]
            if curr_df.empty: continue
            
            # API信息
            molecule = str(curr_df['api_name'].iloc[0])
            formulation = "、".join(list(curr_df['dosage_form'].dropna().unique()))
            
            # 容量和金额
            vol_sum = curr_df['sales_volume_units'].sum()
            val_sum = curr_df['sales_value_usd'].sum()
            api_kg_sum = curr_df['api_kg'].sum()
            
            est_price = val_sum / vol_sum if vol_sum > 0 else 0
            
            # 规格偏好
            stren_grp = curr_df.groupby('strength_raw')['sales_volume_units'].sum().sort_values(ascending=False)
            all_str = "、".join([str(s) for s in stren_grp.index.tolist()[:4]])
            main_str = str(stren_grp.index[0]) if not stren_grp.empty else "-"
            
            # 公司竞争态势
            comp_grp = curr_df.groupby('corporation_name')['sales_volume_units'].sum().sort_values(ascending=False)
            opp_count = len(comp_grp)
            
            prominent_count = 0
            major_opps = []
            if vol_sum > 0:
                for comp, cvol in comp_grp.items():
                    share = cvol / vol_sum
                    if share >= 0.10:
                        prominent_count += 1
                        major_opps.append(str(comp))
                    
            if not major_opps:
                major_opps = [str(x) for x in comp_grp.head(3).index.tolist()]
                
            major_opps_str = "、".join(major_opps)
            
            # 5年 CAGR (销量)
            cagr_str = "-"
            if year_diff > 0:
                past_df = cdf[cdf['year'] == earliest_year]
                past_vol = past_df['sales_volume_units'].sum()
                if past_vol > 0 and vol_sum > 0:
                    cagr = ( (vol_sum / past_vol) ** (1/year_diff) ) - 1
                    cagr_str = f"{cagr:.2%}"
            elif year_diff == 0:
                cagr_str = "N/A"
                
            results.append({
                "country": country,
                "molecule": molecule,
                "form": formulation,
                "strengths": all_str,
                "vol_wan": vol_sum / 10000,
                "val_wan": val_sum / 10000,
                "price": est_price,
                "cagr": cagr_str,
                "opps": opp_count,
                "prominent": prominent_count,
                "major_opps": major_opps_str,
                "main_str": main_str,
                "api_kg": api_kg_sum,
                "vol_sum_raw": vol_sum # for sorting
            })
            
        res_df = pd.DataFrame(results).sort_values(by="vol_sum_raw", ascending=False)
        
        self.table_dashboard.setRowCount(0)
        for _, row in res_df.iterrows():
            r = self.table_dashboard.rowCount()
            self.table_dashboard.insertRow(r)
            
            self.table_dashboard.setItem(r, 0, QTableWidgetItem(row['country']))
            self.table_dashboard.setItem(r, 1, QTableWidgetItem(row['molecule']))
            self.table_dashboard.setItem(r, 2, QTableWidgetItem(row['form']))
            self.table_dashboard.setItem(r, 3, QTableWidgetItem(row['strengths']))
            
            vol_item = QTableWidgetItem(f"{row['vol_wan']:,.1f}")
            vol_item.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
            self.table_dashboard.setItem(r, 4, vol_item)
            
            val_item = QTableWidgetItem(f"{row['val_wan']:,.1f}")
            val_item.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
            self.table_dashboard.setItem(r, 5, val_item)
            
            px_item = QTableWidgetItem(f"${row['price']:,.4f}")
            px_item.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
            self.table_dashboard.setItem(r, 6, px_item)
            
            self.table_dashboard.setItem(r, 7, QTableWidgetItem(row['cagr']))
            self.table_dashboard.setItem(r, 8, QTableWidgetItem(str(row['opps'])))
            self.table_dashboard.setItem(r, 9, QTableWidgetItem(str(row['prominent'])))
            self.table_dashboard.setItem(r, 10, QTableWidgetItem(row['major_opps']))
            self.table_dashboard.setItem(r, 11, QTableWidgetItem(row['main_str']))
            
            api_item = QTableWidgetItem(f"{row['api_kg']:,.1f}")
            api_item.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
            self.table_dashboard.setItem(r, 12, api_item)

    def show_copy_menu(self, pos, canvas_widget):
        menu = QMenu(self)
        copy_action = menu.addAction("📋 复制图表到剪贴板")
        action = menu.exec(canvas_widget.mapToGlobal(pos))
        if action == copy_action:
            self.copy_canvas_to_clipboard(canvas_widget)
            
    def copy_canvas_to_clipboard(self, canvas_widget):
        try:
            pixmap = canvas_widget.grab()
            QApplication.clipboard().setPixmap(pixmap)
            # 临时把主窗口状态栏或者随便弹个提示
            QMessageBox.information(self, "复制成功", "图表已复制到剪贴板！可以直接在微信/PPT里粘贴。")
        except Exception as e:
            QMessageBox.warning(self, "复制失败", str(e))
            
    def copy_table_to_clipboard(self):
        table = self.table_strategy
        if table.rowCount() == 0:
            return
            
        header = []
        for i in range(table.columnCount()):
            header.append(table.horizontalHeaderItem(i).text())
            
        rows = ["\t".join(header)]
        for r in range(table.rowCount()):
            row_data = []
            for c in range(table.columnCount()):
                item = table.item(r, c)
                # 提取数字（如果有用千分位，保留，Excel可以直接贴，去掉也可以，这里原样保留，但Excel直接贴有时需要处理一下）
                # 这里去除 $ 和 逗号，方便Excel直接计算
                text = item.text() if item else ""
                text = text.replace('$', '').replace(',', '')
                row_data.append(text)
            rows.append("\t".join(row_data))
            
        clip_text = "\n".join(rows)
        QApplication.clipboard().setText(clip_text)
        QMessageBox.information(self, "复制成功", "表格数据已复制到剪贴板！可以直接在Excel里粘贴并计算成本！")

    def _build_tree(self, df_part, tree_widget, show_price=False):
        tree_widget.clear()
        if df_part.empty: return
        
        total_vol = df_part['sales_volume_units'].sum()
        if total_vol <= 0: return
        
        # 第1层: 规格
        grp_str = df_part.groupby('strength_raw')
        str_vols = grp_str['sales_volume_units'].sum().sort_values(ascending=False)
        
        for st, svol in str_vols.items():
            if svol <= 0: continue
            spct = svol / total_vol
            
            st_item = QTreeWidgetItem(tree_widget)
            st_item.setText(0, f"💊 {st}")
            st_item.setText(1, f"{svol:,.0f}")
            st_item.setText(2, f"{spct:.1%}")
            # 规格层级折叠还是展开看体验，默认展开
            st_item.setExpanded(True)
            
            # 第2层: 包装
            sub_df = grp_str.get_group(st)
            grp_pack = sub_df.groupby('pack_size')
            pack_vols = grp_pack['sales_volume_units'].sum().sort_values(ascending=False)
            
            for pk, pvol in pack_vols.items():
                if pvol <= 0: continue
                ppct = pvol / svol # 包装占该规格的比例 (也可以算占总盘的比例，这里占该规格的有助于阅读)
                
                pk_item = QTreeWidgetItem(st_item)
                pk_item.setText(0, f"📦 {pk} 装")
                pk_item.setText(1, f"{pvol:,.0f}")
                pk_item.setText(2, f"{ppct:.1%} (同规格占比)")
                
                if show_price and 'sales_value_usd' in sub_df.columns:
                    pack_df = grp_pack.get_group(pk)
                    pval = pack_df['sales_value_usd'].sum()
                    price = pval / pvol if pvol > 0 else 0
                    pk_item.setText(3, f"${price:,.4f}")

    def render_pack_analysis(self):
        df = self.filtered_df
        if 'strength_raw' not in df.columns or 'pack_size' not in df.columns or 'sales_volume_units' not in df.columns:
            return
            
        df_valid = df[df['sales_volume_units'] > 0].copy()
        df_valid['pack_size'] = pd.to_numeric(df_valid['pack_size'], errors='coerce').fillna(0).astype(int)
        df_valid['strength_raw'] = df_valid['strength_raw'].astype(str).str.strip().replace('nan', '未知规格')
        self._pack_valid_df = df_valid
        
        # A. 全盘架构渲染
        self._build_tree(df_valid, self.tree_global)
        
        # A2. 全盘概览可视化 (Sunburst/饼图替代品)
        self.fig_pack.clear()
        ax1 = self.fig_pack.add_subplot(121)
        ax2 = self.fig_pack.add_subplot(122)
        
        # 规格饼图
        str_vols = df_valid.groupby('strength_raw')['sales_volume_units'].sum().sort_values(ascending=False)
        total_vol = str_vols.sum()
        
        # 为了不让饼图太碎，保留前5个，其余合并为Others
        if len(str_vols) > 6:
            top_str = str_vols.iloc[:5]
            others_str_vol = str_vols.iloc[5:].sum()
            str_plot = pd.concat([top_str, pd.Series({'其他规格': others_str_vol})])
        else:
            str_plot = str_vols
            
        wedges1, texts1, autotexts1 = ax1.pie(str_plot, labels=str_plot.index, autopct='%1.1f%%',
                                              textprops=dict(color="black", fontsize=9),
                                              colors=plt.cm.Set3.colors)
        ax1.set_title("全体国家 - 各规格 (Strength) 销量占比", pad=15, fontweight='bold')
        
        # 主力规格的包装拆分环形图 (选取最大的规格作为代表演示，或展示整体包装分布)
        main_str = str_vols.index[0] if not str_vols.empty else ""
        sub_df = df_valid[df_valid['strength_raw'] == main_str]
        
        title2 = f"全体国家 - 总体各包装组合结构占比"
        pack_vols = df_valid.groupby(['strength_raw', 'pack_size'])['sales_volume_units'].sum().sort_values(ascending=False)
        
        if len(pack_vols) > 7:
            top_pack = pack_vols.iloc[:6]
            others_pack_vol = pack_vols.iloc[6:].sum()
            pack_plot = pd.concat([top_pack, pd.Series({('其他', 0): others_pack_vol})])
        else:
            pack_plot = pack_vols
            
        labels2 = [f"{idx[0]} - {idx[1]}装" if idx[0] != '其他' else "其他组合" for idx in pack_plot.index]
        
        wedges2, texts2, autotexts2 = ax2.pie(pack_plot, labels=labels2, autopct='%1.1f%%',
                                              textprops=dict(color="black", fontsize=9),
                                              wedgeprops=dict(width=0.4, edgecolor='w'), # 环形图
                                              colors=plt.cm.Pastel1.colors)
        ax2.set_title(title2, pad=15, fontweight='bold')
        
        self.fig_pack.tight_layout()
        self.canvas_pack.draw()
        
        # B. 国家列表渲染
        self.table_pack_countries.setRowCount(0)
        self.tree_country_detail.clear()
        self.table_pack_comps.setRowCount(0)
        self.tree_comp_detail.clear()
        
        if 'market_region' in df_valid.columns:
            total_global = df_valid['sales_volume_units'].sum()
            c_grp = df_valid.groupby('market_region')['sales_volume_units'].sum().sort_values(ascending=False)
            
            for ctry, cvol in c_grp.items():
                r = self.table_pack_countries.rowCount()
                self.table_pack_countries.insertRow(r)
                
                self.table_pack_countries.setItem(r, 0, QTableWidgetItem(str(ctry)))
                
                v_item = QTableWidgetItem(f"{cvol:,.0f}")
                v_item.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
                self.table_pack_countries.setItem(r, 1, v_item)
                
                cpct = cvol / total_global if total_global > 0 else 0
                self.table_pack_countries.setItem(r, 2, QTableWidgetItem(f"{cpct:.1%}"))
                
    def on_pack_country_clicked(self, row, col):
        ctry_item = self.table_pack_countries.item(row, 0)
        if not ctry_item: return
        self._curr_pack_country = ctry_item.text()
        
        df = self._pack_valid_df
        c_df = df[df['market_region'] == self._curr_pack_country]
        
        # 渲染选中国家的树
        self._build_tree(c_df, self.tree_country_detail)
        
        # 激活刷新下级的公司列表
        self.refresh_pack_companies()
        
    def refresh_pack_companies(self):
        if not hasattr(self, '_curr_pack_country') or not self._curr_pack_country: return
        
        df = self._pack_valid_df
        c_df = df[df['market_region'] == self._curr_pack_country]
        if c_df.empty or 'corporation_name' not in c_df.columns: return
        
        total_ctry = c_df['sales_volume_units'].sum()
        comp_vols = c_df.groupby('corporation_name')['sales_volume_units'].sum().sort_values(ascending=False)
        
        only_others = self.radio_comp_others.isChecked()
        if only_others and len(comp_vols) > 5:
            target_comps = comp_vols.iloc[5:]
        else:
            target_comps = comp_vols
            
        self.table_pack_comps.setRowCount(0)
        self.tree_comp_detail.clear()
        
        for comp, pvol in target_comps.items():
            r = self.table_pack_comps.rowCount()
            self.table_pack_comps.insertRow(r)
            
            self.table_pack_comps.setItem(r, 0, QTableWidgetItem(str(comp)))
            
            v_item = QTableWidgetItem(f"{pvol:,.0f}")
            v_item.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
            self.table_pack_comps.setItem(r, 1, v_item)
            
            ppct = pvol / total_ctry if total_ctry > 0 else 0
            self.table_pack_comps.setItem(r, 2, QTableWidgetItem(f"{ppct:.1%}"))
            
    def on_pack_comp_clicked(self, row, col):
        comp_item = self.table_pack_comps.item(row, 0)
        if not comp_item or not hasattr(self, '_curr_pack_country'): return
        comp = comp_item.text()
        
        df = self._pack_valid_df
        cc_df = df[(df['market_region'] == self._curr_pack_country) & (df['corporation_name'] == comp)]
        
        # 渲染选中公司的树 (附带均价)
        self._build_tree(cc_df, self.tree_comp_detail, show_price=True)
            
    def render_others_analysis(self):
        df = self.filtered_df
        if 'market_region' not in df.columns or 'corporation_name' not in df.columns or 'sales_volume_units' not in df.columns:
            return
            
        # 分国家计算“Others”矩阵 (Step 5 核心)
        results = []
        for country, vdf in df.groupby('market_region'):
            vol_sum = vdf['sales_volume_units'].sum()
            if vol_sum <= 0: continue
            
            # 找到各大公司销量
            c_grp = vdf.groupby('corporation_name')['sales_volume_units'].sum().sort_values(ascending=False)
            
            # 前5公司为巨头
            top5_vol = c_grp.head(5).sum()
            others_sr = c_grp.iloc[5:]
            others_vol = others_sr.sum()
            others_count = len(others_sr)
            
            # 如果没 others 就设0
            others_pct = others_vol / vol_sum if vol_sum > 0 else 0
            
            results.append({
                'Country': country,
                'TotalVol': vol_sum,
                'OthersVol': others_vol,
                'OthersPct': others_pct,
                'OthersCount': others_count
            })
            
        res_df = pd.DataFrame(results).sort_values(by='TotalVol', ascending=False)
        
        # 填充表格
        self.table_others.setRowCount(0)
        for idx, row in res_df.iterrows():
            r = self.table_others.rowCount()
            self.table_others.insertRow(r)
            
            self.table_others.setItem(r, 0, QTableWidgetItem(row['Country']))
            vol_item = QTableWidgetItem(f"{row['TotalVol']:,.0f}")
            vol_item.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
            self.table_others.setItem(r, 1, vol_item)
            
            ovol_item = QTableWidgetItem(f"{row['OthersVol']:,.0f}")
            ovol_item.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
            self.table_others.setItem(r, 2, ovol_item)
            
            self.table_others.setItem(r, 3, QTableWidgetItem(f"{row['OthersPct']:.1%}"))
            self.table_others.setItem(r, 4, QTableWidgetItem(f"{row['OthersCount']:.0f} 家"))
            
        self.table_others.resizeColumnsToContents()

    def render_strategy_analysis(self):
        df = self.filtered_df
        if df.empty: return
        
        req_cols = ['market_region', 'year', 'sales_volume_units', 'sales_value_usd', 'strength_raw', 'pack_size']
        if not all(c in df.columns for c in req_cols): return
        
        max_pack = self.spin_max_pack.value()
        max_str_mg = self.spin_max_str.value()
        pen_y1 = self.spin_pen_y1.value() / 100.0
        pen_y2 = self.spin_pen_y2.value() / 100.0
        pen_y3 = self.spin_pen_y3.value() / 100.0
        
        latest_year = df['year'].max()
        curr_df = df[df['year'] == latest_year].copy()
        if curr_df.empty: return
        
        # 1. 解析数字化的规格 (从 strength_raw 里提纯 mg 数)
        # 用正则提取第一个出现的数字或小数作为该药的 mg 含量
        curr_df['strength_num'] = curr_df['strength_raw'].astype(str).str.extract(r'(\d+(?:\.\d+)?)').astype(float)
        # 提取不到默认填0
        curr_df['strength_num'] = curr_df['strength_num'].fillna(0)
        
        # 2. 确保 pack_size 存在且为数字
        curr_df['pack_size'] = pd.to_numeric(curr_df['pack_size'], errors='coerce').fillna(0).astype(int)
        
        results = []
        for country, cdf in curr_df.groupby('market_region'):
            vol_sum_country = cdf['sales_volume_units'].sum()
            if vol_sum_country <= 0: continue
            
            # 过滤出符合我们 包装能力 AND 浓度能力 的细分市场 (即可达小盘)
            # 只有这部分才是我们真正能吃到渗透率的“局域总盘”
            feasible_df = cdf[
                (cdf['pack_size'] <= max_pack) & 
                (cdf['pack_size'] > 0) & 
                (cdf['strength_num'] <= max_str_mg)
            ]
            
            feasible_vol = feasible_df['sales_volume_units'].sum()
            if feasible_vol <= 0: continue
            
            # 找出我们能做的规格组合里，销量最大的（即最优解推荐使用）
            pack_grp = feasible_df.groupby(['strength_raw', 'pack_size'])['sales_volume_units'].sum().sort_values(ascending=False)
            
            top_combos = pack_grp.head(2)
            combo_strs = []
            for (st, pk), p_vol in top_combos.items():
                combo_strs.append(f"{pk}装")
                
            main_str = top_combos.index[0][0] # 最大的那个规格
            best_packs = "、".join(combo_strs)
            
            # 收益计算: 我们在“可达小盘(Feasible Vol)”身上吃到三个阶段预期的渗透率
            vol_y1 = feasible_vol * pen_y1
            vol_y2 = feasible_vol * pen_y2
            vol_y3 = feasible_vol * pen_y3
            
            # 均价仅作为参考 (老板主要看量来算成本)
            f_val = feasible_df['sales_value_usd'].sum()
            feasible_price = f_val / feasible_vol if feasible_vol > 0 else 0
            
            results.append({
                "country": country,
                "main_str": main_str,
                "best_packs": best_packs,
                "feasible_vol": feasible_vol,
                "vol_y1": vol_y1,
                "vol_y2": vol_y2,
                "vol_y3": vol_y3,
                "price": feasible_price
            })
            
        res_df = pd.DataFrame(results)
        if res_df.empty:
            self.table_strategy.setRowCount(0)
            self.fig_strategy.clear()
            self.fig_strategy.text(0.5, 0.5, "当前限制条件下无符合的抛放市场", ha='center', va='center')
            self.canvas_strategy.draw()
            return

        # 排序：按照 Y3 终极起量最大的排序
        res_df = res_df.sort_values(by="vol_y3", ascending=False)
        
        self.table_strategy.setRowCount(0)
        for _, row in res_df.iterrows():
            r = self.table_strategy.rowCount()
            self.table_strategy.insertRow(r)
            
            self.table_strategy.setItem(r, 0, QTableWidgetItem(row['country']))
            self.table_strategy.setItem(r, 1, QTableWidgetItem(row['main_str']))
            self.table_strategy.setItem(r, 2, QTableWidgetItem(row['best_packs']))
            
            fv_item = QTableWidgetItem(f"{row['feasible_vol']:,.0f}")
            fv_item.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
            self.table_strategy.setItem(r, 3, fv_item)
            
            vy1 = QTableWidgetItem(f"{row['vol_y1']:,.0f}")
            vy1.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
            self.table_strategy.setItem(r, 4, vy1)
            
            vy2 = QTableWidgetItem(f"{row['vol_y2']:,.0f}")
            vy2.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
            self.table_strategy.setItem(r, 5, vy2)
            
            vy3 = QTableWidgetItem(f"{row['vol_y3']:,.0f}")
            vy3.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
            self.table_strategy.setItem(r, 6, vy3)
            
            px_item = QTableWidgetItem(f"${row['price']:,.4f}")
            px_item.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
            self.table_strategy.setItem(r, 7, px_item)
            
        # 绘制终局体积总结横向条形图 (Y3 的量)
        self.fig_strategy.clear()
        
        if not res_df.empty:
            ax = self.fig_strategy.add_subplot(111)
            plot_df = res_df.head(10).sort_values(by="vol_y3", ascending=True) # 条形图倒序画
            
            bars = ax.barh(plot_df['country'], plot_df['vol_y3'], color='#48bb78', height=0.6)
            
            # 在图上标出销量
            for bar in bars:
                width = bar.get_width()
                label_x_pos = width + (res_df['vol_y3'].max() * 0.01)
                ax.text(label_x_pos, bar.get_y() + bar.get_height()/2, f"{width:,.0f} Units", 
                        va='center', fontsize=9, fontweight='bold', color='#2d3748')
                
            ax.set_title(f"TOP 10 核心国家目标达成阶段起盘量 (Y3 渗透率: {self.spin_pen_y3.value()}%)", pad=10, fontweight='bold')
            ax.set_xlabel("Y3 目标起量 (Units)")
            ax.margins(x=0.15) # 腾一点右边距写字
            
            # 隐藏上下右边框
            ax.spines['top'].set_visible(False)
            ax.spines['right'].set_visible(False)
            ax.spines['bottom'].set_alpha(0.3)
            ax.spines['left'].set_alpha(0.3)
        
        self.fig_strategy.tight_layout()
        self.canvas_strategy.draw()

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = ForecastApp()
    window.show()
    sys.exit(app.exec())
