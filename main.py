import sys
import os

from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, 
    QLabel, QPushButton, QStackedWidget, QPlainTextEdit, QProgressBar,
    QSpacerItem, QSizePolicy, QFrame, QLineEdit, QFormLayout,
    QTabWidget, QScrollArea, QComboBox, QMessageBox, QFileDialog
)
from PySide6.QtCore import Qt, QThread, Signal, QSize
from PySide6.QtGui import QIcon, QFont, QPixmap, QPalette, QColor

# ======= 导入底层处理脚本 =======
import step_a_download
import step_b_clean
from step3_standardize import StandardizationEngine
from step4_visualizer import Step4DashboardWidget, CheckableComboBox, ClipboardHelper
from flexible_pivot import FlexiblePivotWidget
from step4_visualizer import AnalysisEngineV24
import core_config
from matplotlib.backends.backend_qtagg import FigureCanvasQTAgg as FigureCanvas
import matplotlib.pyplot as plt
import pandas as pd
import numpy as np

# ======= 1. 后台通用 Worker 线程定义 =======
class Worker(QThread):
    progress_signal = Signal(int, int)
    log_signal = Signal(str)
    finished_signal = Signal(bool, str) # success, result_data

    def __init__(self, task_func, *args, **kwargs):
        super().__init__()
        self.task_func = task_func
        self.args = args
        self.kwargs = kwargs

    def run(self):
        try:
            # Inject dynamic callbacks
            def cb_log(msg):
                self.log_signal.emit(str(msg))
            def cb_progress(curr, total):
                self.progress_signal.emit(curr, total)
                
            self.kwargs['log_callback'] = cb_log
            if 'progress_callback' in self.kwargs.get('_inject_keys', []):
                self.kwargs['progress_callback'] = cb_progress
                
            if '_inject_keys' in self.kwargs:
                del self.kwargs['_inject_keys']

            res = self.task_func(*self.args, **self.kwargs)
            
            # Unpack results if tuple
            if isinstance(res, tuple) and len(res) == 2:
                success, data = res
            else:
                success = bool(res)
                data = ""
            self.finished_signal.emit(success, str(data))
        except Exception as e:
            self.log_signal.emit(f"[-] 执行过程中发生未捕获异常: {e}")
            self.finished_signal.emit(False, str(e))


# ======= 2. 基础步骤页面设计 =======
class BaseStepWidget(QWidget):
    """通用的步骤容器：包含居中执行按钮、进度条、黑暗极客风日志区"""
    completed_signal = Signal(str) # 向外发送执行结果
    
    def __init__(self, step_title, parent=None):
        super().__init__(parent)
        self.step_title = step_title
        self.setup_ui()
        
    def setup_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(40, 40, 40, 40)
        layout.setSpacing(20)
        
        # 页面标题
        title = QLabel(self.step_title)
        title.setStyleSheet("font-size: 24px; font-weight: bold; color: #1a365d;")
        layout.addWidget(title)
        
        # 参数输入区（让子类拓展）
        self.param_layout = QFormLayout()
        layout.addLayout(self.param_layout)
        
        # 居中的执行按钮
        btn_layout = QHBoxLayout()
        self.run_btn = QPushButton("▶ 开始执行本步骤")
        self.run_btn.setFixedSize(200, 45)
        self.run_btn.setCursor(Qt.PointingHandCursor)
        self.run_btn.setStyleSheet("""
            QPushButton {
                background-color: #2b6cb0;
                color: white;
                font-size: 15px;
                font-weight: bold;
                border-radius: 6px;
                border: none;
            }
            QPushButton:hover { background-color: #2c5282; }
            QPushButton:disabled { background-color: #a0aec0; color: #e2e8f0; }
        """)
        self.run_btn.clicked.connect(self.execute)
        btn_layout.addStretch()
        btn_layout.addWidget(self.run_btn)
        btn_layout.addStretch()
        layout.addLayout(btn_layout)
        
        # 进度条
        self.progress_bar = QProgressBar()
        self.progress_bar.setFixedHeight(15)
        self.progress_bar.setTextVisible(False)
        self.progress_bar.setStyleSheet("""
            QProgressBar {
                background-color: #e2e8f0;
                border-radius: 7px;
            }
            QProgressBar::chunk {
                background-color: #38b2ac;
                border-radius: 7px;
            }
        """)
        self.progress_bar.hide()
        layout.addWidget(self.progress_bar)
        
        # 黑暗主题控制台日志区
        self.console = QPlainTextEdit()
        self.console.setReadOnly(True)
        self.console.setStyleSheet("""
            QPlainTextEdit {
                background-color: #1a202c;
                color: #a0aec0;
                font-family: 'Consolas', 'Courier New', monospace;
                font-size: 13px;
                border-radius: 8px;
                padding: 10px;
                line-height: 1.5;
            }
        """)
        layout.addWidget(self.console)
        
        self.worker = None

    def append_log(self, text):
        if "[+]" in text or "[+++]" in text or "成功" in text:
            color = "#48bb78" # Green
        elif "[-]" in text or "失败" in text or "错误" in text:
            color = "#f56565" # Red
        elif "[*]" in text or ">>>" in text:
            color = "#4299e1" # Blue
        else:
            color = "#a0aec0" # Default Gray
            
        html = f"<span style='color:{color};'>{text}</span>"
        self.console.appendHtml(html)
        
    def update_progress(self, curr, total):
        self.progress_bar.show()
        if total > 0:
            self.progress_bar.setMaximum(total)
            self.progress_bar.setValue(curr)
            
    def on_finished(self, success, data):
        self.run_btn.setEnabled(True)
        self.progress_bar.hide()
        if success:
            self.append_log(f"<br/><span style='color:#38a169;font-weight:bold;'>✅ {self.step_title} 处理完成，请进入下一步。</span><br/>")
            self.completed_signal.emit(data)
        else:
            self.append_log(f"<br/><span style='color:#e53e3e;font-weight:bold;'>❌ {self.step_title} 处理失败，请检查日志。</span><br/>")

    def execute(self):
        # 交给子类具体触发 worker 逻辑
        pass


class Step1Widget(BaseStepWidget):
    def __init__(self):
        super().__init__("Step 1: 外部数据下载聚合 (TSM API)")
        self.drugs_input = QPlainTextEdit("Irbesartan Ranolazine Clopidogrel Celecoxib Lamotrigine Topiramate Duloxetine Citalopram Bisoprolol Ticagrelor Memantine Rivaroxaban Tadalafil")
        self.drugs_input.setMinimumHeight(60)
        self.drugs_input.setMaximumHeight(150)
        self.years_input = QLineEdit("2020-2025")
        self.drugs_input.setStyleSheet("padding: 5px;")
        self.years_input.setStyleSheet("padding: 5px;")
        self.param_layout.addRow("目标药物名 (空格/换行分隔):", self.drugs_input)
        self.param_layout.addRow("统计年份区间 (如 2022-2025):", self.years_input)

    def execute(self):
        self.run_btn.setEnabled(False)
        self.console.clear()
        self.append_log("[*] 正在扫描本地数据缓存...")
        
        drugs_text = self.drugs_input.toPlainText().replace('\n', ' ')
        time_period = self.years_input.text()
        
        # 为了不卡死UI，我们使用一个独立的 Worker 来进行纯本地扫描
        # 这里借助我们刚刚在 step_a_download 抽离解析逻辑和寻找文件逻辑，但为了避免重写大量代码并保持主结构，
        # 我们采取一个巧妙的方法：让 download_and_aggregate_tsm 增加一个 check_only 模式或者暴露 check 函数
        # 不过既然咱们在 step_a 已经完善了隔离，只需把 download_and_aggregate_tsm 直接用来执行
        
        # 定义一个专门用来做预扫描或真正执行的高级槽函数
        def _pre_scan_and_ask():
            # 我们在主线程快速过一遍缺失项，因为只有 os.path.exists 检测，速度非常快（毫秒级）
            # 所以直接在主线程稍微计算一下 missing 即可，不会卡死 UI
            import os
            from step_a_download import get_parsed_drug_names, parse_years
            
            parsed_names = get_parsed_drug_names(drugs_text)
            years_to_fetch = parse_years(time_period)
            if not parsed_names or not years_to_fetch:
                 self.append_log("[-] 输入无效，无法解析药名或年份。")
                 self.run_btn.setEnabled(True)
                 return
                 
            BASE_DIR = os.path.dirname(os.path.abspath(__file__))
            output_dir = os.path.abspath(os.path.join(BASE_DIR, "TSM_Downloads", "raw_files"))
            
            endpoints_to_try = ["MIDS", "ATC"]
            local_files = []
            missing_tasks = []
            
            for drug_name in parsed_names:
                for current_year in years_to_fetch:
                    found_cache = False
                    for ep in endpoints_to_try:
                        filename = f"TSM_{current_year}_{drug_name}_{ep}.xlsx"
                        filepath = os.path.join(output_dir, filename)
                        if os.path.exists(filepath) and os.path.getsize(filepath) > 1024:
                            local_files.append((filepath, drug_name))
                            found_cache = True
                            break
                    if not found_cache:
                        missing_tasks.append((drug_name, current_year))
            
            skip_downloads = False
            if missing_tasks and local_files:
                # 混杂情况：部分有，部分没有
                reply = QMessageBox.question(
                    self, 
                    "发现本地数据", 
                    f"扫描到您请求的数据中，本地已有 {len(local_files)} 个缓存，但缺失 {len(missing_tasks)} 个任务。\n\n"
                    f"选择 [Yes] 去下载缺失的数据（可能较慢）。\n"
                    f"选择 [No] 放弃下载缺失项，立刻仅凭借现有的本地数据出表（极速）。",
                    QMessageBox.Yes | QMessageBox.No,
                    QMessageBox.Yes
                )
                if reply == QMessageBox.No:
                    skip_downloads = True
            elif missing_tasks and not local_files:
                self.append_log("[*] 本地无对应缓存，必须请求云端下载...")
            elif local_files and not missing_tasks:
                self.append_log("[*] 所有数据均在本地，开始极速提取...")
                
            # 设置 flag 给底层 worker
            # 我们通过 Python 鸭子类型给模块函数动态打个标记
            step_a_download.download_and_aggregate_tsm.skip_downloads = skip_downloads

            self.append_log("[*] 初始化任务队列...")
            func = step_a_download.download_and_aggregate_tsm
            self.worker = Worker(
                func, 
                drug_names=drugs_text, 
                time_period=time_period
            )
            self.worker.log_signal.connect(self.append_log)
            self.worker.finished_signal.connect(self.on_finished)
            self.worker.start()

        _pre_scan_and_ask()


class Step2Widget(BaseStepWidget):
    def __init__(self):
        super().__init__("Step 2: 规格剂型提取与维度清洗")
        self.input_file_path = ""
        
    def set_input(self, path):
        self.input_file_path = path
        
    def execute(self):
        if not self.input_file_path:
            base_dir = os.path.dirname(os.path.abspath(__file__))
            cache_file_csv = os.path.join(base_dir, "Cache", "step1_latest.csv")
            cache_file_xlsx = os.path.join(base_dir, "Cache", "step1_latest.xlsx")
            
            if os.path.exists(cache_file_csv):
                self.input_file_path = os.path.abspath(cache_file_csv)
                self.append_log("[*] 未捕获到直接下发流，检测到 Step 1 历史缓存 (CSV)，将直接使用...")
            elif os.path.exists(cache_file_xlsx):
                self.input_file_path = os.path.abspath(cache_file_xlsx)
                self.append_log("[*] 未捕获到直接下发流，检测到 Step 1 历史缓存 (Excel)，将直接使用...")
            else:
                self.append_log("[-] 错误：没有检测到有效的输入文件，且无本地历史缓存 (Cache/step1_latest.csv 或 .xlsx)，请先执行 Step 1！")
                return
            
        self.run_btn.setEnabled(False)
        self.console.clear()
        self.append_log(f"[*] 解析前置 Excel 数据源: {self.input_file_path}")
        
        func = step_b_clean.clean_and_cache_data
        self.worker = Worker(func, input_excel_path=self.input_file_path)
        self.worker.log_signal.connect(self.append_log)
        self.worker.finished_signal.connect(self.on_finished)
        self.worker.start()


class Step3Widget(BaseStepWidget):
    def __init__(self):
        super().__init__("Step 3: 业务列标准化与拆分缓存")
        self.input_parquet_path = ""
        
    def set_input(self, path):
        self.input_parquet_path = path

    def execute(self):
        if not self.input_parquet_path:
            base_dir = os.path.dirname(os.path.abspath(__file__))
            cache_file = os.path.join(base_dir, "Cache", "step1_latest.parquet")
            if os.path.exists(cache_file):
                self.input_parquet_path = os.path.abspath(cache_file)
                self.append_log("[*] 未捕获到直接下发流，检测到 Step 2 历史标准化模型，将直接使用...")
            else:
                self.append_log("[-] 错误：找不到前置 Parquet 缓存 (Cache/step1_latest.parquet)，请确认 Step 2 是否成功。")
                return
            
        self.run_btn.setEnabled(False)
        self.console.clear()
        self.append_log(f"[*] 加载 Parquet 核心库: {self.input_parquet_path}")
        
        engine = StandardizationEngine(input_path=self.input_parquet_path)
        func = engine.execute_standardization
        
        self.worker = Worker(func, _inject_keys=['progress_callback'])
        self.worker.log_signal.connect(self.append_log)
        self.worker.progress_signal.connect(self.update_progress)
        
        # Custom finish capture because step3 returns bool without string
        def handle_fin(success, _):
            self.on_finished(success, engine.output_dir)
            
        self.worker.finished_signal.connect(handle_fin)
        self.worker.start()


class EuropeanAnalysisPage(QWidget):
    """第5步：欧洲市场独占深度分析页面，复用【欧洲市场预测promax】的分析能力"""
    def __init__(self, parent=None):
        super().__init__(parent)
        self.analyzer = AnalysisEngineV24()
        self.df_raw = None
        self.current_api = None
        
        self.setup_ui()

    def setup_ui(self):
        self.layout = QVBoxLayout(self)
        self.layout.setContentsMargins(15, 15, 15, 15)
        
        # 顶栏控制
        top_bar = QHBoxLayout()
        self.api_combo = QComboBox()
        self.api_combo.setMinimumWidth(200)
        self.api_combo.currentTextChanged.connect(self.on_api_changed)
        
        self.originator_combo = CheckableComboBox()
        self.originator_combo.setMinimumWidth(250)
        self.originator_combo.lineEdit().setPlaceholderText("可下拉勾选原研企业...")
        # 禁用搜索功能，避免每次键入触发showPopup导致乱窜
        self.originator_combo.lineEdit().setReadOnly(True)
        
        orig_row = QHBoxLayout()
        orig_row.setContentsMargins(0, 0, 0, 0)
        orig_row.addWidget(self.originator_combo)
        
        btn_all = QPushButton("全选")
        btn_clear = QPushButton("清空")
        btn_all.setFixedWidth(40)
        btn_clear.setFixedWidth(40)
        style = "QPushButton { padding: 4px; font-size: 11px; background-color: #cbd5e0; color: #2d3748; border-radius: 3px; font-weight: normal; } QPushButton:hover { background-color: #a0aec0; }"
        btn_all.setStyleSheet(style)
        btn_clear.setStyleSheet(style)
        btn_all.clicked.connect(self.originator_combo.check_all)
        btn_clear.clicked.connect(self.originator_combo.uncheck_all)
        orig_row.addWidget(btn_all)
        orig_row.addWidget(btn_clear)
        orig_container = QWidget()
        orig_container.setLayout(orig_row)
        
        self.run_btn = QPushButton("▶ 执行深度预测")
        self.run_btn.setStyleSheet("background-color: #2b6cb0; color: white; padding: 6px 15px; border-radius: 4px; font-weight: bold;")
        self.run_btn.clicked.connect(self.run_analysis)
        self.run_btn.setEnabled(False)
        
        self.tips_label = QLabel("提示: 先选择 API，可按需勾选锁定部分企业为原研")
        self.tips_label.setStyleSheet("color: #718096; font-style: italic;")

        top_bar.addWidget(QLabel("选择通用名 (API):"))
        top_bar.addWidget(self.api_combo)
        top_bar.addWidget(QLabel("标记原研:"))
        top_bar.addWidget(orig_container)
        top_bar.addWidget(self.run_btn)
        top_bar.addWidget(self.tips_label)
        top_bar.addStretch()
        self.layout.addLayout(top_bar)
        
        # Tabs 面板
        self.tabs = QTabWidget()
        self.tabs.setStyleSheet("""
            QTabBar::tab { padding: 8px 15px; font-weight: bold; }
            QTabBar::tab:selected { background-color: #3182ce; color: white; }
        """)
        self.layout.addWidget(self.tabs)
        
        self.tab_names = [
            "0. 战略决策 (Global)", "1. 趋势分析 (Trend)", "2. 原研渗透率", 
            "3. 全市场格局", "4. 机会矩阵", "5. 智能预测"
        ]
        self.tab_widgets = {}
        for name in self.tab_names:
            w = QWidget()
            w.setLayout(QVBoxLayout())
            self.tabs.addTab(w, name)
            self.tab_widgets[name] = w

    def set_data(self, df: pd.DataFrame):
        self.df_raw = df
        if not df.empty and 'api_name' in df.columns:
            apis = sorted(df['api_name'].dropna().unique().tolist())
            self.api_combo.blockSignals(True)
            self.api_combo.clear()
            self.api_combo.addItems([str(a) for a in apis])
            self.api_combo.blockSignals(False)
            if apis:
                self.api_combo.setCurrentIndex(0)
                self.on_api_changed(apis[0])

    def on_api_changed(self, api_name):
        self.current_api = api_name
        self.run_btn.setEnabled(bool(api_name))
        
        # 动态更新原研企业下拉框
        if not api_name or self.df_raw is None:
            self.originator_combo.clear()
            return
            
        df_api = self.df_raw[self.df_raw['api_name'] == api_name]
        if 'corporation_name' not in df_api.columns or 'sales_value_usd' not in df_api.columns or 'sales_volume_units' not in df_api.columns:
            return
            
        # 计算出厂均价(作为参考保留)，但下拉菜单按企业名称首字母A-Z自动排序
        agg = df_api.groupby('corporation_name').agg(
            Sales=('sales_value_usd', 'sum'),
            Vol=('sales_volume_units', 'sum')
        )
        agg['Price'] = (agg['Sales'] / agg['Vol'].replace({0: np.nan}) * 0.3).fillna(0).round(4)
        sorted_comps = sorted([str(c) for c in agg.index if pd.notna(c)])
        
        self.originator_combo.blockSignals(True)
        self.originator_combo.clear()
        
        import core_config
        
        # 自动根据 ORIGINATOR_CONFIG 匹配已知原研企业
        orig_keywords = core_config.ORIGINATOR_CONFIG.get(str(api_name).upper(), [])
        
        for c in sorted_comps:
            # 检查该企业是否匹配原研关键词
            is_orig = any(kw.upper() in str(c).upper() for kw in orig_keywords)
            self.originator_combo.addCheckableItem(str(c), checked=is_orig)
            
        self.originator_combo.updateText()
        self.originator_combo.blockSignals(False)

    def run_analysis(self):
        if not self.current_api or self.df_raw is None: return
        
        self.tips_label.setText("分析中...")
        QApplication.processEvents()
        
        # 1. 过滤当前 API 数据
        df_api = self.df_raw[self.df_raw['api_name'] == self.current_api].copy()
        
        # 2. 映射到 Promax 所需字段 (国家, 集团/企业, Sales_Clean, Volume_Clean, Factory_Price, 年份)
        mapper = {
            'market_region': '国家',
            'corporation_name': '集团/企业',
            'year': '年份',
            'sales_value_usd': 'Sales_Clean',
            'sales_volume_units': 'Volume_Clean',
            'price_per_unit': 'Factory_Price' # or derived as 0.3 * unit
        }
        for k, v in mapper.items():
            if k in df_api.columns:
                df_api[v] = df_api[k]

        # 计算出厂价估计模型
        if 'Volume_Clean' in df_api.columns and 'Sales_Clean' in df_api.columns:
            df_api['Volume_Clean'] = pd.to_numeric(df_api['Volume_Clean'], errors='coerce').fillna(0)
            df_api['Sales_Clean'] = pd.to_numeric(df_api['Sales_Clean'], errors='coerce').fillna(0)
            df_api['Factory_Price'] = (df_api['Sales_Clean'] / df_api['Volume_Clean'].replace({0: np.nan}) * 0.3).round(4)
        
        # Filter Target Markets
        from step4_visualizer import Target_Markets, ORIGINATOR_CONFIG
        if '国家' in df_api.columns:
            df_api['国家'] = df_api['国家'].astype(str).str.upper().str.strip()
            df_api = df_api[df_api['国家'].isin(Target_Markets)]
        
        df_clean = df_api[df_api['Sales_Clean'] > 0].copy()
        
        # !! 核心修正：由于2025年只有2个月数据，全局将2025年的数据年化 (x6) 供各模块使用 !!
        mask_25 = df_clean['年份'].astype(str).str.contains('2025')
        if mask_25.any():
            df_clean.loc[mask_25, 'Sales_Clean'] *= 6
            df_clean.loc[mask_25, 'Volume_Clean'] *= 6
            # Factory Price Remains Unchanged
            
        if df_clean.empty:
            self.tips_label.setText("当前API在欧洲市场无可用数据！")
            return
            
        # 3. 将手工选择并打钩的企业作为原研
        self.analyzer.selected_companies_for_originator = self.originator_combo.get_checked_items()

        # 4. 初始化国家范围 (Top 15)
        c_stats = df_clean.groupby('国家')['Sales_Clean'].sum().sort_values(ascending=False)
        self.analyzer.current_countries = c_stats.head(15).index.tolist()
        self.analyzer.df_sales_clean = df_clean

        # 5. 生成所有图表并嵌入 QTabWidget
        self._render_tab(self.tab_widgets["0. 战略决策 (Global)"], self.analyzer.draw_summary_table, df_clean)
        self._render_tab(self.tab_widgets["1. 趋势分析 (Trend)"], self.analyzer.draw_trend, df_clean)
        self._render_tab(self.tab_widgets["2. 原研渗透率"], self.analyzer.draw_penetration, df_clean)
        # 解除翻页国家限制，画满画板，自动滚动长图，此时包含原研企业数据
        self._render_tab(self.tab_widgets["3. 全市场格局"], lambda f, d: self.analyzer.draw_pie_batch(f, d, self.analyzer.current_countries), df_clean)
        # 使用在 Promax 写好的冒泡图替代动态矩阵
        self._render_tab(self.tab_widgets["4. 机会矩阵"], self.analyzer.draw_bubble_matrix, df_clean)
        self._render_tab(self.tab_widgets["5. 智能预测"], self.analyzer.draw_prediction, df_clean)

        self.tips_label.setText(f"分析完成！当前指定原研企业: {', '.join(self.analyzer.selected_companies_for_originator) or '未指定 (将视作全仿制药市场分析)'}")

    def _render_tab(self, parent_widget, draw_func, df):
        layout = parent_widget.layout()
        # 清空旧内容
        while layout.count():
            item = layout.takeAt(0)
            if item.widget():
                item.widget().deleteLater()
                
        # 加一个内部滚动层
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        container = QWidget()
        container_layout = QVBoxLayout(container)
        
        fig = plt.Figure(figsize=(10, 8), dpi=100) # 更大高度
        ax = fig.add_subplot(111)
        ax.axis('off') # default clear
        
        # 增加原生带图标的复制按钮放在顶部，更容易被发现
        btn_layout = QHBoxLayout()
        btn_layout.addStretch()
        copy_btn = QPushButton("📋 一键粘贴图表至微信/PPT")
        copy_btn.setFixedSize(220, 35)
        copy_btn.setCursor(Qt.PointingHandCursor)
        copy_btn.setStyleSheet("""
            QPushButton { background-color: #ed8936; color: white; font-weight: bold; border-radius: 6px; font-size: 13px;}
            QPushButton:hover { background-color: #dd6b20; }
        """)
        copy_btn.clicked.connect(lambda _, f=fig: self._copy_figure(f))
        btn_layout.addWidget(copy_btn)
        container_layout.addLayout(btn_layout)
        
        canvas = FigureCanvas(fig)
        canvas.wheelEvent = lambda event: event.ignore() # 防止内部图表拦截滚轮
        
        report_text = ""
        try:
            # draw_* 返回的是要附带的文字报告 (或者 tuple)
            res = draw_func(fig, df)
            
            # 【修复闪动无法滚动】根据画布的实际内容高度强制设定 canvas 和容器高度
            # 拿到最终画板计算需要的英寸高度
            required_h = int(fig.get_size_inches()[1] * fig.dpi)
            min_h = max(800, required_h)
            canvas.setMinimumHeight(min_h)
            container.setMinimumHeight(min_h + 200) # 为文字和按钮预留空间
            
            if isinstance(res, tuple):
                report_text = res[0]
                leg_map = res[1]
                # Bind pick event for trend chart
                fig._highlighted_leg = None
                
                def on_pick(event):
                    legline = event.artist
                    if legline not in leg_map: return
                    c = leg_map[legline]
                    
                    if getattr(fig, '_highlighted_leg', None) == c:
                        # 再次点击，恢复所有
                        fig._highlighted_leg = None
                        for axes in fig.axes:
                            for line in axes.get_lines():
                                line.set_alpha(1.0)
                                line.set_linewidth(2)
                        for tl in leg_map.keys():
                            tl.set_alpha(1.0)
                    else:
                        # 首次点击，高亮自己，虚化别人
                        fig._highlighted_leg = c
                        for axes in fig.axes:
                            for line in axes.get_lines():
                                lbl = str(line.get_label())
                                if lbl.startswith('_'): continue
                                if lbl.startswith(c):
                                    line.set_alpha(1.0)
                                    line.set_linewidth(3)
                                    line.set_zorder(10)
                                else:
                                    line.set_alpha(0.1)
                                    line.set_linewidth(1)
                                    line.set_zorder(1)
                        for tl in leg_map.keys():
                            tl.set_alpha(1.0 if tl == legline else 0.1)
                            
                    canvas.draw_idle()
                canvas.mpl_connect('pick_event', on_pick)
            else:
                report_text = res
        except Exception as e:
            report_text = f"绘制该图表时发生异常: {str(e)}"
            ax.text(0.5, 0.5, report_text, ha='center', va='center')
            
        container_layout.addWidget(canvas, stretch=3)
        
        if report_text and isinstance(report_text, str):
            text_edit = QPlainTextEdit(report_text)
            text_edit.setReadOnly(True)
            text_edit.setMinimumHeight(120)
            text_edit.setMaximumHeight(200)
            text_edit.setStyleSheet("background-color: #f8f9fa; color: #2d3748; font-family: Consolas;")
            container_layout.addWidget(text_edit, stretch=1)

        scroll.setWidget(container)
        layout.addWidget(scroll)

    def _copy_figure(self, fig):
        if ClipboardHelper.copy_figure(fig):
            QMessageBox.information(self, "复制成功", "图表已成功高保真复制到剪贴板！可以直接粘贴。")
        else:
            QMessageBox.warning(self, "复制失败", "无法复制图表到剪贴板。可能是无可用内存或平台不支持。")

# ======= 3. 主窗口统筹 =======
class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("以岭药业万洲制药市场分析 (Analytics Shell 1.0)")
        self.resize(1280, 800)
        self.center_window()
        self.setup_ui()
        
    def center_window(self):
        qr = self.frameGeometry()
        cp = self.screen().availableGeometry().center()
        qr.moveCenter(cp)
        self.move(qr.topLeft())

    def setup_ui(self):
        self.setStyleSheet("""
            QMainWindow {
                background-color: #f7fafc;
            }
        """)
        
        # 核心 Widget
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)
        main_layout.setContentsMargins(0, 0, 0, 0)
        main_layout.setSpacing(0)
        
        # 1. 顶部 Header
        header = QFrame()
        header.setFixedHeight(70)
        header.setStyleSheet("background-color: #ffffff; border-bottom: 1px solid #e2e8f0;")
        header_layout = QHBoxLayout(header)
        header_layout.setContentsMargins(20, 0, 20, 0)
        
        # Logo & Title
        logo_label = QLabel()
        logo_path = "logo.png" if os.path.exists("logo.png") else ("logo.jpg" if os.path.exists("logo.jpg") else "")
        if logo_path:
            pixmap = QPixmap(logo_path).scaled(40, 40, Qt.KeepAspectRatio, Qt.SmoothTransformation)
            logo_label.setPixmap(pixmap)
        else:
            logo_label.setText("📦")
            logo_label.setStyleSheet("font-size: 28px;")
            
        title_label = QLabel("万洲制药市场全景分析系统")
        title_label.setStyleSheet("font-size: 18px; font-weight: 800; color: #1a202c; letter-spacing: 1px;")
        
        header_layout.addWidget(logo_label)
        header_layout.addSpacing(10)
        header_layout.addWidget(title_label)
        
        self.toggle_main_sidebar_btn = QPushButton("☰ 折叠/展开主菜单")
        self.toggle_main_sidebar_btn.setStyleSheet("""
            QPushButton {
                background-color: transparent; 
                color: #2b6cb0; 
                font-weight: bold; 
                font-size: 13px;
                padding: 5px 10px;
                border: 1px solid #cbd5e0;
                border-radius: 4px;
            }
            QPushButton:hover {
                background-color: #ebf8ff;
            }
        """)
        self.toggle_main_sidebar_btn.setCursor(Qt.PointingHandCursor)
        self.toggle_main_sidebar_btn.clicked.connect(self.toggle_main_sidebar)
        header_layout.addSpacing(20)
        header_layout.addWidget(self.toggle_main_sidebar_btn)
        
        header_layout.addStretch()
        
        # Progress Indicator Text
        self.progress_indicator = QLabel("[1] ➔ [2] ➔ [3] ➔ [4]")
        self.progress_indicator.setStyleSheet("font-size: 15px; font-weight: bold; color: #a0aec0; letter-spacing: 2px;")
        header_layout.addWidget(self.progress_indicator)
        
        main_layout.addWidget(header)
        
        # 2. 中间区域: 左侧菜单 + 右侧 Stack 
        body_layout = QHBoxLayout()
        main_layout.addLayout(body_layout)
        
        # Sidebar
        self.main_sidebar = QFrame()
        self.main_sidebar.setFixedWidth(240)
        self.main_sidebar.setStyleSheet("""
            QFrame {
                background-color: #ffffff;
                border-right: 1px solid #e2e8f0;
            }
            QPushButton {
                text-align: left;
                padding: 15px;
                font-size: 14px;
                font-weight: bold;
                color: #4a5568;
                border: none;
                border-radius: 6px;
                margin: 5px 10px;
            }
            QPushButton:hover {
                background-color: #edf2f7;
            }
            QPushButton:checked {
                background-color: #ebf8ff;
                color: #3182ce;
                border-left: 4px solid #3182ce;
            }
            QPushButton:disabled {
                color: #cbd5e0;
            }
        """)
        sidebar_layout = QVBoxLayout(self.main_sidebar)
        sidebar_layout.setContentsMargins(0, 20, 0, 0)
        sidebar_layout.setAlignment(Qt.AlignTop)
        
        self.nav_btns = []
        nav_titles = [
            "1. 数据底座提取", 
            "2. 剂型规格清洗", 
            "3. 核心分片标准化", 
            "🔮 [可选] 多维综合透视矩阵",
            "4. 市场情报看板 (Dashboard)",
            "5. 欧洲市场深度分析"
        ]
        for i, text in enumerate(nav_titles):
            btn = QPushButton(text)
            btn.setCheckable(True)
            btn.clicked.connect(lambda checked, idx=i: self.switch_page(idx))
            sidebar_layout.addWidget(btn)
            self.nav_btns.append(btn)
            
        self.nav_btns[0].setChecked(True)
        body_layout.addWidget(self.main_sidebar)
        
        # Stacked Widget
        self.stack = QStackedWidget()
        body_layout.addWidget(self.stack)
        
        # 初始化页面
        self.page1 = Step1Widget()
        self.page2 = Step2Widget()
        self.page3 = Step3Widget()
        self.page_pivot = FlexiblePivotWidget() # 插入点: index = 3
        self.page4 = Step4DashboardWidget() # index = 4
        self.page5 = EuropeanAnalysisPage() # index = 5
        
        self.stack.addWidget(self.page1)
        self.stack.addWidget(self.page2)
        self.stack.addWidget(self.page3)
        self.stack.addWidget(self.page_pivot)
        self.stack.addWidget(self.page4)
        self.stack.addWidget(self.page5)
        
        # 启动时自动扫描 Cache/API_Views 发现新药并增补到 core_config.py
        self._auto_scan_api_views()

    def _auto_scan_api_views(self):
        """启动时扫描 Cache/API_Views 目录, 发现新药名自动增补到 core_config.py 的 ORIGINATOR_CONFIG"""
        try:
            base_dir = os.path.dirname(os.path.abspath(__file__))
            api_views_dir = os.path.join(base_dir, "Cache", "API_Views")
            config_path = os.path.join(base_dir, "core_config.py")
            
            if not os.path.isdir(api_views_dir) or not os.path.isfile(config_path):
                return
            
            # 1. 提取所有药物名（从文件名 core_cache_DRUGNAME.parquet）
            discovered_drugs = set()
            for f in os.listdir(api_views_dir):
                if f.startswith("core_cache_") and f.endswith(".parquet"):
                    drug_name = f.replace("core_cache_", "").replace(".parquet", "").strip().upper()
                    if drug_name:
                        discovered_drugs.add(drug_name)
            
            if not discovered_drugs:
                return
            
            # 2. 读取当前 core_config.py 中已有的 ORIGINATOR_CONFIG 键
            existing_drugs = set(k.upper() for k in core_config.ORIGINATOR_CONFIG.keys())
            
            # 3. 找出新药（不在已有配置中的）
            new_drugs = sorted(discovered_drugs - existing_drugs)
            
            if not new_drugs:
                return  # 没有新增需要
            
            # 4. 读取 core_config.py 文件内容，找到 ORIGINATOR_CONFIG 的结束位置并插入新条目
            with open(config_path, 'r', encoding='utf-8') as f:
                content = f.read()
            
            # 找到 ORIGINATOR_CONFIG 中最后一个 } 的位置
            import re
            # 匹配 ORIGINATOR_CONFIG = { ... } 块 — 找到 } 之前插入新行
            pattern = r'(ORIGINATOR_CONFIG\s*=\s*\{)(.*?)(^\})'
            match = re.search(pattern, content, re.DOTALL | re.MULTILINE)
            
            if not match:
                return
            
            # 在 } 之前插入新条目
            existing_block = match.group(2)
            new_entries = []
            for drug in new_drugs:
                # 确保不重复（大小写不敏感）
                if f'"{drug}"' not in existing_block and f"'{drug}'" not in existing_block:
                    new_entries.append(f'    "{drug}":' + ' ' * max(1, 16 - len(drug)) + '[],' + f'  # TODO: 请填写原研企业关键词')
            
            if not new_entries:
                return
            
            # 把新行插入到 } 之前
            insert_text = "\n    # --- 以下为自动扫描发现的新药，请手动补充原研企业 ---\n" + "\n".join(new_entries) + "\n"
            new_content = content[:match.start(3)] + insert_text + content[match.start(3):]
            
            with open(config_path, 'w', encoding='utf-8') as f:
                f.write(new_content)
            
            print(f"[AutoScan] 已自动发现并增补 {len(new_entries)} 个新药到 core_config.py: {[d for d in new_drugs if any(d in e for e in new_entries)]}")
            
            # 5. 同时更新内存中的 ORIGINATOR_CONFIG
            for drug in new_drugs:
                if drug not in core_config.ORIGINATOR_CONFIG:
                    core_config.ORIGINATOR_CONFIG[drug] = []
                    
        except Exception as e:
            print(f"[AutoScan] 扫描 API_Views 时出错（不影响主程序）: {e}")

    def toggle_main_sidebar(self):
        self.main_sidebar.setVisible(not self.main_sidebar.isVisible())
        
        # 绑定流水线流转事件
        self.page1.completed_signal.connect(self.on_step1_done)
        self.page2.completed_signal.connect(self.on_step2_done)
        self.page3.completed_signal.connect(self.on_step3_done)
        
    def switch_page(self, index):
        for i, btn in enumerate(self.nav_btns):
            btn.setChecked(i == index)
        self.stack.setCurrentIndex(index)
        
        # 更新指示器高亮逻辑
        states = ['[1]', '[2]', '[3]', '[Pivot]', '[4]', '[5]']
        states[index] = f"<span style='color:#3182ce;'>{states[index]}</span>"
        self.progress_indicator.setText(" ➔ ".join(states))
        
        # 激活 Pivot 组件时尝试自动重新装载最新下发的底层数据
        if index == 3:
            # 如果从没读过数据或用户切到这里，顺手帮加载
            if self.page_pivot.df_raw is None:
                self.page_pivot.load_data()

        # Step 4 / 5 特定逻辑
        if index == 4 or index == 5:
            self.showMaximized()
            # 尝试加载所有 API cache file 塞给视图层
            try:
                test_dir = 'Cache/API_Views/'
                if self.page3 and hasattr(self.page3, 'worker') and self.page3.worker:
                    test_dir = self.page3.worker.kwargs.get('output_dir', 'Cache/API_Views/')
                
                # Use absolute path
                base_dir = os.path.dirname(os.path.abspath(__file__))
                test_dir = os.path.join(base_dir, test_dir.lstrip('/'))
                
                if not hasattr(self, '_cached_combined_df') or self._cached_combined_df is None:
                    if os.path.exists(test_dir):
                        files = os.listdir(test_dir)
                        dfs = []
                        for f in files:
                            if f.endswith('.parquet'):
                                p = os.path.join(test_dir, f)
                                dfs.append(pd.read_parquet(p))
                        
                        if dfs:
                            self._cached_combined_df = pd.concat(dfs, ignore_index=True)
                
                if self._cached_combined_df is not None and not self._cached_combined_df.empty:
                    # Filter data to only show APIs requested in Step 1
                    searched_drugs_text = self.page1.drugs_input.toPlainText().replace('\n', ' ').strip()
                    if searched_drugs_text:
                        from step_a_download import get_parsed_drug_names
                        parsed_names = get_parsed_drug_names(searched_drugs_text)
                        if parsed_names and 'api_name' in self._cached_combined_df.columns:
                            filtered_df = self._cached_combined_df[self._cached_combined_df['api_name'].isin(parsed_names)]
                        else:
                            filtered_df = self._cached_combined_df
                    else:
                        filtered_df = self._cached_combined_df

                    if index == 4:
                        self.page4.set_dataframe(filtered_df)
                        self.page4.on_filter_changed({})
                    elif index == 5:
                        self.page5.set_data(filtered_df)
            except Exception as e:
                import traceback
                print("Failed to auto-load databse for plotting:", traceback.format_exc())

    def on_step1_done(self, file_path):
        self.page2.set_input(file_path)
        self.switch_page(1)
        self.page2.execute()
        
    def on_step2_done(self, parquet_path):
        if parquet_path:
            self.page3.set_input(parquet_path)
            self.switch_page(2)
            self.page3.execute()
        else:
            # 走到这里意味着步骤2被纯缓存旁路跳过了（全部已存在最新缓存）
            self.append_log("[*] 检测到清洗阶段全部使用旁路缓存，自动跳转至底层合成 (Step 4 可点击切换查看) ...")
            # 既然跳过了说明API_Views里都有最新的数据，我们直接帮用户切到Step4并加载
            self.switch_page(4)
        
    def on_step3_done(self, out_dir):
        self._cached_combined_df = None # 置空缓存，强制重新读取
        self.switch_page(3)


if __name__ == "__main__":
    # 启用现代高分屏渲染质量
    QApplication.setHighDpiScaleFactorRoundingPolicy(Qt.HighDpiScaleFactorRoundingPolicy.PassThrough)
    
    app = QApplication(sys.argv)
    
    window = MainWindow()
    window.show()
    sys.exit(app.exec())
