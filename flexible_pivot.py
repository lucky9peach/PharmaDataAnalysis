import os
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QLabel, QPushButton,
    QTabWidget, QTableWidget, QTableWidgetItem, QHeaderView,
    QComboBox, QMessageBox, QFileDialog, QScrollArea, QSplitter
)
from PySide6.QtCore import Qt
from matplotlib.backends.backend_qtagg import FigureCanvasQTAgg as FigureCanvas
import core_config
from step4_visualizer import ClipboardHelper

class FlexiblePivotWidget(QWidget):
    """
    独立且灵活的多维数据分析透视面板。
    直读 Step 1 刚下载合并出来的 Cache/step1_latest.xlsx。
    提供核心维度的交叉透视表、简单分布图、以及支持一键完整导出至 Excel 功能。
    """
    def __init__(self, parent=None):
        super().__init__(parent)
        self.df_raw = None
        self.current_api = None
        self.setup_ui()

    def setup_ui(self):
        self.layout = QVBoxLayout(self)
        self.layout.setContentsMargins(15, 15, 15, 15)

        # =============== 顶部控制栏 ===============
        top_bar = QHBoxLayout()
        
        self.btn_load = QPushButton("🔄 加载最新 Step 1 数据")
        self.btn_load.setStyleSheet("background-color: #4a5568; color: white; padding: 6px 15px; border-radius: 4px; font-weight: bold;")
        self.btn_load.clicked.connect(self.load_data)
        
        self.api_combo = QComboBox()
        self.api_combo.setMinimumWidth(200)
        self.api_combo.currentTextChanged.connect(self.on_api_changed)
        
        self.btn_export = QPushButton("💾 导出本页透视表至 Excel")
        self.btn_export.setStyleSheet("background-color: #38a169; color: white; padding: 6px 15px; border-radius: 4px; font-weight: bold;")
        self.btn_export.clicked.connect(self.export_to_excel)
        self.btn_export.setEnabled(False)

        top_bar.addWidget(self.btn_load)
        top_bar.addSpacing(20)
        top_bar.addWidget(QLabel("当前聚焦 API (通用名):"))
        top_bar.addWidget(self.api_combo)
        top_bar.addStretch()
        top_bar.addWidget(self.btn_export)

        self.layout.addLayout(top_bar)

        # =============== 核心 Tab 区域 ===============
        self.tabs = QTabWidget()
        self.tabs.setStyleSheet("""
            QTabBar::tab { padding: 10px 20px; font-weight: bold; font-size: 14px;}
            QTabBar::tab:selected { background-color: #3182ce; color: white; }
        """)
        self.layout.addWidget(self.tabs)

        self.tab_names = [
            "🌍 地理市场分布 (Country)", 
            "💊 主力剂型分布 (Formulation)", 
            "🔀 终极交叉透视 (国家 x 剂型)"
        ]
        
        self.tab_widgets = {}
        for name in self.tab_names:
            w = QWidget()
            w.setLayout(QVBoxLayout())
            
            # 内部左右分割 (左图右表 / 上图下表)
            splitter = QSplitter(Qt.Horizontal)
            
            # 画图区
            fig_container = QWidget()
            fig_layout = QVBoxLayout(fig_container)
            fig = plt.Figure(figsize=(6, 5), dpi=100)
            canvas = FigureCanvas(fig)
            fig_layout.addWidget(canvas)
            fig_container.fig = fig
            fig_container.canvas = canvas
            fig_container.ax = fig.add_subplot(111)
            # 防止图表截获滚轮事件，让滚动条可以全区域工作
            canvas.wheelEvent = lambda event: event.ignore()
            
            # 复制图表按钮
            btn_copy_fig = QPushButton("📋 复制此图表")
            btn_copy_fig.setStyleSheet("background-color: #ed8936; color: white; font-weight: bold; border-radius: 4px;")
            btn_copy_fig.clicked.connect(lambda _, f=fig: self._copy_figure(f))
            fig_layout.addWidget(btn_copy_fig)
            
            # 数据表区
            table_container = QWidget()
            table_layout = QVBoxLayout(table_container)
            table = QTableWidget()
            table.setAlternatingRowColors(True)
            table.setStyleSheet("""
                QTableWidget { font-size: 13px; border: 1px solid #e2e8f0; }
                QHeaderView::section { background-color: #edf2f7; font-weight: bold; padding: 4px;}
            """)
            table_layout.addWidget(table)
            table_container.table = table
            
            # 复制名单功能按钮 (给 Tab 1 特供)
            if "地理市场" in name:
                btn_copy_names = QPushButton("📋 一键复制所有买家国家名单")
                btn_copy_names.clicked.connect(self._copy_country_names)
                table_layout.addWidget(btn_copy_names)

            splitter.addWidget(fig_container)
            splitter.addWidget(table_container)
            splitter.setSizes([500, 500])
            
            w.layout().addWidget(splitter)
            self.tabs.addTab(w, name)
            
            self.tab_widgets[name] = {
                'fig': fig,
                'canvas': canvas,
                'ax': fig_container.ax,
                'table': table,
                'df_pivot': None # 用于挂载当前 tab 计算出的透视表数据，供导出 Excel
            }
            
        # 设置 matplotlib 字体
        import platform
        if platform.system() == 'Windows':
            plt.rcParams['font.sans-serif'] = ['SimHei', 'Microsoft YaHei']
        elif platform.system() == 'Darwin':
            plt.rcParams['font.sans-serif'] = ['Arial Unicode MS']
        plt.rcParams['axes.unicode_minus'] = False

    def _copy_figure(self, fig):
        if ClipboardHelper.copy_figure(fig):
            QMessageBox.information(self, "复制成功", "图表已复制到剪贴板！可以直接粘贴进微信或PPT。")
        else:
            QMessageBox.warning(self, "复制失败", "图表复制失败，可能由于平台不支持。")
            
    def _copy_country_names(self):
        tab_data = self.tab_widgets["🌍 地理市场分布 (Country)"]
        df_pivot = tab_data['df_pivot']
        if df_pivot is not None and not df_pivot.empty:
            countries = df_pivot.index.tolist()
            text = "\n".join([str(c) for c in countries])
            
            from PySide6.QtGui import QClipboard
            from PySide6.QtWidgets import QApplication
            clipboard = QApplication.clipboard()
            clipboard.setText(text)
            QMessageBox.information(self, "复制成功", f"已成功复制 {len(countries)} 个国家名称到剪贴板！")

    def load_data(self):
        base_dir = os.path.dirname(os.path.abspath(__file__))
        file_path_csv = os.path.join(base_dir, "Cache", "step1_latest.csv")
        file_path_xlsx = os.path.join(base_dir, "Cache", "step1_latest.xlsx")
        
        file_path = ""
        if os.path.exists(file_path_csv):
            file_path = file_path_csv
        elif os.path.exists(file_path_xlsx):
            file_path = file_path_xlsx
            
        if not file_path:
            QMessageBox.critical(self, "错误", f"找不到数据源文件 (需要先完成 Step 1 下载):\nCache/step1_latest.csv")
            return
            
        try:
            from PySide6.QtWidgets import QApplication
            from PySide6.QtCore import Qt
            QApplication.setOverrideCursor(Qt.WaitCursor)
            
            if file_path.endswith('.csv'):
                df = pd.read_csv(file_path, encoding='utf-8-sig', low_memory=False)
            else:
                df = pd.read_excel(file_path)
            
            # --- 核心降维清洗：只保留多维分析关心的列并翻译成易读中文 ---
            # 兼容：有时下载的是原始中文表头、有时可能是英文表头
            mapping_rules = {
                '检索药名': 'API',
                
                # 地理
                '国家': '国家/地区',
                'market_region': '国家/地区',
                
                # 药品固有属性
                '剂型': '细分剂型',
                'formulation': '细分剂型',
                'NFC1': '大类别',
                '规格': '规格',
                
                # 公司属性
                '集团/企业': '企业(MAH)',
                'mah': '企业(MAH)',
                
                # 年份
                '年份': '年份',
                'year': '年份',
                
                # 销量与金额 (因为不同药下载列名带有空格和单位比如 "最小单包装销售数量 粒")
                '最小单包装销售数量 粒': '最小单包装销量(Unit)',
                '最小单包装销售数量': '最小单包装销量(Unit)',
                'sales_volume_units': '最小单包装销量(Unit)',
                
                '大包装销售数量 粒': '大包装销量(Box)',
                '大包装销售数量': '大包装销量(Box)',
                
                '销售额': '销售金额(USD)',
                'sales_value_usd': '销售金额(USD)',
                
                '公斤': '原料药消耗量(KG)',
                'volume_api_kg': '原料药消耗量(KG)'
            }
            
            # 使用 contains 模糊重命名因为列名经常带不同的单位后缀如 "最小单包装销售数量 支"
            for orig_col in df.columns:
                if '国家' in orig_col or orig_col == 'market_region':
                    df.rename(columns={orig_col: '国家/地区'}, inplace=True)
                elif '剂型' in orig_col or orig_col == 'formulation':
                    df.rename(columns={orig_col: '细分剂型'}, inplace=True)
                elif orig_col == '规格' or orig_col == 'strength_raw':
                    df.rename(columns={orig_col: '规格'}, inplace=True)
                elif '集团/企业' == orig_col or orig_col == 'mah':
                    df.rename(columns={orig_col: '企业(MAH)'}, inplace=True)
                elif '最小单包装销售数量' in orig_col or orig_col == 'sales_volume_units':
                    df.rename(columns={orig_col: '最小单包装销量(Unit)'}, inplace=True)
                elif '大包装销售数量' in orig_col:
                    df.rename(columns={orig_col: '大包装销量(Box)'}, inplace=True)
                elif orig_col == '销售额' or orig_col == 'sales_value_usd':
                    df.rename(columns={orig_col: '销售金额(USD)'}, inplace=True)
                elif '公斤' in orig_col or orig_col == 'volume_api_kg':
                    df.rename(columns={orig_col: '原料药消耗量(KG)'}, inplace=True)
                elif orig_col == '检索药名':
                    df.rename(columns={orig_col: 'API'}, inplace=True)
                    
            df_renamed = df
            
            # 手工补齐如果缺失的列，保持行不错乱
            essential_cols = {
                'API': '未知API', 
                '国家/地区': '未知国家', 
                '细分剂型': '未知剂型', 
                '规格': '未知',
                '企业(MAH)': '未知企业'
            }
            for col, fill_val in essential_cols.items():
                if col not in df_renamed.columns:
                    df_renamed[col] = fill_val
                else:
                    df_renamed[col] = df_renamed[col].fillna(fill_val)
                    
            # 核心数值列初始化
            numeric_cols = ['最小单包装销量(Unit)', '大包装销量(Box)', '销售金额(USD)', '原料药消耗量(KG)']
            for num_col in numeric_cols:
                if num_col not in df_renamed.columns:
                    df_renamed[num_col] = 0.0
                else:
                    df_renamed[num_col] = pd.to_numeric(df_renamed[num_col], errors='coerce').fillna(0.0)

            # 衍生计算列！这里保障核心推导逻辑的精确度
            # 1. 求转换系数(每盒多少粒) = 最小单包装销量 / 大包装销量
            # 使用 numpy divide 避免除 0 报错 (结果变 infinity)
            df_renamed['装量系数(粒/盒)'] = np.where(
                df_renamed['大包装销量(Box)'] > 0, 
                df_renamed['最小单包装销量(Unit)'] / df_renamed['大包装销量(Box)'], 
                np.nan
            )
            
            # 2. 单价(每粒多少钱) = 销售额 / 最小单包装销量
            df_renamed['颗粒单价(USD/Unit)'] = np.where(
                df_renamed['最小单包装销量(Unit)'] > 0,
                df_renamed['销售金额(USD)'] / df_renamed['最小单包装销量(Unit)'],
                np.nan
            )
            
            # 3. 出厂价 = 单价 * 0.3
            df_renamed['预估出厂价(USD)'] = df_renamed['颗粒单价(USD/Unit)'] * 0.3
            
            # 重新整理精度的四舍五入
            df_renamed['颗粒单价(USD/Unit)'] = df_renamed['颗粒单价(USD/Unit)'].round(6)
            df_renamed['预估出厂价(USD)'] = df_renamed['预估出厂价(USD)'].round(6)

            self.df_raw = df_renamed
            
            # 更新下拉框
            apis = sorted(df_renamed['API'].dropna().unique().tolist())
            self.api_combo.blockSignals(True)
            self.api_combo.clear()
            self.api_combo.addItems([str(a) for a in apis])
            self.api_combo.blockSignals(False)
            
            if apis:
                self.api_combo.setCurrentIndex(0)
                self.on_api_changed(apis[0])
                
            QMessageBox.information(self, "加载成功", f"成功读取最新的下发数据，共 {len(df_renamed)} 行。")
            self.btn_export.setEnabled(True)
            
        except Exception as e:
            import traceback
            QMessageBox.critical(self, "读取失败", f"发生了以下错误:\n{str(e)}\n\n{traceback.format_exc()}")
        finally:
            from PySide6.QtWidgets import QApplication
            QApplication.restoreOverrideCursor()

    def on_api_changed(self, api_name):
        self.current_api = api_name
        if not self.current_api or self.df_raw is None: return
        
        # 截取该 API 的数据
        df_api = self.df_raw[self.df_raw['API'] == self.current_api].copy()
        
        # 重新渲染三大 Tab
        self.render_tab1_geography(df_api)
        self.render_tab2_formulation(df_api)
        self.render_tab3_crosstab(df_api)

    def _fill_table(self, table_widget: QTableWidget, df: pd.DataFrame, is_crosstab=False):
        table_widget.clear()
        
        if df.empty:
            table_widget.setRowCount(0)
            table_widget.setColumnCount(0)
            return

        # 把 Index 变成普通列，方便展示
        df_display = df.reset_index()
        
        table_widget.setRowCount(df_display.shape[0])
        table_widget.setColumnCount(df_display.shape[1])
        table_widget.setHorizontalHeaderLabels([str(c) for c in df_display.columns])
        
        for row in range(df_display.shape[0]):
            for col in range(df_display.shape[1]):
                val = df_display.iloc[row, col]
                
                # 格式化数值展示
                if pd.isna(val):
                    display_str = ""
                elif isinstance(val, (int, np.integer)):
                    display_str = f"{val:,}"
                elif isinstance(val, (float, np.floating)):
                    if val > 1000:
                        display_str = f"{val:,.0f}"
                    else:
                        display_str = f"{val:.2f}"
                else:
                    display_str = str(val)
                    
                item = QTableWidgetItem(display_str)
                # 数值靠右，文本靠左
                if isinstance(val, (int, float, np.number)):
                    item.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
                else:
                    item.setTextAlignment(Qt.AlignLeft | Qt.AlignVCenter)
                    
                table_widget.setItem(row, col, item)
                
        table_widget.resizeColumnsToContents()
        table_widget.horizontalHeader().setStretchLastSection(True)

    def render_tab1_geography(self, df):
        tab = self.tab_widgets["🌍 地理市场分布 (Country)"]
        ax = tab['ax']
        ax.clear()

        # 透视计算 (按国家汇总销量十年来总计)
        df_agg = df.groupby('国家/地区')[['最小单包装销量(Unit)', '销售金额(USD)']].sum().sort_values(by='最小单包装销量(Unit)', ascending=False)
        tab['df_pivot'] = df_agg # 保存结果准备导出
        
        self._fill_table(tab['table'], df_agg)
        
        # 画图 - Top 15 国家的销量柱状图
        top_N = df_agg.head(15)
        if not top_N.empty:
            labels = top_N.index.tolist()
            vals = top_N['最小单包装销量(Unit)'].tolist()
            
            y_pos = np.arange(len(labels))
            
            # 给深色好看的柱子
            ax.bar(y_pos, vals, align='center', color='#3182ce', alpha=0.8, edgecolor='none')
            ax.set_xticks(y_pos)
            ax.set_xticklabels(labels, rotation=45, ha='right', fontsize=9)
            ax.set_ylabel('累计历史销售总数量 (Unit)', fontsize=10)
            ax.set_title(f'【{self.current_api}】Top 15 地理市场销量分布 (历史总计)', fontsize=12, pad=15)
            
            # 去除顶部和右侧外框
            ax.spines['top'].set_visible(False)
            ax.spines['right'].set_visible(False)
            
            tab['fig'].tight_layout()
            tab['canvas'].draw()

    def render_tab2_formulation(self, df):
        tab = self.tab_widgets["💊 主力剂型分布 (Formulation)"]
        fig = tab['fig']
        ax = tab['ax']
        fig.clf()
        ax = fig.add_subplot(111)
        tab['ax'] = ax

        # 剂型透视
        df_agg = df.groupby('细分剂型')[['最小单包装销量(Unit)', '销售金额(USD)', '原料药消耗量(KG)']].sum().sort_values(by='最小单包装销量(Unit)', ascending=False)
        tab['df_pivot'] = df_agg
        
        self._fill_table(tab['table'], df_agg)
        
        # 画个饼图看份额 (取 Top 6，剩下全算 Others)
        if not df_agg.empty:
            N = 6
            top_N = df_agg.head(N)
            others = df_agg.iloc[N:].sum()
            
            labels = top_N.index.tolist()
            sizes = top_N['最小单包装销量(Unit)'].tolist()
            
            if not others.empty and others['最小单包装销量(Unit)'] > 0:
                labels.append('其他小众剂型')
                sizes.append(others['最小单包装销量(Unit)'])
                
            # 过滤掉 0 大小的保证不死轴
            clean_labels = [l for l, s in zip(labels, sizes) if s > 0]
            clean_sizes = [s for s in sizes if s > 0]

            if clean_sizes:
                colors = plt.cm.tab20c(np.linspace(0, 1, len(clean_sizes)))
                wedges, texts, autotexts = ax.pie(
                    clean_sizes, labels=clean_labels, autopct='%1.1f%%',
                    startangle=140, colors=colors,
                    wedgeprops=dict(width=0.4, edgecolor='w') # 画成环形图比较现代
                )
                plt.setp(autotexts, size=9, weight="bold")
                plt.setp(texts, size=10)
                
                ax.set_title(f'【{self.current_api}】主力剂型销量占比情况', fontsize=12, pad=15)
                tab['fig'].tight_layout()
                tab['canvas'].draw()

    def render_tab3_crosstab(self, df):
        tab = self.tab_widgets["🔀 终极交叉透视 (国家 x 剂型)"]
        fig = tab['fig']
        ax = tab['ax']
        
        # 为了彻底清除以前画的 colorbar 相关 axes，最安全的做法是清空整个 figure 重新加子图
        fig.clf()
        ax = fig.add_subplot(111)
        tab['ax'] = ax
        
        # 双维交叉表 (Rows=国家, Cols=剂型, Values=销量)
        # 用 pivot_table 
        pt = pd.pivot_table(
            df, 
            values='最小单包装销量(Unit)', 
            index='国家/地区', 
            columns='细分剂型', 
            aggfunc='sum', 
            fill_value=0
        )
        
        # 顺便加个总计列好排序国家
        pt['总计'] = pt.sum(axis=1)
        pt = pt.sort_values(by='总计', ascending=False)
        # 去掉行列太多零的情况，美化
        pt = pt.loc[:, (pt != 0).any(axis=0)]
        
        tab['df_pivot'] = pt
        self._fill_table(tab['table'], pt, is_crosstab=True)
        
        # 交叉分析的图：画一个热力散点气泡图（为了美观取代传统 Heatmap）
        if not pt.empty:
            # 取 Top 15 个国家 和 Top 10 个剂型 画图，否则太密
            top_pt = pt.head(15).drop(columns=['总计']).copy()
            # 选最畅销的10个剂型
            top_forms = top_pt.sum(axis=0).sort_values(ascending=False).head(10).index
            top_pt = top_pt[top_forms]
            
            y_labels = top_pt.index.tolist()
            x_labels = top_pt.columns.tolist()
            
            xv, yv = np.meshgrid(np.arange(len(x_labels)), np.arange(len(y_labels)))
            # 展平矩阵
            values = top_pt.values.flatten()
            
            # 过滤 0 的气泡不画
            mask = values > 0
            
            # 使用平方根把很大和很小的体积缩小差距
            sizes = np.sqrt(values[mask])
            # 把 sizes 定标到合适的像素大小范围 (20 -> 800)
            if sizes.max() > sizes.min():
                norm_sizes = 20 + 780 * (sizes - sizes.min()) / (sizes.max() - sizes.min())
            else:
                norm_sizes = sizes * 0 + 200
                
            scatter = ax.scatter(xv.flatten()[mask], yv.flatten()[mask], s=norm_sizes, alpha=0.6, c=values[mask], cmap='viridis', edgecolors='none')
            
            ax.set_xticks(np.arange(len(x_labels)))
            ax.set_xticklabels(x_labels, rotation=45, ha='right', fontsize=9)
            
            ax.set_yticks(np.arange(len(y_labels)))
            ax.set_yticklabels(y_labels, fontsize=10)
            
            ax.set_title(f'【{self.current_api}】主力国家与主售剂型分布雷达', fontsize=12, pad=10)
            ax.invert_yaxis() # 国家排序 排名高的在上面
            
            fig = tab['fig']
            fig.colorbar(scatter, ax=ax, label='最小单包装销量 (Unit)', anchor=(0, 0.5), shrink=0.7)
            
            fig.tight_layout()
            tab['canvas'].draw()

    def export_to_excel(self):
        if not self.current_api:
            return
            
        file_path, _ = QFileDialog.getSaveFileName(
            self, 
            "导出所有多维分析透视表", 
            f"灵活透视分析_{self.current_api}.xlsx", 
            "Excel Files (*.xlsx)"
        )
        if not file_path:
            return
            
        try:
            with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
                # 遍历三个 Tab 中的 DataFrame，写入三个 Sheet
                for name_tab in self.tab_names:
                    df_pivot = self.tab_widgets[name_tab]['df_pivot']
                    # 处理 Sheet 名，去掉 Emoji 和过长字符，不能超过 31
                    sheet_name = name_tab.split(" ")[1][:30] 
                    if df_pivot is not None and not df_pivot.empty:
                        df_pivot.to_excel(writer, sheet_name=sheet_name)
                        
                # 再把本 API 的原始清洗底层数据也附上去
                if self.df_raw is not None:
                    df_api = self.df_raw[self.df_raw['API'] == self.current_api]
                    df_api.to_excel(writer, sheet_name="底层元数据明细", index=False)
                    
            QMessageBox.information(self, "导出成功", f"恭喜，您已成功导出全部透视表数据至：\n{file_path} \n\n您可以直接发给老板了！")
        except Exception as e:
            QMessageBox.critical(self, "导出失败", f"导出过程中出错，请勿在 Excel 刚打开时强行覆盖它。\n{str(e)}")
