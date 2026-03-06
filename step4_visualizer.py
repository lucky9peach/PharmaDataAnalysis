import pandas as pd
import numpy as np
import io
import os
import sys
import subprocess
import platform
import matplotlib
import matplotlib.cm as cm
import matplotlib.colors as mcolors
import core_config
import scipy.stats as stats

# 欧洲经济区 + 英国 完整清单
EEA_AND_UK_MARKETS = [
    "AUSTRIA", "BELGIUM", "BULGARIA", "CROATIA", "CYPRUS",
    "CZECH REPUBLIC", "DENMARK", "ESTONIA", "FINLAND", "FRANCE",
    "GERMANY", "GREECE", "HUNGARY", "ICELAND", "IRELAND",
    "ITALY", "LATVIA", "LIECHTENSTEIN", "LITHUANIA", "LUXEMBOURG",
    "MALTA", "NETHERLANDS", "NORWAY", "POLAND", "PORTUGAL",
    "ROMANIA", "SLOVAKIA", "SLOVENIA", "SPAIN", "SWEDEN",
    "UNITED KINGDOM"
]
US_MARKETS = ["UNITED STATES", "US", "USA", "美国"]
# 完整可展示的市场（EEA+UK+US，Step4侧边栏只展示这些）
ALLOWED_DISPLAY_MARKETS = set([m.upper() for m in EEA_AND_UK_MARKETS] + [m.upper() for m in US_MARKETS])

# 解决跨平台中文字体显示方块的问题 (苹果和 Windows 研判)
if platform.system() == 'Darwin':
    matplotlib.rcParams['font.sans-serif'] = ['Arial Unicode MS', 'PingFang SC']
else:
    matplotlib.rcParams['font.sans-serif'] = ['Microsoft YaHei', 'SimHei']
matplotlib.rcParams['axes.unicode_minus'] = False

from PySide6.QtWidgets import (QWidget, QVBoxLayout, QHBoxLayout, QLabel, 
                               QComboBox, QPushButton, QScrollArea, 
                               QGraphicsDropShadowEffect, QFrame, QGridLayout, QApplication, QSizePolicy)
from PySide6.QtCore import Qt
from PySide6.QtGui import QColor, QImage, QStandardItemModel, QStandardItem
from matplotlib.backends.backend_qtagg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.figure import Figure
import mplcursors

class CheckableComboBox(QComboBox):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setModel(QStandardItemModel(self))
        self.view().pressed.connect(self.handleItemPressed)
        self.model().dataChanged.connect(self.updateText)
        self._changed = False
        
        self.setEditable(True)
        self.lineEdit().setReadOnly(True)
        self.lineEdit().installEventFilter(self)
        self.updateText()
        
    def eventFilter(self, widget, event):
        if widget == self.lineEdit() and event.type() == event.Type.MouseButtonRelease:
            self.showPopup()
            return True
        return super().eventFilter(widget, event)
        
    def handleItemPressed(self, index):
        item = self.model().itemFromIndex(index)
        if item and item.isEnabled():
            if item.checkState() == Qt.Checked:
                item.setCheckState(Qt.Unchecked)
            else:
                item.setCheckState(Qt.Checked)
            self._changed = True
            
    def hidePopup(self):
        if not self._changed:
            super().hidePopup()
        self._changed = False

    def updateText(self):
        checked = self.get_checked_items()
        if checked:
            self.lineEdit().setText(", ".join(checked))
        else:
            self.lineEdit().setText("全选 / 未选...")

    def get_checked_items(self):
        checkedItems = []
        for i in range(self.count()):
            item = self.model().item(i)
            if item and item.checkState() == Qt.Checked:
                checkedItems.append(item.text())
        return checkedItems

class FloatingPopup(QWidget):
    """自定义带阴影的高级悬浮提示窗"""
    def __init__(self, source, meth, purpose, parent=None):
        # 必须不带 parent (或作为 ToolTip/Popup) 才能浮出主窗口限制
        super().__init__(None)
        self.setWindowFlags(Qt.ToolTip | Qt.FramelessWindowHint)
        self.setAttribute(Qt.WA_TranslucentBackground)
        
        # 容器
        container = QFrame(self)
        container.setObjectName("PopupContainer")
        container.setStyleSheet("""
            QFrame#PopupContainer {
                background-color: rgba(255, 255, 255, 0.96);
                border: 1px solid #cbd5e0;
                border-radius: 8px;
            }
            QLabel {
                color: #2d3748;
                font-size: 12px;
            }
            QLabel.Title {
                color: #3182ce;
                font-weight: bold;
                font-size: 12px;
            }
        """)
        
        shadow = QGraphicsDropShadowEffect(self)
        shadow.setBlurRadius(15)
        shadow.setColor(QColor(0, 0, 0, 40))
        shadow.setOffset(0, 4)
        container.setGraphicsEffect(shadow)
        
        layout = QVBoxLayout(container)
        layout.setSpacing(6)
        layout.setContentsMargins(15, 15, 15, 15)
        
        def add_item(title, text):
            tl = QLabel(title)
            tl.setProperty("class", "Title")
            vl = QLabel(text)
            vl.setWordWrap(True)
            layout.addWidget(tl)
            layout.addWidget(vl)
            
        add_item("📊 数据来源 (Source)", source)
        add_item("⚙️ 计算逻辑 (Methodology)", meth)
        add_item("🎯 商业用途 (Business Purpose)", purpose)
        
        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(10, 10, 10, 10) # 留出阴影空间
        main_layout.addWidget(container)

class ClipboardHelper:
    """跨平台剪贴板系统，静态类"""
    @staticmethod
    def copy_figure(figure):
        try:
            # 1. 尝试使用 Qt 原生方法
            buf = io.BytesIO()
            figure.savefig(buf, format='png', dpi=150, bbox_inches='tight')
            buf.seek(0)
            
            # 使用 QApplication 实例获取剪贴板
            clipboard = QApplication.clipboard()
            if clipboard:
                image = QImage.fromData(buf.getvalue())
                clipboard.setImage(image)
                return True
        except Exception as e:
            print(f"Qt 原生剪贴板失败: {e}")
            pass

        # 2. 回退机制：根据操作系统选择
        if sys.platform == 'win32':
            try:
                import win32clipboard
                # Windows 环境的特定处理
                buf = io.BytesIO()
                figure.savefig(buf, format='bmp', dpi=150, bbox_inches='tight')
                # 取 BMP 的数据部分（去掉前14个字节的文件头）
                data = buf.getvalue()[14:]
                win32clipboard.OpenClipboard()
                win32clipboard.EmptyClipboard()
                win32clipboard.SetClipboardData(win32clipboard.CF_DIB, data)
                win32clipboard.CloseClipboard()
                return True
            except Exception as e:
                print(f"Windows win32clipboard 回退失败: {e}")
                
        elif sys.platform == 'darwin':
            try:
                # macOS 使用 AppleScript
                import tempfile
                with tempfile.NamedTemporaryFile(suffix='.png', delete=False) as temp_file:
                    temp_path = temp_file.name
                    figure.savefig(temp_path, format='png', dpi=150, bbox_inches='tight')
                
                script = f'''
                    set the clipboard to (read (POSIX file "{temp_path}") as TIFF picture)
                '''
                subprocess.run(['osascript', '-e', script], check=True)
                os.remove(temp_path)
                return True
            except Exception as e:
                print(f"macOS AppleScript 回退失败: {e}")
                
        return False

class EMAManager:
    """管理 EMA 数据读取、缓存与过滤的实体类"""
    def __init__(self):
        self.ema_data = pd.DataFrame()
        
    def load_data(self, file_path=None):
        if file_path and os.path.exists(file_path):
            try:
                # Based on the user's promax.py logic for reading EMA
                df = pd.read_excel(file_path, header=19, usecols=[0, 1, 3, 4])
                df.columns = ['Product', 'Substance', 'Country', 'MAH']
                df['Country'] = df['Country'].astype(str).str.upper().str.strip()
                df['Substance'] = df['Substance'].astype(str).str.upper()
                self.ema_data = df
                return True
            except Exception as e:
                print(f"EMA data load failed: {e}")
                return False
        return False
        
    def get_market_data(self, filters=None):
        # 此处供外界获取 EMA 加工数据
        return self.ema_data

# 2. FilterSidebar (筛选侧边栏)
class FilterSidebar(QWidget):
    """侧边栏过滤器，支持多维度关联筛选"""
    from PySide6.QtCore import Signal
    filter_changed = Signal(dict)
    
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setup_ui()
        
    def setup_ui(self):
        self.layout = QVBoxLayout(self)
        self.layout.setSpacing(10)
        self.layout.setContentsMargins(12, 16, 12, 16)
        
        # 专业深蓝/灰/白 QSS 配色方案
        self.setStyleSheet("""
            QWidget {
                background-color: #ffffff;
                border-right: 1px solid #dcdde1;
            }
            QLabel {
                color: #2c3e50;
                font-weight: bold;
                font-size: 12px;
            }
            QComboBox {
                background-color: #f8f9fa;
                border: 1px solid #ced4da;
                border-radius: 5px;
                padding: 5px;
                color: #2c3e50;
            }
            QPushButton {
                background-color: #1a365d;
                color: white;
                border-radius: 6px;
                padding: 8px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #2c5282;
            }
        """)

        self.api_combo = CheckableComboBox()
        self.market_combo = QComboBox()
        self.date_combo = QComboBox()
        self.dosage_combo = QComboBox()
        # 原研药企业多选框 (可不选)
        self.originator_combo = CheckableComboBox()
        
        self.layout.addWidget(QLabel("通用名单 (API)"))
        self.layout.addWidget(self.api_combo)
        
        self.layout.addWidget(QLabel("市场目标 (Market)"))
        self.layout.addWidget(self.market_combo)
        
        # 快捷按钮：选择全部欧洲经济区
        self.eea_quick_btn = QPushButton("⚡ 全部欧洲经济区 (EEA+UK)")
        self.eea_quick_btn.setStyleSheet(
            "background-color: #2b6cb0; color: white; border-radius: 5px; "
            "padding: 6px; font-size: 11px; font-weight: bold;"
        )
        self.eea_quick_btn.setCursor(Qt.PointingHandCursor)
        self.eea_quick_btn.clicked.connect(self._select_all_eea)
        self.layout.addWidget(self.eea_quick_btn)
        
        self.layout.addWidget(QLabel("年份选择 (Date Range)"))
        self.layout.addWidget(self.date_combo)
        self.layout.addWidget(QLabel("规格剂型 (Pack / Dosage)"))
        self.layout.addWidget(self.dosage_combo)
        
        self.layout.addWidget(QLabel("原研药企业 (Originator, 可不选)"))
        self.layout.addWidget(self.originator_combo)
        
        # 触发按钮
        self.apply_btn = QPushButton("▶ 应用全局筛选")
        self.apply_btn.clicked.connect(self.emit_filters)
        self.layout.addWidget(self.apply_btn)
        
        self.layout.addStretch()
        
        self.clear_cache_btn = QPushButton("清空所有缓存并关闭")
        self.clear_cache_btn.setStyleSheet("background-color: #e53e3e; color: white; border-radius: 6px; padding: 8px; font-weight: bold;")
        self.clear_cache_btn.setCursor(Qt.PointingHandCursor)
        self.clear_cache_btn.clicked.connect(self.clear_cache_and_close)
        
        self.close_btn = QPushButton("不删缓存直接关闭")
        self.close_btn.setStyleSheet("background-color: #718096; color: white; border-radius: 6px; padding: 8px; font-weight: bold;")
        self.close_btn.setCursor(Qt.PointingHandCursor)
        self.close_btn.clicked.connect(self.close_without_deleting)
        
        self.layout.addWidget(self.clear_cache_btn)
        self.layout.addWidget(self.close_btn)
        
    def _select_all_eea(self):
        """快捷选中「欧洲经济区(EEA+UK)」全市场"""
        idx = self.market_combo.findText("【欧洲经济区 (EEA+UK)】")
        if idx >= 0:
            self.market_combo.setCurrentIndex(idx)

    def clear_cache_and_close(self):
        import shutil, os
        from PySide6.QtWidgets import QApplication
        if os.path.exists("Cache"):
            shutil.rmtree("Cache", ignore_errors=True)
        if os.path.exists("TSM_Downloads/raw_files"):
            shutil.rmtree("TSM_Downloads/raw_files", ignore_errors=True)
        QApplication.quit()

    def close_without_deleting(self):
        from PySide6.QtWidgets import QApplication
        QApplication.quit()
        
    def emit_filters(self):
        filters = {
            "api_name": self.api_combo.get_checked_items(),
            "market_region": self.market_combo.currentText(),
            "year": self.date_combo.currentText(),
            "dosage_form": self.dosage_combo.currentText(),
            "originator_companies": self.originator_combo.get_checked_items(),
        }
        self.filter_changed.emit(filters)


# 4. AnalysisCard (基础图表卡片)
class AnalysisCard(QFrame):
    """带阴影/圆角的基础卡片类（充当图表容器），并附带复制按钮与高级提示"""
    def __init__(self, title, source="", meth="", purpose="", parent=None):
        super().__init__(parent)
        self.title = title
        self.source = source
        self.meth = meth
        self.purpose = purpose
        self.popup = None
        self.setup_ui()
        
    def setup_ui(self):
        self.setObjectName("AnalysisCard")
        # 圆角10px结构设计与右上角按钮绝对定位支持
        self.setStyleSheet("""
            QFrame#AnalysisCard {
                background-color: #ffffff;
                border-radius: 10px;
                border: 1px solid #e1e5e8;
            }
            QLabel#CardTitle {
                color: #1a365d;
                font-weight: bold;
                font-size: 15px;
                padding: 5px;
            }
            QPushButton#CopyBtn {
                background-color: transparent;
                color: #718096;
                font-size: 12px;
                padding: 4px;
                border: none;
                border-radius: 4px;
            }
            QPushButton#CopyBtn:hover {
                background-color: #f7fafc;
                color: #2c5282;
            }
            QLabel#InfoBadge {
                color: #a0aec0;
                font-size: 14px;
                font-weight: bold;
                padding: 0 5px;
            }
            QLabel#InfoBadge:hover {
                color: #3182ce;
            }
        """)
        
        # 注入 QGraphicsDropShadowEffect
        shadow = QGraphicsDropShadowEffect(self)
        shadow.setBlurRadius(15)
        shadow.setColor(QColor(0, 0, 0, 25))
        shadow.setOffset(0, 4)
        self.setGraphicsEffect(shadow)
        
        self.layout = QVBoxLayout(self)
        self.layout.setContentsMargins(15, 15, 15, 15)
        
        # 标题栏：包含标题和复制按钮
        title_layout = QHBoxLayout()
        title_layout.setContentsMargins(0, 0, 0, 0)
        
        self.title_label = QLabel(self.title)
        self.title_label.setObjectName("CardTitle")
        title_layout.addWidget(self.title_label)
        
        # 添加入口
        self.info_badge = QLabel("[ i ]")
        self.info_badge.setObjectName("InfoBadge")
        self.info_badge.setCursor(Qt.PointingHandCursor)
        self.info_badge.enterEvent = self.show_popup
        self.info_badge.leaveEvent = self.hide_popup
        title_layout.addWidget(self.info_badge)
        
        title_layout.addStretch()
        
        # 复制按钮，位于右上角
        self.copy_btn = QPushButton("复制图表")
        self.copy_btn.setObjectName("CopyBtn")
        self.copy_btn.setToolTip("点此将图表复制到剪贴板")
        self.copy_btn.setCursor(Qt.PointingHandCursor)
        self.copy_btn.clicked.connect(self.on_copy_clicked)
        title_layout.addWidget(self.copy_btn)
        
        self.layout.addLayout(title_layout)
        
        self.setMinimumHeight(450)
        self.figure = Figure(figsize=(8, 6), dpi=100)
        self.figure.patch.set_facecolor('#ffffff')
        self.canvas = FigureCanvas(self.figure)
        self.layout.addWidget(self.canvas)
        
    def show_popup(self, event):
        if not self.popup:
            self.popup = FloatingPopup(self.source, self.meth, self.purpose)
        
        # 将局部坐标转换为全局屏幕坐标
        pos = self.info_badge.mapToGlobal(self.info_badge.rect().bottomRight())
        self.popup.move(pos.x() + 5, pos.y() + 5)
        self.popup.show()
        
    def hide_popup(self, event):
        if self.popup:
            self.popup.hide()

    def on_copy_clicked(self):
        # 仅执行剪贴板复制逻辑，状态提示外包给上层
        ClipboardHelper.copy_figure(self.figure)
        
    def generate_chart(self, df: pd.DataFrame, filters: dict):
        raise NotImplementedError("子类必须重写该方法")


# 5. GlobalStrategicTierCard (波士顿矩阵)
class GlobalStrategicTierCard(AnalysisCard):
    """策略梯队 (波士顿分析矩阵)"""
    def __init__(self, parent=None):
        super().__init__(
            title="全球战略梯队 (波士顿矩阵)", 
            source="TSM 全球销售底稿数据 (Step3 拆解)",
            meth="以年度销量构筑X轴，预估出厂价格构筑Y轴，全球份额决定气泡尺寸",
            purpose="用于界定我方及竞品产品线处于高量低价还是高价低量的不同竞争层级",
            parent=parent
        )
        
    def generate_chart(self, df: pd.DataFrame, filters: dict):
        self.figure.clear()
        ax = self.figure.add_subplot(111)
        
        if df.empty or 'sales_volume_units' not in df.columns or 'sales_value_usd' not in df.columns:
            ax.text(0.5, 0.5, "有效数据不足", ha='center', va='center', color='#7f8c8d')
            self.canvas.draw()
            return
            
        grp = df.groupby('corporation_name').agg({
            'sales_volume_units': 'sum',
            'sales_value_usd': 'sum'
        })
        grp['factory_price_est'] = (grp['sales_value_usd'] / grp['sales_volume_units'].replace({0: pd.NA}) * 0.3).fillna(0.0)
        grp = grp.nlargest(20, 'sales_volume_units').reset_index()
        
        if grp.empty:
            ax.text(0.5, 0.5, "无可呈现数据", ha='center', va='center', color='#7f8c8d')
            self.canvas.draw()
            return

        # Handle size scaling
        max_usd = grp['sales_value_usd'].max()
        if max_usd > 0:
            sizes = (grp['sales_value_usd'] / max_usd) * 2000 + 50
        else:
            sizes = 150

        scatter = ax.scatter(
            grp['sales_volume_units'], 
            grp['factory_price_est'], 
            s=sizes,
            alpha=0.75, 
            c='#2c5282',
            edgecolors='white'
        )
        
        ax.set_xlabel("年度总销量 (Units)", color='#4a5568')
        ax.set_ylabel("加权预估出厂单价 (USD)", color='#4a5568')
        ax.set_title("【巨头象限分布阵列】Top 20 厂商规模与定价权结构", fontsize=10, color='#1a365d')
        ax.grid(True, linestyle='--', alpha=0.4, color='#a0aec0')
        ax.spines['top'].set_visible(False)
        ax.spines['right'].set_visible(False)
        
        ax.axhline(grp['factory_price_est'].median(), color='gray', linestyle='--', alpha=0.5)
        ax.axvline(grp['sales_volume_units'].median(), color='gray', linestyle='--', alpha=0.5)
        
        # mplcursors 悬停信息
        cursor = mplcursors.cursor(scatter, hover=True)
        @cursor.connect("add")
        def on_add(sel):
            idx = sel.index
            corp = grp.iloc[idx].get('corporation_name', '未知主体')
            vol = grp.iloc[idx]['sales_volume_units']
            price = grp.iloc[idx]['factory_price_est']
            usd = grp.iloc[idx]['sales_value_usd']
            sel.annotation.set_text(f"企业: {corp}\n总销量: {vol:,.0f}\n均价: ${price:,.2f}\n创收: ${usd:,.0f}")
            sel.annotation.get_bbox_patch().set(fc="white", alpha=0.9, ec="#2c5282")
            
        self.figure.tight_layout()
        self.canvas.draw()


# 6. TrendAnalysisCard (量价趋势)
class TrendAnalysisCard(AnalysisCard):
    """整体市场的量价推演趋势图"""
    def __init__(self, parent=None):
        super().__init__(
            title="大盘整体量价走势推演", 
            source="宏观年度截面聚合销量明细库",
            meth="通过按年份对销量求和(Sum)、对预估出厂价求平均(Mean)，以双轴折线平滑连接",
            purpose="观测品类市场的整体膨胀或萎缩周期，作为生命节点(导入、成长、衰退)判断的前置指标",
            parent=parent
        )
        
    def generate_chart(self, df: pd.DataFrame, filters: dict):
        self.figure.clear()
        ax1 = self.figure.add_subplot(111)
        
        if df.empty or 'year' not in df.columns:
            ax1.text(0.5, 0.5, "当前暂无趋势数据", ha='center', va='center', color='#7f8c8d')
            self.canvas.draw()
            return
            
        trend_data = df.groupby('year').agg({
            'sales_volume_units': 'sum',
            'sales_value_usd': 'sum'
        }).reset_index().sort_values('year')
        trend_data['factory_price_est'] = (trend_data['sales_value_usd'] / trend_data['sales_volume_units'].replace({0: pd.NA}) * 0.3).fillna(0.0)
        
        ax2 = ax1.twinx()
        line1 = ax1.plot(trend_data['year'], trend_data['sales_volume_units'], marker='o', color='#2c5282', linewidth=2, label='总销量')
        line2 = ax2.plot(trend_data['year'], trend_data['factory_price_est'], marker='s', color='#718096', linewidth=2, linestyle='--', label='均价')
        
        ax1.set_xlabel("统计年份", color='#4a5568')
        ax1.set_ylabel("总销量 (Units)", color='#2c5282')
        ax2.set_ylabel("平均预估出厂单价 (USD)", color='#718096')
        
        lines = line1 + line2
        labels = [l.get_label() for l in lines]
        ax1.legend(lines, labels, loc='upper center', bbox_to_anchor=(0.5, 1.1), ncol=2, frameon=False)
        
        ax1.grid(True, linestyle=':', alpha=0.5, color='#a0aec0')
        ax1.spines['top'].set_visible(False)
        ax2.spines['top'].set_visible(False)
        
        cursor = mplcursors.cursor(lines, hover=True)
        @cursor.connect("add")
        def on_add(sel):
            val_type = "销量" if sel.artist == line1[0] else "均价"
            sel.annotation.set_text(f"年份: {sel.target[0]:.0f}\n{val_type}: {sel.target[1]:,.2f}")
            sel.annotation.get_bbox_patch().set(fc="white", alpha=0.9, ec="#cbd5e0")
            
        self.figure.tight_layout()
        self.canvas.draw()


# 7. ProductComparisonCard (产品间销量对比柱状图+价格折线图) [新增]
class ProductComparisonCard(AnalysisCard):
    """产品对比与竞争价格结构"""
    def __init__(self, parent=None):
        super().__init__(
            title="产品销量份额与定价结构对比", 
            source="各企业销量与单价分布矩阵",
            meth="头部Top10企业的销量总和展示为柱状图；同时叠加密集算术平均价作为主轴折线图",
            purpose="发现哪些企业是在采取低价倾销换取份额策略，哪些是在保持高溢价的稳盘策略",
            parent=parent
        )
        
    def generate_chart(self, df: pd.DataFrame, filters: dict):
        self.figure.clear()
        ax1 = self.figure.add_subplot(111)
        
        if df.empty or 'corporation_name' not in df.columns:
            ax1.text(0.5, 0.5, "无可对比数据源", ha='center', va='center', color='#7f8c8d')
            self.canvas.draw()
            return

        # 聚合对比分析：柱状图 (销量), 折线图 (平均单价)
        comp_data = df.groupby('corporation_name').agg({
            'sales_volume_units': 'sum',
            'sales_value_usd': 'sum'
        }).reset_index().sort_values('sales_volume_units', ascending=False).head(10)
        comp_data['factory_price_est'] = (comp_data['sales_value_usd'] / comp_data['sales_volume_units'].replace({0: pd.NA}) * 0.3).fillna(0.0)
        
        x = np.arange(len(comp_data))
        ax2 = ax1.twinx()
        
        bars = ax1.bar(x, comp_data['sales_volume_units'], color='#3182ce', alpha=0.9, width=0.5, label='企业总销量')
        line = ax2.plot(x, comp_data['factory_price_est'], color='#2d3748', marker='D', linewidth=2, label='平均定价')
        
        ax1.set_xticks(x)
        ax1.set_xticklabels(comp_data['corporation_name'], rotation=35, ha='right', color='#4a5568')
        ax1.set_ylabel("集团销量 (Units)", color='#3182ce')
        ax2.set_ylabel("预估出厂均价 (USD)", color='#2d3748')
        
        ax1.spines['top'].set_visible(False)
        ax2.spines['top'].set_visible(False)
        
        handles1, labels1 = ax1.get_legend_handles_labels()
        handles2, labels2 = ax2.get_legend_handles_labels()
        ax1.legend(handles1 + handles2, labels1 + labels2, loc='upper right', frameon=False)
        
        cursor = mplcursors.cursor(list(bars.patches) + line, hover=True)
        @cursor.connect("add")
        def on_add(sel):
            idx = int(round(sel.target[0]))
            if 0 <= idx < len(comp_data):
                corp = comp_data.iloc[idx]['corporation_name']
                is_this_bar = type(sel.artist).__name__ == "Rectangle"
                if is_this_bar: 
                    sel.annotation.set_text(f"企业: {corp}\n销量: {sel.target[1]:,.0f}")
                else:
                    sel.annotation.set_text(f"企业: {corp}\n价格: ${sel.target[1]:,.2f}")
                sel.annotation.get_bbox_patch().set(fc="white", alpha=0.9, ec="#a0aec0")
            
        self.figure.tight_layout()
        self.canvas.draw()


# 8. SingleProductTrendCard (单产品连续多年销量/价格折线图) [新增]
class SingleProductTrendCard(AnalysisCard):
    """单独产品多年连续跟踪与演化"""
    def __init__(self, parent=None):
        super().__init__(
            title="单竞争产品生命周期纵深研判", 
            source="动态隔离单一头部企业的发展轨迹",
            meth="提取销量最高的头牌企业，连续追踪其历年的销量变迁（蓝线）与单价跳高（灰线）",
            purpose="通过纵向长周期评估对手生命周期，以推演我们后续产品的接力切入点",
            parent=parent
        )
        
    def generate_chart(self, df: pd.DataFrame, filters: dict):
        self.figure.clear()
        ax1 = self.figure.add_subplot(111)
        
        if df.empty or 'year' not in df.columns:
            ax1.text(0.5, 0.5, "检索无可用区间数据", ha='center', va='center', color='#7f8c8d')
            self.canvas.draw()
            return
            
        # 提取当前 DataFrame 中头部份额最大的产品作为焦点
        target_corp = df['corporation_name'].mode()[0] if not df.empty else "未知"
        single_df = df[df['corporation_name'] == target_corp]

        trend_data = single_df.groupby('year').agg({
            'sales_volume_units': 'sum',
            'sales_value_usd': 'sum'
        }).reset_index().sort_values('year')
        trend_data['factory_price_est'] = (trend_data['sales_value_usd'] / trend_data['sales_volume_units'].replace({0: pd.NA}) * 0.3).fillna(0.0)
        
        ax2 = ax1.twinx()
        
        line1 = ax1.plot(trend_data['year'], trend_data['sales_volume_units'], marker='o', 
                         color='#2b6cb0', linewidth=2.5, label='单品销量')
        line2 = ax2.plot(trend_data['year'], trend_data['factory_price_est'], marker='^', 
                         color='#e2e8f0', markerfacecolor='#4a5568', markeredgecolor='#4a5568', 
                         linewidth=2, linestyle='-.', label='单品均价')
        
        ax1.set_xlabel("统计年份", color='#4a5568')
        ax1.set_ylabel("销量 (Units)", color='#2b6cb0')
        ax2.set_ylabel("价格 (USD)", color='#4a5568')
        ax1.set_title(f"【焦点企业:{target_corp}】 动态推演", fontsize=10, color='#1a365d')
        
        ax1.grid(True, linestyle='-', alpha=0.3, color='#e2e8f0')
        ax1.spines['top'].set_visible(False)
        ax2.spines['top'].set_visible(False)
        
        lines = line1 + line2
        labels = [l.get_label() for l in lines]
        ax1.legend(lines, labels, loc='upper left', frameon=False)
        
        cursor = mplcursors.cursor(lines, hover=True)
        @cursor.connect("add")
        def on_add(sel):
            val_type = "销量" if sel.artist == line1[0] else "价格"
            sel.annotation.set_text(f"年份: {sel.target[0]:.0f}\n{val_type}: {sel.target[1]:,.2f}")
            sel.annotation.get_bbox_patch().set(fc="white", alpha=0.9, ec="#cbd5e0")
            
        self.figure.tight_layout()
        self.canvas.draw()


# ==========================================
# 额外四种分析卡片 (Stub Definitions)
# ==========================================

class CompetitiveLandscapeCard(AnalysisCard):
    """竞争格局 (HHI垄断指数)"""
    def __init__(self, parent=None):
        super().__init__(
            title="市场竞争格局与集中度 (HHI)", 
            source="整体企业市场销量份额计算",
            meth="提取销量 Top5 企业排名，同时基于全景份额百分比的平方和测算 HHI 垄断指数",
            purpose="评估市场是被寡头垄断 (HHI>2500)，还是处于高度分散竞争 (HHI<1500)",
            parent=parent
        )
        
    def generate_chart(self, df: pd.DataFrame, filters: dict):
        self.figure.clear()
        ax = self.figure.add_subplot(111)
        
        if df.empty or 'corporation_name' not in df.columns or 'sales_volume_units' not in df.columns:
            ax.text(0.5, 0.5, "有效数据不足", ha='center', va='center', color='#7f8c8d')
            self.canvas.draw()
            return
            
        corp_sales = df.groupby('corporation_name')['sales_volume_units'].sum()
        total_vol = corp_sales.sum()
        if total_vol <= 0:
            ax.text(0.5, 0.5, "总销量为0无法测算", ha='center', va='center', color='#7f8c8d')
            self.canvas.draw()
            return
            
        shares = corp_sales / total_vol * 100
        hhi = (shares ** 2).sum()
        
        top5 = corp_sales.sort_values(ascending=True).tail(5)
        
        bars = ax.barh(top5.index, top5.values, color='#4299e1', alpha=0.85, height=0.6)
        
        # Annotate values
        for bar in bars:
            width = bar.get_width()
            ax.text(width, bar.get_y() + bar.get_height()/2, f" {width:,.0f}", 
                    ha='left', va='center', fontsize=9, color='#2d3748')
                    
        ax.set_xlabel("企业总销量 (Units)", color='#4a5568')
        ax.spines['top'].set_visible(False)
        ax.spines['right'].set_visible(False)
        ax.grid(True, axis='x', linestyle='--', alpha=0.4)
        
        # Display HHI gauge text as title
        status = "高度垄断" if hhi > 2500 else "中度集中" if hhi > 1500 else "高度竞争分散"
        color = "#e53e3e" if hhi > 2500 else "#d69e2e" if hhi > 1500 else "#38a169"
        ax.set_title(f"【HHI指数: {hhi:,.1f}】 {status}", color=color, fontsize=12, fontweight='bold')
        
        self.figure.tight_layout()
        self.canvas.draw()

class USAPIConsumptionCard(AnalysisCard):
    def __init__(self, parent=None):
        super().__init__(
            title="美国API占比全球API消耗", 
            source="API_KG（公斤）用量聚合",
            meth="不随其他筛选条件变化，提取所选API的（美国API公斤数/全球API公斤数）",
            purpose="直白体现美国市场占全球的消费比重",
            parent=parent
        )
        self.full_df = pd.DataFrame()

    def set_full_df(self, df: pd.DataFrame):
        self.full_df = df

    def generate_chart(self, df: pd.DataFrame, filters: dict):
        self.figure.clear()
        ax = self.figure.add_subplot(111)
        
        # We use full_df to ignore market/year filters
        api_list = filters.get("api_name", [])
        if not api_list:
            ax.text(0.5, 0.5, "请先选择至少一个API", ha='center', va='center', color='#7f8c8d')
            self.canvas.draw()
            return
            
        plot_df = self.full_df[self.full_df["api_name"].isin(api_list)] if 'api_name' in self.full_df.columns else pd.DataFrame()
        if plot_df.empty or 'api_kg' not in plot_df.columns or 'market_region' not in plot_df.columns:
            ax.text(0.5, 0.5, "无足够公斤数据(api_kg)", ha='center', va='center', color='#7f8c8d')
            self.canvas.draw()
            return
            
        us_mask = plot_df["market_region"].str.contains("美国|US|USA|UNITED STATES", case=False, na=False)
        us_kg = plot_df[us_mask]['api_kg'].sum()
        total_kg = plot_df['api_kg'].sum()
        
        if total_kg <= 0:
            ax.text(0.5, 0.5, "所选API总体积为0", ha='center', va='center', color='#7f8c8d')
            self.canvas.draw()
            return
            
        other_kg = total_kg - us_kg
        
        wedges, texts, autotexts = ax.pie(
            [us_kg, other_kg], labels=["美国", "全球其他区域"], 
            autopct=lambda p: f'{p:1.1f}%\n({p/100.*total_kg:,.0f} kg)', 
            colors=['#e53e3e', '#2b6cb0'], startangle=90, wedgeprops={'linewidth': 1, 'edgecolor': 'white'}
        )
        for t in texts: t.set_fontsize(12)
        for at in autotexts: at.set_color('white'); at.set_fontsize(12)
            
        ax.set_title("美国整体占比 (按API消耗公斤数)", fontsize=12, color='#1a365d')
        ax.axis('equal')
        self.figure.tight_layout()
        self.canvas.draw()

class CountryMarketSharePieCard(AnalysisCard):
    """某国家市场占比的饼状图，支持排除原研药"""
    def __init__(self, parent=None):
        super().__init__(
            title="区域市场份额分布", 
            source="筛选后的区域销量数据",
            meth="展示指定市场份额的饼图，支持排除特定企业(包括全部原研厂家)",
            purpose="直观审视在无原研厂或剔除指定厂家条件下的真实市场割据",
            parent=parent
        )
        self.exclude_combo = CheckableComboBox()
        self.exclude_combo.setMinimumWidth(200)
        self.exclude_combo.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        self.exclude_combo.model().dataChanged.connect(self.on_exclude_changed)
        
        self.layout.itemAt(0).layout().insertWidget(1, QLabel("手动排除企业:"))
        self.layout.itemAt(0).layout().insertWidget(2, self.exclude_combo)
        
        # 将默认的 FigureCanvas 包裹到滚动条区域中
        self.layout.removeWidget(self.canvas)
        self.scroll_area = QScrollArea()
        self.scroll_area.setWidgetResizable(True)
        self.scroll_area.setFrameShape(QFrame.NoFrame)
        self.scroll_area.setWidget(self.canvas)
        self.layout.addWidget(self.scroll_area)
        
        self.setMinimumHeight(600)  # 卡片外部高度固定，内部靠 scroll area
        
        self.last_df = pd.DataFrame()
        self.last_filters = {}

    def on_exclude_changed(self, *args):
        if not self.last_df.empty:
            self.generate_chart(self.last_df, self.last_filters, populate_combo=False)

    def generate_chart(self, df: pd.DataFrame, filters: dict, **kwargs):
        populate_combo = kwargs.get('populate_combo', True)
        self.last_df = df
        self.last_filters = filters
        self.figure.clear()

        if df.empty or 'corporation_name' not in df.columns:
            ax = self.figure.add_subplot(111)
            ax.text(0.5, 0.5, "有效数据不足", ha='center', va='center', color='#7f8c8d', fontsize=13)
            self.canvas.draw()
            return

        if populate_combo:
            corps = sorted(df['corporation_name'].dropna().unique().tolist())
            self.exclude_combo.blockSignals(True)
            self.exclude_combo.clear()
            item_orig = QStandardItem("【排除所有原研药】")
            item_orig.setFlags(Qt.ItemIsUserCheckable | Qt.ItemIsEnabled)
            item_orig.setData(Qt.Unchecked, Qt.CheckStateRole)
            self.exclude_combo.model().appendRow(item_orig)
            for c in corps:
                item = QStandardItem(c)
                item.setFlags(Qt.ItemIsUserCheckable | Qt.ItemIsEnabled)
                item.setData(Qt.Unchecked, Qt.CheckStateRole)
                self.exclude_combo.model().appendRow(item)
            self.exclude_combo.updateText()
            self.exclude_combo.blockSignals(False)

            # 如果侧边栏已配置原研药，自动帮勾上「排除所有原研药」
            if filters.get('originator_companies'):
                m = self.exclude_combo.model()
                if m.item(0):
                    m.item(0).setCheckState(Qt.Checked)
                self.exclude_combo.updateText()

        plot_df = df.copy()

        # 1) 侧边栏原研药排除（originator_companies）
        orig_companies = filters.get('originator_companies', [])
        if orig_companies:
            plot_df = plot_df[~plot_df['corporation_name'].isin(orig_companies)]

        # 2) 卡片内部排除 combo
        exclude_corps = self.exclude_combo.get_checked_items()
        if "【排除所有原研药】" in exclude_corps:
            if 'is_originator' in plot_df.columns:
                plot_df = plot_df[plot_df['is_originator'] == False]
        to_drop = [c for c in exclude_corps if c != "【排除所有原研药】"]
        if to_drop:
            plot_df = plot_df[~plot_df['corporation_name'].isin(to_drop)]

        if plot_df.empty:
            ax = self.figure.add_subplot(111)
            ax.text(0.5, 0.5, "排除后无可用数据", ha='center', va='center', color='#7f8c8d', fontsize=13)
            self.canvas.draw()
            return

        # 配色表（最多18种）
        PIE_COLORS = [
            '#2c7bb6','#d7191c','#1a9641','#fdae61','#abd9e9',
            '#f46d43','#74add1','#a50026','#313695','#fee090',
            '#4dac26','#b8e186','#7b3294','#c2a5cf','#008837',
            '#a6dba0','#e66101','#fdb863'
        ]

        mkt = filters.get("market_region", "")
        unique_countries = sorted(plot_df['market_region'].dropna().unique()) if 'market_region' in plot_df.columns else []

        def _build_pie_data(sub_df, max_slices=10):
            """聚合、排序，超出 max_slices 的合并为'其他'"""
            agg = sub_df.groupby('corporation_name')['sales_volume_units'].sum()
            agg = agg[agg > 0].sort_values(ascending=False)
            if agg.empty:
                return pd.Series(dtype=float)
            if len(agg) > max_slices:
                top = agg.head(max_slices)
                others = pd.Series([agg.iloc[max_slices:].sum()], index=['其他/Others'])
                agg = pd.concat([top, others])
            return agg

        def _draw_single_pie(ax_sub, pie_data, title, small=False):
            """在指定 Axes 上绘制带图例的饼图"""
            colors = PIE_COLORS[:len(pie_data)]
            wedges, texts = ax_sub.pie(
                pie_data.values,
                labels=None,          # 不在扇区上写标签
                colors=colors,
                startangle=140,
                wedgeprops={'linewidth': 0.8, 'edgecolor': 'white'},
                pctdistance=0.82
            )
            ax_sub.axis('equal')

            # 图例写在饼图右侧
            total = pie_data.sum()
            legend_labels = [
                f"{name[:22]}: {val/total*100:.1f}%"
                for name, val in zip(pie_data.index, pie_data.values)
            ]
            fontsize = 7 if small else 9
            ax_sub.legend(
                wedges, legend_labels,
                loc='center left',
                bbox_to_anchor=(1.0, 0.5),
                fontsize=fontsize,
                frameon=False
            )
            title_fs = 9 if small else 11
            ax_sub.set_title(title, fontsize=title_fs, color='#1a365d', fontweight='bold', pad=8)

        # ── 多国家 sub-plot 模式 ──
        if len(unique_countries) > 1:
            cols = 3
            rows = (len(unique_countries) + cols - 1) // cols
            
            fig_h = max(6, rows * 4.5)
            self.figure.set_size_inches(13, fig_h)
            self.canvas.setMinimumHeight(int(fig_h * 100))

            for idx, c in enumerate(unique_countries):
                country_df = plot_df[plot_df['market_region'] == c]
                pie_data = _build_pie_data(country_df, max_slices=8)
                if pie_data.empty:
                    continue
                ax_sub = self.figure.add_subplot(rows, cols, idx + 1)
                _draw_single_pie(ax_sub, pie_data, c, small=True)

            self.figure.suptitle(
                "各国家市场装配格局（饼图）",
                fontsize=13, color='#1a365d', fontweight='bold', y=0.98
            )
            self.figure.subplots_adjust(wspace=0.8, hspace=0.4, left=0.05, right=0.95, top=0.9, bottom=0.1)
            self.canvas.draw()
            return

        # ── 单国家 / 全局饼图模式 ──
        pie_data = _build_pie_data(plot_df, max_slices=12)
        if pie_data.empty:
            ax = self.figure.add_subplot(111)
            ax.text(0.5, 0.5, "销量为0，无法绘制饼图", ha='center', va='center', color='#7f8c8d', fontsize=13)
            self.canvas.draw()
            return

        self.figure.set_size_inches(10, 6)
        self.canvas.setMinimumHeight(600) # 单张图恢复默认画布高度保护
        ax = self.figure.add_subplot(111)
        mkt_str = mkt if mkt and mkt not in ["全部", ""] else "全域"
        _draw_single_pie(ax, pie_data, f"【{mkt_str}】市场份额分析", small=False)
        self.figure.subplots_adjust(left=0.1, right=0.6, top=0.9, bottom=0.1)
        self.canvas.draw()

class SKUShareCard(AnalysisCard):
    def __init__(self, parent=None):
        super().__init__("核心 SKU 分布 (前五大企业占比)", parent=parent)
        self.meth = "抽取销量Top5的制药企业，分别抓取各家销量最高的Top5核心规格剂型，展示其相对结构分布"

    def generate_chart(self, df: pd.DataFrame, filters: dict):
        self.figure.clear()
        ax = self.figure.add_subplot(111)
        
        if df.empty or 'corporation_name' not in df.columns or 'strength_raw' not in df.columns:
            ax.text(0.5, 0.5, "缺乏规格明细数据", ha='center', va='center', color='#7f8c8d')
            self.canvas.draw()
            return

        top5_corps = df.groupby('corporation_name')['sales_volume_units'].sum().nlargest(5).index
        if len(top5_corps) == 0:
            ax.text(0.5, 0.5, "无足够销量数据", ha='center', va='center', color='#7f8c8d')
            self.canvas.draw()
            return
            
        plot_df = df[df['corporation_name'].isin(top5_corps)]
        sku_grouped = plot_df.groupby(['corporation_name', 'strength_raw'])['sales_volume_units'].sum().reset_index()
        
        ax.invert_yaxis()
        y_pos = np.arange(len(top5_corps))
        bottom = np.zeros(len(top5_corps))
        cmap = cm.get_cmap('tab20', 20)
        
        for i, corp in enumerate(top5_corps):
            corp_data = sku_grouped[sku_grouped['corporation_name'] == corp]
            top_skus = corp_data.nlargest(5, 'sales_volume_units')
            rest = corp_data[~corp_data['strength_raw'].isin(top_skus['strength_raw'])]
            
            if not rest.empty:
                rest_sum = rest['sales_volume_units'].sum()
                top_skus = pd.concat([top_skus, pd.DataFrame({'corporation_name': [corp], 'strength_raw': ['其他/Other'], 'sales_volume_units': [rest_sum]})])

            top_skus = top_skus.sort_values(by='sales_volume_units', ascending=False)
            
            for j, (_, row) in enumerate(top_skus.iterrows()):
                sku = str(row['strength_raw'])
                val = row['sales_volume_units']
                short_sku = (sku[:15] + '..') if len(sku) > 15 else sku
                
                bar = ax.barh(i, val, left=bottom[i], height=0.6, color=cmap(j), edgecolor='white', alpha=0.9)
                bottom[i] += val
                
                if val > (bottom[i] * 0.05):
                    cx = bottom[i] - (val / 2)
                    ax.text(cx, i, short_sku, ha='center', va='center', color='#2d3748', fontsize=7, rotation=0)

        ax.set_yticks(y_pos)
        ax.set_yticklabels([str(c) for c in top5_corps], color='#2d3748')
        ax.set_xlabel("销量 (Units)", color='#4a5568')
        ax.set_title("Top5 企业核心规格型号分布", fontsize=11, color='#1a365d')
        
        ax.spines['top'].set_visible(False)
        ax.spines['right'].set_visible(False)
        
        self.figure.tight_layout()
        self.canvas.draw()

# 3. Step4DashboardWidget (主仪表盘)


# ==========================================
# 新增的高维度分析卡片 (100种药物海选视图)
# ==========================================

class MarketAttractivenessMatrixCard(AnalysisCard):
    """市场吸引力气泡矩阵 (Market Attractiveness Matrix)"""
    def __init__(self, parent=None):
        super().__init__(
            title="全品种市场吸引力矩阵 (蓝海探测器)", 
            source="多品种销量与增长率交叉",
            meth="X轴为市场规模(Units), Y轴为期间复合增长率(CAGR), 气泡大小为参与竞争企业数",
            purpose="在百大API中快速定位那些市场极大且保持强劲增长，且竞争者相对稀少的蓝海目标",
            parent=parent
        )

    def generate_chart(self, df: pd.DataFrame, filters: dict):
        self.figure.clear()
        ax = self.figure.add_subplot(111)
        if df.empty or 'api_name' not in df.columns or 'year' not in df.columns:
            ax.text(0.5, 0.5, "缺乏必要的数据维度(api_name, year)", ha='center', va='center', color='#7f8c8d')
            self.canvas.draw()
            return
            
        # 计算每个 API 的最早和最晚年份的销量
        api_stats = []
        for api, grp in df.groupby('api_name'):
            years = sorted(grp['year'].unique())
            if len(years) >= 2:
                v_start = grp[grp['year'] == years[0]]['sales_volume_units'].sum()
                v_end = grp[grp['year'] == years[-1]]['sales_volume_units'].sum()
                if v_start > 0 and v_end > 0:
                    cagr = ((v_end / v_start) ** (1 / (years[-1] - years[0])) - 1) * 100
                else:
                    cagr = 0
            else:
                cagr = 0
            
            total_vol = grp['sales_volume_units'].sum()
            comp_count = grp['corporation_name'].nunique() if 'corporation_name' in grp.columns else 1
            if total_vol > 0:
                api_stats.append({'api_name': api, 'vol': total_vol, 'cagr': cagr, 'comp': comp_count})
        
        if not api_stats:
            ax.text(0.5, 0.5, "有效跨年对比数据不足", ha='center', va='center', color='#7f8c8d')
            self.canvas.draw()
            return
            
        stat_df = pd.DataFrame(api_stats)
        
        scatter = ax.scatter(
            stat_df['vol'], stat_df['cagr'], 
            s=stat_df['comp'] * 50 + 20, 
            alpha=0.6, c=stat_df['cagr'], cmap='coolwarm', edgecolors='white'
        )
        
        ax.set_xscale('log') # 销量往往悬殊，用对数轴更好
        if len(stat_df) > 0 and len(stat_df['cagr'].dropna()) > 0:
            c_min = stat_df['cagr'].min()
            c_max = stat_df['cagr'].max()
            ax.set_ylim(min(-10, c_min-10), max(100, c_max + 20))
        
        ax.set_xlabel("市场总销量大小 (Log Scale)", color='#4a5568')
        ax.set_ylabel("复合增长率 CAGR (%)", color='#4a5568')
        ax.grid(True, linestyle='--', alpha=0.3)
        ax.axhline(0, color='gray', linestyle='-', linewidth=1, alpha=0.5)
        ax.spines['top'].set_visible(False)
        ax.spines['right'].set_visible(False)
        
        cursor = mplcursors.cursor(scatter, hover=True)
        @cursor.connect("add")
        def on_add(sel):
            idx = sel.index
            row = stat_df.iloc[idx]
            sel.annotation.set_text(f"API: {row['api_name']}\n总销量: {row['vol']:,.0f}\nCAGR: {row['cagr']:.1f}%\n竞争者数: {row['comp']}")
            sel.annotation.get_bbox_patch().set(fc="white", alpha=0.9, ec="#2c5282")
            
        self.figure.colorbar(scatter, ax=ax, label="CAGR 热度")
        self.figure.tight_layout()
        self.canvas.draw()


class TopGrowthAPIsCard(AnalysisCard):
    """高复合增长率 Top API 排行"""
    def __init__(self, parent=None):
        super().__init__(
            title="高成长明星赛道 Top 15 (增量王)", 
            source="全局品种的期初至期末复合增速换算",
            meth="提取销量基数超过1000的药品，计算多年CAGR后排序提取前15名",
            purpose="捕捉近期突然爆发或稳定快速扩张的黑马品种作为仿制优先选项",
            parent=parent
        )

    def generate_chart(self, df: pd.DataFrame, filters: dict):
        self.figure.clear()
        ax = self.figure.add_subplot(111)
        if df.empty or 'api_name' not in df.columns or 'year' not in df.columns:
            ax.text(0.5, 0.5, "缺乏必要的数据维度(api_name, year)", ha='center', va='center', color='#7f8c8d')
            self.canvas.draw()
            return
            
        api_cagr = []
        for api, grp in df.groupby('api_name'):
            total_vol = grp['sales_volume_units'].sum()
            if total_vol < 1000: continue
                
            years = sorted(grp['year'].unique())
            if len(years) >= 2:
                v_start = grp[grp['year'] == years[0]]['sales_volume_units'].sum()
                v_end = grp[grp['year'] == years[-1]]['sales_volume_units'].sum()
                if v_start > 0 and v_end > 0:
                    cagr = ((v_end / v_start) ** (1 / (years[-1] - years[0])) - 1) * 100
                    api_cagr.append({'api_name': api, 'cagr': cagr})
                    
        if not api_cagr:
            ax.text(0.5, 0.5, "有效增速数据不足", ha='center', va='center', color='#7f8c8d')
            self.canvas.draw()
            return
            
        cagr_df = pd.DataFrame(api_cagr).sort_values('cagr', ascending=False).head(15)
        cagr_df = cagr_df.sort_values('cagr', ascending=True) # barh 倒序画
        
        bars = ax.barh(cagr_df['api_name'], cagr_df['cagr'], color='#dd6b20', alpha=0.85)
        
        for bar in bars:
            width = bar.get_width()
            ax.text(width, bar.get_y() + bar.get_height()/2, f" {width:,.1f}%", 
                    ha='left', va='center', fontsize=8, color='#2d3748')
                    
        ax.set_xlabel("复合年增长率 CAGR (%)", color='#4a5568')
        ax.spines['top'].set_visible(False)
        ax.spines['right'].set_visible(False)
        ax.grid(True, axis='x', linestyle=':', alpha=0.5)
        
        self.figure.tight_layout()
        self.canvas.draw()


class PriceErosionTrendsCard(AnalysisCard):
    """价格侵蚀漏斗分析 / 多品种价格红海追踪"""
    def __init__(self, parent=None):
        super().__init__(
            title="品种集采风暴与价格侵蚀跌幅", 
            source="各 API 多年份均价",
            meth="将所有所选主打 API 第一年的均价强制归一化为100%，以线图展现后续几年的残留价值百分比",
            purpose="如果跌幅极大表明已进入集采肉搏或专利悬崖底端，需警惕低价地狱；平稳线表明拥有价格护城河",
            parent=parent
        )

    def generate_chart(self, df: pd.DataFrame, filters: dict):
        self.figure.clear()
        ax = self.figure.add_subplot(111)
        if df.empty or 'api_name' not in df.columns or 'year' not in df.columns or 'factory_price_est' not in df.columns:
            ax.text(0.5, 0.5, "缺乏均价或年份维度", ha='center', va='center', color='#7f8c8d')
            self.canvas.draw()
            return
            
        top_apis = df.groupby('api_name')['sales_volume_units'].sum().nlargest(10).index
        if len(top_apis) == 0:
            ax.text(0.5, 0.5, "无可呈现品种", ha='center', va='center', color='#7f8c8d')
            self.canvas.draw()
            return
            
        cmap = cm.get_cmap('tab10', 10)
        lines_info = []
        for i, api in enumerate(top_apis):
            grp = df[df['api_name'] == api].groupby('year').agg({'sales_volume_units': 'sum', 'sales_value_usd': 'sum'}).reset_index().sort_values('year')
            grp['factory_price_est'] = (grp['sales_value_usd'] / grp['sales_volume_units'].replace({0: pd.NA}) * 0.3).fillna(0.0)
            if len(grp) >= 2:
                base_price = grp['factory_price_est'].iloc[0]
                if base_price > 0:
                    grp['norm_price'] = (grp['factory_price_est'] / base_price) * 100
                    line, = ax.plot(grp['year'], grp['norm_price'], marker='.', linewidth=2, color=cmap(i), label=(api[:10]+'..'))
                    lines_info.append(line)
        
        if not lines_info:
            ax.text(0.5, 0.5, "无连续年份的价格数据以计算折扣", ha='center', va='center', color='#7f8c8d')
            self.canvas.draw()
            return
            
        ax.axhline(100, color='gray', linestyle='--', alpha=0.7)
        ax.set_xlabel("年份", color='#4a5568')
        ax.set_ylabel("相比基准年相对价格 (%)", color='#4a5568')
        
        years_all = df['year'].dropna().unique()
        if len(years_all) > 0:
            ax.set_xticks(sorted(years_all))
            
        ax.legend(loc='lower left', frameon=False, fontsize=8, ncol=2)
        ax.spines['top'].set_visible(False)
        ax.spines['right'].set_visible(False)
        ax.grid(True, linestyle='-', alpha=0.2)
        
        self.figure.tight_layout()
        self.canvas.draw()


class MarketRegionalDistributionCard(AnalysisCard):
    """全球核心战区（欧美）依赖度面积"""
    def __init__(self, parent=None):
        super().__init__(
            title="Top API 全球战区销量依附度", 
            source="各个 API 按国家拆解销量",
            meth="将数据划归为欧盟大五国、美国和其他，展现绝对体量的横向堆叠条形图",
            purpose="明确该赛道的主要消费区域，避免主攻美国却发现目标药完全属于欧盟特权大品种的尴尬",
            parent=parent
        )

    def generate_chart(self, df: pd.DataFrame, filters: dict):
        self.figure.clear()
        ax = self.figure.add_subplot(111)
        if df.empty or 'api_name' not in df.columns or 'market_region' not in df.columns:
            ax.text(0.5, 0.5, "缺乏市场区域数据", ha='center', va='center', color='#7f8c8d')
            self.canvas.draw()
            return
            
        top_apis = df.groupby('api_name')['sales_volume_units'].sum().nlargest(10).index
        if len(top_apis) == 0:
            ax.text(0.5, 0.5, "无数据", ha='center', va='center', color='#7f8c8d')
            self.canvas.draw()
            return
            
        plot_df = df[df['api_name'].isin(top_apis)].copy()
        
        def classify_region(mkt):
            mkt = str(mkt).upper()
            if any(us in mkt for us in ['US', 'UNITED STATES', '美国']): return '美国 (US)'
            if mkt in [c.upper() for c in core_config.EU_BIG5]: return '欧洲五大国 (EU Big5)'
            return '其他区域 (Others)'
            
        plot_df['mkt_class'] = plot_df['market_region'].apply(classify_region)
        pivot = plot_df.pivot_table(index='api_name', columns='mkt_class', values='sales_volume_units', aggfunc='sum', fill_value=0)
        
        pivot['Total'] = pivot.sum(axis=1)
        pivot = pivot.sort_values('Total', ascending=True).drop(columns=['Total'])
        
        colors = {'美国 (US)': '#e53e3e', '欧洲五大国 (EU Big5)': '#2b6cb0', '其他区域 (Others)': '#a0aec0'}
        cols = [c for c in ['美国 (US)', '欧洲五大国 (EU Big5)', '其他区域 (Others)'] if c in pivot.columns]
        
        if len(cols) > 0 and not pivot.empty:
            pivot[cols].plot(kind='barh', stacked=True, ax=ax, color=[colors.get(c, 'black') for c in cols], width=0.7, alpha=0.85)
            
        ax.set_ylabel("")
        ax.set_xlabel("绝对销量规模 (Units)", color='#4a5568')
        ax.legend(title="", loc='lower right', frameon=False, fontsize=8)
        ax.spines['top'].set_visible(False)
        ax.spines['right'].set_visible(False)
        ax.grid(True, axis='x', linestyle='--', alpha=0.4)
        
        self.figure.tight_layout()
        self.canvas.draw()


class HHICrossAPIComparisonCard(AnalysisCard):
    """赛道垄断横向对比"""
    def __init__(self, parent=None):
        super().__init__(
            title="多赛道寡头格局透视表 (HHI 横向对比)", 
            source="各大 API 的独立 HHI 指数计算",
            meth="为这100多个API单独计算HHI指数，挑选销量前15大的API水平展示其产业集中度",
            purpose="红色极高代表被原研或头号巨头绝对通吃，绿色偏低意味着草根崛起处于一片混战，切入更为轻松",
            parent=parent
        )

    def generate_chart(self, df: pd.DataFrame, filters: dict):
        self.figure.clear()
        ax = self.figure.add_subplot(111)
        if df.empty or 'api_name' not in df.columns or 'corporation_name' not in df.columns:
            ax.text(0.5, 0.5, "缺乏竞争格局数据", ha='center', va='center', color='#7f8c8d')
            self.canvas.draw()
            return
            
        top_apis = df.groupby('api_name')['sales_volume_units'].sum().nlargest(15).index
        if len(top_apis) == 0:
            ax.text(0.5, 0.5, "无数据", ha='center', va='center', color='#7f8c8d')
            self.canvas.draw()
            return
            
        hhi_data = []
        for api in top_apis:
            grp = df[df['api_name'] == api]
            corp_sales = grp.groupby('corporation_name')['sales_volume_units'].sum()
            total_vol = corp_sales.sum()
            if total_vol > 0:
                shares = corp_sales / total_vol * 100
                hhi = (shares ** 2).sum()
                hhi_data.append({'api_name': api, 'hhi': hhi})
                
        if not hhi_data:
            ax.text(0.5, 0.5, "无法计算足够 HHI", ha='center', va='center', color='#7f8c8d')
            self.canvas.draw()
            return
            
        stat_df = pd.DataFrame(hhi_data).sort_values('hhi', ascending=True)
        colors = ['#48bb78' if h < 1500 else '#ecc94b' if h < 2500 else '#f56565' for h in stat_df['hhi']]
        bars = ax.barh(stat_df['api_name'], stat_df['hhi'], color=colors, alpha=0.85)
        
        ax.axvline(1500, color='#38a169', linestyle='--', alpha=0.5, label="充分竞争")
        ax.axvline(2500, color='#e53e3e', linestyle='-.', alpha=0.5, label="高度寡头垄断")
        
        for bar in bars:
            width = bar.get_width()
            ax.text(width, bar.get_y() + bar.get_height()/2, f" {width:,.0f}", 
                    ha='left', va='center', fontsize=8, color='#2d3748')
                    
        ax.set_xlabel("HHI 市场集中度指数", color='#4a5568')
        ax.legend(loc='lower right', frameon=False, fontsize=8)
        ax.spines['top'].set_visible(False)
        ax.spines['right'].set_visible(False)
        
        self.figure.tight_layout()
        self.canvas.draw()

# ==========================================
class Step4DashboardWidget(QWidget):
    """主仪表盘入口，聚合管理与响应"""
    def __init__(self, parent=None):
        super().__init__(parent)
        self.ema_manager = EMAManager()
        self.current_data = pd.DataFrame() 
        self.setup_ui()
        
    def setup_ui(self):
        # 面板级冷色调背景配色 (深蓝/灰/白 体系的基础底色)
        self.setStyleSheet("""
            Step4DashboardWidget {
                background-color: #f4f6f8;
            }
        """)
        
        self.layout = QHBoxLayout(self)
        self.layout.setSpacing(20)
        self.layout.setContentsMargins(20, 20, 20, 20)
        
        # 1. 挂载侧边栏
        self.sidebar = FilterSidebar()
        self.sidebar.filter_changed.connect(self.on_filter_changed)
        self.sidebar.setFixedWidth(280)
        self.layout.addWidget(self.sidebar)
        
        # 2. 右侧创建支持动态扩展高度的 ScrollArea
        self.scroll_area = QScrollArea()
        self.scroll_area.setWidgetResizable(True)
        self.scroll_area.setFrameShape(QFrame.NoFrame)
        self.scroll_area.setStyleSheet("background-color: transparent;")
        
        self.charts_container = QWidget()
        self.charts_container.setStyleSheet("background-color: transparent;")
        
        # 采用网格控制卡片
        self.charts_layout = QGridLayout(self.charts_container)
        self.charts_layout.setSpacing(20)
        self.charts_layout.setContentsMargins(0, 0, 0, 0)
        
        # 初始化上述 四 个分析卡片
        self.card_tier = GlobalStrategicTierCard()
        self.card_trend = TrendAnalysisCard()
        self.card_comparison = ProductComparisonCard()
        self.card_single_trend = SingleProductTrendCard()
        
        # 新增卡片
        self.card_hhi = CompetitiveLandscapeCard()
        self.card_hhi.source = "全量企业市场销量份额"
        self.card_hhi.meth = "计算各企业销量占比百分比平方和生成 HHI 指数"
        self.card_hhi.purpose = "评估市场是被寡头垄断，还是处于高度分散竞争"

        self.card_us_consumption = USAPIConsumptionCard()

        self.card_distributor = CountryMarketSharePieCard()
        self.card_distributor.source = "由左侧指定的各个国家内竞争格局明细数据"
        self.card_distributor.meth = "提取企业在当地的总销量展示分布，同时可以通过手动下拉框排除任何一家厂商（如原研企业）以重新平摊数据"
        self.card_distributor.purpose = "当剔除核心垄断竞品时评估市场的切入空间及剩余对手分布情况"

        self.card_sku = SKUShareCard()
        self.card_sku.source = "统一规格解析层级的明细"
        self.card_sku.meth = "提取头部选手Top 5核心规格型号的绝对销量依赖面"
        self.card_sku.purpose = "为跟进提供研发依据，避免仿制无人问津的冷门边缘规格"
        
        # 布局排列
        self.charts_layout.addWidget(self.card_tier, 0, 0)
        self.charts_layout.addWidget(self.card_trend, 0, 1)
        
        self.charts_layout.addWidget(self.card_comparison, 1, 0)
        self.charts_layout.addWidget(self.card_single_trend, 1, 1)

        self.charts_layout.addWidget(self.card_hhi, 2, 0)
        self.charts_layout.addWidget(self.card_us_consumption, 2, 1)

        self.charts_layout.addWidget(self.card_distributor, 3, 0)
        self.charts_layout.addWidget(self.card_sku, 3, 1)

        self.card_attr = MarketAttractivenessMatrixCard()
        self.card_top_growth = TopGrowthAPIsCard()
        self.card_erosion = PriceErosionTrendsCard()
        self.card_regional = MarketRegionalDistributionCard()
        self.card_hhi_cross = HHICrossAPIComparisonCard()
        
        self.charts_layout.addWidget(self.card_attr, 4, 0)
        self.charts_layout.addWidget(self.card_top_growth, 4, 1)
        self.charts_layout.addWidget(self.card_erosion, 5, 0)
        self.charts_layout.addWidget(self.card_regional, 5, 1)
        self.charts_layout.addWidget(self.card_hhi_cross, 6, 0, 1, 2)
        
        self.scroll_area.setWidget(self.charts_container)
        self.layout.addWidget(self.scroll_area)
        
    def set_dataframe(self, df: pd.DataFrame):
        """挂载前端数据的公有钉子"""
        self.current_data = df
        self.card_us_consumption.set_full_df(self.current_data)

        if not self.current_data.empty:
            # 1) API 多选
            if 'api_name' in self.current_data.columns:
                apis = sorted(self.current_data['api_name'].dropna().unique().tolist())
                self.sidebar.api_combo.clear()
                for a in apis:
                    item = QStandardItem(str(a))
                    item.setFlags(Qt.ItemIsUserCheckable | Qt.ItemIsEnabled)
                    item.setData(Qt.Unchecked, Qt.CheckStateRole)
                    self.sidebar.api_combo.model().appendRow(item)

            # 2) 市场下拉：只展示 EEA+UK+美国，无其他
            market_combo = self.sidebar.market_combo
            market_combo.clear()
            market_combo.addItem("全部")
            market_combo.addItem("【欧洲经济区 (EEA+UK)】")
            market_combo.addItem("【欧洲六大国 (EU Big5+UK)】")
            market_combo.addItem("【美国】")
            market_combo.insertSeparator(market_combo.count())
            if 'market_region' in self.current_data.columns:
                all_mkt = sorted(self.current_data['market_region'].dropna().unique().tolist())
                # 只加入 EEA+UK+US 中已有数据的国家
                for m in all_mkt:
                    if str(m).upper() in ALLOWED_DISPLAY_MARKETS:
                        market_combo.addItem(str(m))

            # 3) 日期和剂型
            if 'year' in self.current_data.columns:
                years = sorted([str(int(x)) for x in self.current_data['year'].dropna().unique()])
                self.sidebar.date_combo.clear()
                self.sidebar.date_combo.addItem("全部")
                self.sidebar.date_combo.addItems(years)

            if 'dosage_form' in self.current_data.columns:
                dosages = sorted([str(x) for x in self.current_data['dosage_form'].dropna().unique()])
                self.sidebar.dosage_combo.clear()
                self.sidebar.dosage_combo.addItem("全部")
                self.sidebar.dosage_combo.addItems(dosages)

            # 4) 原研药企业几个候选（基于 ORIGINATOR_CONFIG 与实际数据进行执行级过滤）
            self.sidebar.originator_combo.clear()
            if 'api_name' in self.current_data.columns and 'corporation_name' in self.current_data.columns:
                avail_corps = self.current_data['corporation_name'].dropna().unique().tolist()
                detected_origs = set()
                for api, kws in core_config.ORIGINATOR_CONFIG.items():
                    for corp in avail_corps:
                        for kw in kws:
                            if kw.upper() in str(corp).upper():
                                detected_origs.add(corp)
                for orig in sorted(detected_origs):
                    item = QStandardItem(orig)
                    item.setFlags(Qt.ItemIsUserCheckable | Qt.ItemIsEnabled)
                    item.setData(Qt.Unchecked, Qt.CheckStateRole)
                    self.sidebar.originator_combo.model().appendRow(item)

    def on_filter_changed(self, filters: dict):
        """中继器处理：拦截 Filter 参数，下发给各个图表系统"""
        if self.current_data.empty:
            return

        filtered_df = self.current_data.copy()

        api_list = filters.get("api_name", [])
        if api_list and "api_name" in filtered_df.columns:
            filtered_df = filtered_df[filtered_df["api_name"].isin(api_list)]

        if filters.get("market_region") and "market_region" in filtered_df.columns and filters["market_region"] not in ["", "全部"]:
            mkt = filters["market_region"]
            if mkt == "【欧洲经济区 (EEA+UK)】":
                filtered_df = filtered_df[filtered_df["market_region"].isin(EEA_AND_UK_MARKETS)]
            elif mkt == "【欧洲六大国 (EU Big5+UK)】":
                filtered_df = filtered_df[filtered_df["market_region"].isin(core_config.EU_BIG5 + ["UNITED KINGDOM"])]
            elif mkt == "【美国】":
                filtered_df = filtered_df[filtered_df["market_region"].isin(US_MARKETS)]
            else:
                filtered_df = filtered_df[filtered_df["market_region"] == mkt]

        if filters.get("year") and "year" in filtered_df.columns and filters["year"] not in ["", "全部"]:
            try:
                filtered_df = filtered_df[filtered_df["year"] == int(filters["year"])]
            except ValueError:
                pass

        if filters.get("dosage_form") and "dosage_form" in filtered_df.columns and filters["dosage_form"] not in ["", "全部"]:
            filtered_df = filtered_df[filtered_df["dosage_form"] == filters["dosage_form"]]

        # 同步重绘所有子分析卡片
        self.card_tier.generate_chart(filtered_df, filters)
        self.card_trend.generate_chart(filtered_df, filters)
        self.card_comparison.generate_chart(filtered_df, filters)
        self.card_single_trend.generate_chart(filtered_df, filters)
        self.card_hhi.generate_chart(filtered_df, filters)
        self.card_us_consumption.generate_chart(filtered_df, filters)
        self.card_distributor.generate_chart(filtered_df, filters)
        self.card_sku.generate_chart(filtered_df, filters)
        self.card_attr.generate_chart(filtered_df, filters)
        self.card_top_growth.generate_chart(filtered_df, filters)
        self.card_erosion.generate_chart(filtered_df, filters)
        self.card_regional.generate_chart(filtered_df, filters)
        self.card_hhi_cross.generate_chart(filtered_df, filters)


# ==========================================
# 提取自 欧洲市场预测promax.py 的分析引擎引擎配置库
# ==========================================
MARKER_MAP = {
    'o': '●', # Circle
    's': '■', # Square
    '^': '▲', # Triangle Up
    'D': '◆', # Diamond
    'v': '▼', # Triangle Down
    '<': '◀', 
    '>': '▶',
    'p': '⬟', 
    '*': '★', 
    'h': '⬢', 
    'X': '✖',
    'd': '♦'
}
MARKERS_LIST = list(MARKER_MAP.keys())

# 配色
DISTINCT_COLORS = [
    '#1f77b4', '#ff7f0e', '#2ca02c', '#d62728', '#9467bd', 
    '#8c564b', '#e377c2', '#7f7f7f', '#bcbd22', '#17becf',
    '#000080', '#800000', '#008080', '#808000', '#000000'
]

# 咨询风格配色
COLORS = {
    'primary': '#0052CC', 'secondary': '#36B37E', 'danger': '#FF5630', 
    'warning': '#FFAB00', 'neutral': '#6B778C', 
    'tiers': {
        'Tier 1': '#0052CC', 'Tier 2': '#36B37E', 
        'Tier 3': '#FFAB00', 'Tier 4': '#FF5630', 'Wait': '#6B778C'
    }
}

SALES_CACHE_FILE = "sales_data_cache.pkl"
EMA_CACHE_FILE = "ema_data_cache.pkl"

Target_Markets = [
    "UNITED KINGDOM", "GERMANY", "FRANCE", "ITALY", "SPAIN", 
    "NETHERLANDS", "BELGIUM", "SWEDEN", "SWITZERLAND", "AUSTRIA", 
    "POLAND", "PORTUGAL", "GREECE", "CZECH REPUBLIC", "HUNGARY",
    "ROMANIA", "IRELAND", "DENMARK", "FINLAND", "NORWAY"
]
EU5 = ["UNITED KINGDOM", "GERMANY", "FRANCE", "ITALY", "SPAIN"]

# ==========================================
# 2. 原研药知识库 (Originator Config)
# ==========================================
ORIGINATOR_CONFIG = {
    "APIXABAN": ["BRISTOL", "PFIZER", "SQUIBB"],
    "CITALOPRAM": ["LUNDBECK"],
    "RIVAROXABAN": ["BAYER"],
    "TADALAFIL": ["LILLY", "ICOS", "GLAXO"],
    "TICAGRELOR": ["ASTRAZENECA"],
    "TOBISILATE": ["ETHAMSYLATE"], 
}

# ==========================================
# 2.5 欧洲国家市场特征词典 (Country Profiles)
# ==========================================
COUNTRY_PROFILES = {
    "GERMANY": {
        "Type": "高度集约，保险主导 (Sickness Funds)",
        "Mechanism": "贴牌折扣协议 (Discount Agreements)",
        "Substitution": "药房强制替换为签订了折扣协议的仿制药",
        "Insight": "极度看重价格，赢者通吃格局(Tender驱动)。一旦失去折扣协议，本地份额几乎归零。新玩家需在两年一次的招标期抓住断供机会，通常是大玩家(Sandoz/Ratiopharm/Aliud/1A)的战场。"
    },
    "UNITED KINGDOM": {
        "Type": "自由定价但有利润封顶，药房主导",
        "Mechanism": "药房利润分成限制 (Drug Tariff)",
        "Substitution": "医生按INN(通用名)开处方，药店自行选择采购品牌",
        "Insight": "市场极度分散。药房会选择进货价最低、折扣最高的品牌以赚取最大差价。关系网在批发商(Wholesalers)层面最重要，而不是医生。"
    },
    "SWEDEN": {
        "Type": "单一支付方招标 (Monthly Tender)",
        "Mechanism": "每月大洗牌，赢者拿走几乎100%增量",
        "Substitution": "强制最高报销范围内最便宜的仿制药 (PvA)",
        "Insight": "极其极致的唯价格论市场。由于是定期重新竞标，一旦“Product of the Month”中标当月通吃，落标则销量暴跌。小公司往往靠极低价突袭。"
    },
    "FRANCE": {
        "Type": "内部参考定价，仿制药群体目标",
        "Mechanism": "参考组定价 (Tarif Forfaitaire de Responsabilité - TFR)",
        "Substitution": "药师有替代权，但患者需补差价如果选择贵药（近期政策更严）",
        "Insight": "法国人喜欢“国货”(如 Biogaran, Zentiva)。仿制药渗透率曾被认为较低，但政府正强制推高。首仿有先发优势，价格战不如德国惨烈，通常温和降价。"
    },
    "ITALY": {
        "Type": "区域化采购，原研偏好严重",
        "Mechanism": "大区公立医院招标 + 零售药房参考价",
        "Substitution": "药师可替换，但原研药品牌忠诚度极高",
        "Insight": "意大利是典型的“品牌仿制药”(Branded Generics)市场。患者和医生特别认牌子，导致纯靠低价很难立刻撬动份额，销售队伍的客情维护依然不可或缺。"
    },
    "SPAIN": {
        "Type": "快速价格联动大区制度",
        "Mechanism": "参考价格系统 (Sistema de Precios de Referencia)",
        "Substitution": "平价原则，所有药必须降到同组最低价才能报销",
        "Insight": "一旦有首仿以极低价上市，原研药和所有现有品牌必须在极短时间内跟着降价到同一水平。这导致大家最后拼的是药店的隐性折扣(Bonus)和供货稳定性。"
    },
    "POLAND": {
        "Type": "患者自付比例较高，品牌认知导向",
        "Mechanism": "报销清单与固定报销额",
        "Substitution": "药师必须告知最便宜的替代药，但患者有决定权",
        "Insight": "作为东欧大国，价格敏感度高，但患者对本地传统大厂（如 Polpharma, Adamed）极具信任。外来者往往需要找本地Top代理商做渠道包销。"
    },
    "NETHERLANDS": {
        "Type": "极端的保险公司偏好系统 (Preference Policy)",
        "Mechanism": "荷兰四大医保巨头掌握绝对话语权",
        "Substitution": "只报销保险公司指定的首选药物 (Preference Drug)",
        "Insight": "比德国还极端的“寡头采买”。一旦被最大的几家保险单子剔除，就无药可救。荷兰几乎消灭了所有的温和药房溢价空间。"
    }
}



class AnalysisEngineV24:
    def __init__(self):
        self.selected_companies_for_originator = []
        self.current_countries = []
        self.highlighted_country = None
        self.df_sales_clean = None
        self.pie_batch_index = 0
        self.pie_batch_size = 6

    def draw_summary_table(self, fig, df):
        ax = fig.add_subplot(111)
        ax.axis('off')
        
        data = []
        for c in self.current_countries:
            sub = df[df['国家']==c]
            total_vol = sub['Volume_Clean'].sum()
            total_sales = sub['Sales_Clean'].sum()
            avg_price = (total_sales / total_vol * 0.3) if total_vol > 0 else 0
            
            orig = sub[sub['集团/企业'].isin(self.selected_companies_for_originator)]['Sales_Clean'].sum()
            orig_share = orig / total_sales if total_sales > 0 else 0
            data.append({'C': c, 'Vol': total_vol, 'Price': avg_price, 'Orig': orig_share})
            
        df_metrics = pd.DataFrame(data)
        if df_metrics.empty: return "无数据"
        
        med_vol = df_metrics['Vol'].median()
        med_price = df_metrics['Price'].median()
        
        def get_tier(row):
            if row['Orig'] > 0.8: return "Tier 4"
            if row['Vol'] > med_vol and row['Price'] > med_price: return "Tier 1"
            if row['Vol'] > med_vol: return "Tier 2"
            if row['Price'] > med_price: return "Tier 3"
            return "Wait"
            
        df_metrics['Tier'] = df_metrics.apply(get_tier, axis=1)
        df_metrics = df_metrics.sort_values(['Vol', 'Price'], ascending=False)
        
        # 移除表头
        table_data = [] # No Header
        cell_colors = []
        
        for _, r in df_metrics.iterrows():
            row_c = [r['C'], r['Tier'], f"{r['Vol']/1e6:.1f}M", f"{r['Price']:.2f}", f"{r['Orig']:.0%}"]
            table_data.append(row_c)
            base_c = COLORS['tiers'].get(r['Tier'], "#FFFFFF")
            # Transparent colors for table
            if r['Tier'] == "Tier 1": base_c = "#E3F2FD" 
            if r['Tier'] == "Tier 2": base_c = "#E8F5E9"
            if r['Tier'] == "Tier 3": base_c = "#FFF3E0"
            if r['Tier'] == "Tier 4": base_c = "#FFEBEE"
            if r['Tier'] == "Wait": base_c = "#F5F5F5"
            cell_colors.append([base_c]*5)
            
        # 调整表格样式
        the_table = ax.table(cellText=table_data, cellColours=cell_colors, loc='center', cellLoc='left', colWidths=[0.25, 0.2, 0.2, 0.15, 0.2])
        the_table.auto_set_font_size(False); the_table.set_fontsize(10); the_table.scale(1, 1.8)
        
        text = "【战略决策建议】\n"
        t1 = df_metrics[df_metrics['Tier'].str.contains("Tier 1")]['C'].tolist()
        t4 = df_metrics[df_metrics['Tier'].str.contains("Tier 4")]['C'].tolist()
        text += f"• 核心首选 (Tier 1): {', '.join(t1)} (量大价优)\n"
        if t4: text += f"• 战略机会 (Tier 4): {', '.join(t4)} (原研垄断)\n\n"
        text += "📊 【数据口径】: 抽取最新周期的整体销量(Volume)和预估出厂价(Price)。原研占比反映了理论替代空间。\n"
        text += f"💡 【汇报话术】: “老板，基于量价双优的原则，{', '.join(t1[:2])}等国是我们的Tier-1核心票仓；而像{', '.join(t4[:2])}虽然目前量不大，但是原研药占据绝对主导，缺乏仿制药竞争，存在切入蓝海的巨大降维打击空间。”"
        return text

    def draw_trend(self, fig, df):
        ax1 = fig.add_subplot(211); ax2 = fig.add_subplot(212, sharex=ax1)
        df_sub = df[df['国家'].isin(self.current_countries)]
        trend = df_sub.groupby(['国家','年份']).apply(lambda x: pd.Series({
            'Vol': x['Volume_Clean'].sum(), 
            'Price': (x['Sales_Clean'].sum()/x['Volume_Clean'].sum()*0.3) if x['Volume_Clean'].sum()>0 else 0
        })).reset_index()
        trend['年份'] = trend['年份'].astype(str).str.replace(r'\.0$', '', regex=True)
        
        labels1 = []
        
        for i, c in enumerate(self.current_countries):
            sub = trend[trend['国家']==c].sort_values('年份')
            if sub.empty: continue
            yrs = sub['年份'].tolist(); vols = sub['Vol'].tolist(); prs = sub['Price'].tolist()
            
            # Calculate CAGR for Vol
            cagr_label = ""
            if len(vols) >= 2 and vols[0] > 0:
                n = len(vols) - 1
                cagr = (vols[-1] / vols[0]) ** (1/n) - 1
                if cagr > 0.15: stage = "快速放量期"
                elif cagr > 0: stage = "导入/平台期"
                else: stage = "衰退期"
                cagr_label = f" ({stage} {cagr:.0%})"

            color = DISTINCT_COLORS[i % len(DISTINCT_COLORS)]
            marker = MARKERS_LIST[i % len(MARKERS_LIST)]
            
            # Highlight logic
            alpha = 1.0
            lw = 2
            if self.highlighted_country and c != self.highlighted_country:
                alpha = 0.1 
                lw = 1
            elif self.highlighted_country == c:
                lw = 4
            
            kw = {'color': color, 'marker': marker, 'markersize': 6, 'linewidth': lw, 'alpha': alpha}
            
            ax1.plot(yrs, vols, label=c + cagr_label, **kw)
            ax2.plot(yrs, prs, label=c + cagr_label, **kw)
            labels1.append(c)
        
        ax1.set_title("Volume Trend (Raw YTD)"); ax2.set_title("Price Trend")
        
        # Draw legend on ax1 and make pickable
        leg = ax1.legend(loc='center left', bbox_to_anchor=(1.02, 0.5), fontsize=8, frameon=False)
        leg_map = {}
        if leg:
            for legline, country in zip(leg.get_lines(), labels1):
                legline.set_picker(True)
                legline.set_pickradius(10)
                leg_map[legline] = country
                
        fig.subplots_adjust(left=0.1, right=0.7, top=0.9, bottom=0.1)
        
        text = "【图表解读】\n   展示历史销量与价格走势。右侧图例附带自动生命周期判定。点击图例可高亮对应国家。\n"
        text += "📊 【数据口径】: 取用历年 Volume_Clean 汇总，价格基于销售额/数量折算。\n"
        text += "💡 【汇报话术】: “老板，看左侧的价格曲线是否已经走平。如果价格已经进入‘L型’底部（平台期），说明价格战基本见底，是我们可以带着成本优势吃利润的安全区。重点关注标为【快速放量期】的国家。”"
        return text, leg_map

    def draw_penetration(self, fig, df):
        ax = fig.add_subplot(111)
        res = []
        opps = []
        for c in self.current_countries:
            sub = df[df['国家']==c]
            total = sub['Sales_Clean'].sum()
            orig = sub[sub['集团/企业'].isin(self.selected_companies_for_originator)]['Sales_Clean'].sum()
            if total>0:
                gen_share = (total-orig)/total*100
                res.append({'C':c, 'G':gen_share, 'O':orig/total*100})
                if gen_share < 20: opps.append(f"{c}(原研{orig/total:.0%})")
        
        if not res: return "无数据"
        pd.DataFrame(res).set_index('C').sort_values('G').plot(kind='barh', stacked=True, ax=ax, color=[COLORS['secondary'], COLORS['danger']])
        ax.legend(["Generic", "Originator"], bbox_to_anchor=(1.02, 1), loc='upper left')
        fig.subplots_adjust(right=0.8)
        
        text = "【图表解读】\n   红色=原研药占比，绿色=仿制药占比。\n"
        if opps: text += f"   • 最佳替代切入点: {', '.join(opps)}。\n\n"
        text += "📊 【数据口径】: 匹配 Originator Config 配置中的企业作为原研药。计算结果为纯销售额的份额占比。\n"
        text += "💡 【汇报话术】: “老板，红色柱子越长的地方，原研品牌粘性越强（如意大利）。这类市场我们需要找强势本地代理商；红色非常短的市场，说明已经是纯粹的红海大混战了。”"
        return text

    def draw_pie_batch(self, fig, df, countries):
        if not countries:
            return "无数据"
        cols = 3
        rows = (len(countries) + cols - 1) // cols
        fig.set_size_inches(12, max(6, rows * 5))
        axes = fig.subplots(rows, cols).flatten()
        fig.subplots_adjust(wspace=0.6, hspace=0.8)
        
        leaders = []
        for i, ax in enumerate(axes):
            if i < len(countries):
                c = countries[i]
                sub = df[df['国家']==c]
                d = sub.groupby('集团/企业')['Sales_Clean'].sum().sort_values(ascending=False)
                if d.sum() > 0:
                    total_sales = d.sum()
                    shares = d / total_sales * 100
                    hhi = (shares ** 2).sum()
                    top1 = shares.iloc[0] / 100
                    
                    leaders.append(f"{c}(Top1 {top1:.0%})")
                    # Keep Top 5 + Others logic
                    if len(d)>5: d = pd.concat([d.iloc[:5], pd.Series([d.iloc[5:].sum()], index=['Others'])])
                    
                    wedges, texts, autotexts = ax.pie(d, labels=None, autopct='%1.0f%%', textprops={'fontsize':8}, radius=0.7)
                    ax.axis('equal')
                    # Add legend
                    legend_labels = [f"{label} ({val/d.sum()*100:.1f}%)" for label, val in zip(d.index, d.values)]
                    ax.legend(wedges, legend_labels, loc="center left", bbox_to_anchor=(0.95, 0.5), fontsize=7, frameon=False)
                    
                    title = f"{c}\n(HHI: {hhi:.0f})"
                    ax.set_title(title, pad=15)
                    
                    if hhi > 2500:
                        ax.text(0, -1.3, "高度寡头垄断", color='white', ha='center', fontsize=8, fontweight='bold', bbox=dict(facecolor='red', alpha=0.8, edgecolor='none', boxstyle='round,pad=0.3'))
                    elif hhi < 1500:
                        ax.text(0, -1.3, "市场分散，存在机会", color='white', ha='center', fontsize=8, fontweight='bold', bbox=dict(facecolor='green', alpha=0.8, edgecolor='none', boxstyle='round,pad=0.3'))
                else: ax.axis('off')
            else: ax.axis('off')
            
        fig.subplots_adjust(wspace=0.6, hspace=0.8)
        
        text = f"【仿制药格局（已剔除原研药干扰）】\n   • 领头羊: {'; '.join(leaders[:3])}。\n\n"
        text += "📊 【数据口径】: 仅计算非原研药企业的存量博弈份额。HHI(市场集中度)指数 = 仿制药企业市占率平方和。\n"
        text += "💡 【汇报话术】: “老板，HHI > 2500的地方（标红）代表已经形成了极强的寡头垄断，比如有本地大厂控盘，贸然进入大概率当炮灰；我们挑 HHI < 1500 的‘绿色’分散市场做横向盘整最有利可图。”"
        return text

    def draw_dynamic_matrix(self, fig, df, x_metric, y_metric):
        ax = fig.add_subplot(111)
        data = []
        df_sub = df[df['国家'].isin(self.current_countries)]
        
        for c in self.current_countries:
            sub = df_sub[df_sub['国家']==c]
            if sub.empty: continue
            
            tot_sales = sub['Sales_Clean'].sum()
            tot_vol = sub['Volume_Clean'].sum()
            
            orig_sales = sub[sub['集团/企业'].isin(self.selected_companies_for_originator)]['Sales_Clean'].sum()
            orig_share = orig_sales / tot_sales if tot_sales > 0 else 0
            
            gen_sub = sub[~sub['集团/企业'].isin(self.selected_companies_for_originator)]
            comps = gen_sub['集团/企业'].nunique()
            gen_sales = gen_sub['Sales_Clean'].sum()
            if gen_sales > 0:
                shares = gen_sub.groupby('集团/企业')['Sales_Clean'].sum() / gen_sales * 100
                hhi = (shares**2).sum()
            else:
                hhi = 0
                
            ts = sub.groupby('年份').apply(lambda x: pd.Series({
                'V': x['Volume_Clean'].sum(),
                'P': x['Sales_Clean'].sum()/x['Volume_Clean'].sum() if x['Volume_Clean'].sum()>0 else 0
            })).sort_index()
            
            cagr = 0; price_drop = 0
            if len(ts) >= 2:
                v0 = ts.iloc[0]['V']; v1 = ts.iloc[-1]['V']
                p0 = ts.iloc[0]['P']; p1 = ts.iloc[-1]['P']
                if v0 > 0: cagr = (v1/v0)**(1/(len(ts)-1)) - 1
                if p0 > 0: price_drop = (p0 - p1) / p0  # Positive means dropped
                
            data.append({
                'C': c,
                '市场绝对规模': tot_sales,
                '原研替代空间': 1 - orig_share,
                '价格降幅': price_drop,
                '市场增速': cagr,
                '竞争者数量': comps,
                '集中度HHI': hhi,
                'Vol': tot_vol # bubble size
            })
            
        df_mat = pd.DataFrame(data)
        if df_mat.empty: return "无数据"
        
        x = df_mat[x_metric]
        y = df_mat[y_metric]
        sizes = df_mat['Vol']
        sf = 2000 / sizes.max() if sizes.max() > 0 else 1
        
        self._matrix_countries = df_mat['C'].tolist()
        
        # Make sure sizes are at least visible
        s_arr = [max(20, v) for v in sizes*sf]
        scatter = ax.scatter(x, y, s=s_arr, alpha=0.6, edgecolors='k', picker=True)
        
        for i, r in df_mat.iterrows():
            ax.text(r[x_metric], r[y_metric], r['C'], fontsize=9, ha='center', va='center')
            
        ax.axvline(x.median(), ls=':', color='gray')
        ax.axhline(y.median(), ls=':', color='gray')
        ax.set_xlabel(x_metric)
        ax.set_ylabel(y_metric)
        
        if x_metric in ["原研替代空间", "价格降幅", "市场增速"]:
            ax.xaxis.set_major_formatter(plt.FuncFormatter(lambda val, pos: f"{val:.0%}"))
        if y_metric in ["原研替代空间", "价格降幅", "市场增速"]:
            ax.yaxis.set_major_formatter(plt.FuncFormatter(lambda val, pos: f"{val:.0%}"))
            
        text = f"【多因素矩阵】  X轴: {x_metric} | Y轴: {y_metric} | 气泡大小: 整体销量\n\n"
        text += "📊 【数据口径】: X和Y轴数据由上方控件动态生成，自动过滤空数据。\n"
        text += "💡 【汇报话术】: “老板，这个图是我们的战略导航。选【规模】对比【增速】，能帮我们找明星级重磅市场；选【原研替代空间】对比【竞争者数量】能帮我们寻找降维打击的‘蓝海’；直接点击这上面的重磅国家气泡，就能自动跳到该国的对手盘进行尽调（Drill-down）！”"
        return text
    def draw_prediction(self, fig, df, sim_price=0.4, sim_hhi=1200):
        # 原版混合预测Hybrid
        ax = fig.add_subplot(111)
        df_sub = df[df['国家'].isin(self.current_countries)]
        summary = []
        res_data = []
        
        for c in self.current_countries:
            sub = df_sub[df_sub['国家']==c].groupby('年份').apply(lambda x: pd.Series({
                'Vol': x['Volume_Clean'].sum()
            })).reset_index().sort_values('年份')
            
            if len(sub) < 3: continue
            try:
                x = sub['年份'].astype(str).str.extract(r'(\d{4})')[0].astype(int)
                y = sub['Vol']
                
                # Hybrid Prediction Logic
                x_fit = x[x < 2025]; y_fit = y[x < 2025]
                
                if len(x_fit) >= 2:
                    slope, intercept, r_val, _, _ = stats.linregress(x_fit, y_fit)
                    pred_linear = slope * 2025 + intercept
                    actual_2025_h1 = y[x==2025].values[0] if 2025 in x.values else 0
                    pred_runrate = actual_2025_h1 * 2.0
                    r_sq = r_val ** 2
                    weight_linear = 0.7 if r_sq > 0.8 else 0.3
                    final_pred = weight_linear * pred_linear + (1 - weight_linear) * pred_runrate
                    
                    if final_pred > 0:
                        pace = actual_2025_h1 / final_pred
                        res_data.append({'C': c, 'Pace': pace * 100})
                        status = "超预期" if pace > 0.55 else "滞后"
                        summary.append(f"• {c}: 预测{final_pred/1e6:.1f}M，当前{status}({pace:.0%})")
            except: pass

        if res_data:
            df_p = pd.DataFrame(res_data).set_index('C').sort_values('Pace')
            # 颜色逻辑：大于50%绿色，小于50%橙色
            colors = [COLORS['secondary'] if p > 50 else COLORS['warning'] for p in df_p['Pace']]
            df_p['Pace'].plot(kind='barh', ax=ax, color=colors)
            
            max_p = df_p['Pace'].max()
            ax.set_xlim(0, max(100, max_p * 1.1))
            
            for container in ax.containers:
                ax.bar_label(container, fmt='%.0f%%', padding=3)
                
            ax.axvline(50, ls='--', color='red', label='H1 Target (50%)')
            ax.set_title("2025 Forecast Achievement (Hybrid Model: Trend + RunRate)")
            ax.legend()
            
        fig.subplots_adjust(right=0.8)
        text = "【智能预测】 基于历史趋势与H1跑率的混合预测。红色虚线为50%半年度达标线。\n"
        text += "\n".join(summary)
        
        text += f"\n\n【What-If 情景模拟】 假定定价为原研药的 {sim_price:.0%}，目标环境 HHI={sim_hhi}:\n"
        expected_share = (1 - sim_price) * (2000 / max(sim_hhi, 500)) * 0.15
        expected_share = max(0.0, float(min(expected_share, 0.45)))
        text += f"   ➤ 预计在上述环境下，本产品第3年市场份额可达: {expected_share:.1%}\n\n"
        text += "💡 【汇报话术】: “老板，按照上面我拖拽设定的激进价格打法（原研价折让率），结合该国的垄断难度（HHI滑块），系统用测算公式预估我们第三年能抢下这么多份额，这就给我们明年制定欧洲整体销量盘子（Budget）托住了底。”"
        
        return text

    def draw_bubble_matrix(self, fig, df):
        ax = fig.add_subplot(111)
        df_sub = df[df['国家'].isin(self.current_countries)]
        
        data = []
        for c in self.current_countries:
            sub = df_sub[df_sub['国家']==c]
            if sub.empty: continue
            
            tot_vol = sub['Volume_Clean'].sum()
            tot_sales = sub['Sales_Clean'].sum()
            avg_price = (tot_sales / tot_vol * 0.3) if tot_vol > 0 else 0
            
            data.append({
                'Country': c,
                'Volume': tot_vol,
                'Sales': tot_sales,
                'Price': avg_price
            })
            
        if not data:
            ax.text(0.5, 0.5, "无数据", ha='center', va='center', color='gray')
            return "无数据"
            
        df_p = pd.DataFrame(data)
        
        x = df_p['Price']
        y = df_p['Volume']
        sizes = df_p['Sales']
        
        max_size = sizes.max()
        sf = 3000 / max_size if max_size > 0 else 1
        s_arr = [max(50, v) for v in sizes * sf]
        
        # Color based on distinct colors for pretty visualization
        colors = [DISTINCT_COLORS[i % len(DISTINCT_COLORS)] for i in range(len(df_p))]
        
        scatter = ax.scatter(x, y, s=s_arr, alpha=0.6, c=colors, edgecolors='white', picker=True)
        
        for i, r in df_p.iterrows():
            ax.text(r['Price'], r['Volume'], r['Country'], fontsize=9, ha='center', va='center')
            
        med_x = x.median() if not x.empty else 0
        med_y = y.median() if not y.empty else 0
        
        ax.axvline(med_x, ls='--', color='gray', alpha=0.7, label=f'Price Median ({med_x:.2f})')
        ax.axhline(med_y, ls='--', color='gray', alpha=0.7, label=f'Volume Median ({med_y:,.0f})')
        
        ax.set_xlabel("预估出厂单价 (Factory Price, USD)")
        ax.set_ylabel("总销量 (Volume, Units)")
        ax.set_title("核心市场战略机遇矩阵 (气泡大小: 销售额)")
        ax.grid(True, linestyle=':', alpha=0.4)
        ax.legend(bbox_to_anchor=(1.02, 1), loc='upper left')
        
        fig.subplots_adjust(bottom=0.15, right=0.8)
        text = "【市场预测气泡矩阵】 横坐标: 出厂均价 | 纵坐标: 销量 | 气泡大小: 总体销售额\n\n"
        text += "📊 【数据口径】: 抽取所有历史总销量与对应出厂单价截面进行投影映射。虚线为中位线。\n"
        text += "💡 【汇报话术】: “老板，这个图右上角的市场（价格双高），就是最优质的‘肥肉市场’，利润丰厚且容量巨大。左下角的是典型低价内卷且规模鸡肋的‘红海’，建议直接放弃。”"
        
        return text

