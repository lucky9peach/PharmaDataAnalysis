import sys
import os
import pandas as pd
from pathlib import Path

from PySide6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
                               QLabel, QPushButton, QSplitter, QTableWidget, QTableWidgetItem,
                               QHeaderView, QMessageBox, QTreeWidget, QTreeWidgetItem, QFileDialog,
                               QLineEdit, QMenu, QDialog, QFormLayout, QComboBox, QDialogButtonBox)
from PySide6.QtCore import Qt, QFileInfo, QDateTime

CACHE_DIR = Path("Cache")

class CacheManagerApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("缓存管理中心 (Cache Manager)")
        self.setGeometry(100, 100, 1200, 700)
        
        # Ensure Cache dir exists
        if not CACHE_DIR.exists():
            CACHE_DIR.mkdir(parents=True, exist_ok=True)
            
        self.current_file_path = None
        self.current_df = pd.DataFrame()
        
        self.setup_ui()
        self.refresh_file_list()

    def setup_ui(self):
        main_widget = QWidget()
        self.setCentralWidget(main_widget)
        main_layout = QVBoxLayout(main_widget)
        
        # Top toolbar
        toolbar = QHBoxLayout()
        
        btn_refresh = QPushButton("🔄 刷新缓存列表")
        btn_refresh.clicked.connect(self.refresh_file_list)
        btn_refresh.setStyleSheet("padding: 8px; font-weight: bold;")
        
        btn_delete_sel = QPushButton("🗑️ 删除选中文件")
        btn_delete_sel.clicked.connect(self.delete_selected_files)
        btn_delete_sel.setStyleSheet("padding: 8px; color: #c53030;")
        
        btn_bulk_del_rows = QPushButton("✂️ 批量按条件删行")
        btn_bulk_del_rows.clicked.connect(self.bulk_delete_rows)
        btn_bulk_del_rows.setStyleSheet("padding: 8px; color: #dd6b20;")
        
        btn_clear_all = QPushButton("🧨 清空所有缓存")
        btn_clear_all.clicked.connect(self.clear_all_caches)
        btn_clear_all.setStyleSheet("padding: 8px; color: white; background-color: #e53e3e; font-weight: bold;")
        
        toolbar.addWidget(btn_refresh)
        toolbar.addWidget(btn_delete_sel)
        toolbar.addWidget(btn_bulk_del_rows)
        toolbar.addStretch()
        toolbar.addWidget(btn_clear_all)
        
        main_layout.addLayout(toolbar)
        
        # Main split view
        splitter = QSplitter(Qt.Horizontal)
        
        # Left side: File List (Tree)
        left_widget = QWidget()
        left_layout = QVBoxLayout(left_widget)
        left_layout.setContentsMargins(0, 0, 0, 0)
        
        self.tree_files = QTreeWidget()
        self.tree_files.setHeaderLabels(["文件名", "大小", "修改时间"])
        self.tree_files.setSelectionMode(QTreeWidget.ExtendedSelection) # Support multi-select
        self.tree_files.itemClicked.connect(self.on_file_clicked)
        self.tree_files.setContextMenuPolicy(Qt.CustomContextMenu)
        self.tree_files.customContextMenuRequested.connect(self.show_file_context_menu)
        
        left_layout.addWidget(QLabel("<b>📦 缓存文件列表 (Parquet / Excel)</b>"))
        left_layout.addWidget(self.tree_files)
        
        # Right side: Data Editor
        right_widget = QWidget()
        right_layout = QVBoxLayout(right_widget)
        right_layout.setContentsMargins(0, 0, 0, 0)
        
        self.lbl_current_file = QLabel("<b>未选择文件</b>")
        self.lbl_current_file.setStyleSheet("padding: 5px; background-color: #edf2f7; border-radius: 4px;")
        
        self.table_data = QTableWidget()
        self.table_data.itemChanged.connect(self.on_cell_changed)
        
        # We need a flag to prevent itemChanged triggering while we are loading data
        self.is_loading_data = False
        
        btn_save = QPushButton("💾 保存修改 (覆写)")
        btn_save.setStyleSheet("padding: 10px; background-color: #38a169; color: white; font-weight: bold;")
        btn_save.clicked.connect(self.save_current_file)
        
        right_layout.addWidget(self.lbl_current_file)
        right_layout.addWidget(self.table_data)
        right_layout.addWidget(btn_save)
        
        splitter.addWidget(left_widget)
        splitter.addWidget(right_widget)
        splitter.setSizes([400, 800])
        
        main_layout.addWidget(splitter)

    def refresh_file_list(self):
        self.tree_files.clear()
        self.lbl_current_file.setText("<b>未选择文件</b>")
        self.table_data.setRowCount(0)
        self.table_data.setColumnCount(0)
        self.current_file_path = None
        self.current_df = pd.DataFrame()
        
        if not CACHE_DIR.exists(): return
        
        # Find all cache files recursively (.parquet, .xlsx, .csv)
        valid_exts = {'.parquet', '.xlsx', '.csv'}
        
        # Group by sub-directories
        dirs = {}
        for root, _, files in os.walk(CACHE_DIR):
            for file in files:
                path = Path(root) / file
                if path.suffix.lower() in valid_exts:
                    rel_dir = str(path.parent.relative_to(CACHE_DIR))
                    if rel_dir == '.': rel_dir = 'Root (根目录)'
                    if rel_dir not in dirs: dirs[rel_dir] = []
                    dirs[rel_dir].append(path)
                    
        for d_name, d_files in dirs.items():
            dir_item = QTreeWidgetItem(self.tree_files)
            dir_item.setText(0, f"📁 {d_name}")
            dir_item.setExpanded(True)
            
            for f in sorted(d_files, key=lambda x: x.name):
                info = QFileInfo(str(f))
                size_kb = info.size() / 1024.0
                mod_time = info.lastModified().toString("yyyy-MM-dd HH:mm:ss")
                
                f_item = QTreeWidgetItem(dir_item)
                icon = "��" if f.suffix.lower() == '.parquet' else "📗"
                f_item.setText(0, f"{icon} {f.name}")
                f_item.setText(1, f"{size_kb:.1f} KB")
                f_item.setText(2, mod_time)
                # Store the absolute path in UserRole for retrieval
                f_item.setData(0, Qt.UserRole, str(f))

        self.tree_files.resizeColumnToContents(0)

    def on_file_clicked(self, item, column):
        path_str = item.data(0, Qt.UserRole)
        if not path_str: return # It's a directory node
        
        self.load_file(path_str)
        
    def show_file_context_menu(self, pos):
        item = self.tree_files.itemAt(pos)
        if not item: return
        
        path_str = item.data(0, Qt.UserRole)
        if not path_str: return # Directoy
        
        menu = QMenu()
        open_action = menu.addAction("📖 读取并查看")
        delete_action = menu.addAction("🗑️ 删除此文件")
        
        action = menu.exec(self.tree_files.viewport().mapToGlobal(pos))
        if action == open_action:
            self.load_file(path_str)
        elif action == delete_action:
            reply = QMessageBox.question(self, '确认删除', f"确定要彻底删除文件 {Path(path_str).name} 吗？", QMessageBox.Yes | QMessageBox.No)
            if reply == QMessageBox.Yes:
                try:
                    os.remove(path_str)
                    self.refresh_file_list()
                    if self.current_file_path == path_str:
                        self.lbl_current_file.setText("<b>未选择文件</b>")
                        self.table_data.setRowCount(0)
                        self.table_data.setColumnCount(0)
                        self.current_file_path = None
                    QMessageBox.information(self, "成功", "文件已删除！")
                except Exception as e:
                    QMessageBox.warning(self, "错误", f"删除失败: {e}")

    def load_file(self, path_str):
        self.lbl_current_file.setText(f"⏳ 正在读取: {Path(path_str).name} ...")
        QApplication.processEvents()
        
        try:
            path = Path(path_str)
            if path.suffix.lower() == '.parquet':
                self.current_df = pd.read_parquet(path)
            elif path.suffix.lower() == '.xlsx':
                self.current_df = pd.read_excel(path)
            elif path.suffix.lower() == '.csv':
                self.current_df = pd.read_csv(path)
            else:
                raise ValueError("不支持的文件格式")
                
            self.current_file_path = path_str
            self.render_dataframe()
            msg = f"<b>📄 当前编辑:</b> {path.name} (共 {len(self.current_df)} 行, {len(self.current_df.columns)} 列)"
            self.lbl_current_file.setText(msg)
            
        except Exception as e:
            self.lbl_current_file.setText(f"❌ 读取错误: {Path(path_str).name}")
            QMessageBox.critical(self, "读取失败", f"无法打开缓存文件:\n{str(e)}")
            self.current_file_path = None
            self.current_df = pd.DataFrame()

    def render_dataframe(self):
        self.is_loading_data = True
        df = self.current_df
        
        self.table_data.setRowCount(0)
        self.table_data.setColumnCount(len(df.columns))
        self.table_data.setHorizontalHeaderLabels([str(c) for c in df.columns])
        
        # Dataframes can be huge, we might want to just show root 1000 lines if we don't paginate
        MAX_ROWS = 5000 
        display_df = df.head(MAX_ROWS)
        
        self.table_data.setRowCount(len(display_df))
        
        for r_idx, row in display_df.iterrows():
            for c_idx, col_name in enumerate(display_df.columns):
                val = row[col_name]
                if pd.isna(val):
                    val_str = ""
                else:
                    val_str = str(val)
                item = QTableWidgetItem(val_str)
                self.table_data.setItem(r_idx, c_idx, item)
                
        self.table_data.resizeColumnsToContents()
        self.is_loading_data = False
        
        if len(df) > MAX_ROWS:
            QMessageBox.information(self, "数据截断", f"文件行数 ({len(df)}) 超过最大显示限制，目前仅显示并允许编辑前 {MAX_ROWS} 行。若要编辑后续内容，请转移至代码层操作。")

    def on_cell_changed(self, item):
        if self.is_loading_data or self.current_df.empty or not self.current_file_path: return
        
        row = item.row()
        col = item.column()
        new_val = item.text()
        col_name = self.current_df.columns[col]
        
        # Try to cast to the original column's dtype if possible, or just keep as string if it fails
        orig_dtype = self.current_df[col_name].dtype
        try:
            if pd.api.types.is_numeric_dtype(orig_dtype):
                if '.' in new_val:
                    cast_val = float(new_val)
                else:
                    cast_val = int(new_val)
            else:
                cast_val = new_val
        except:
            cast_val = new_val # Fallback
            
        self.current_df.iat[row, col] = cast_val

    def save_current_file(self):
        if not self.current_file_path or self.current_df.empty:
            QMessageBox.warning(self, "警告", "当前没有正在编辑的数据可以保存！")
            return
            
        path = Path(self.current_file_path)
        reply = QMessageBox.question(self, '确认保存', f"确定要覆写缓存文件 {path.name} 吗？\n这是危险操作，覆写后不可逆转。", QMessageBox.Yes | QMessageBox.No)
        
        if reply == QMessageBox.Yes:
            self.lbl_current_file.setText(f"⏳ 正在保存: {path.name} ...")
            QApplication.processEvents()
            
            try:
                if path.suffix.lower() == '.parquet':
                    self.current_df.to_parquet(path, index=False)
                elif path.suffix.lower() == '.xlsx':
                    self.current_df.to_excel(path, index=False)
                elif path.suffix.lower() == '.csv':
                    self.current_df.to_csv(path, index=False)
                    
                QMessageBox.information(self, "保存成功", f"文件 {path.name} 覆写成功！")
                self.refresh_file_list() # Update timestamp
                msg = f"<b>📄 当前编辑:</b> {path.name} (已在 {QDateTime.currentDateTime().toString('HH:mm:ss')} 保存)"
                self.lbl_current_file.setText(msg)
                
            except Exception as e:
                self.lbl_current_file.setText(f"❌ 保存失败: {path.name}")
                QMessageBox.critical(self, "保存失败", f"无法写入缓存文件:\n{str(e)}")

    def delete_selected_files(self):
        selected_items = self.tree_files.selectedItems()
        if not selected_items:
            QMessageBox.warning(self, "警告", "请先从左侧列表中选中要删除的文件。")
            return
            
        paths_to_delete = []
        for item in selected_items:
            path_str = item.data(0, Qt.UserRole)
            if path_str: paths_to_delete.append(path_str)
            
        if not paths_to_delete: return
            
        reply = QMessageBox.question(self, '批量删除', f"确定要彻底删除选中的 {len(paths_to_delete)} 个缓存文件吗？", QMessageBox.Yes | QMessageBox.No)
        
        if reply == QMessageBox.Yes:
            success = 0
            errors = []
            for path_str in paths_to_delete:
                try:
                    os.remove(path_str)
                    success += 1
                    if self.current_file_path == path_str:
                        self.lbl_current_file.setText("<b>未选择文件</b>")
                        self.table_data.setRowCount(0)
                        self.table_data.setColumnCount(0)
                        self.current_file_path = None
                except Exception as e:
                    errors.append(f"{Path(path_str).name}: {e}")
                    
            self.refresh_file_list()
            
            if not errors:
                QMessageBox.information(self, "删除成功", f"成功删除了 {success} 个文件！")
            else:
                err_msg = "\n".join(errors)
                QMessageBox.warning(self, "部分删除失败", f"成功: {success} 个\n失败:\n{err_msg}")

    def clear_all_caches(self):
        reply = QMessageBox.critical(self, '💣 清空警告', 
                                     "🧨 这是极度危险的操作！🧨\n这将会永久删除 Cache/ 目录下的所有本地数据库、Parquet 和 Excel 缓存。\n如果你没有源数据，应用将无法运行。\n\n您确定要彻底清空所有缓存吗？", 
                                     QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
                                     
        if reply == QMessageBox.Yes:
            success = 0
            # Double confirm
            reply2 = QMessageBox.warning(self, '最后确认', "真的要清空吗？", QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
            if reply2 == QMessageBox.Yes:
                for root, _, files in os.walk(CACHE_DIR):
                    for file in files:
                        p = Path(root) / file
                        if p.suffix.lower() in {'.parquet', '.xlsx', '.csv'}:
                            try:
                                os.remove(p)
                                success += 1
                            except: pass
                            
                self.refresh_file_list()
                QMessageBox.information(self, "清理完成", f"系统已核弹级清理，共销毁了 {success} 个缓存文件。")
                self.lbl_current_file.setText("<b>未选择文件</b>")
                self.table_data.setRowCount(0)
                self.table_data.setColumnCount(0)
                self.current_file_path = None

    def bulk_delete_rows(self):
        selected_items = self.tree_files.selectedItems()
        if not selected_items:
            QMessageBox.warning(self, "警告", "请先从左侧列表中选中要操作的文件（支持多选）。")
            return
            
        paths_to_process = [item.data(0, Qt.UserRole) for item in selected_items if item.data(0, Qt.UserRole)]
        if not paths_to_process: return
        
        dialog = QDialog(self)
        dialog.setWindowTitle("批量条件删行")
        layout = QFormLayout(dialog)
        
        col_input = QLineEdit("year")
        layout.addRow("目标列名(如 year):", col_input)
        
        op_combo = QComboBox()
        op_combo.addItems(["==", "!=", ">", "<", ">=", "<="])
        layout.addRow("逻辑运算符:", op_combo)
        
        val_input = QLineEdit("2025")
        layout.addRow("想删除的条件值:", val_input)
        
        lbl_warn = QLabel("<b>注意：</b>此操作将遍历所选文件，若其中存在该列，则匹配此条件的<b>行将被彻底删除（覆写保存）！</b>")
        lbl_warn.setStyleSheet("color: red;")
        lbl_warn.setWordWrap(True)
        layout.addRow(lbl_warn)
        
        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel, dialog)
        buttons.accepted.connect(dialog.accept)
        buttons.rejected.connect(dialog.reject)
        layout.addRow(buttons)
        
        if dialog.exec() != QDialog.Accepted:
            return
            
        col_name = col_input.text().strip()
        op = op_combo.currentText()
        val_str = val_input.text().strip()
        
        if not col_name or not val_str:
            QMessageBox.warning(self, "错误", "列名与条件值不能为空！")
            return
            
        # 尝试转数字
        try:
            val_num = float(val_str) if '.' in val_str else int(val_str)
        except:
            val_num = val_str
            
        success_count = 0
        deleted_records = 0
        errors = []
        
        self.lbl_current_file.setText(f"⏳ 正在批量处理 {len(paths_to_process)} 个文件 ...")
        QApplication.processEvents()
        
        for path_str in paths_to_process:
            path = Path(path_str)
            try:
                # 读取
                if path.suffix.lower() == '.parquet':
                    df = pd.read_parquet(path)
                elif path.suffix.lower() == '.xlsx':
                    df = pd.read_excel(path)
                elif path.suffix.lower() == '.csv':
                    df = pd.read_csv(path)
                else:
                    continue
                    
                if col_name not in df.columns:
                    continue # 没这列就算了
                    
                orig_len = len(df)
                
                # 构建过滤掩码
                if op == "==":
                    mask = ~(df[col_name] == val_num)
                elif op == "!=":
                    mask = ~(df[col_name] != val_num)
                elif op == ">":
                    mask = ~(df[col_name] > val_num)
                elif op == "<":
                    mask = ~(df[col_name] < val_num)
                elif op == ">=":
                    mask = ~(df[col_name] >= val_num)
                elif op == "<=":
                    mask = ~(df[col_name] <= val_num)
                
                new_df = df[mask]
                del_count = orig_len - len(new_df)
                
                if del_count > 0:
                    # 回写
                    if path.suffix.lower() == '.parquet':
                        new_df.to_parquet(path, index=False)
                    elif path.suffix.lower() == '.xlsx':
                        new_df.to_excel(path, index=False)
                    elif path.suffix.lower() == '.csv':
                        new_df.to_csv(path, index=False)
                        
                    deleted_records += del_count
                    success_count += 1
                    
            except Exception as e:
                errors.append(f"{path.name}: {e}")
                
        self.lbl_current_file.setText("<b>未选择文件</b>")
        self.refresh_file_list()
        
        if errors:
            QMessageBox.warning(self, "完成，但有错误", f"处理完成。共在 {success_count} 个文件中删除了 {deleted_records} 行数据。\n部分报错:\n{chr(10).join(errors)}")
        else:
            QMessageBox.information(self, "批量删行成功", f"成功对 {success_count} 个受影响的文件进行了覆写，累计删除了 {deleted_records} 行数据！")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = CacheManagerApp()
    window.show()
    sys.exit(app.exec())
