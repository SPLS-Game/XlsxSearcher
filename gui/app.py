"""XlsxSearcher 主界面 - PyQt5 版本"""
import os
import sys
import threading
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QTreeWidget, QTreeWidgetItem, QLabel, QLineEdit, QPushButton,
    QSplitter, QStatusBar, QProgressBar, QMessageBox, QFileDialog
)
from PyQt5.QtCore import Qt, QThread, pyqtSignal
from PyQt5.QtGui import QFont

from core.indexer import IndexManager
from core.scanner import XlsxScanner
from core.searcher import Searcher
from utils.file_utils import open_file, open_in_explorer, copy_to_clipboard


class ScanWorker(QThread):
    """扫描工作线程"""
    finished = pyqtSignal(int, int, int)  # added, updated, deleted
    error = pyqtSignal(str)

    def __init__(self, directory, scanner, index_manager):
        super().__init__()
        self.directory = directory
        self.scanner = scanner
        self.index_manager = index_manager

    def run(self):
        try:
            added, updated, deleted = self.scanner.scan_directory_incremental(
                self.directory, self.index_manager
            )
            self.finished.emit(added, updated, deleted)
        except Exception as e:
            self.error.emit(str(e))


class XlsxSearcherApp(QMainWindow):
    def __init__(self):
        super().__init__()

        # 核心组件
        self.index_manager = IndexManager()
        self.scanner = XlsxScanner()
        self.searcher = Searcher(self.index_manager)

        # 状态变量
        self.scan_directory = None
        self.search_results = []
        self.is_scanning = False

        self._init_ui()

    def _init_ui(self):
        """初始化UI"""
        # 窗口设置
        self.setWindowTitle("XlsxSearcher - Excel子表搜索工具")
        self.setMinimumSize(1000, 700)
        self.resize(1000, 700)

        # 中心部件
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)

        # 顶部搜索区域
        top_widget = QWidget()
        top_layout = QHBoxLayout(top_widget)
        top_layout.setContentsMargins(10, 10, 10, 5)
        main_layout.addWidget(top_widget)

        # 目录选择和显示
        self.dir_label = QLabel("未选择目录")
        self.dir_label.setStyleSheet("color: gray")
        top_layout.addWidget(QLabel("扫描目录:"))
        top_layout.addWidget(self.dir_label)
        top_layout.addWidget(QPushButton("选择目录", clicked=self._select_directory))
        top_layout.addWidget(QPushButton("重新扫描", clicked=self._rescan))
        top_layout.addWidget(QPushButton("清空索引", clicked=self._clear_index))

        # 搜索区域
        search_widget = QWidget()
        search_layout = QHBoxLayout(search_widget)
        search_layout.setContentsMargins(10, 5, 10, 10)
        main_layout.addWidget(search_widget)

        # 子表名称搜索
        self.sheet_entry = QLineEdit()
        self.sheet_entry.setPlaceholderText("子表名称")
        self.sheet_entry.textChanged.connect(self._do_search)
        search_layout.addWidget(QLabel("子表名称:"))
        search_layout.addWidget(self.sheet_entry)

        # 文件名搜索
        self.filename_entry = QLineEdit()
        self.filename_entry.setPlaceholderText("文件名")
        self.filename_entry.textChanged.connect(self._do_search)
        search_layout.addWidget(QLabel("文件名:"))
        search_layout.addWidget(self.filename_entry)

        # 状态栏
        self.status_bar = QStatusBar()
        self.setStatusBar(self.status_bar)
        self.status_bar.showMessage("就绪")

        # 结果表格 (使用 QTreeWidget 替代 Treeview)
        self.result_tree = QTreeWidget()
        self.result_tree.setHeaderLabels(["文件名", "子表名称", "文件路径"])
        self.result_tree.setColumnWidth(0, 200)
        self.result_tree.setColumnWidth(1, 150)
        self.result_tree.setColumnWidth(2, 500)
        self.result_tree.setAlternatingRowColors(True)

        # 绑定事件
        self.result_tree.itemClicked.connect(self._on_select)
        self.result_tree.itemDoubleClicked.connect(self._open_file)

        main_layout.addWidget(self.result_tree)

        # 底部操作按钮
        bottom_widget = QWidget()
        bottom_layout = QHBoxLayout(bottom_widget)
        bottom_layout.setContentsMargins(10, 5, 10, 10)
        main_layout.addWidget(bottom_widget)

        self.btn_open = QPushButton("打开文件")
        self.btn_open.clicked.connect(self._open_file)
        self.btn_open.setEnabled(False)

        self.btn_locate = QPushButton("定位文件")
        self.btn_locate.clicked.connect(self._locate_file)
        self.btn_locate.setEnabled(False)

        self.btn_copy = QPushButton("复制路径")
        self.btn_copy.clicked.connect(self._copy_path)
        self.btn_copy.setEnabled(False)

        bottom_layout.addWidget(self.btn_open)
        bottom_layout.addWidget(self.btn_locate)
        bottom_layout.addWidget(self.btn_copy)
        bottom_layout.addStretch()

    def _check_existing_index(self):
        """检查是否存在已保存的扫描目录（可选显示）"""
        stats = self.index_manager.get_stats()
        if stats['file_count'] > 0:
            self.status_bar.showMessage(f"已加载索引：{stats['file_count']}文件, {stats['sheet_count']}子表")

    def _select_directory(self):
        """选择扫描目录"""
        directory = QFileDialog.getExistingDirectory(
            self, "选择要扫描的目录", os.path.expanduser("~")
        )
        if directory:
            self.scan_directory = directory
            self.dir_label.setText(self._truncate_path(directory))
            self._start_scan()

    def _rescan(self):
        """重新扫描"""
        if self.scan_directory:
            self._start_scan()

    def _start_scan(self):
        """开始扫描（在线程中执行）"""
        if self.is_scanning:
            return

        self.is_scanning = True
        self.status_bar.showMessage("正在扫描...")

        # 禁用按钮
        for btn in [self.btn_open, self.btn_locate, self.btn_copy]:
            btn.setEnabled(False)

        # 启动扫描线程
        self.scan_worker = ScanWorker(
            self.scan_directory, self.scanner, self.index_manager
        )
        self.scan_worker.finished.connect(self._on_scan_complete)
        self.scan_worker.error.connect(self._on_scan_error)
        self.scan_worker.start()

    def _on_scan_complete(self, added, updated, deleted):
        """扫描完成回调"""
        self.is_scanning = False

        stats = self.index_manager.get_stats()
        self.status_bar.showMessage(
            f"索引完成: {stats['file_count']}文件, {stats['sheet_count']}子表"
        )

        # 执行初始搜索
        self._do_search()

    def _on_scan_error(self, error_msg):
        """扫描错误回调"""
        self.is_scanning = False
        self.status_bar.showMessage("扫描出错")
        QMessageBox.critical(self, "错误", f"扫描失败: {error_msg}")

    def _clear_index(self):
        """清空索引"""
        reply = QMessageBox.question(
            self, "确认", "确定要清空所有索引数据吗？"
        )
        if reply == QMessageBox.Yes:
            self.index_manager.clear_index()
            self.result_tree.clear()
            self.status_bar.showMessage("索引已清空")

    def _do_search(self):
        """执行搜索"""
        sheet_keyword = self.sheet_entry.text().strip()
        filename_keyword = self.filename_entry.text().strip()

        if not sheet_keyword and not filename_keyword:
            # 显示所有已索引的文件
            self.search_results = self.index_manager.get_all_files()
            # 需要获取每个文件的子表
            all_results = []
            for file_info in self.search_results:
                sheet_results = self.index_manager.search_by_filename(file_info['filename'])
                all_results.extend(sheet_results)
            self.search_results = all_results
        else:
            self.search_results = self.searcher.search(sheet_keyword, filename_keyword)

        self._update_results()

    def _update_results(self):
        """更新搜索结果表格"""
        self.result_tree.clear()

        for result in self.search_results:
            sheet_display = result.get('sheet_name') or result.get('sheet_names', '')
            item = QTreeWidgetItem([
                result['filename'],
                str(sheet_display),
                result['filepath']
            ])
            self.result_tree.addTopLevelItem(item)

        self.status_bar.showMessage(f"找到 {len(self.search_results)} 条结果")

    def _on_select(self, item, column):
        """选中结果项"""
        self.btn_open.setEnabled(True)
        self.btn_locate.setEnabled(True)
        self.btn_copy.setEnabled(True)

    def _get_selected_filepath(self):
        """获取选中的文件路径"""
        selected_items = self.result_tree.selectedItems()
        if selected_items:
            index = self.result_tree.indexOfTopLevelItem(selected_items[0])
            if index < len(self.search_results):
                return self.search_results[index]['filepath']
        return None

    def _open_file(self):
        """打开文件"""
        filepath = self._get_selected_filepath()
        if filepath and os.path.exists(filepath):
            try:
                open_file(filepath)
            except Exception as e:
                QMessageBox.critical(self, "错误", f"无法打开文件: {e}")
        else:
            QMessageBox.warning(self, "警告", "文件不存在或已被移动")

    def _locate_file(self):
        """在资源管理器中定位文件"""
        filepath = self._get_selected_filepath()
        if filepath and os.path.exists(filepath):
            try:
                open_in_explorer(filepath)
            except Exception as e:
                QMessageBox.critical(self, "错误", f"无法定位文件: {e}")
        else:
            QMessageBox.warning(self, "警告", "文件不存在或已被移动")

    def _copy_path(self):
        """复制文件路径"""
        filepath = self._get_selected_filepath()
        if filepath:
            if copy_to_clipboard(filepath):
                QMessageBox.information(self, "成功", "路径已复制到剪贴板")
            else:
                QMessageBox.critical(self, "错误", "复制失败")
        else:
            QMessageBox.warning(self, "警告", "请先选择文件")

    def _truncate_path(self, path, max_length=50):
        """截断路径显示"""
        if len(path) <= max_length:
            return path
        return "..." + path[-max_length:]

    def run(self):
        """运行应用"""
        self.show()


def run_app():
    """启动应用程序"""
    app = QApplication(sys.argv)
    window = XlsxSearcherApp()
    window.show()
    sys.exit(app.exec_())
