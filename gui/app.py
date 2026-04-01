"""XlsxSearcher 主界面"""
import os
import sys
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import threading

from core.indexer import IndexManager
from core.scanner import XlsxScanner
from core.searcher import Searcher
from utils.file_utils import open_file, open_in_explorer, copy_to_clipboard


class XlsxSearcherApp:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("XlsxSearcher - Excel子表搜索工具")
        self.root.geometry("900x600")
        self.root.minsize(800, 500)

        # 核心组件
        self.index_manager = IndexManager()
        self.scanner = XlsxScanner()
        self.searcher = Searcher(self.index_manager)

        # 状态变量
        self.scan_directory = None
        self.search_results = []
        self.is_scanning = False

        self._init_ui()

        # 检查是否有已保存的扫描目录
        self._check_existing_index()

    def _check_existing_index(self):
        """检查是否存在已保存的扫描目录"""
        stats = self.index_manager.get_stats()
        if stats['file_count'] > 0:
            # 询问用户是否要搜索已索引的文件
            result = messagebox.askyesno(
                "欢迎使用",
                f"已找到索引数据：{stats['file_count']}个文件，{stats['sheet_count']}个子表。\n\n是否选择新的扫描目录？",
                icon='question'
            )
            if result:
                self._select_directory()
            else:
                # 执行初始搜索以显示所有文件
                self._do_search()
        else:
            self._select_directory()

    def _init_ui(self):
        """初始化UI"""
        # 顶部搜索区域
        top_frame = ttk.Frame(self.root, padding=10)
        top_frame.pack(fill=tk.X)

        # 目录选择和显示
        ttk.Label(top_frame, text="扫描目录:").pack(side=tk.LEFT)
        self.dir_label = ttk.Label(top_frame, text="未选择目录", foreground="gray")
        self.dir_label.pack(side=tk.LEFT, padx=5)

        ttk.Button(top_frame, text="选择目录", command=self._select_directory).pack(side=tk.LEFT, padx=5)
        ttk.Button(top_frame, text="重新扫描", command=self._rescan).pack(side=tk.LEFT, padx=5)
        ttk.Button(top_frame, text="清空索引", command=self._clear_index).pack(side=tk.LEFT, padx=5)

        # 搜索区域
        search_frame = ttk.LabelFrame(self.root, text="搜索", padding=10)
        search_frame.pack(fill=tk.X, padx=10, pady=5)

        # 子表名称搜索
        ttk.Label(search_frame, text="子表名称:").grid(row=0, column=0, sticky=tk.W, pady=5)
        self.sheet_entry = ttk.Entry(search_frame, width=30)
        self.sheet_entry.grid(row=0, column=1, padx=5, pady=5)
        self.sheet_entry.bind('<KeyRelease>', lambda e: self._do_search())

        # 文件名搜索
        ttk.Label(search_frame, text="文件名:").grid(row=0, column=2, sticky=tk.W, padx=10, pady=5)
        self.filename_entry = ttk.Entry(search_frame, width=30)
        self.filename_entry.grid(row=0, column=3, padx=5, pady=5)
        self.filename_entry.bind('<KeyRelease>', lambda e: self._do_search())

        # 状态标签
        self.status_label = ttk.Label(search_frame, text="", foreground="blue")
        self.status_label.grid(row=0, column=4, padx=10, pady=5)

        # 结果表格
        result_frame = ttk.Frame(self.root, padding=10)
        result_frame.pack(fill=tk.BOTH, expand=True, padx=10)

        # 创建表格
        columns = ("filename", "sheet_name", "filepath")
        self.tree = ttk.Treeview(result_frame, columns=columns, show="headings", selectmode="browse")

        self.tree.heading("filename", text="文件名")
        self.tree.heading("sheet_name", text="子表名称")
        self.tree.heading("filepath", text="文件路径")

        self.tree.column("filename", width=200)
        self.tree.column("sheet_name", width=150)
        self.tree.column("filepath", width=400)

        # 滚动条
        vsb = ttk.Scrollbar(result_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)
        self.tree.pack(fill=tk.BOTH, expand=True)

        # 双击打开文件
        self.tree.bind('<Double-Button-1>', lambda e: self._open_file())

        # 底部操作按钮
        bottom_frame = ttk.Frame(self.root, padding=10)
        bottom_frame.pack(fill=tk.X)

        self.btn_open = ttk.Button(bottom_frame, text="打开文件", command=self._open_file, state=tk.DISABLED)
        self.btn_open.pack(side=tk.LEFT, padx=5)

        self.btn_locate = ttk.Button(bottom_frame, text="定位文件", command=self._locate_file, state=tk.DISABLED)
        self.btn_locate.pack(side=tk.LEFT, padx=5)

        self.btn_copy = ttk.Button(bottom_frame, text="复制路径", command=self._copy_path, state=tk.DISABLED)
        self.btn_copy.pack(side=tk.LEFT, padx=5)

        # 绑定选择事件
        self.tree.bind('<<TreeviewSelect>>', self._on_select)

    def _select_directory(self):
        """选择扫描目录"""
        directory = filedialog.askdirectory(title="选择要扫描的目录")
        if directory:
            self.scan_directory = directory
            self.dir_label.config(text=self._truncate_path(directory))
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
        self.status_label.config(text="正在扫描...")

        # 禁用按钮
        for btn in [self.btn_open, self.btn_locate, self.btn_copy]:
            btn.config(state=tk.DISABLED)

        # 在新线程中执行扫描
        thread = threading.Thread(target=self._scan_worker, args=(self.scan_directory,))
        thread.daemon = True
        thread.start()

    def _scan_worker(self, directory):
        """扫描工作线程"""
        try:
            added, updated, deleted = self.scanner.scan_directory_incremental(directory, self.index_manager)

            # 更新UI（在线程中）
            self.root.after(0, self._on_scan_complete, added, updated, deleted)
        except Exception as e:
            self.root.after(0, self._on_scan_error, str(e))

    def _on_scan_complete(self, added, updated, deleted):
        """扫描完成回调"""
        self.is_scanning = False

        stats = self.index_manager.get_stats()
        self.status_label.config(
            text=f"索引完成: {stats['file_count']}文件, {stats['sheet_count']}子表"
        )

        # 执行初始搜索
        self._do_search()

    def _on_scan_error(self, error_msg):
        """扫描错误回调"""
        self.is_scanning = False
        self.status_label.config(text="扫描出错")
        messagebox.showerror("错误", f"扫描失败: {error_msg}")

    def _clear_index(self):
        """清空索引"""
        result = messagebox.askyesno("确认", "确定要清空所有索引数据吗？")
        if result:
            self.index_manager.clear_index()
            self.tree.delete(*self.tree.get_children())
            self.status_label.config(text="索引已清空")
            self._select_directory()

    def _do_search(self):
        """执行搜索"""
        sheet_keyword = self.sheet_entry.get().strip()
        filename_keyword = self.filename_entry.get().strip()

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
        self.tree.delete(*self.tree.get_children())

        for result in self.search_results:
            sheet_display = result.get('sheet_name') or result.get('sheet_names', '')
            self.tree.insert('', tk.END, values=(
                result['filename'],
                sheet_display,
                result['filepath']
            ))

        self.status_label.config(text=f"找到 {len(self.search_results)} 条结果")

    def _on_select(self, event):
        """选中结果项"""
        selection = self.tree.selection()
        if selection:
            self.btn_open.config(state=tk.NORMAL)
            self.btn_locate.config(state=tk.NORMAL)
            self.btn_copy.config(state=tk.NORMAL)
        else:
            self.btn_open.config(state=tk.DISABLED)
            self.btn_locate.config(state=tk.DISABLED)
            self.btn_copy.config(state=tk.DISABLED)

    def _get_selected_filepath(self):
        """获取选中的文件路径"""
        selection = self.tree.selection()
        if not selection:
            return None

        item = self.tree.item(selection[0])
        values = item['values']
        if values:
            return values[2]  # filepath
        return None

    def _open_file(self):
        """打开文件"""
        filepath = self._get_selected_filepath()
        if filepath and os.path.exists(filepath):
            try:
                open_file(filepath)
            except Exception as e:
                messagebox.showerror("错误", f"无法打开文件: {e}")
        else:
            messagebox.showwarning("警告", "文件不存在或已被移动")

    def _locate_file(self):
        """在资源管理器中定位文件"""
        filepath = self._get_selected_filepath()
        if filepath and os.path.exists(filepath):
            try:
                open_in_explorer(filepath)
            except Exception as e:
                messagebox.showerror("错误", f"无法定位文件: {e}")
        else:
            messagebox.showwarning("警告", "文件不存在或已被移动")

    def _copy_path(self):
        """复制文件路径"""
        filepath = self._get_selected_filepath()
        if filepath:
            if copy_to_clipboard(filepath):
                messagebox.showinfo("成功", "路径已复制到剪贴板")
            else:
                messagebox.showerror("错误", "复制失败")
        else:
            messagebox.showwarning("警告", "请先选择文件")

    def _truncate_path(self, path, max_length=50):
        """截断路径显示"""
        if len(path) <= max_length:
            return path
        return "..." + path[-max_length:]

    def run(self):
        """运行应用"""
        self.root.mainloop()