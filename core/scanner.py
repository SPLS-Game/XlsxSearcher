"""xlsx文件扫描器 - 递归扫描目录并提取子表名称"""
import os
import time
from typing import List, Tuple, Callable
from openpyxl import load_workbook

class XlsxScanner:
    def __init__(self):
        self.supported_extensions = ['.xlsx', '.xlsm']

    def is_xlsx_file(self, filepath: str) -> bool:
        """检查是否为xlsx文件"""
        ext = os.path.splitext(filepath)[1].lower()
        return ext in self.supported_extensions

    def get_sheet_names(self, filepath: str) -> List[str]:
        """获取xlsx文件的所有子表名称"""
        try:
            # 只读取工作表名称，不加载整个文件，提升性能
            wb = load_workbook(filepath, read_only=True, data_only=True)
            sheet_names = wb.sheetnames
            wb.close()
            return sheet_names
        except Exception as e:
            # 忽略无法读取的文件
            print(f"警告: 无法读取文件 {filepath}: {e}")
            return []

    def scan_file(self, filepath: str) -> Tuple[str, float, List[str]] | None:
        """扫描单个文件，返回(文件名, 修改时间, 子表列表)"""
        if not self.is_xlsx_file(filepath):
            return None

        try:
            stat = os.stat(filepath)
            modified_time = stat.st_mtime
            filename = os.path.basename(filepath)
            sheet_names = self.get_sheet_names(filepath)
            return (filename, modified_time, sheet_names)
        except Exception as e:
            print(f"警告: 扫描文件失败 {filepath}: {e}")
            return None

    def scan_directory(self, directory: str, progress_callback: Callable = None) -> List[Tuple[str, str, float, List[str]]]:
        """
        递归扫描目录下的所有xlsx文件
        返回: [(filename, filepath, modified_time, sheet_names), ...]
        """
        results = []

        for root, dirs, files in os.walk(directory):
            # 跳过隐藏目录
            dirs[:] = [d for d in dirs if not d.startswith('.')]

            for filename in files:
                if not self.is_xlsx_file(filename):
                    continue

                filepath = os.path.join(root, filename)
                result = self.scan_file(filepath)

                if result:
                    results.append((result[0], filepath, result[1], result[2]))

                if progress_callback:
                    progress_callback(len(results))

        return results

    def scan_directory_incremental(self, directory: str, index_manager, progress_callback: Callable = None) -> Tuple[int, int, int]:
        """
        增量扫描目录，只更新有变化的文件
        返回: (新增文件数, 更新文件数, 删除文件数)
        """
        # 获取所有需要扫描的文件
        all_files = {}
        for root, dirs, files in os.walk(directory):
            dirs[:] = [d for d in dirs if not d.startswith('.')]
            for filename in files:
                if self.is_xlsx_file(filename):
                    filepath = os.path.join(root, filename)
                    all_files[filepath] = True

        # 获取索引中已存在的文件
        indexed_files = {}
        for file_info in index_manager.get_all_files():
            indexed_files[file_info['filepath']] = file_info['modified_time']

        added = 0
        updated = 0
        deleted = 0

        # 处理新增和更新的文件
        for filepath in all_files:
            try:
                stat = os.stat(filepath)
                modified_time = stat.st_mtime
                filename = os.path.basename(filepath)

                existing = index_manager.get_file_info(filepath)

                if existing is None:
                    # 新文件
                    sheet_names = self.get_sheet_names(filepath)
                    if sheet_names:  # 只索引包含子表的文件
                        index_manager.add_file(filename, filepath, modified_time, sheet_names)
                        added += 1
                elif existing[1] != modified_time:
                    # 文件已更新
                    sheet_names = self.get_sheet_names(filepath)
                    index_manager.add_file(filename, filepath, modified_time, sheet_names)
                    updated += 1

                if progress_callback:
                    progress_callback(added + updated)

            except Exception as e:
                print(f"警告: 处理文件失败 {filepath}: {e}")

        # 处理已删除的文件
        for filepath in indexed_files:
            if filepath not in all_files:
                index_manager.delete_file(filepath)
                deleted += 1

        return (added, updated, deleted)