"""索引管理器 - 使用SQLite存储xlsx文件索引"""
import sqlite3
import os
import sys
from typing import List, Dict, Tuple

class IndexManager:
    def __init__(self, db_path: str = None):
        if db_path is None:
            # 在用户目录创建数据库
            user_home = os.path.expanduser("~")
            if sys.platform == 'darwin':
                # macOS: 使用 Application Support 目录
                app_data_dir = os.path.join(user_home, "Library", "Application Support", "XlsxSearcher")
            elif sys.platform == 'win32':
                # Windows: 直接在用户目录下
                app_data_dir = os.path.join(user_home, "XlsxSearcher")
            else:
                # Linux: 使用 .local/share 目录
                app_data_dir = os.path.join(user_home, ".local", "share", "XlsxSearcher")
            os.makedirs(app_data_dir, exist_ok=True)
            db_path = os.path.join(app_data_dir, "index.db")
        self.db_path = db_path
        self._init_db()

    def _init_db(self):
        """初始化数据库表"""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        # 创建xlsx文件索引表
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS xlsx_files (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                filename TEXT NOT NULL,
                filepath TEXT UNIQUE NOT NULL,
                modified_time REAL NOT NULL,
                sheet_count INTEGER DEFAULT 0
            )
        ''')
        # 创建子表索引表
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS sheets (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                file_id INTEGER NOT NULL,
                sheet_name TEXT NOT NULL,
                FOREIGN KEY (file_id) REFERENCES xlsx_files(id) ON DELETE CASCADE
            )
        ''')
        # 创建索引
        cursor.execute('CREATE INDEX IF NOT EXISTS idx_filename ON xlsx_files(filename)')
        cursor.execute('CREATE INDEX IF NOT EXISTS idx_filepath ON xlsx_files(filepath)')
        cursor.execute('CREATE INDEX IF NOT EXISTS idx_sheet_name ON sheets(sheet_name)')
        cursor.execute('CREATE INDEX IF NOT EXISTS idx_file_id ON sheets(file_id)')
        conn.commit()
        conn.close()

    def get_file_info(self, filepath: str) -> Tuple:
        """获取文件信息"""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        cursor.execute('SELECT id, modified_time FROM xlsx_files WHERE filepath = ?', (filepath,))
        result = cursor.fetchone()
        conn.close()
        return result  # (id, modified_time) or None

    def add_file(self, filename: str, filepath: str, modified_time: float, sheet_names: List[str]):
        """添加文件及其子表到索引"""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        # 插入文件信息
        cursor.execute(
            'INSERT OR REPLACE INTO xlsx_files (filename, filepath, modified_time, sheet_count) VALUES (?, ?, ?, ?)',
            (filename, filepath, modified_time, len(sheet_names))
        )
        file_id = cursor.lastrowid
        # 如果是REPLACE，先删除旧的子表
        cursor.execute('DELETE FROM sheets WHERE file_id = ?', (file_id,))
        # 插入子表信息
        for sheet_name in sheet_names:
            cursor.execute('INSERT INTO sheets (file_id, sheet_name) VALUES (?, ?)', (file_id, sheet_name))
        conn.commit()
        conn.close()

    def delete_file(self, filepath: str):
        """从索引中删除文件"""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        cursor.execute('SELECT id FROM xlsx_files WHERE filepath = ?', (filepath,))
        row = cursor.fetchone()
        if row:
            file_id = row[0]
            cursor.execute('DELETE FROM sheets WHERE file_id = ?', (file_id,))
            cursor.execute('DELETE FROM xlsx_files WHERE id = ?', (file_id,))
        conn.commit()
        conn.close()

    def get_all_files(self) -> List[Dict]:
        """获取所有已索引的文件"""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        cursor.execute('SELECT id, filename, filepath, modified_time, sheet_count FROM xlsx_files')
        rows = cursor.fetchall()
        conn.close()
        return [
            {'id': r[0], 'filename': r[1], 'filepath': r[2], 'modified_time': r[3], 'sheet_count': r[4]}
            for r in rows
        ]

    def search_by_sheet_name(self, keyword: str) -> List[Dict]:
        """按子表名称模糊搜索"""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        cursor.execute('''
            SELECT f.filename, f.filepath, s.sheet_name
            FROM sheets s
            JOIN xlsx_files f ON s.file_id = f.id
            WHERE s.sheet_name LIKE ?
            ORDER BY f.filename
        ''', (f'%{keyword}%',))
        rows = cursor.fetchall()
        conn.close()
        return [{'filename': r[0], 'filepath': r[1], 'sheet_name': r[2]} for r in rows]

    def search_by_filename(self, keyword: str) -> List[Dict]:
        """按文件名模糊搜索"""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        cursor.execute('''
            SELECT f.filename, f.filepath, GROUP_CONCAT(s.sheet_name, ', ')
            FROM xlsx_files f
            LEFT JOIN sheets s ON f.id = s.file_id
            WHERE f.filename LIKE ?
            GROUP BY f.id
            ORDER BY f.filename
        ''', (f'%{keyword}%',))
        rows = cursor.fetchall()
        conn.close()
        return [{'filename': r[0], 'filepath': r[1], 'sheet_names': r[2]} for r in rows]

    def search(self, sheet_keyword: str = None, filename_keyword: str = None) -> List[Dict]:
        """综合搜索"""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()

        if sheet_keyword and filename_keyword:
            # 组合搜索
            cursor.execute('''
                SELECT f.filename, f.filepath, s.sheet_name
                FROM sheets s
                JOIN xlsx_files f ON s.file_id = f.id
                WHERE s.sheet_name LIKE ? AND f.filename LIKE ?
                ORDER BY f.filename
            ''', (f'%{sheet_keyword}%', f'%{filename_keyword}%'))
        elif sheet_keyword:
            cursor.execute('''
                SELECT f.filename, f.filepath, s.sheet_name
                FROM sheets s
                JOIN xlsx_files f ON s.file_id = f.id
                WHERE s.sheet_name LIKE ?
                ORDER BY f.filename
            ''', (f'%{sheet_keyword}%',))
        elif filename_keyword:
            cursor.execute('''
                SELECT f.filename, f.filepath, GROUP_CONCAT(s.sheet_name, ', ')
                FROM xlsx_files f
                LEFT JOIN sheets s ON f.id = s.file_id
                WHERE f.filename LIKE ?
                GROUP BY f.id
                ORDER BY f.filename
            ''', (f'%{filename_keyword}%',))
        else:
            return []

        rows = cursor.fetchall()
        conn.close()

        if sheet_keyword and not filename_keyword:
            return [{'filename': r[0], 'filepath': r[1], 'sheet_name': r[2]} for r in rows]
        else:
            return [{'filename': r[0], 'filepath': r[1], 'sheet_names': r[2]} for r in rows]

    def clear_index(self):
        """清空所有索引"""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        cursor.execute('DELETE FROM sheets')
        cursor.execute('DELETE FROM xlsx_files')
        conn.commit()
        conn.close()

    def get_stats(self) -> Dict:
        """获取索引统计信息"""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        cursor.execute('SELECT COUNT(*) FROM xlsx_files')
        file_count = cursor.fetchone()[0]
        cursor.execute('SELECT COUNT(*) FROM sheets')
        sheet_count = cursor.fetchone()[0]
        conn.close()
        return {'file_count': file_count, 'sheet_count': sheet_count}