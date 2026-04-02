"""文件操作工具"""
import os
import subprocess
import sys

def open_file(filepath: str):
    """使用默认程序打开文件"""
    if sys.platform == 'win32':
        os.startfile(filepath)
    elif sys.platform == 'darwin':
        subprocess.run(['open', filepath])
    else:
        subprocess.run(['xdg-open', filepath])

def open_in_explorer(filepath: str):
    """在文件管理器中定位文件"""
    if sys.platform == 'win32':
        # Windows: 使用资源管理器选中文件
        # 必须使用列表形式，确保路径正确处理
        subprocess.Popen(['explorer', '/select,', filepath])
    elif sys.platform == 'darwin':
        # macOS: 在 Finder 中选中文件
        subprocess.run(['open', '-R', filepath])
    else:
        # Linux: 在文件管理器中打开目录
        directory = os.path.dirname(filepath)
        subprocess.run(['xdg-open', directory])

def copy_to_clipboard(text: str):
    """复制文本到剪贴板"""
    try:
        import pyperclip
        pyperclip.copy(text)
        return True
    except ImportError:
        # 备用方案：使用tkinter
        try:
            import tkinter as tk
            root = tk.Tk()
            root.withdraw()
            root.clipboard_clear()
            root.clipboard_append(text)
            root.update()
            root.destroy()
            return True
        except Exception:
            return False
    except Exception:
        return False