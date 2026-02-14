import chardet
import os
import sys

class File_Manager:
    def __Detect_Encoding__(file_path: str) -> str:
        with open(file_path, 'rb') as f:
            raw_data = f.read()
            result = chardet.detect(raw_data)
            return result['encoding']

    def Read_File(file_path: str) -> str | None:
        try:
            encoding = File_Manager.__Detect_Encoding__(file_path)
            with open(file_path, "r", encoding=encoding) as f:
                return f.read()
        except Exception as e:
            print(e)
            return None
        
    # 从完整的文件路径中获取文件名
    def Get_FileName(file_path: str) -> str | None:
        if '/' not in file_path:
            return None
        return file_path.split("/")[-1]

    
    # 检查文件是否存在
    def File_Exist(file_path: str) -> bool:
        return os.path.exists(file_path)

    # 用系统文件管理器打开目录
    def open_directory(dir_path):
        if not os.path.exists(dir_path):
            print(f"目录不存在: {dir_path}")
            return
        
        try:
            if sys.platform == 'win32':  # Windows
                os.startfile(dir_path)
            elif sys.platform == 'darwin':  # macOS
                subprocess.run(['open', dir_path])
            else:  # Linux
                subprocess.run(['xdg-open', dir_path])
        except Exception as e:
            print(f"打开目录失败: {e}")

    # 用系统默认程序打开文件
    def open_file(file_path):
        if not os.path.exists(file_path):
            print(f"文件不存在: {file_path}")
            return
        
        try:
            if sys.platform == 'win32':  # Windows
                os.startfile(file_path)
            elif sys.platform == 'darwin':  # macOS
                subprocess.run(['open', file_path])
            else:  # Linux
                subprocess.run(['xdg-open', file_path])
        except Exception as e:
            print(f"打开文件失败: {e}")
    

# Uint Test
if __name__ == "__main__":
    print(File_Manager.File_Exist("chores/main3.py"))