import chardet
import os

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
    

# Uint Test
if __name__ == "__main__":
    print(File_Manager.File_Exist("chores/main3.py"))