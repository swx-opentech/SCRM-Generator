# @note:系统给用户提供了一个“工程文件”存储的形式，以json格式存储在文件内
import json
from typing import Literal
from file_manager import File_Manager
import os

class Project_Manager:
    def New_Project(project_name: str, project_path: str, software_name: str, software_version: str) -> tuple[bool, str]:
        # project_path 不包含具体文件名
        project_path = os.path.join(project_path, f"{project_name}.scrm")
        if os.path.exists(project_path):
            return False, f"存在同名的文件！"
        source_data = {
            "programme" : "SCRM_Generator",
            "software_name" : software_name,
            "software_version" : software_version,
            "source_code_paths" : []
            }
        try:
            with open(project_path, "w", encoding="utf-8") as f:
                json.dump(source_data, f, indent=4, ensure_ascii=False)
                return True, project_path
        except Exception as e:
            return False, f"项目创建失败: {e}"

    
    def Check_Valid_Path(project_path: str) -> bool:
        if not project_path or not isinstance(project_path, str):
            return False
        if not os.path.exists(project_path) or not os.path.isdir(project_path):
            return False
        if not os.access(project_path, os.R_OK | os.W_OK):
            return False
        return True

    def Load_Project(project_path: str) -> tuple[bool, dict|str]:
        if not os.path.exists(project_path) or not os.path.isfile(project_path):# 检查文件是否存在且为文件
            return False, f"文件不存在或不是有效文件: {project_path}"
        
        try:
            with open(project_path, "r", encoding="utf-8") as f: # 读取文件内容并解析 JSON
                data = json.load(f)
            return True, data
        except Exception as e:
             return True, f"读取文件时发生错误: {e}"
        
    def Modify_Project(project_path: str, data: dict) -> tuple[bool, str]:
        if not data.get("software_name") or not data.get("software_version"):
            return False, "信息不完整！"
        
        if not os.path.exists(project_path): # 检查文件是否存在
            return False, f"文件不存在: {project_path}"

        try:
            with open(project_path, "w", encoding="utf-8") as f:
                json.dump(data, f, indent=4, ensure_ascii=False)
                return True, "信息修改成功！"
        except Exception as e:
            return False, f"修改失败: {e}"



if __name__ == "__main__":
    print(Project_Manager.Load_Project("D:/Wenxuan/桌面\沈文轩测试软件.scrm"))