import sys
import re
from PyQt5.QtWidgets import QApplication, QMainWindow, QDialog, QFileDialog, QMessageBox
from PyQt5.QtCore import Qt, QObject, pyqtSignal
from datetime import datetime
import os

from docx_writer import DocumentWriter
from file_manager import File_Manager
from project_manager import Project_Manager
from ui_main import Ui_MainWindow
from ui_newpro import Ui_NewPro



class MainWindow(QMainWindow, Ui_MainWindow):
    on_edit_status = False
    on_edit_codes_list = None
    on_edit_project_data = None
    on_edit_project_path = None
    def __init__(self):
        super().__init__()
        self.setupUi(self)
        self.NewBtn.clicked.connect(self.New_Project)
        self.NewBtn_Menu.triggered.connect(self.New_Project)
        self.LoadBtn.clicked.connect(self.Open_Project)
        self.LoadBtn_Menu.triggered.connect(self.Open_Project)
        self.SavePro.clicked.connect(self.Save)
        self.SavePro_Menu.triggered.connect(self.Save)
        self.SoftwareName.textChanged.connect(self.Update_SoftwareName)
        self.SoftwareVersion.textChanged.connect(self.Update_SoftwareVersion)
        self.ClosePro.clicked.connect(self.CloseProject)
        self.ClosePro_Menu.triggered.connect(self.CloseProject)

        self.AddFile.clicked.connect(self.Add_CodeFile)
        self.AddFile_Menu.triggered.connect(self.Add_CodeFile)
        self.DeleteFile.clicked.connect(self.Delete_CodeFile)
        self.DeleteFile_Menu.triggered.connect(self.Delete_CodeFile)
        self.UpFile.clicked.connect(self.Up_CodeFile)
        self.DownFile.clicked.connect(self.Down_CodeFile)

        self.GenerateBtn.clicked.connect(self.Generate_File)

    def New_Project(self):
        NewProDialog = NewPro()
        NewProDialog.project_created.connect(self.on_project_created)
        NewProDialog.exec_()

    def Open_Project(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "选择工程文件", ".", "SCRM Files (*.scrm)")
        print(f"[MainWindow] 用户选择了文件！路径为: {file_path}")
        if file_path:
            self.Load_Project(file_path)

    def on_project_created(self, project_path: str):
        print(f"[MainWindow] 接收到工程创建成功的信号，路径为: {project_path}")
        self.Load_Project(project_path)
        

    def Load_Project(self, project_path: str):
        success_flag, project_data = Project_Manager.Load_Project(project_path)
        if not success_flag:
            QMessageBox.critical(None, "错误信息", f"工程加载失败！错误信息：{project_data}")
            return
        self.on_edit_project_data = project_data # 把信息存储到类中
        self.on_edit_codes_list = self.on_edit_project_data.get("source_code_paths")  # 是引用赋值，两个会同时变动
        self.on_edit_project_path = project_path

        # 启用文本框并显示基本信息
        self.SoftwareName.setText(project_data.get("software_name"))
        self.SoftwareVersion.setText(project_data.get("software_version"))
        self.SoftwareName.setEnabled(True)
        self.SoftwareVersion.setEnabled(True)

        # 启用代码文件编辑按钮
        self.AddFile.setEnabled(True)
        self.AddFile_Menu.setEnabled(True)
        self.DeleteFile.setEnabled(True)
        self.DeleteFile_Menu.setEnabled(True)
        self.UpFile.setEnabled(True)
        self.DownFile.setEnabled(True)
        self.SavePro.setEnabled(True)
        self.SavePro_Menu.setEnabled(True)

        # 重要：检查列表里面文件是否都真实地存在？？？
        missing_file_lst = []
        for code_file in self.on_edit_codes_list:
            if not File_Manager.File_Exist(code_file): # 文件不存在
                missing_file_lst.append(code_file)
                
        if missing_file_lst:
            file_names = "\n".join([os.path.basename(f) for f in missing_file_lst[:5]])
            if len(missing_file_lst) > 5:
                file_names += f"\n...等 {len(missing_file_lst)} 个文件"
            
            QMessageBox.warning(self,  "警告信息",  f"发现 {len(missing_file_lst)} 个文件不存在，将自动移除！\n\n{file_names}", QMessageBox.Yes)

            self.on_edit_codes_list = [f for f in self.on_edit_codes_list if f not in missing_file_lst]
            self.on_edit_project_data["source_code_paths"] = self.on_edit_codes_list
            self.on_edit_status = True

        self.Update_CodeBoxUI(0)

        self.NewBtn.setEnabled(False)
        self.NewBtn_Menu.setEnabled(False)
        self.LoadBtn.setEnabled(False)
        self.LoadBtn_Menu.setEnabled(False)
        self.ClosePro.setEnabled(True)
        self.ClosePro_Menu.setEnabled(True)
        self.GenerateBtn.setEnabled(True)
        self.INFO(f"加载工程成功！{project_path}")


    def CloseProject(self):
        if self.on_edit_status: # 还没有退出编辑
            reply = QMessageBox.question(self, "确认关闭", "工程尚未保存，确定要关闭吗？", QMessageBox.Yes | QMessageBox.No | QMessageBox.Cancel)
            if reply == QMessageBox.Cancel or reply == QMessageBox.No:
                self.INFO("用户取消了关闭操作")
                return
            else:
                self.INFO("用户强制关闭工程")
            
        # 禁用文本框并显示基本信息
        self.SoftwareName.setText("")
        self.SoftwareVersion.setText("")
        self.SoftwareName.setEnabled(False)
        self.SoftwareVersion.setEnabled(False)

        # 禁用代码文件编辑按钮
        self.AddFile.setEnabled(False)
        self.AddFile_Menu.setEnabled(False)
        self.DeleteFile.setEnabled(False)
        self.DeleteFile_Menu.setEnabled(False)
        self.UpFile.setEnabled(False)
        self.DownFile.setEnabled(False)
        self.SavePro.setEnabled(False)
        self.SavePro_Menu.setEnabled(False)

        # 类变量成员全部清空
        self.on_edit_status = False
        self.on_edit_codes_list = None
        self.on_edit_project_data = None
        self.on_edit_project_path = None

        self.CodeFiles.clear()

        self.NewBtn.setEnabled(True)
        self.NewBtn_Menu.setEnabled(True)
        self.LoadBtn.setEnabled(True)
        self.LoadBtn_Menu.setEnabled(True)
        self.ClosePro.setEnabled(False)
        self.ClosePro_Menu.setEnabled(False)
        self.GenerateBtn.setEnabled(False)

        self.INFO("关闭工程成功！")


    def Add_CodeFile(self):
        file_paths, _ = QFileDialog.getOpenFileNames(self, "选择代码文件", ".")
        
        if file_paths:
            added_files = []
            for file_path in file_paths:
                if file_path in self.on_edit_codes_list:
                    QMessageBox.warning(None, "警告信息", f"文件 {file_path} 已存在！")
                    continue
                self.on_edit_status = True
                self.on_edit_codes_list.append(file_path)
                added_files.append(file_path)
            
            self.Update_CodeBoxUI(focus=self.CodeFiles.count())
            self.INFO(f"成功添加了{len(added_files)}个文件")
            print(f"[MainWindow] 添加代码文件：用户添加了文件！路径为: {added_files}, 现在有{self.on_edit_project_data.get('source_code_paths')}")

    def Delete_CodeFile(self):
        select_index = self.CodeFiles.currentRow()
        
        if select_index < 0:
            return
        
        total = len(self.on_edit_codes_list)
        deleted = self.on_edit_codes_list[select_index]
        
        del self.on_edit_codes_list[select_index]        # 删除
        
        # 计算新焦点
        if total == 1:
            new_focus = -1
        elif select_index == 0:
            new_focus = 0
        elif select_index == total - 1:
            new_focus = select_index - 1
        else:
            new_focus = select_index
        
        self.on_edit_status = True
        self.Update_CodeBoxUI(focus = new_focus)
        self.INFO(f"已删除: {deleted}")


    def Up_CodeFile(self):
        select_index = self.CodeFiles.currentRow()

        if select_index < 0:
            QMessageBox.warning(self, "警告", "请先选择要移动的文件")
            return
        
        # 检查是否已经在最上面
        if select_index == 0:
            QMessageBox.warning(self, "提示", "已经在最顶部，无法上移")
            return

        current_index = select_index
        
        self.on_edit_codes_list[current_index], self.on_edit_codes_list[current_index - 1] = \
            self.on_edit_codes_list[current_index - 1], self.on_edit_codes_list[current_index]
        
        self.on_edit_status = True        # 更新状态
        new_focus = current_index - 1     # 新的焦点位置（上移后，原项目的位置变成 current_index - 1）
        
        self.Update_CodeBoxUI(focus=new_focus)
        
        self.INFO(f"已上移: {self.on_edit_codes_list[new_focus]}")
    

    def Down_CodeFile(self):
        current_row = self.CodeFiles.currentRow()
        total_items = len(self.on_edit_codes_list)
        
        if current_row < 0:
            QMessageBox.warning(self, "警告", "请先选择要移动的文件")
            return
        
        # 检查是否已经在最底部
        if current_row == total_items - 1:
            QMessageBox.warning(self, "提示", "已经在最底部，无法下移")
            return
        
        self.on_edit_codes_list[current_row], self.on_edit_codes_list[current_row + 1] = \
            self.on_edit_codes_list[current_row + 1], self.on_edit_codes_list[current_row]
      
        self.on_edit_status = True  # 标记修改
        self.Update_CodeBoxUI(focus=current_row + 1) # 更新UI，焦点设置到下移后的位置
        
        self.INFO(f"已下移: {self.on_edit_codes_list[current_row + 1]}")


    def Save(self):
        if not self.on_edit_project_path or not self.on_edit_project_data:
            QMessageBox.warning(None, "警告信息", "没有正在编辑的工程！")
            return
        success_flag, msg = Project_Manager.Modify_Project(self.on_edit_project_path, self.on_edit_project_data)
        if not success_flag:
            QMessageBox.critical(None, "错误信息", f"保存失败:{msg}")
            return
        self.on_edit_status = False # 保存完毕，退出修改
        print(f"[MainWindow] path = {self.on_edit_project_path}  ||  data = {self.on_edit_project_data} 保存成功！")
        self.INFO("保存成功！")

    def Generate_File(self):
        if self.on_edit_status: # 还没有退出编辑
            QMessageBox.warning(self, "警告信息", "工程没有保存，请保存后再生成文件！", QMessageBox.Yes)
            return
        
        if not self.on_edit_codes_list: # 啥都没有
            QMessageBox.warning(self, "警告信息", "没有代码文件可以生成！", QMessageBox.Yes)
            return
        
        self.ProgressBar.setValue(0)
        total_cnt = len(self.on_edit_codes_list)

        # 开始生成
        DocumentWriter.total_codeblock = total_cnt
        DocumentWriter.current_codeblock = 0
        DocumentWriter.software_name = self.on_edit_project_data.get('software_name')
        DocumentWriter.software_version = self.on_edit_project_data.get('software_version')

        # 增加功能：检查模板文件
        if not File_Manager.File_Exist("source-code_model.docx"): # 使用的是相对路径
            QMessageBox.critical(self, "错误", f"找不到模板docx文件！请将模板文件放置在和程序相同的目录下", QMessageBox.Yes)
            return

        export_dir = os.path.dirname(self.on_edit_project_path)
        export_filename = f"{export_dir}/{DocumentWriter.software_name}{DocumentWriter.software_version} - 源代码.docx"

        doc = DocumentWriter.Create()
        for cnt, file_path in enumerate(self.on_edit_codes_list):
            success_flag, msg = DocumentWriter.Generate_One(doc, file_path)
            if not success_flag:
                QMessageBox.warning(self, "警告信息", f"文件{file_path}生成错误：{msg}，文件生成失败！", QMessageBox.Yes)
                self.INFO("生成时发生了错误，生成失败！")
                self.ProgressBar.setValue(0)
                return
            self.INFO(msg)
            self.ProgressBar.setValue(int((cnt + 1)/ total_cnt * 100))
        DocumentWriter.Save(doc, export_filename)
        
        # 处理提示框
        msg_box = QMessageBox(QMessageBox.Information, "提示信息", f"导出成功，导出目录为：\n{export_dir}", parent=self)
        btn_view_dir = msg_box.addButton("打开目录", QMessageBox.ActionRole)
        btn_view_file = msg_box.addButton("打开文件", QMessageBox.ActionRole)
        btn_ok = msg_box.addButton("确定", QMessageBox.AcceptRole)
        msg_box.setDefaultButton(btn_ok)
        msg_box.exec_()

        if msg_box.clickedButton() == btn_view_file:
            File_Manager.open_file(export_filename)
        elif msg_box.clickedButton() == btn_view_dir:
            File_Manager.open_directory(export_dir)

        DocumentWriter.current_codeblock = 0
        DocumentWriter.total_codeblock = 0
        DocumentWriter.software_name = "[软件名称]"
        DocumentWriter.software_version = "[软件版本]"
        
        

    def Update_SoftwareName(self):
        if self.SoftwareName.text() == self.on_edit_project_data.get('software_name'): # 未修改
            return
        self.on_edit_project_data.update({"software_name" : self.SoftwareName.text()})
        self.on_edit_status = True
        self.INFO(f"修改了软件名：{self.on_edit_project_data.get('software_name')}")

    def Update_SoftwareVersion(self):
        if self.SoftwareVersion.text() == self.on_edit_project_data.get('software_version'): # 未修改
            return
        self.on_edit_project_data.update({"software_version" : self.SoftwareVersion.text()})
        self.on_edit_status = True
        self.INFO(f"修改了版本号：{self.on_edit_project_data.get('software_version')}")


    #----------------------下面是UI的一些改变---------------------
    def Update_CodeBoxUI(self, focus: int):
        self.CodeFiles.clear()
        for codefile in self.on_edit_codes_list:
            self.CodeFiles.addItem(codefile)
        if focus!=-1 and 0 <= focus < self.CodeFiles.count():
            self.CodeFiles.setCurrentRow(focus)

    def INFO(self, msg: str):
        time_str = datetime.now().strftime("%H:%M:%S")
        self.Status.setText(f"状态信息：[{time_str}] {msg}")

    #----------------------上面是UI的一些改变---------------------

class NewPro(QDialog, Ui_NewPro):
    project_created = pyqtSignal(str)
    def __init__(self):
        super().__init__()
        self.setupUi(self)
        self.ChoosePath.clicked.connect(self.choose_path)
        self.OKNew.clicked.connect(self.Make_Project_File)
        self.Cancel.clicked.connect(self.close)

    def choose_path(self):
        directory = QFileDialog.getExistingDirectory(self, "选择文件夹", ".") # 弹出目录选择对话框
        
        if directory: # 如果用户选择了路径，则打印或处理
            print(f"[NewPro] 选择的路径是：{directory}")
            self.ProjectPath.setText(directory)

    def Make_Project_File(self):
        project_name = self.ProjectName.text()
        software_name = self.SoftwareName.text()
        software_version = self.SoftwareVersion.text()
        project_directory = self.ProjectPath.text()

        if not project_name.strip() or not software_name.strip() or not software_version.strip() or not project_directory.strip():
            QMessageBox.warning(None, "警告信息", "工程名称、软件名称、版本号不能为空！")
            return

        if not bool(re.match(r"^V?\d+(\.\d+)*$", software_version)):
            QMessageBox.warning(None, "警告信息", "版本号撰写有误! 标准格式: V数字(含小数点)或数字, 如V1.0或1.0")
            return
        
        if not Project_Manager.Check_Valid_Path(project_directory):
            QMessageBox.warning(None, "警告信息", "选择的路径有误！检查路径是否存在及其读写权限！")
        
        success_flag, path_msg = Project_Manager.New_Project(project_name, project_directory, software_name, software_version)
        if success_flag:
            QMessageBox.information(None, "提示信息", "工程创建成功！即将打开工程！")
            print(f"[NewPro] 工程创建成功, path = {path_msg}")
        else:
            QMessageBox.critical(None, "错误信息", f"工程创建失败！错误信息：{path_msg}")
            print(f"[NewPro] 工程创建失败, msg = {path_msg}")
            return
        self.project_created.emit(path_msg)
        self.close()


if __name__ == "__main__":
    QApplication.setAttribute(Qt.AA_EnableHighDpiScaling)
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())