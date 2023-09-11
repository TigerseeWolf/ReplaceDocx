import os

from PyQt5.QtWidgets import *
from gui.ReplaceWindow import Ui_ReplaceWindow
from PyQt5.QtCore import Qt
import sys
from docxtpl import DocxTemplate
import pandas as pd


class ReplaceDocx(QDialog, Ui_ReplaceWindow):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setupUi(self)
        self.excel_path = ''
        self.docx_path = ''
        self.df_data = pd.DataFrame()
        self.docx = None
        self.choseExcelButton.clicked.connect(self.chose_excel_path)
        self.choseWrodButton.clicked.connect(self.chose_word_path)
        self.replaceButton.clicked.connect(self.replace)

    def load_excel(self):
        """导入EXCEL数据"""
        if not self.excel_path:
            QMessageBox().critical(self, "提示", '错误：Excel数据文件路径为空！', QMessageBox.Yes | QMessageBox.No, QMessageBox.Yes)
            return
        try:
            self.df_data = pd.read_excel(self.excel_path)
        except:
            QMessageBox().critical(self, "提示", '错误：Excel数据文件读取失败！', QMessageBox.Yes | QMessageBox.No, QMessageBox.Yes)

    def chose_excel_path(self):
        self.excel_path, fd = QFileDialog.getOpenFileName(self, "请选择Excel数据文件", "./", "Excel(*.xlsx);;")
        self.lineEdit_excel_path.setText(self.excel_path)
        self.load_excel()

    def load_word(self):
        """载入word模板"""
        if not self.docx_path:
            QMessageBox().critical(self, "提示", '错误：Word模板文件路径为空！', QMessageBox.Yes | QMessageBox.No, QMessageBox.Yes)
            return
        self.docx = DocxTemplate(self.docx_path)

    def chose_word_path(self):
        self.docx_path, fd = QFileDialog.getOpenFileName(self, "请选择Word模板文件", "./", "Word(*.docx);;")
        self.lineEdit_word_path.setText(self.docx_path)
        self.load_word()

    def replace(self):
        """替换文件"""
        if self.df_data.empty:
            QMessageBox().critical(self, "提示", '数据读取为空！', QMessageBox.Yes | QMessageBox.No, QMessageBox.Yes)
            return
        if not self.docx_path:
            QMessageBox().critical(self, "提示", 'Word模板为空！', QMessageBox.Yes | QMessageBox.No, QMessageBox.Yes)
            return

        if not os.path.exists('./输出文档'):
            os.mkdir('./输出文档')  # 创建文件夹

        output_name = self.lineEdit_output_name.text()

        self.replaceButton.setEnabled(False)
        self.replaceButton.update()
        for i in range(len(self.df_data)):
            tpl = self.docx
            dict_context = self.df_data.iloc[i, :].to_dict()
            tpl.render(dict_context)
            if output_name:
                name = output_name.format(**dict_context)
            else:
                name = self.docx_path.split('/')[-1].split(".")[0] + "_" + str(i + 1)
            tpl.save(f"./输出文档/{name}.docx")
            self.replaceButton.setText(f"第{i + 1}/{len(self.df_data)}个")
            self.replaceButton.update()

        QMessageBox.about(self, "提示", "批量替换完成")
        self.replaceButton.setEnabled(True)
        self.replaceButton.setText("替换")


if __name__ == "__main__":
    # 固定的，PyQt5程序都需要QApplication对象。sys.argv是命令行参数列表，确保程序可以双击运行
    QApplication.setAttribute(Qt.AA_EnableHighDpiScaling)
    app = QApplication(sys.argv)
    # 初始化
    myWin = ReplaceDocx()
    # 将窗口控件显示在屏幕上
    myWin.show()
    # 程序运行，sys.exit方法确保程序完整退出。
    sys.exit(app.exec_())
