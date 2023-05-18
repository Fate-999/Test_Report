import os
import sys
import App
import Copy_Test_Case
import Find_Issue_Number
import Copy_Issue_To_Report
import Handle_Issue_Table
import Merge_Report_Table
import Judge_New_Issue
import Refresh_All_Table
from PyQt5 import QtWidgets, QtCore
from PyQt5.QtCore import QDate, QThread, pyqtSignal
from PyQt5.QtWidgets import QApplication, QWidget, QPushButton, QFileDialog, QHBoxLayout, QVBoxLayout, QRadioButton, \
    QDateEdit, QMessageBox


class QmyWidget(QWidget):

    def __init__(self, parent=None):
        super().__init__(parent)  # 调用父类的构造函数，创建QWidget窗体
        # 设置窗体大小及标题
        self.resize(800, 600)
        self.setWindowTitle("单体报告自动化工具")
        self.setupUi()

    def setupUi(self):
        # 创建布局
        self.all_layout = QVBoxLayout()
        self.layout1 = QHBoxLayout()
        self.layout2 = QHBoxLayout()
        self.layout3 = QHBoxLayout()
        self.layout4 = QHBoxLayout()
        self.layout8 = QHBoxLayout()
        self.layout5 = QHBoxLayout()
        self.layout6 = QHBoxLayout()
        self.layout7 = QHBoxLayout()
        # 创建文件Items
        self.select_result = QPushButton("选择文件", self)
        self.select_result.clicked.connect(self.select_result_url, )
        self.select_issue = QPushButton("选择文件", self)
        self.select_issue.clicked.connect(self.select_issue_url, )
        self.test_report_btn = QPushButton("选择文件夹", self)
        self.test_report_btn.clicked.connect(self.select_dir)
        self.result = QtWidgets.QLabel("请输入result文件地址：", self)
        self.result_url = QtWidgets.QLabel(" ", self)
        self.issue = QtWidgets.QLabel("请输入issue文件地址：", self)
        self.issue_url = QtWidgets.QLabel(" ", self)
        self.yes_or_no = QtWidgets.QLabel("请选择是否为完整报告", self)
        self.date = QtWidgets.QLabel("请输入BIOS Release日期", self)
        self.test_report = QtWidgets.QLabel("请选择TestReport存放地址", self)
        self.runing_status = QtWidgets.QLabel("等待数据..", self)
        self.test_report_url = QtWidgets.QLabel(" ", self)
        self.template = QtWidgets.QLabel("请选择模板文件", self)
        self.template_url = QtWidgets.QLabel(" ", self)
        self.select_template_btn = QPushButton("选择文件", self)
        self.select_template_btn.clicked.connect(self.select_template_url, )
        self.run = QPushButton("Run", self)
        self.run.clicked.connect(self.run_start, )
        self.reset_btn = QPushButton("reset", self)
        self.reset_btn.clicked.connect(self.reset, )
        self.muti_result = QtWidgets.QLabel("请选择result文件夹", self)
        self.muti_result_url = QtWidgets.QLabel(" ", self)
        self.muti_result_btn = QPushButton("选择文件夹", self)
        self.muti_result_btn.clicked.connect(self.select_dir2)
        self.muti_result.hide()
        self.muti_result_btn.hide()
        self.muti_result_url.hide()
        # 单选按钮 是
        self.yes_btn = QRadioButton()
        self.yes_btn.setChecked(True)
        self.yes_btn.setToolTip("Yes")
        self.yes_btn.setText("是")
        self.yes_btn.clicked.connect(self.selectOption)

        # 单选按钮 否
        self.no_btn = QRadioButton()
        self.no_btn.setChecked(False)
        self.no_btn.setToolTip("No")
        self.no_btn.setText("否")
        self.no_btn.clicked.connect(self.selectOption)

        # 日期组件
        self.start_date = QDateEdit(QDate(2023, 5, 1))
        self.start_date.setDisplayFormat("yyyy/MM/dd")  # 设置日期格式
        self.start_date.setMinimumDate(QDate.currentDate().addDays(-365))  # 设置最小日期
        self.start_date.setMaximumDate(QDate.currentDate())  # 设置最大日期
        self.start_date.setCalendarPopup(True)

        # 将组件添加到布局中
        self.layout1.addWidget(self.result)
        self.layout1.addWidget(self.result_url)
        self.layout1.addWidget(self.select_result)
        self.layout2.addWidget(self.issue)
        self.layout2.addWidget(self.issue_url)
        self.layout2.addWidget(self.select_issue)
        self.layout3.addWidget(self.yes_or_no)
        self.layout3.addWidget(self.yes_btn)
        self.layout3.addWidget(self.no_btn)
        self.layout8.addWidget(self.muti_result)
        self.layout8.addWidget(self.muti_result_url)
        self.layout8.addWidget(self.muti_result_btn)
        self.layout4.addWidget(self.date)
        self.layout4.addWidget(self.start_date)
        self.layout5.addWidget(self.test_report)
        self.layout5.addWidget(self.test_report_url)
        self.layout5.addWidget(self.test_report_btn)
        self.layout6.addWidget(self.template)
        self.layout6.addWidget(self.template_url)
        self.layout6.addWidget(self.select_template_btn)
        self.layout7.addWidget(self.runing_status)
        self.layout7.addWidget(self.reset_btn)
        self.layout7.addWidget(self.run)
        self.all_layout.addLayout(self.layout1)
        self.all_layout.addLayout(self.layout2)
        self.all_layout.addLayout(self.layout3)
        self.all_layout.addLayout(self.layout8)
        self.all_layout.addLayout(self.layout4)
        self.all_layout.addLayout(self.layout5)
        self.all_layout.addLayout(self.layout6)
        self.all_layout.addLayout(self.layout7)
        # 为窗体添加布局
        self.setLayout(self.all_layout)

    def select_result_url(self):
        """选择文件对话框"""
        # QFileDialog组件定义
        fileDialog = QFileDialog(self)
        # QFileDialog组件设置
        fileDialog.setWindowTitle("请选择Result文件")  # 设置对话框标题
        fileDialog.setFileMode(QFileDialog.AnyFile)  # 设置能打开文件的格式
        fileDialog.setNameFilter("EXCEL (*.xlsx*)")  # 按文件名过滤
        file_path = fileDialog.exec()  # 窗口显示，返回文件路径
        if file_path and fileDialog.selectedFiles():
            print("选择文件成功：{}".format(fileDialog.selectedFiles()[0]))
            self.result_url.setText(fileDialog.selectedFiles()[0])

    def select_issue_url(self):
        """选择文件对话框"""
        # QFileDialog组件定义
        fileDialog = QFileDialog(self)
        # QFileDialog组件设置
        fileDialog.setWindowTitle("请选择Result文件")  # 设置对话框标题
        fileDialog.setFileMode(QFileDialog.AnyFile)  # 设置能打开文件的格式
        fileDialog.setNameFilter("EXCEL (*.xlsx*)")  # 按文件名过滤
        file_path = fileDialog.exec()  # 窗口显示，返回文件路径
        if file_path and fileDialog.selectedFiles():
            print("选择文件成功：{}".format(fileDialog.selectedFiles()[0]))
            self.issue_url.setText(fileDialog.selectedFiles()[0])

    def select_template_url(self):
        """选择文件对话框"""
        # QFileDialog组件定义
        fileDialog = QFileDialog(self)
        # QFileDialog组件设置
        fileDialog.setWindowTitle("请选择Result文件")  # 设置对话框标题
        fileDialog.setFileMode(QFileDialog.AnyFile)  # 设置能打开文件的格式
        fileDialog.setNameFilter("EXCEL (*.xlsx*)")  # 按文件名过滤
        file_path = fileDialog.exec()  # 窗口显示，返回文件路径
        if file_path and fileDialog.selectedFiles():
            print("选择文件成功：{}".format(fileDialog.selectedFiles()[0]))
            self.template_url.setText(fileDialog.selectedFiles()[0])

    def selectOption(self):
        if self.yes_btn.isChecked():
            self.muti_result.hide()
            self.muti_result_btn.hide()
            self.muti_result_url.hide()
            self.result.show()
            self.result_url.show()
            self.select_result.show()
            print('选择了 是 ')
        elif self.no_btn.isChecked():
            self.muti_result_url.show()
            self.muti_result.show()
            self.muti_result_btn.show()
            self.result.hide()
            self.result_url.hide()
            self.select_result.hide()
            print('选择了 否 ')

    def select_dir(self):
        """选择文件夹对话框架"""
        dir_path = QFileDialog.getExistingDirectory(self, '标题', os.getcwd())
        if dir_path:
            print("选择文件夹成功：{}".format(dir_path))
            self.test_report_url.setText(dir_path+"/Test_Report.xlsx")

    def select_dir2(self):
        """选择文件夹对话框架"""
        dir_path = QFileDialog.getExistingDirectory(self, '标题', os.getcwd())
        if dir_path:
            print("选择文件夹成功：{}".format(dir_path))
            self.muti_result_url.setText(dir_path)

    def run_start(self):
        Time = self.start_date.text()
        result_url = self.result_url.text()
        issue_url = self.issue_url.text()
        report_url = self.test_report_url.text()
        template_url = self.template_url.text()
        mutiresult_url = self.muti_result_url.text()
        app = App.Launch_App()

        if self.yes_btn.isChecked():
            try:
                Handle_Issue_Table.Handle_issue_table(app, issue_url)
                Copy_Test_Case.Copy_test_case(app, result_url,report_url,template_url)
                Find_Issue_Number.Search_bugid_paste_allissue(app,report_url)
                Copy_Issue_To_Report.Copy_Case(app, issue_url,report_url)
                Judge_New_Issue.Judge(app, Time,report_url)
                Refresh_All_Table.Refresh_all(app,report_url)

            except Exception as e:
                print(e)

            finally:
                self.runing_status.setText("处理完成！")
                app.quit()
        else:
            try:
                Merge_Report_Table.Merge_Report_Table(app,mutiresult_url)
                Handle_Issue_Table.Handle_issue_table(app, issue_url)
                Copy_Test_Case.Copy_multi_test_case(app,report_url,mutiresult_url,template_url)
                Find_Issue_Number.Search_bugid_paste_allissue(app, report_url)
                Copy_Issue_To_Report.Copy_Case(app, issue_url, report_url)
                Judge_New_Issue.Judge(app, Time, report_url)
                Refresh_All_Table.Refresh_all(app, report_url)
            except Exception as e:
                print(e)
            finally:
                self.runing_status.setText("处理完成！")
                app.quit()

    def reset(self):
        self.runing_status.setText("等待数据..")
        self.test_report_url.setText(" ")
        self.issue_url.setText(" ")
        self.result_url.setText(" ")
        self.template_url.setText(" ")
        self.start_date.setDate(QDate(2023, 5, 1))

if __name__ == '__main__':
    app = QApplication(sys.argv)
    myMain = QmyWidget()
    myMain.show()
    sys.exit(app.exec_())
