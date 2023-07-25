import typing

from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtWidgets import QMainWindow, QFileDialog, QWidget, QFormLayout, QLineEdit, QComboBox, QPushButton, \
    QMessageBox, QAction
from PyQt5.QtCore import QStringListModel, Qt
import os
import re
import shutil
import sys



class Ui_mainWindow(QMainWindow):
    def setupUi(self, mainWindow):
        mainWindow.setObjectName("mainWindow")
        mainWindow.setEnabled(True)
        mainWindow.resize(1440, 960)

        self.centralwidget = QtWidgets.QWidget(mainWindow)
        self.centralwidget.setObjectName("centralwidget")
        self.pushButton = QtWidgets.QPushButton(self.centralwidget)
        self.pushButton.setGeometry(QtCore.QRect(0, 0, 480, 220))
        font = QtGui.QFont()
        font.setPointSize(15)
        self.pushButton.setFont(font)
        self.pushButton.setObjectName("pushButton")
        self.pushButton_2 = QtWidgets.QPushButton(self.centralwidget)
        self.pushButton_2.setGeometry(QtCore.QRect(0, 220, 480, 220))
        font = QtGui.QFont()
        font.setPointSize(15)
        self.pushButton_2.setFont(font)
        self.pushButton_2.setObjectName("pushButton_2")
        self.pushButton_3 = QtWidgets.QPushButton(self.centralwidget)
        self.pushButton_3.setGeometry(QtCore.QRect(0, 440, 480, 220))
        font = QtGui.QFont()
        font.setPointSize(15)
        self.pushButton_3.setFont(font)
        self.pushButton_3.setObjectName("pushButton_3")
        self.pushButton_4 = QtWidgets.QPushButton(self.centralwidget)
        self.pushButton_4.setGeometry(QtCore.QRect(0, 660, 480, 220))
        font = QtGui.QFont()
        font.setPointSize(15)
        self.pushButton_4.setFont(font)
        self.pushButton_4.setObjectName("pushButton_4")
        self.textBrowser = QtWidgets.QTextBrowser(self.centralwidget)
        self.textBrowser.setGeometry(QtCore.QRect(480, 0, 622, 880)) 
        font = QtGui.QFont()
        font.setPointSize(15)
        self.textBrowser.setFont(font)
        self.textBrowser.setObjectName("textBrowser")
        self.listView = QtWidgets.QListView(self.centralwidget)
        self.listView.setGeometry(QtCore.QRect(1100, 0, 342, 722)) 
        self.listView.setObjectName("listView")
        self.listView.setEditTriggers(QtWidgets.QAbstractItemView.NoEditTriggers)
        self.pushButton_7 = QtWidgets.QPushButton(self.centralwidget)
        self.pushButton_7.setGeometry(QtCore.QRect(1100, 720, 342, 80))
        font = QtGui.QFont()
        font.setPointSize(15)
        self.pushButton_7.setFont(font)
        self.pushButton_7.setObjectName("pushButton_7")
        self.pushButton_8 = QtWidgets.QPushButton(self.centralwidget)
        self.pushButton_8.setGeometry(QtCore.QRect(1100, 800, 342, 80))
        font = QtGui.QFont()
        font.setPointSize(15)
        self.pushButton_8.setFont(font)
        self.pushButton_8.setObjectName("pushButton_8")
        mainWindow.setCentralWidget(self.centralwidget)
        self.menubar = QtWidgets.QMenuBar(mainWindow)
        self.menubar.setGeometry(QtCore.QRect(0, 0, 1440, 17)) 
        self.menubar.setObjectName("menubar")
        self.menu = QtWidgets.QMenu(self.menubar)
        self.menu.setObjectName("help_menu")

        about_developer_action = QAction('关于开发者', self)
        about_developer_action.triggered.connect(self.show_about_developer)
        self.menu.addAction(about_developer_action)
        help_action = QAction('使用帮助', self)
        help_action.triggered.connect(self.show_help)
        self.menu.addAction(help_action)

        mainWindow.setMenuBar(self.menubar)
        self.statusbar = QtWidgets.QStatusBar(mainWindow)
        self.statusbar.setObjectName("statusbar")
        mainWindow.setStatusBar(self.statusbar)
        self.menubar.addAction(self.menu.menuAction())

        self.retranslateUi(mainWindow)
        QtCore.QMetaObject.connectSlotsByName(mainWindow)

        self.pushButton.clicked.connect(self.open_problem_shots_dialog)
        self.pushButton_2.clicked.connect(self.open_feedback_dialog)
        self.pushButton_3.clicked.connect(self.setting_info)
        self.pushButton_4.clicked.connect(self.start_making)
        self.pushButton_7.clicked.connect(self.recover_backup)
        self.pushButton_8.clicked.connect(self.reset)

        self.feedback_file = ""
        self.problem_shots_dir = ""
        self.class_name=""
        self.subject_name=""
        self.date=""
        self.week=""

        self.step_completed = [False, False, False]

    def show_about_developer(self):
        QMessageBox.information(self, '关于开发者', '这个错题集生成器由Leonard开发，Lrj参与了Demo前期功能制作，旨在帮助TR的助教同学们一定程度上免于耗时的错题集制作\n项目地址：https://github.com/Leonardlsa/problems-folder-generator')

    def show_help(self):
        QMessageBox.information(self, '使用帮助', '''<h3>使用方法：</h3>
<blockquote>
    <ol>
        <li>将讲义题目整齐截图png格式（截图长度尽量一致即可），按题号标注图片名，均以2位数字为名（如01，08，15，36），整理至文件夹A</li>
        <li>将反馈表以第一列姓名，第二列错题，第一行写表头的格式规范填写<sup>[^规范]</sup><br>
            [^规范]:
            <ol>
                <li>姓名栏设置居中直接填写，之前不要打空格</li>
                <li>错题号以数字填写（不用补两位），以中文逗号分隔</li>
                <li>不要打任何多余的符号/空格</li>
                <li>第二列仅输入"全对"/"未交"/"未带"/以中文逗号分隔的数字串，中的一种</li>
            </ol>
        </li>
        <li>将讲义文件夹A，错题反馈表B，本程序置于同一目录</li>
        <li>运行程序，并按要求填写文件夹A，错题反馈表B，学科，日期，开课学期数即可自动完成</li>
    </ol>
</blockquote>
''')

    def retranslateUi(self, mainWindow):
        _translate = QtCore.QCoreApplication.translate
        mainWindow.setWindowTitle(_translate("mainWindow", "错题集生成器"))
        self.pushButton.setText(_translate("mainWindow", "选定习题截图所在文件夹"))
        self.pushButton_2.setText(_translate("mainWindow", "选定本科目今日错题反馈表"))
        self.pushButton_3.setText(_translate("mainWindow", "填写反馈表信息"))
        self.pushButton_4.setText(_translate("mainWindow", "开始制作错题集"))
        self.pushButton_7.setText(_translate("mainWindow", "恢复备份"))
        self.pushButton_8.setText(_translate("mainWindow", "开始新的任务"))
        self.menu.setTitle(_translate("mainWindow", "帮助(建议先看看)"))

    def open_problem_shots_dialog(self):
        chosen = False
        if self.problem_shots_dir:
            chosen = True
        self.problem_shots_dir =QFileDialog.getExistingDirectory(self, "选择目录", options=QFileDialog.ShowDirsOnly)
        if self.problem_shots_dir:
            if chosen:
                self.textBrowser.append("已<font color='blue'>重新</font>选择习题截图所在文件夹：")
            else:
                self.textBrowser.append("已选择习题截图所在文件夹：")
            self.textBrowser.append("<font color='red'>"+ str(self.problem_shots_dir)+"</font>"+"下。\n")
            for _ in range(0, self.listView.modelColumn()):
                self.listView.model.removeRows(_)
            pictures = []
            for root, dirs, files in os.walk(self.problem_shots_dir, topdown=False):
                for f in files:
                    if f.endswith(".jpg") or f.endswith(".png"):
                        if f[:-4].isdigit():
                            if len(f)==5:
                                os.rename(os.path.join(root, f), os.path.join(root, "0"+f))# 为了让文件名按照数字顺序排列，将1.jpg改为01.jpg,且默认题号至多两位
                                f="0"+f
                            pictures.append(str(f))
            pictures.sort()
            self.listView.setModel(QStringListModel(pictures))
            if len(pictures)==0:
                self.textBrowser.append("<font color='red'>警告：</font>所选文件夹下没有图片！\n")
            else:
                self.pushButton.setStyleSheet("background-color: rgb(127,255,127);")
                self.step_completed[0] = True
                self.textBrowser.append("所选文件夹下共有<font color='blue'>"+str(len(pictures))+"</font>张图片。详细图片列表请看右侧栏\n")

    def open_feedback_dialog(self):
        chosen = False
        if self.feedback_file:
            chosen = True
        self.feedback_file, _ = QFileDialog.getOpenFileName(self, "选择文件", filter="Excel Files (*.xlsx)", options=QFileDialog.ReadOnly)
        if self.feedback_file:
            if chosen:
                self.textBrowser.append("已<font color='blue'>重新</font>选择本科目今日错题反馈表：")
            else:
                self.textBrowser.append("已选择本科目今日错题反馈表：")
            self.textBrowser.append("<font color='red'>" + str(self.feedback_file) + "</font>" + "。\n")
            self.pushButton_2.setStyleSheet("background-color: rgb(127,255,127);")
            self.step_completed[1] = True

    def setting_info(self):
        pop_wgt=QWidget()
        form_wgt=QWidget(pop_wgt)
        form_wgt.setGeometry(0, 0, 500, 240)
        layout=QFormLayout()

        class_line=QLineEdit()
        class_line.setPlaceholderText("请输入班级")
        class_line.setMaxLength(10)
        # class_line.setStyleSheet("height: 40px;")
        layout.addRow("班级：", class_line)

        subject_line=QComboBox()
        subject_line.addItems(["数学", "物理", "化学", "生物"])
        layout.addRow("科目：", subject_line)

        date=QLineEdit()
        date.setPlaceholderText("格式如0720")
        date.setMaxLength(4)
        layout.addRow("日期：", date)

        week=QLineEdit()
        week.setPlaceholderText("第几周，如3")
        week.setMaxLength(2)
        layout.addRow("周数：", week)

        date_pattern=r'\d{4}'
        match=re.search(date_pattern, self.feedback_file.split("/")[-1])
        if match:
            date.setText(match.group())

        import configparser
        config = configparser.ConfigParser()
        config.read("config.ini")
        if config.has_section("info"):
            if config.has_option("info", "class"):
                class_line.setText(config.get("info", "class"))
                self.class_name = config.get("info", "class")

        form_wgt.setLayout(layout)

        complete_button=QPushButton("完成",parent=pop_wgt)
        complete_button.clicked.connect(lambda: self.complete_setting_info(class_line, subject_line, date, week,pop_wgt))
        complete_button.move(200, 240)

        pop_wgt.setStyleSheet(
            """
            * {
                font-family: "Microsoft YaHei UI";
                font-size: 20px; /* 设置字体大小 */
                height: 40px; /* 设置高度 */
            }
            """
        )
        pop_wgt.resize(500, 300)

        pop_wgt.setWindowTitle("填写反馈表信息")
        pop_wgt.setWindowModality(Qt.ApplicationModal)
        pop_wgt.show()

    def complete_setting_info(self, class_line, subject_line, date, week,pop_wgt):
        if class_line.text() and subject_line.currentText() and date.text() and week.text():
            self.class_name=class_line.text()
            self.textBrowser.append("已填写反馈表信息：")
            self.textBrowser.append("<font color='blue'>班级：</font>"+class_line.text())
            self.textBrowser.append("<font color='blue'>科目：</font>"+subject_line.currentText())
            self.textBrowser.append("<font color='blue'>日期：</font>"+date.text())
            self.textBrowser.append("<font color='blue'>周数：</font>"+week.text()+"\n")
            self.class_name = class_line.text()
            self.subject = subject_line.currentText()
            self.date = date.text()
            self.week = week.text()

            import configparser
            config = configparser.ConfigParser()
            config.read("config.ini")
            if not config.has_section("info"):
                config.add_section("info")
            config.set("info", "class", self.class_name)
            config.write(open("config.ini", "w"))

            self.pushButton_3.setStyleSheet("background-color: rgb(127,255,127);")
            self.step_completed[2] = True
            pop_wgt.close()
        else:
            QMessageBox.warning(self, "警告", "请填写完整信息！")

    def start_making(self):
        if self.step_completed[0] and self.step_completed[1] and self.step_completed[2]:
            self.textBrowser.append("开始生成错题集...\n")
            from core import generate
            try:
                generate(self.problem_shots_dir,self.feedback_file,self.class_name, self.subject,self.week,self.date,self)
                self.textBrowser.append("<font color='green'>错题集已生成！</font>\n")
                self.pushButton_4.setStyleSheet("background-color: rgb(127,255,127);")
                if os.path.exists("backup\\"):
                    shutil.rmtree("backup\\")
            except Exception as e:
                self.textBrowser.append("<font color='red'>错误：</font>"+str(e)+"\n")
                self.textBrowser.append("<font color='red'>建议立即恢复备份</font>\n")
        else:
            QMessageBox.warning(self, "警告", "请先完成前三个步骤！")

    def reset(self):
        self.textBrowser.clear()
        for _ in range(0, self.listView.modelColumn()):
            self.listView.model.removeRows(_)
        self.pushButton.setStyleSheet("background-color: rgb(255,255,255);")
        self.pushButton_2.setStyleSheet("background-color: rgb(255,255,255);")
        self.pushButton_3.setStyleSheet("background-color: rgb(255,255,255);")
        self.pushButton_4.setStyleSheet("background-color: rgb(255,255,255);")
        self.step_completed = [False, False, False]
        self.problem_shots_dir = ""
        self.feedback_file = ""
        self.class_name = ""
        self.subject = ""
        self.date = ""
        self.week = ""

    def recover_backup(self):
        self.textBrowser.append("正在恢复备份...\n")
        if os.path.exists("backup\\"):
            if os.path.exists("学生\\"):
                shutil.rmtree("学生\\")
            shutil.copytree("backup\\", "学生\\")
            self.textBrowser.append("<font color='green'>备份已恢复！</font>\n")
            shutil.rmtree("backup\\")
        else:
            self.textBrowser.append("<font color='red'>备份不存在！</font>\n这很可能发生于你本学期第一次使用错题集生成器\n")

