from mainwindow import *
app = QtWidgets.QApplication(sys.argv)
mainWindow = QtWidgets.QMainWindow()
ui = Ui_mainWindow()
ui.setupUi(mainWindow)
mainWindow.setFixedSize(1440, 960)
ui.textBrowser.append("欢迎使用错题集生成器！\n")
mainWindow.show()

sys.exit(app.exec_())