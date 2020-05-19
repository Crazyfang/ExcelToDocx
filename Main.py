import sys
import Surface
from PyQt5.QtWidgets import QApplication, QMainWindow, QFileDialog, QMessageBox
from PyQt5.QtCore import QThread, pyqtSignal
import os


class ThreadTransfer(QThread):
    signOut = pyqtSignal(str, float)

    def __init__(self, path, type):
        super(ThreadTransfer, self).__init__()
        self.filepath = path
        self.type = type

    def run(self):
        self.signOut.emit('程序开始处理', 90)
        # manage = ConverseClass()

        if os.path.isdir(self.filepath):
            # result = manage.read_file_from_directory(self.filepath)
            # for i in result:
            #     self.signOut.emit(i, 90)
            pass
        else:
            # result = manage.read_file_from_single_file_path(self.filepath)
            # self.signOut.emit(result, 90)
            pass
        self.signOut.emit('处理完成', 100)


class ExcelToWord(QMainWindow, Surface.Ui_MainWindow):
    """
    Interface the user watch
    """

    def __init__(self):
        QMainWindow.__init__(self)
        Surface.Ui_MainWindow.__init__(self)
        self.setupUi(self)

        # 1-选取文件 2-选取文件夹
        self.type = 1
        self.workthread = None
        self.progressBar.setValue(0)
        self.pushButton_FileOrDirectory.clicked.connect(self.select_excel_file)
        self.radioButton_File.toggled.connect(self.change_type)
        self.radioButton_Directory.toggled.connect(self.change_type)
        self.pushButton_Start.clicked.connect(self.start_process)

    def select_excel_file(self):
        if 1 == self.type:
            filename_choose, file_type = QFileDialog.getOpenFileName(self, '打开',
                                                                     os.path.join(os.path.expanduser("~"), 'Desktop'),
                                                                     'Excel文件 (*.xlsx);;All Files (*)')
            self.lineEdit_FileOrDirectory.setText(filename_choose)
        else:
            directory_path = QFileDialog.getExistingDirectory(self, '选取文件夹', './')
            self.lineEdit_FileOrDirectory.setText(directory_path)

    def start_process(self):
        if not self.lineEdit_FileOrDirectory.text():
            QMessageBox.information(self, '提示', '请选择文件或者文件夹!')
        else:
            pass
        self.workthread = ThreadTransfer(self.lineEdit_FileOrDirectory.text(),
                                         self.type)

        self.workthread.signOut.connect(self.list_add)
        self.Button_Start.setEnabled(False)
        self.Button_Start.setText('正在处理')
        self.workthread.start()

    def list_add(self, message):
        for file_name in message:
            self.listWidget_Adjust.addItem(file_name)

    def get_file_name_list(self):
        item_count = self.listWidget_Adjust.count()
        file_name_list = []
        for i in item_count:
            file_name_list.append(self.listWidget_Adjust.item(i).text())

    def change_type(self):
        if self.radioButton_File.isChecked():
            self.type = 1
        else:
            self.type = 2


if __name__ == '__main__':
    app = QApplication(sys.argv)
    # MainWindow = QMainWindow()
    ui = ExcelToWord()
    # ui.setupUi(MainWindow)
    ui.show()
    sys.exit(app.exec_())
