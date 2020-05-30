import sys
import Surface
from PyQt5.QtWidgets import QApplication, QMainWindow, QFileDialog, QMessageBox
from PyQt5.QtCore import QThread, pyqtSignal
import os
from Function import FuncOfConvert
from DocxCombine import main


class ThreadTransfer(QThread):
    signOut = pyqtSignal(str, float)

    def __init__(self, path, select_type):
        super(ThreadTransfer, self).__init__()
        self.filepath = path
        self.type = select_type

    def run(self):
        manage = FuncOfConvert()

        if os.path.isdir(self.filepath):
            result = manage.get_file_list(self.filepath)
            print(result)
            file_count = len(result)
            for index, i in enumerate(result):
                extension = os.path.splitext(i)[-1][1:]
                if extension == 'xlsx':
                    manage.read_data_from_excel(i)
                else:
                    manage.read_data_from_xls(i)
                manage.win32test(i)
                self.signOut.emit(os.path.splitext(i)[0] + '.docx', (index + 1) / file_count * 100 - 1)
            pass
        else:
            extension = os.path.splitext(self.filepath)[-1][1:]
            if extension == 'xlsx':
                manage.read_data_from_excel(self.filepath)
            else:
                manage.read_data_from_xls(self.filepath)
            manage.win32test(self.filepath)
            self.signOut.emit(os.path.splitext(self.filepath)[0] + '.docx', 90)
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
        self.pushButton_Combine.clicked.connect(self.combine_file)

    def select_excel_file(self):
        if 1 == self.type:
            filename_choose, file_type = QFileDialog.getOpenFileName(self, '打开',
                                                                     os.path.join(os.path.expanduser("~"), 'Desktop'),
                                                                     'Excel文件 (*.xlsx);;All Files (*)')
            self.lineEdit_FileOrDirectory.setText(filename_choose)
        else:
            directory_path = QFileDialog.getExistingDirectory(self, '选取文件夹', './')
            self.lineEdit_FileOrDirectory.setText(directory_path)

    def combine_file(self):
        file_name_list = self.get_file_name_list()
        if file_name_list:
            process_sign = main(file_name_list)
            if process_sign[0]:
                QMessageBox.information(self, '提示', '合并成功,文件名称合并.docx!')
            else:
                QMessageBox.information(self, '提示', '合并失败,{},请检查后重试!'.format(process_sign[1]))
        else:
            QMessageBox.information(self, '提示', '请先进行转换操作再进行合并功能!')

    def start_process(self):
        if not self.lineEdit_FileOrDirectory.text():
            QMessageBox.information(self, '提示', '请选择文件或者文件夹!')
        else:
            pass
        self.workthread = ThreadTransfer(self.lineEdit_FileOrDirectory.text(),
                                         self.type)

        self.workthread.signOut.connect(self.list_add)
        self.pushButton_Start.setEnabled(False)
        self.pushButton_Start.setText('正在处理')
        self.workthread.start()

    def list_add(self, file_name, number):
        if number == 100:
            self.pushButton_Start.setEnabled(True)
            self.pushButton_Start.setText('开始处理')
        else:
            self.listWidget_Adjust.addItem(file_name)
        self.progressBar.setValue(number)

    def get_file_name_list(self):
        item_count = self.listWidget_Adjust.count()
        file_name_list = []
        for i in range(item_count):
            file_name_list.append(self.listWidget_Adjust.item(i).text())

        return file_name_list

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
