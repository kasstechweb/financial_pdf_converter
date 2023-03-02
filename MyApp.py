import os
from PyQt6.QtWidgets import (QWidget, QPushButton, QFileDialog,
                                QLabel, QGridLayout, QMessageBox, QListWidget)
import shutil
from convert import convert
class MyApp(QWidget):
    balance_sheet = ''
    income_statement_sheet = ''
    def __init__(self):
        super().__init__()
        self.window_width, self.window_height = 600, 100
        self.setMinimumSize(self.window_width, self.window_height)

        self.setWindowTitle('Financial Pdf Converter')

        # layout = QVBoxLayout()
        outer_layout = QGridLayout()
        top_layout = QGridLayout()
        bottom_layout = QGridLayout()        

        # file name labels
        self.label1 = QLabel('No file selected')
        self.label2 = QLabel('No file selected')

        # upload buttons
        self.btn1 = QPushButton('Select Balance Sheet (S100)')
        self.btn1.clicked.connect(self.uploadFile1)
        self.btn2 = QPushButton('Select Income Statement (S125)')
        self.btn2.clicked.connect(self.uploadFile2)
        
        # top part layout
        top_layout.addWidget(self.label1, 0, 0, 1, 2)
        top_layout.addWidget(self.btn1, 0, 2)

        top_layout.addWidget(self.label2, 1, 0, 1, 2)
        top_layout.addWidget(self.btn2, 1, 2)


        self.convert_btn = QPushButton('Convert')
        self.convert_btn.clicked.connect(self.start_convert)
        top_layout.addWidget(self.convert_btn, 2, 0, 1, 3)

        self.label3 = QLabel('')
        top_layout.addWidget(self.label3, 3, 0, 1, 3)

        self.list_widget = QListWidget(self)
        dir_path = os.getcwd() + '\\output\\'
        files_list = []
        for path in os.listdir(dir_path):
            if os.path.isfile(os.path.join(dir_path, path)):
                files_list.append(path)
        self.list_widget.addItems(files_list)
        bottom_layout.addWidget(self.list_widget, 3, 0, 4, 2)

        # create buttons
        open_button = QPushButton('Open')
        open_button.clicked.connect(self.open_file)
        remove_button = QPushButton('Remove')
        remove_button.clicked.connect(self.remove)

        # clear_button = QPushButton('Clear')
        # clear_button.clicked.connect(self.clear)
        folder_button = QPushButton('Open Folder')
        folder_button.clicked.connect(self.open_folder)

        bottom_layout.addWidget(open_button, 3, 2)
        bottom_layout.addWidget(remove_button, 4, 2)
        # bottom_layout.addWidget(clear_button, 5, 2)
        bottom_layout.addWidget(folder_button, 6, 2)

        top_layout.setVerticalSpacing(10)
        bottom_layout.setVerticalSpacing(10)
        bottom_layout.setHorizontalSpacing(10)
        outer_layout.setVerticalSpacing(40)
        # outer_layout.setHorizontalSpacing(30)
        outer_layout.addLayout(top_layout, 0, 0)
        outer_layout.addLayout(bottom_layout, 1, 0)
        self.setLayout(outer_layout)

    def uploadFile1(self):
        # get a single file path from the user
        response = QFileDialog.getOpenFileName(
            parent=self,
            caption='Select a file',
            directory=os.getcwd(),
            filter='PDF File (*.pdf);;'
        )
        # unpack the tuple
        filename, selected_filter = response
        # check if user canceled
        if filename:
            dest_path = os.getcwd() + '\\temp\\'
            shutil.copy(filename, dest_path) # uncomment to move file
            # update the label with file path
            filename_no_path = os.path.basename(filename)
            self.label1.setText(filename_no_path)
            self.balance_sheet = filename_no_path
        # print(str(response))

    def uploadFile2(self):
        # get a single file path from the user
        response = QFileDialog.getOpenFileName(
            parent=self,
            caption='Select a file',
            directory=os.getcwd(),
            filter='PDF File (*.pdf);;'
        )
        # unpack the tuple
        filename, selected_filter = response
        # check if user canceled
        if filename:
            dest_path = os.getcwd() + '\\temp\\'
            shutil.copy(filename, dest_path) # uncomment to move file
            # update the label with file path
            filename_no_path = os.path.basename(filename)
            # update the label with file path
            self.label2.setText(filename_no_path)
            self.income_statement_sheet = filename_no_path
        # print('upload file 2')


    def start_convert(self):
        # print(self.balance_sheet)
        # self.convert_btn.setText('Converting, Please wait...')
        self.label3.setText('Converting, Please wait...')
        self.label1.setText('No file selected')
        self.label2.setText('No file selected')
        
        # balance_sheet = 'temp/' + self.balance_sheet
        if self.balance_sheet != '' and self.income_statement_sheet != '':
            output_file = convert(self.balance_sheet, self.income_statement_sheet)
            # print(output_file)
            if output_file:
                self.label3.setText('after')
                # insert item at the top of the list
                self.list_widget.insertItem(0, output_file + '.xlsx')
                # delete old files from temp
                folder = 'temp/'
                for filename in os.listdir(folder):
                    file_path = os.path.join(folder, filename)
                    try:
                        if os.path.isfile(file_path) or os.path.islink(file_path):
                            os.unlink(file_path)
                        elif os.path.isdir(file_path):
                            shutil.rmtree(file_path)
                    except Exception as e:
                        print('Failed to delete %s. Reason: %s' % (file_path, e))
                # self.convert_btn.setText('Convert') # change button text back
        else:
            # self.convert_btn.setText('Convert')
            dlg = QMessageBox(self)
            dlg.setWindowTitle("Error")
            dlg.setText("Please select Balance Sheet (S100) and Income Statement (S125)")
            dlg.exec()

    def open_file(self):
        current_row = self.list_widget.currentRow()
        item = self.list_widget.item(current_row)
        # print(item.text())
        os.startfile( os.getcwd() + '\\output\\' + item.text())


    def open_folder(self):
        path = os.getcwd() + '\\output\\'
        path = os.path.realpath(path)
        os.startfile(path)
        # print(path)

    def remove(self):
        current_row = self.list_widget.currentRow()
        if current_row >= 0:
            current_item = self.list_widget.takeItem(current_row)
            # print(current_item.text())
            # delete file from output
            os.unlink('output/'+ current_item.text())
            del current_item
