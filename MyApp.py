import os
from PyQt6.QtWidgets import (QWidget, QPushButton, QFileDialog,
                                QLabel, QGridLayout, QVBoxLayout, QHBoxLayout, QListWidget)
import shutil
import convert
class MyApp(QWidget):
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
        self.btn1 = QPushButton('Upload Balance Sheet (S100)')
        self.btn1.clicked.connect(self.uploadFile1)
        self.btn2 = QPushButton('Upload Income Statement (S125)')
        self.btn2.clicked.connect(self.uploadFile2)
        
        # top part layout
        top_layout.addWidget(self.label1, 0, 0, 1, 2)
        top_layout.addWidget(self.btn1, 0, 2)

        top_layout.addWidget(self.label2, 1, 0, 1, 2)
        top_layout.addWidget(self.btn2, 1, 2)


        convert_btn = QPushButton('Convert')
        convert_btn.clicked.connect(self.convert)
        top_layout.addWidget(convert_btn, 2, 0, 1, 3)

        self.list_widget = QListWidget(self)
        dir_path = os.getcwd() + '\\output\\'
        files_list = []
        for path in os.listdir(dir_path):
            if os.path.isfile(os.path.join(dir_path, path)):
                files_list.append(path)
        self.list_widget.addItems(files_list)
        bottom_layout.addWidget(self.list_widget, 3, 0, 4, 2)

        # create buttons
        add_button = QPushButton('Open')
        # add_button.clicked.connect(self.add)
        remove_button = QPushButton('Remove')
        remove_button.clicked.connect(self.remove)

        clear_button = QPushButton('Clear')
        # clear_button.clicked.connect(self.clear)
        folder_button = QPushButton('Open Folder')
        folder_button.clicked.connect(self.open_folder)

        bottom_layout.addWidget(add_button, 3, 2)
        bottom_layout.addWidget(remove_button, 4, 2)
        bottom_layout.addWidget(clear_button, 5, 2)
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
             # get the current working directory
            dest_path = os.getcwd() + '\\temp\\'
            # move the file to the current working directory
            shutil.move(filename, dest_path) # uncomment to move file
            # update the label with file path
            filename_no_path = os.path.basename(filename)
            print(filename_no_path)
            # self.list_widget.addItem(filename_no_path)
            self.label1.setText(filename_no_path)
        print(str(response))

    def uploadFile2(self):
        # get a single file path from the user
        response = QFileDialog.getOpenFileName(
            parent=self,
            caption='Select a file',
            directory=os.getcwd(),
            filter='All Files (*.*)'
        )
        # unpack the tuple
        filename, selected_filter = response
        # check if user canceled
        if filename:
            # update the label with file path
            self.label2.setText(filename)
        print('upload file 2')


    def convert(self):

        # insert item at the top of the list
        self.list_widget.insertItem(0, 'test')
        # self.list_widget.addItem('test')
        print('convert')

    def open_folder(self):
        path = os.getcwd() + '\\output\\'
        path = os.path.realpath(path)
        os.startfile(path)

    def remove(self):
        current_row = self.list_widget.currentRow()
        if current_row >= 0:
            current_item = self.list_widget.takeItem(current_row)
            print(current_item.text())
            del current_item
