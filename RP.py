import os
import sys
from PySide6.QtWidgets import QApplication, QMainWindow, QPushButton, QLineEdit, QLabel, QGridLayout, QWidget, QFileDialog, QMessageBox
from RPMainWindow import Ui_MainWindow

class RPMainWindow(QMainWindow, Ui_MainWindow):
    def __init__(self, parent=None):
        super(RPMainWindow, self).__init__(parent)
        self.setupUi(self)
        self.action_openfolder.triggered.connect(self.open_folder)

        self.action_quitapp.triggered.connect(self.close)

    def open_folder(self):
        folder_path = QFileDialog.getExistingDirectory(self, "Open Folder", os.path.expanduser("~"))
        self.lineEdit_folderpath.setText(folder_path)

    def closeEvent(self, event):
        reply = QMessageBox.question(self, 'Message',
            "Are you sure to quit?", QMessageBox.Yes | 
            QMessageBox.No, QMessageBox.No)

        if reply == QMessageBox.Yes:
            event.accept()
        else:
            event.ignore()

if __name__ == "__main__":
    app = QApplication(sys.argv)
    MainWindow = RPMainWindow()
    MainWindow.show()
    sys.exit(app.exec())