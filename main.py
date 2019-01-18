import sys
from pathlib import Path
from enum import Enum
import subprocess

from PyPDF2 import PdfFileMerger
from PyPDF2.pagerange import PageRange

from PyQt5.QtCore import Qt, QFile, QTextStream
from PyQt5.QtGui import QIcon
from PyQt5.QtWidgets import (QApplication, QMainWindow, QTextEdit, 
                            QSizePolicy, QScrollArea, QHBoxLayout, 
                            QPushButton, QVBoxLayout, QWidget, 
                            QLabel, QLineEdit, QDesktopWidget,
                            QFrame, QToolBar, QMessageBox, QDialog,
                            QComboBox)

from qttools import layout_items

PAGE_RANGE_HELP = '''Remember, page indices start with zero.

        Page range expression examples:
            :     all pages.                   -1    last page.
            22    just the 23rd page.          :-1   all but the last page.
            0:3   the first three pages.       -2    second-to-last page.
            :3    the first three pages.       -2:   last two pages.
            5:    from the sixth page onward.  -3:-1 third & second to last.

        The third, "stride" or "step" number is also recognized.
            ::2       0 2 4 ... to the end.    3:0:-1    3 2 1 but not 0.
            1:10:2    1 3 5 7 9                2::-1     2 1 0.
            ::-1      all pages in reverse order.'''


class UserInterface:
    pass


class ApplicationWindow(QMainWindow):
    def __init__(self):
        super().__init__()

        self.ui = UserInterface()
        #self.ui.setupUi(self)
        #self.ui.action_execute.triggered.connect(self.execute)
        self.setup()
        self.set_stylesheet(r'stylesheets\noah.qss')


    def set_stylesheet(self, path):
        '''
        sets the stylesheet for the application
        '''
        app = QApplication.instance()
        if app is None:
            raise RuntimeError('No Qt Application found.')

        file = QFile(path)
        file.open(QFile.ReadOnly | QFile.Text)
        stream = QTextStream(file)
        app.setStyleSheet(stream.readAll())


    def help_message2(self):
        '''
        displays the help message
        '''
        mb = QMessageBox()
        mb.setIcon(QMessageBox.Information)
        mb.setWindowTitle('Page Range Information')
        mb.setText(PAGE_RANGE_HELP)
        mb.setDetailedText(
            'Visit the documentation for more details:\n\n'
            'https://pythonhosted.org/PyPDF2/Easy%20Concatenation%20Script.html')
        mb.setStandardButtons(QMessageBox.Ok)
        mb.exec()


    def help_message(self):
        '''
        displays just the Qtextedit
        '''
        mb = CustomMessageBox(self)
        mb.setWindowTitle('Page Range Information')
        mb.setWindowIcon(QIcon(r'icons\help.png'))
        mb.setText(PAGE_RANGE_HELP)
        mb.show() # since we gave mb a parent, we can show it. Otherwise, mb.exec()


    def resetcolor(self):
        '''
        resets the color for all the pdf widgets
        '''
        for pdf in self.pdfs():
            pdf.resetcolor()


    def rotation_required(self):
        '''
        checks if rotation of any pages is required
        '''
        reqd = True
        for pdf in self.pdfs():
            if not doc.rotation == Rotations.NONE:
                reqd = False
        return reqd


    def pdfs(self):
        '''
        generator for the pdf documents in the applicaiton
        '''
        for widget in layout_items(self.ui.docframe.layout()):
            if type(widget) is PdfDocument:
                yield(widget)


    def execute(self):
        '''
        runs the stuff
        '''
        merger = PdfFileMerger()

        for pdf in self.pdfs():
            valid = True

            try:
                fp = str(pdf.path)
            except FileNotFoundError as e:
                valid = False
                print(e.args)

            try:
                pg = pdf.pages                
            except ValueError as e:
                valid = False
                print(e.args)

            if valid:
                merger.append(fileobj=fp, pages=pg, rotation=pdf.rotation)


        try:
            with open('output.pdf', 'wb') as output:
                merger.write(output)
                if len(merger.pages) > 0:
                    subprocess.Popen('output.pdf', shell=True)

        except PermissionError:
            print('close the file you bafoon')


    def insert_pdf_view(self):
        '''
        adds a pdf widget to the pdf document layout
        '''
        doc = PdfDocument()
        #doc.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        self.ui.docframe.layout().insertWidget(self.ui.docframe.layout().count()-2, doc)


    def setup(self):
        '''
        sets up the layout stuff
        '''

        # sets the window properties
        self.setWindowTitle('Noah\'s PDF Tools')
        self.setWindowIcon(QIcon(r'icons\pdf.png'))

        # size to 70% the space available
        self.resize(QDesktopWidget().availableGeometry(self).size() * 0.7);

        # add toolbar
        tb = QToolBar()

        # execute action
        action = tb.addAction(QIcon(r'icons\execute.png'), 'Execute')
        action.triggered.connect(self.execute)
        self.addToolBar(tb)

        # execute action
        action = tb.addAction(QIcon(r'icons\recolor.png'), 'Reset Color')
        action.triggered.connect(self.resetcolor)
        self.addToolBar(tb)

        # help action
        action = tb.addAction(QIcon(r'icons\help.png'), 'Help')
        action.triggered.connect(self.help_message)
        self.addToolBar(tb)     

        # make the central widget
        self.ui.mainwidget = QWidget()
        self.ui.mainlayout = QVBoxLayout(self.ui.mainwidget)
        self.setCentralWidget(self.ui.mainwidget)

        # make the scroll area
        self.ui.docscroll = QScrollArea()
        self.ui.docscroll.setWidgetResizable(True)
        self.ui.mainlayout.addWidget(self.ui.docscroll)

        self.ui.docframe = QFrame(self.ui.docscroll)
        self.ui.docframe.setLayout(QVBoxLayout())

        self.ui.docscroll.setWidget(self.ui.docframe)

        # make the first document widget
        self.ui.docframe.layout().addWidget(PdfDocument())
        
        # add the 'add document' button
        self.ui.add_doc_btn = QPushButton()
        self.ui.add_doc_btn.setText('&Add document')
        self.ui.add_doc_btn.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Minimum)
        self.ui.add_doc_btn.clicked.connect(self.insert_pdf_view)
        self.ui.docframe.layout().addWidget(self.ui.add_doc_btn)

        # add the stretch to the end
        self.ui.docframe.layout().addStretch()


class CustomMessageBox(QDialog):

    def __init__(self, parent=None):
        '''
        I was tired of not being able to resize the QMessageBox
        and also I just wanted to display the help text for
        the PageRange class
        '''
        super().__init__(parent)
        if not parent is None: self.setModal(0)
        self.setup()


    def setup(self):
        self.setWindowTitle('Dialog Box')

        # make the central widget
        self.layout = QVBoxLayout(self)

        self.textEdit = QTextEdit()
        self.layout.addWidget(self.textEdit)


    def setText(self, txt):
        self.textEdit.setText(txt)


class Rotations(Enum):
    NONE = 0
    LEFT = 90
    RIGHT = 270
    UPSIDEDOWN = 180
    

class PdfDocument(QWidget):

    def __init__(self):
        super().__init__()
        self.setup_ui()


    def setup_ui(self):
        #self.setMinimumHeight(200)

        self.setAcceptDrops(True)

        self.layout = QHBoxLayout(self)

        # file path box
        self.path_edit = QLineEdit()
        self.path_label = QLabel('&File Path')
        self.path_label.setBuddy(self.path_edit)

        self.layout.addWidget(self.path_label)
        self.layout.addWidget(self.path_edit)
        
        # pages box
        self.pages_edit = QLineEdit()
        self.pages_label = QLabel('&Pages')
        self.pages_label.setBuddy(self.pages_edit)

        self.layout.addWidget(self.pages_label)
        self.layout.addWidget(self.pages_edit)

        # rotate box
        self.rotate_box = QComboBox()
        self.rotate_box.addItem('None', Rotations.NONE)
        self.rotate_box.addItem('Left 90', Rotations.LEFT)
        self.rotate_box.addItem('Right 90', Rotations.RIGHT)
        self.rotate_box.addItem('180', Rotations.UPSIDEDOWN)
        self.rotate_box.activated.connect(self.set_rotation)

        self.rotate_label = QLabel('&Rotation')
        self.rotate_label.setBuddy(self.rotate_box)

        self.layout.addWidget(self.rotate_label)
        self.layout.addWidget(self.rotate_box)

        self.rotation = Rotations.NONE.value

        # close button
        self.close_btn = QPushButton('&Delete')
        self.close_btn.clicked.connect(self.deleteLater)
        
        self.layout.addWidget(self.close_btn)


    def set_rotation(self, index):
        self.rotation = self.rotate_box.itemData(index).value


    @property
    def pages(self):
        try:
            rng = PageRange(self.pages_edit.text())
            self.colorgood(self.pages_edit)
            return rng
        except:
            self.colorbad(self.pages_edit)
            raise ValueError(f'Cannot make valid page range from: {self.pages_edit.text()}')


    @property
    def path(self):
        p = Path(self.path_edit.text())
        if p.exists() and p.suffix.lower() == '.pdf' and not p.name == '':
            self.colorgood(self.path_edit)
            return p
        else:
            self.colorbad(self.path_edit)
            raise FileNotFoundError(f'Path not found {p}. Path should be to PDF.')


    def recolor(self, widget):
        '''
        recolors the widget -- recolors itself if no parameters are given
        '''
        widget.style().unpolish(widget)
        widget.style().polish(widget)


    def resetcolor(self):
        '''
        resets the widget color to it's default
        '''
        for widget in layout_items(self.layout):
            widget.setProperty('valid_field', None)
            self.recolor(widget)


    def colorgood(self, w:QWidget):
        w.setProperty('valid_field', True)
        self.recolor(w)


    def colorbad(self, w:QWidget):
        w.setProperty('valid_field', False)
        self.recolor(w)


    def dragMoveEvent(self, e):
        #e.accept()
        pass


    def dragEnterEvent(self, e):
        if e.mimeData().hasUrls():
            e.acceptProposedAction()


    def dropEvent(self, e):
        if e.mimeData().hasUrls():
            for url in e.mimeData().urls():
                path = url.path()
                if path[0] == '/':
                    path = path[1:]
                    path = Path(path)

                self.path_edit.setText(str(path))
                self.pages_edit.setText(':')


def main():
    app = QApplication(sys.argv)
    application = ApplicationWindow()
    application.show()
    sys.exit(app.exec_())

if __name__ == "__main__":
    main()