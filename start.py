from PyQt5.QtWidgets import QApplication, QMainWindow, QFileDialog
from layout import Ui_MainWindow
from PIL import Image
import os
import winshell
import sys


class App(Ui_MainWindow):
    def __init__(self, main_window):
        Ui_MainWindow.__init__(self)
        self.setupUi(main_window)

        self.button_submit.clicked.connect(self.on_submit)
        self.button_search_tile.clicked.connect(self.on_file)

    def on_submit(self):
        os.mkdir(self.text_name.text())
        os.chdir(self.text_name.text())
        generate_launcher(self.text_name.text(), self.text_url.text())
        generate_shortcut(self.text_name.text())
        generate_manifest(self.text_name.text(), self.text_name.text(), self.text_path.text())

    def on_file(self):
        file_dialog = QFileDialog()
        options = file_dialog.Options()
        options |= file_dialog.DontUseNativeDialog
        fname = file_dialog.getOpenFileName(file_dialog, 'Open file',
                                            'c:\\', "Image files (*.jpg *.gif)")
        if fname:
            self.text_path.setText(fname[0])


def generate_manifest(name, show_name, filename, text_color='light', background_color='#FFFFFF'):
    sizes = [(150, 150), (70, 70), (44, 44)]
    file = open('{0}.VisualElementsManifest.xml'.format(name), 'w')
    if show_name:
        show_name = 'on'
    else:
        show_name = 'off'
    image = Image.open(filename)
    generate_ico(image, name)
    for size in sizes:
        small_image = image.resize(size, Image.BILINEAR)
        small_image.save('{0}.{1}.png'.format(name, size[0]), 'PNG')
    file.write("<Application xmlns:xsi='http://www.w3.org/2001/XMLSchema-instance'>\n    <VisualElements\n        "
               "ShowNameOnSquare150x150Logo='{0}'\n        Square150x150Logo='{1}.150.png'\n        "
               "Square70x70Logo='{1}.70.png'\n        Square44x44Logo='{1}.44.png'\n        ForegroundText='{2}'\n    "
               "    BackgroundColor='{3}'/>\n</Application>".format(show_name, name, text_color, background_color))
    file.close()


def generate_ico(image, name):
    image.save('{0}.ico'.format(name), 'ICO')


def generate_shortcut(name, path=None):
    if path is None:
        path = '{0}.bat'.format(name)
    winshell.CreateShortcut(Path=os.path.join(os.path.abspath(os.curdir), '{0}.lnk'.format(name)),
                            Target=path,
                            Icon=(os.path.join(os.path.abspath(os.curdir), '{0}.ico'.format(name)), 0))
    winshell.move_file(source_path='{0}.lnk'.format(name),
                       target_path=os.path.join(winshell.programs(), '{0}.lnk'.format(name)))


def generate_launcher(name, url):
    file = open(name + '.bat', 'w')
    file.write('start {0}'.format(url))
    file.close()


if __name__ == '__main__':
    app = QApplication(sys.argv)
    main_window = QMainWindow()
    ui = App(main_window)
    main_window.show()
    sys.exit(app.exec_())
