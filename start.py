from tkinter import Tk
from tkinter.filedialog import askopenfilename
from PIL import Image
import os
import winshell


def generate_manifest(name, show_name, filename, textcolor='light', backgroundcolor='#FFFFFF'):
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
               "    BackgroundColor='{3}'/>\n</Application>".format(show_name, name, textcolor, backgroundcolor))
    file.close()


def generate_ico(image, name):
    image.save('{0}.ico'.format(name), 'ICO')


def generate_shortcut(name):
    winshell.CreateShortcut(os.path.join(os.path.abspath(os.curdir), '{0}.lnk'.format(name)), Target='{0}.bat'.format(name), Icon=(os.path.join(os.path.abspath(os.curdir), '{0}.ico'.format(name)), 0))
    winshell.move_file('{0}.lnk'.format(name), os.path.join(winshell.start_menu(), '{0}.lnk'.format(name)))


def generate_launcher(name, url):
    file = open(name + '.bat', 'w')
    file.write('start {0}'.format(url))
    file.close()


def ask_user():
    while True:
        name = input('What is the name of the service/program you are trying to add?')
        if name != '':
            break
    Tk().withdraw()
    path = askopenfilename()
    # path = 'D:/Tiles/Tiles/google contacts/Google Contacts.png'
    print(path)
    while True:
        web = input('Is this a webpage?')
        if web == 'Y' or web == 'N':
            break
    if web == 'Y':
        while True:
            website = input('What is the full URL?')
            if website != '':
                break
        generate_launcher(name, website)
        generate_manifest(name, True, path, 'light', '#FFFFFF')
        generate_shortcut(name)


ask_user()
