import zipfile
import pyperclip
import sys



try:
    # for Python2
    from xml.etree.cElementTree import XML
    from Tkinter import *   ## notice capitalized T in Tkinter
    from Tkinter import ttk
    from tkFileDialog import askopenfilename

except ImportError:
    # for Python3
    from xml.etree.ElementTree import XML
    from tkinter import *   ## notice lowercase 't' in tkinter here
    from tkinter import ttk
    from tkinter.filedialog import askopenfilename



"""
Module that extract text from MS XML Word document (.docx), convert it to unicode and modify it.
Among modifications :
- add a header and a footer
- remove some word at the beginning of some lines starting with "N°" and replace with emoji's numbers
- add some Whatsapp's formatting syntax (

(Inspired by <https://gist.github.com/etienned/7539105> which only extract paragraph from word file)
"""

WORD_NAMESPACE = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
PARA = WORD_NAMESPACE + 'p'
TEXT = WORD_NAMESPACE + 't'
errmsg = 'Error!'

def quit():
    global root
    root.quit()
    sys.exit()

def OpenFile():
    """
    GUI + copy result in clipboard
    """
    name = askopenfilename(initialdir="D:/yassi/Downloads",
                           filetypes =[("Word File", "*.docx")],
                           title = "Choose a file."
                           )
    if name :
        pyperclip.copy(get_docx_text(name))
        textlabel.set("Text is in clipboard")
        # print(get_docx_text(name))

def get_docx_text(path):
    """
    Take the path of a docx file as argument, modify it.
    """
    document = zipfile.ZipFile(path)
    xml_content = document.read('word/document.xml')
    document.close()
    tree = XML(xml_content)
    #Header
    paragraphs = ['﷽','📖 *Al-Adaab Al-Mufrad* – _La véritable éducation (Imam Bukhaari RAH)_']

    for paragraph in tree.getiterator(PARA):
        # EXTRACTION
        texts = [node.text
                 for node in paragraph.getiterator(TEXT)
                 if node.text]

        # MODIFICATION
        # Basic titles of my text
        if ''.join(texts).startswith('Chapitre') :
            texts = ["📄*" + ''.join(texts) + "*"]
        elif ''.join(texts).startswith('Traduction') :
            texts = ["_📌" + ''.join(texts) + "_"]
        elif ''.join(texts).startswith('Commentaires') :
            texts = ["_📌" + ''.join(texts) + "_"]

        # Special lines that have "Hadith N°" inside the text
        elif "Hadith N°" in ''.join(texts):
            texts.pop(0)
            # In case we have ['381 - ', 'Hadith N°1\xa0: ', 'D’après Abû ',....
            if "1" in ''.join(texts[0]) :
                texts.pop(0)
                texts = ["1️⃣" + ''.join(texts)]
            elif "2" in ''.join(texts[0]) :
                texts.pop(0)
                texts = ["2️⃣" + ''.join(texts)]
            elif "3" in ''.join(texts[0]) :
                texts.pop(0)
                texts = ["3️⃣" + ''.join(texts)]
            elif "4" in ''.join(texts[0]) :
                texts.pop(0)
                texts = ["4️⃣" + ''.join(texts)]
            elif "5" in ''.join(texts[0]) :
                texts.pop(0)
                texts = ["5️⃣" + ''.join(texts)]
            elif "6" in ''.join(texts[0]) :
                texts.pop(0)
                texts = ["6️⃣" + ''.join(texts)]
            else :
                # In case we have ['382 - ', 'Hadith N°', '2', '\xa0: ',....
                texts.pop(0)
                if "1" in ''.join(texts[0]):
                    texts.pop(0)
                    texts = ["1️⃣" + ''.join(texts)]
                elif "2" in ''.join(texts[0]):
                    texts.pop(0)
                    texts = ["2️⃣" + ''.join(texts)]
                elif "3" in ''.join(texts[0]):
                    texts.pop(0)
                    texts = ["3️⃣" + ''.join(texts)]
                elif "4" in ''.join(texts[0]):
                    texts.pop(0)
                    texts = ["4️⃣" + ''.join(texts)]
                elif "5" in ''.join(texts[0]):
                    texts.pop(0)
                    texts = ["5️⃣" + ''.join(texts)]
                elif "6" in ''.join(texts[0]):
                    texts.pop(0)
                    texts = ["6️⃣" + ''.join(texts)]
            print("resultat :")
            print(texts)
            if ":" in ''.join(texts)[:6] :
                texts = [''.join(texts).replace(':','',1)]
            if " " in ''.join(texts)[:6] :
                texts = [''.join(texts).replace(' ','',1)]
            print(texts)

        # in all other cases (ih not title, if does not have "Hadith N°)
        else :
            texts = ["▶" + ''.join(texts)]

        # After IF
        if texts:
            # Whatsapp formating bold for citation
            texts = [''.join(texts).replace('«','*«').replace('»','»*')]

            # Deleting non-breaking spaces & cie
            texts = [''.join(texts).replace('\xa0', '')]
            texts = [''.join(texts).replace('\u2009', '')]

            # Deleting useless spaces & cie
            texts = [''.join(texts).replace(' ﷺ', 'ﷺ')]
            texts = [''.join(texts).replace(' ؓ', 'ؓ')]
            #texts = [''.join(texts).replace(': *«', ':*«')]

            paragraphs.append(''.join(texts))
        else:
            paragraphs.append('')

    # Footer
    paragraphs.append("")
    paragraphs.append("_Qu’Allah SWT nous accorde le véritable amour pour le Prophèteﷺ et qu'il nous aide à mettre en pratique ses précieux conseils._ *Amine*")
    paragraphs.append("")
    paragraphs.append("_*Mufti Yassine Gangat*_")

    return '\n\n'.join(paragraphs)

root = Tk(  )
Title = root.title( "DocxToWhatsapp Converter")
textlabel = StringVar()
textlabel.set(".docx => Whatsapp Converter")
label = ttk.Label(root, textvariable =textlabel,foreground="red",font=("Helvetica", 16))
label.pack()
Button(text='File Open', command=OpenFile).pack(fill=X)
Button(text='Exit', command=quit).pack(fill=X)

#Menu Bar
menu = Menu(root)
root.config(menu=menu)
file = Menu(menu)

file.add_command(label = 'Open', command = OpenFile)
file.add_command(label = 'Exit', command = quit)
menu.add_cascade(label = 'File', menu = file)


root.mainloop()

