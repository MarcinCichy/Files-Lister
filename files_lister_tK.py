# =====================================================================================================================
# Program ma ułatwić życie osobom, które nie znają poleceń DOSa lub tych, które nie chcą 'babrać' sie w linię poleceń.
# Dokładnie chodzi o komendę DIR *.* /w >lista.txt,
# które zapisuje do pliku 'lista.txt' nazwy wszystkich plików z folderu, w którym aktualnie się znajdujemy. 
# W przypadku tworzenia programów do wycięcia detali na laserze ułatwia to wykonanie spisu elementów do wycięcia.
# Spis powstaje z plików zawierających rysunki poszczególnych detali.
# 
# Najczęściej używanymi formatami plików są (poza *.geo): *.dxf, *.dwg, *.pdf.
# Dodatkowo można wybrać dowolny, inny format pliku, poprzez wpisanie jego roszerzenia.
# Listę można wydrukować.
#
# Wersja 1.0 - tworzenie listy pików
# Wersja 1.1 - dodano drukowanie listy
# Wersja 1.2 - dodano automatyczne zapisywanie pliku lista.txt w listowanym katalogu 
# Wersja 1.3 - interfejs GUI wykonany przy pomocy Tkinter
# Wersja 1.3.2 - dodano przycisk "Aktualizuj", służący do zapisania nowej wersji pliku "lista.txt"
# po edycji zawartości okna tekstowego
# =====================================================================================================================
# The program is to make life easier for people who don't know DOS commands or those who don't want to
# mess with the command line.
# Exactly the DIR command *. * / W> list.txt,
# which saves to the file 'list.txt' the names of all files from the current folder.
# In the case of creating programs for cutting details on the laser, it facilitates the preparation of a list of
# elements to be cut.
# The list is created from files containing drawings of individual details.
#
# The most commonly used file formats are (except *.geo): * .dxf, * .dwg, * .pdf.
# Additionally, you can choose any other file format by entering its extension.
# The list can be printed.
#
# Version 1.0 - creating a list of peaks
# Version 1.1 - added list printing
# Version 1.2 - added automatic saving of the list.txt file in the directory listed
# Version 1.3 - used Tkinter to create GUI
# Version 1.3.2 - the "Aktualizuj" button has been added, which is used to save a new version of the "lista.txt" file
# after editing the contents of the text window
# ==========================================================================================================================

from tkinter import filedialog
from tkinter import *
import sys
import os
import fnmatch
from pathlib import Path
# import win32ui
import win32api
# import win32print


class Application(Frame):
    def __init__(self, master):
        super(Application, self).__init__(master)
        self.grid()
        self.create_widgets()

    # utwórz wszystkie potrzebne widgety GUI
    def create_widgets(self): 

        frm_First_Frame = Frame(self)  # , bg = 'blue'
        frm_First_Frame.grid(row = 0, column = 0,  ipadx = 5, ipady = 5, sticky = NW)
        frm_Second_Frame = Frame(self)  # , bg = 'red'
        frm_Second_Frame.grid(row = 0, column = 3, ipadx = 5, ipady = 5, sticky = NW)
        frm_Third_Frame = Frame(self)  # , bg = 'green'
        frm_Third_Frame.grid(row = 5, column =3, ipadx = 5, ipady = 5, sticky = NW)

        # utwórz przycisk do otwierania i wyboru katalgów
        self.btn_OpenFolder = Button(frm_First_Frame)
        self.btn_OpenFolder['text'] = 'Otwórz katalog'
        self.btn_OpenFolder['command'] = self.choose_folder
        self.btn_OpenFolder.grid(row = 0, column = 0, padx = 5, pady = 5, sticky = NW)

        # utwórz ramkę, która okala przyciski wyboru i je grupuje
        labelframe = LabelFrame (frm_First_Frame, text = 'Rozszerzenie:')
        labelframe.grid(row = 1, column = 0, padx = 5, pady = 2, sticky = NW)

        # utwórz zmienną, która ma reprezentowac pojedyncze wybrane rozszerzenie pliku
        self.extension = StringVar()
        self.extension.set(None)
        
        # utwórz przycisk wyboru (radio button)
        Radiobutton (labelframe, text = 'DXF', variable = self.extension, value = 'dxf', command = self.set_filetype).grid(row = 0, column = 0, padx = 2, pady = 2, sticky = W)
        Radiobutton (labelframe, text = 'DWG', variable = self.extension, value = 'dwg', command = self.set_filetype).grid(row = 1, column = 0, padx = 2, pady = 2, sticky = W)
        Radiobutton (labelframe, text = 'PDF', variable = self.extension, value = 'pdf', command = self.set_filetype).grid(row = 2, column = 0, padx = 2, pady = 2, sticky = W)
        Radiobutton (labelframe, text = 'INNE', variable = self.extension, value = 'inne', command = self.set_filetype).grid(row = 3, column = 0, padx = 2, pady = 2, sticky = W)
        self.entry_ext = Entry(labelframe, width = 5)                       # utworzenie pola tekstowego, do wprowadzania dowolnego rozszerzenia pliku
        self.entry_ext.grid(row = 3, column = 2,padx = 5, sticky = W)

        # utwórz etykietę do wypisywania nazwy wybranego folderu
        self.lbl_FolderName = Label(frm_First_Frame)
        self.lbl_FolderName['text'] = 'Nazwa folderu:'
        self.lbl_FolderName.grid(row = 2, column = 0, padx = 5, pady = 2, sticky = W)
        self.lbl_FolderPathValue = Label(frm_First_Frame)
        self.lbl_FolderPathValue.grid(row = 3, column = 0, padx = 5, pady = 2, sticky = W)
        
        # utwórz etykietę opisującą pole tekstowe
        self.lbl_ListOfFiles = Label(frm_Second_Frame)
        self.lbl_ListOfFiles['text'] = 'Lista plików:'
        self.lbl_ListOfFiles.grid(row = 0, column = 0, padx = 2, pady = 8, sticky = NW)

        # utwórz pole tekstowe do wypisywania listy plków
        self.txt_ListOfFiles = Text(frm_Second_Frame, width = 80, height = 25) # ,width = 100, height = 100
        self.txt_ListOfFiles.grid(row = 1, column = 0, padx = 2, pady = 10, sticky = NE)
        scrollbar = Scrollbar(frm_Second_Frame, orient = 'vertical', command = self.txt_ListOfFiles.yview)
        scrollbar.grid(row =1,  column = 2, sticky = 'NS')
        self.txt_ListOfFiles.config(yscrollcommand=scrollbar.set)
        

        # utwórz przycisk do drukowania listy plików
        self.btn_PrintListOfFiles = Button(frm_Third_Frame)
        self.btn_PrintListOfFiles['text'] = 'Drukuj'
        self.btn_PrintListOfFiles['command'] = self.print_file
        self.btn_PrintListOfFiles.grid(row = 0, column = 2, padx = 1, pady = 2, sticky = W)

        # utwórz przycisk do aktualizowania pliku 'lista.txt' po zmianie zawartości okna tekstowego
        self.btn_PrintListOfFiles = Button(frm_Third_Frame)
        self.btn_PrintListOfFiles['text'] = 'Aktualizuj'
        self.btn_PrintListOfFiles['command'] = self.save_file_txt
        self.btn_PrintListOfFiles.grid(row = 0, column = 0, padx = 1, pady = 2, sticky = W)

    # metoda do wyboru katalogu do wyświetlenia plików i wyświetleni eplików w oknie
    def choose_folder(self):
        global folderpath
        folderpath = filedialog.askdirectory()
        self.lbl_FolderPathValue['text'] = os.path.basename(folderpath)
        
        #wypisanie katalogu
        self.txt_ListOfFiles.delete(0.0, END)
        self.txt_ListOfFiles.insert(END,"*"*60 + '\n')  
        self.txt_ListOfFiles.insert(END,(folderpath.replace('/','\\')) + '\n')   # wyświetl ścieżkę dostępu i jednocześnie zamienia w niej znak '/' na '\' zgodnie z systemem Winodws
        self.txt_ListOfFiles.insert(END,"*"*60 + '\n')  
        self.txt_ListOfFiles.insert(END,"" '\n')
        
        extension = self.set_filetype()
        count = 1
        for file_name in os.listdir(folderpath):                                        # pętla w celu odczytania wszystkich plików w folderze      
            count_str = str(count) 
            if fnmatch.fnmatch(file_name, '*.' + extension):                            # ale warunek powoduje, że będą to tylko pliki o wskazanym rozszerzeniu -> file_type
                self.txt_ListOfFiles.insert(END, count_str + '. ' + file_name + '\n')   # dodaj do pola textEdit kolejno odczytane z folderu nazwy pików, a przed nimi liczbę porządkową 'count_str'
                self.txt_ListOfFiles.insert(END,"-"*60+ '\n')                           # oraz linie rozdzielające pomiędzy nazwami
                count += 1                                                              # dzięki temu, że jest umiejscowione w tytm miejscu zlicza tylo pliki spełniejące warunek
                                                                                        # inaczej zlicza wszytskie pliki w katalogu i wyświetla liczbe porządkową pliku spełniającego warunek
        self.save_file_txt()                                                            # wywołanie metody do automatycznego zapisywania pliku txt z listą plików

    # metoda do ustawiania typu plków do wyświetlenia
    def set_filetype(self):
        extension = self.extension.get()
        if extension == 'inne':
            extension = self.entry_ext.get() 
        else:
            self.entry_ext.delete(0, 'end') 
            extension = self.extension.get()
        return extension

    # automatyczne zapisywanie pliku lista.txt w sprawdanym katalogu
    def save_file_txt(self):
        fp = folderpath.replace('/','\\')
        file_name = (fp+"\lista.txt")
        file = open(file_name,'w')
        text = self.txt_ListOfFiles.get(1.0, 'end')
        file.write(text)
        file.close()

    # metoda do drukowania pliku z listą plików
    def print_file(self):     
        fp = folderpath.replace('/','\\')
        file_name = (fp+"\lista.txt")
        win32api.ShellExecute(0, "print", file_name, None, '*.*', 0)
    # def print_file(self):
    #     printer_name = win32print.GetDefaultPrinter()
    #     win32print.OpenPrinter(printer_name)
    #     print(printer_name)
    #     # folderpath = fp.replace('/','\\')
    #     # file_name = (folderpath+"\lista.txt")
    #     # # win32api.ShellExecute(0, "print", file_name, None, '*.*', 0)
    #     # win32api.ShellExecute(0 , "open" , "rundll32.exe" , file_name , None , 1)
    #     #help(win32api)

root = Tk()
root.title('Files Lister v.1.3.2')
root.geometry('900x550')
path = os.path.dirname(__file__)
# root.iconbitmap(path+"\QF_bold.ico")
app = Application(root)
root.mainloop()