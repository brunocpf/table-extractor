#!c:\Python27\python.exe
# -*- coding: utf-8 -*-

from Tkinter import *
import tkFileDialog as fdlg
from ttk import Frame, Button, Style
import ttk
import tkFont
import json
import sys
import io
import csv
import extractor
import threading
import tkMessageBox as mbox


DEFAULT_SETTINGS = "defaultSettings.json"


def byteify(input):
    if isinstance(input, dict):
        return {byteify(key): byteify(value)
                for key, value in input.iteritems()}
    elif isinstance(input, list):
        return [byteify(element) for element in input]
    elif isinstance(input, unicode):
        return input.encode('utf-8')
    else:
        return input


class ExtractorSettings:
    @staticmethod
    def decodeInt(text):
        try:
            return int(float(text))
        except ValueError:
            return 0

    @staticmethod
    def decodeFloat(text):
        try:
            return float(text)
        except ValueError:
            return 0.0

    @staticmethod
    def decodeHeights(text):
        try:
            exec("d = {" + text.replace("Demais", "-1") + "}")
            if -1 not in d:
                if '-1' in d:
                    d[-1] = d.pop('-1')
                d[-1] = 0
            return d
        except:
            return {-1: 0}

    @staticmethod
    def decodeMultilineColumns(text):
        try:
            exec("l = [" + text + "]")
            return l
        except:
            return []

    @staticmethod
    def encodeHeights(d):
        if '-1' in d:
            d[-1] = d.pop('-1')
        if -1 in d:
            d['Demais'] = d.pop(-1)
        return str(d).replace("{", "").replace("}","").replace("'", "")
    
    @staticmethod
    def encodeMultilineColumns(d):
        return str(d).replace("[","").replace("]","")




def popExtractionCompleted(event=None):
    mbox.showinfo("", "Extração completa.")


class Window(Frame):
  
    def __init__(self, parent, w, h):
        Frame.__init__(self, parent)
        self.extracting = False
        self.parent = parent
        self.bind('<<Info>>', popExtractionCompleted)
        #self.loadSettings()
        self.loadSettings(initial = True)
        self.initUI(w, h)

    def initUI(self, w, h):
        self.parent.title("Extrator")
        self.centerWindow(w, h)
        self.pack(fill=BOTH, expand=1)
        self.createFrames()

    def centerWindow(self, w, h):
        sw = self.parent.winfo_screenwidth()
        sh = self.parent.winfo_screenheight()
        x = (sw - w)/2
        y = (sh - h)/2
        self.parent.geometry('%dx%d+%d+%d' % (w, h, x, y))
    
    def createFrames(self):
        self.filesFrame = LabelFrame(self, text='Arquivos')
        self.filesFrame.grid(sticky=E+W, padx = 10)

        pdfFileLabel = ttk.Label(self.filesFrame, text = 'Selecione um arquivo PDF: ')
        pdfFileLabel.grid(row = 0, column = 0, padx=10, pady=10, sticky = E)
        self.pdfFileEntry = ttk.Entry(self.filesFrame, width = 42)
        self.pdfFileEntry.grid(row = 0, column = 1, columnspan = 9, sticky = W+E)
        pdfFileButton = ttk.Button(self.filesFrame, text='Escolher arquivo', command=lambda i='open', e=self.pdfFileEntry: self._file_dialog(i, e))
        pdfFileButton.grid(row = 0, column = 10, sticky = W+E, padx = 5)

        settingsFileLabel = ttk.Label(self.filesFrame, text = 'Selecione o arquivo de configurações para carregar: ')
        settingsFileLabel.grid(row = 1, column = 0, padx=10, pady=10, sticky = E)
        self.settingsFileEntry = ttk.Entry(self.filesFrame, width = 42)

        self.settingsFileEntry.grid(row = 1, column = 1, columnspan = 9, sticky = W+E)
        settingsFileButton = ttk.Button(self.filesFrame, text='Escolher arquivo', command=lambda: self.loadSettings())
        settingsFileButton.grid(row = 1, column = 10, sticky = W+E, padx = 5)

        self.settingsFrame = LabelFrame(self, text='Configurações', padx=5, pady=5)
        self.settingsFrame.grid(sticky=N+S+E+W, padx = 10)

        self.refresh()
        
        saveSettingsButton = ttk.Button(self, text='Salvar configurações', command=lambda: self.saveSettings())
        saveSettingsButton.grid(sticky = W+E)

        extractButton = ttk.Button(self, text='<  Extrair tabela  >', command=lambda: self.extractTable())
        extractButton.grid(sticky = N+S+W+E)

        self.pblabel = ttk.Label(self, text = 'Status: Esperando')
        self.pblabel.grid(column = 0, sticky = N+S+W+E)
        self.pb = ttk.Progressbar(self, value = 0, orient="horizontal", mode="determinate")
        self.pb.grid(sticky = N+S+W+E)
        
        #self.status = ttk.Label(self, text="Status: Esperando")
        #self.status.grid(sticky=N+S+E+W)


    def refresh(self):
            startPageLabel = Label(self.settingsFrame, text="Página inicial: ")
            startPageLabel.grid(row = 2, column = 0, padx = 10, pady = 5, sticky = E)
            self.startPageEntry = ttk.Entry(self.settingsFrame, width = 4)
            self.startPageEntry.grid(row = 2, column = 1, sticky = W)
            self.startPageEntry.insert(END, self.startPage)

            endPageLabel = Label(self.settingsFrame, text="Página final: ")
            endPageLabel.grid(row = 3, column = 0, padx = 10, sticky = E)
            self.endPageEntry = ttk.Entry(self.settingsFrame, width = 4)
            self.endPageEntry.grid(row = 3, column = 1, sticky = W)
            self.endPageEntry.insert(END, self.endPage)

            yOffsetLabel = Label(self.settingsFrame, text="Offset vertical: ")
            yOffsetLabel.grid(row = 2, column = 4, padx = 10, sticky = E)
            self.yOffsetEntry = ttk.Entry(self.settingsFrame, width = 8)
            self.yOffsetEntry.grid(row = 2, column = 5, sticky = W)
            self.yOffsetEntry.insert(END, self.yOffset)

            heightsLabel = Label(self.settingsFrame, text="Alturas: ")
            heightsLabel.grid(row = 4, column = 0, padx = 10, pady = 5, sticky = E)
            self.heightsEntry = ttk.Entry(self.settingsFrame, width = 48)
            self.heightsEntry.grid(row = 4, column = 1, sticky = W)
            self.heightsEntry.insert(END, ExtractorSettings.encodeHeights(self.heights))

            multilineLabel = Label(self.settingsFrame, text="Colunas de multiplas linhas: ")
            multilineLabel.grid(row = 5, column = 0, padx = 10, pady = 5, sticky = E)
            self.multilineEntry = ttk.Entry(self.settingsFrame, width = 48)
            self.multilineEntry.grid(row = 5, column = 1, sticky = W)
            self.multilineEntry.insert(END, ExtractorSettings.encodeMultilineColumns(self.multilineColumns))

            multilineSeparatorLabel = Label(self.settingsFrame, text="Sufixo de linhas: ")
            multilineSeparatorLabel.grid(row = 6, column = 0, padx = 10, pady = 5, sticky = E)
            self.multilineSeparatorEntry = ttk.Entry(self.settingsFrame, width = 48)
            self.multilineSeparatorEntry.grid(row = 6, column = 1, sticky = W)
            self.multilineSeparatorEntry.insert(END, self.multilineSeparator)


            self.columnsFrame = LabelFrame(self.settingsFrame, text='Colunas', padx=5, pady=5)
            self.columnsFrame.grid(sticky=N+S+E+W, padx = 10, columnspan=6)

            self.columnXEntries = [x for x in range(len(self.columnsDict))]
            self.columnWEntries = [x for x in range(len(self.columnsDict))]
            self.columnNameEntries = [x for x in range(len(self.columnsDict))]
            for i in range(len(self.columnsDict)):
                columnsLabel = Label(self.columnsFrame, text="x: ")
                columnsLabel.grid(row = i, column = 0, padx = 10, pady = 5, sticky = W+E)
                self.columnXEntries[i] = ttk.Entry(self.columnsFrame, width = 24)
                self.columnXEntries[i].grid(row = i, column = 1, sticky = W)
                self.columnXEntries[i].insert(END, self.columnsDict[i]['x'])

                columnsLabel = Label(self.columnsFrame, text="Largura: ")
                columnsLabel.grid(row = i, column = 2, padx = 10, pady = 5, sticky = W+E)
                self.columnWEntries[i] = ttk.Entry(self.columnsFrame, width = 24)
                self.columnWEntries[i].grid(row = i, column = 3, sticky = W)
                self.columnWEntries[i].insert(END, self.columnsDict[i]['W'])

                columnsLabel = Label(self.columnsFrame, text="Nome: ")
                columnsLabel.grid(row = i, column = 4, padx = 10, pady = 5, sticky = W+E)
                self.columnNameEntries[i] = ttk.Entry(self.columnsFrame, width = 24)
                self.columnNameEntries[i].grid(row = i, column = 5, sticky = W)
                self.columnNameEntries[i].insert(END, self.columnsDict[i]['name'])

    def loadSettings(self, path = None, initial = False):
        if initial:
            path = DEFAULT_SETTINGS
        elif path is None:
            path = self._file_dialog('open', self.settingsFileEntry, True)
        if path is None: return
        with io.open(path, encoding='utf-8') as data_file:    
            data = json.load(data_file)
            data = byteify(data)
            self.startPage = data['startPage']
            self.endPage = data['endPage']
            self.heights = data['heights']
            self.yOffset = data['yOffset']
            self.columnsDict = data['columns']
            self.multilineColumns = data['multilineColumns']
            self.multilineSeparator = data['multilineSeparator']
            if not initial:
                for widget in self.settingsFrame.winfo_children():
                    widget.destroy()
                self.refresh()

    def saveSettings(self):
        path = self._file_dialog('save', None, True)
        if path is None: return
        columns = [x for x in range(len(self.columnsDict))]
        for i in range(len(self.columnsDict)):
            columns[i] = {}
            columns[i]['x'] = ExtractorSettings.decodeFloat(self.columnXEntries[i].get())
            columns[i]['W'] = ExtractorSettings.decodeFloat(self.columnWEntries[i].get())
            columns[i]['name'] = self.columnNameEntries[i].get()

        data = {"startPage": ExtractorSettings.decodeInt(self.startPageEntry.get()),
                "endPage": ExtractorSettings.decodeInt(self.endPageEntry.get()),
                "heights": ExtractorSettings.decodeHeights(self.heightsEntry.get()),
                "yOffset": ExtractorSettings.decodeFloat(self.yOffsetEntry.get()),
                "multilineColumns": ExtractorSettings.decodeMultilineColumns(self.multilineEntry.get()),
                "multilineSeparator": self.multilineSeparatorEntry.get(),
                "columns": columns
                }
        with io.open(path, 'w', encoding='utf-8') as data_file:
            data_file.write(unicode(json.dumps(data, ensure_ascii=False)))

    def disableWidgets(self, frame):
        for widget in frame.winfo_children():
            if isinstance(widget, ttk.Entry) or isinstance(widget, ttk.Button):
                widget.configure(state='disabled')

    def enableWidgets(self, frame):
        for widget in frame.winfo_children():
            if isinstance(widget, ttk.Entry) or isinstance(widget, ttk.Button):
                widget.configure(state='enabled')

    def extractingThread(self, data, path, inputPath):
        ext = extractor.SinapiExtractor(inputPath, data['startPage'], data['endPage'],
                                        data['yOffset'], data['heights'], data['columns'],
                                        data['multilineColumns'], data['multilineSeparator'])

        rows = ext.generateRows(self)
        with io.open(path, 'wb') as csv_file:
            writer = csv.writer(csv_file, delimiter=';', quoting=csv.QUOTE_ALL, lineterminator='\n')
            writer.writerow([name['name'] for name in data['columns']])
            for row in rows:
                writer.writerow([x for x in row])
        csv_file.close()
        self.enableWidgets(self)
        self.enableWidgets(self.filesFrame)
        self.enableWidgets(self.settingsFrame)
        self.enableWidgets(self.columnsFrame)
        self.pb['value'] = 0
        self.pblabel['text'] = 'Status: Esperando'
        self.pb.update()
        self.pblabel.update()
        self.extracting = False
        #self.event_generate('<<Info>>')


    def extractTable(self):
        path = self._file_dialog('save', None, isExtracted = True)
        if path is None or '': return
        inputPath = self.pdfFileEntry.get()
        if inputPath is None or '': return
        self.extracting = True
        columns = [x for x in range(len(self.columnsDict))]
        for i in range(len(self.columnsDict)):
            columns[i] = {}
            columns[i]['x'] = ExtractorSettings.decodeFloat(self.columnXEntries[i].get())
            columns[i]['W'] = ExtractorSettings.decodeFloat(self.columnWEntries[i].get())
            columns[i]['name'] = self.columnNameEntries[i].get()
        data = {"startPage": ExtractorSettings.decodeInt(self.startPageEntry.get()),
                "endPage": ExtractorSettings.decodeInt(self.endPageEntry.get()),
                "heights": ExtractorSettings.decodeHeights(self.heightsEntry.get()),
                "yOffset": ExtractorSettings.decodeFloat(self.yOffsetEntry.get()),
                "multilineColumns": ExtractorSettings.decodeMultilineColumns(self.multilineEntry.get()),
                "multilineSeparator": self.multilineSeparatorEntry.get(),
                "columns": columns
                }
        self.pblabel['text']= "Status: Extraindo..."
        self.pblabel.update()
        self.pb.update()
        self.disableWidgets(self)
        self.disableWidgets(self.filesFrame)
        self.disableWidgets(self.settingsFrame)
        self.disableWidgets(self.columnsFrame)
        t = threading.Thread(target = self.extractingThread, args = (data, path, inputPath))
        t.start()

    def refreshProgress(self, value, maximum):
        self.pb.config(value = value, maximum = maximum)
        self.pb.update()

    def _file_dialog(self, type, ent, isSettings = False, isExtracted = False):
        fn = None
        if isSettings:
            if ent:
                initFile = ent.get()
            else:
                initFile = DEFAULT_SETTINGS
            opts = {'initialfile': initFile,
                    'filetypes': (('Arquivos json', '.json'),)}
        elif isExtracted:
            opts = {'initialfile': "tabela.csv",
                    'filetypes': (('Arquivos csv', '.csv'),)}
        else:
            opts = {'initialfile': ent.get(),
                    'filetypes': (('Arquivos pdf', '.pdf'),)}
         
        if type == 'open':
            opts['title'] = 'Selecione um arquivo para abrir...'
            fn = fdlg.askopenfilename(**opts)
        else:
            opts['title'] = 'Selecione um arquivo para salvar...'
            fn = fdlg.asksaveasfilename(**opts)
 
        if fn and ent:
            ent.delete(0, END)
            ent.insert(END, fn)
        if fn == '': return None
        return fn


def main():
    root = Tk()
    root.resizable(width=FALSE, height=FALSE)
    ex = Window(root, 692, 550)
    root.mainloop()


if __name__ == '__main__':
    reload(sys)
    sys.setdefaultencoding('utf-8')
    main()