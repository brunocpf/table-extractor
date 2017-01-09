# -*- coding: utf-8 -*-
from pdfminer.pdfparser import PDFParser
from pdfminer.pdfdocument import PDFDocument
from pdfminer.pdfpage import PDFPage
from pdfminer.pdfpage import PDFTextExtractionNotAllowed
from pdfminer.pdfinterp import PDFResourceManager
from pdfminer.pdfinterp import PDFPageInterpreter
from pdfminer.pdfdevice import PDFDevice
from pdfminer.layout import LAParams
from pdfminer.converter import PDFPageAggregator
import pdfminer
from operator import itemgetter
import re
import csv
import sys
import io


class SinapiExtractor(object):
    def __init__(self, path, startPage, endPage, yOffset, height, columns, multilineColumns, multilineSeparator):
        """
        Initializes the parser structure.
        @params
            path                 - Required : path of the PDF file used as input
            startPage..endPage   - Required : range of the pages to parse
        """
        fp = open(path, 'rb')
        self.parser = PDFParser(fp)
        self.document = PDFDocument(self.parser)
        if not self.document.is_extractable:
            raise PDFTextExtractionNotAllowed
        self.rsrcmgr = PDFResourceManager()
        self.device = PDFDevice(self.rsrcmgr)
        self.laparams = LAParams()
        self.device = PDFPageAggregator(self.rsrcmgr, laparams=self.laparams)
        self.interpreter = PDFPageInterpreter(self.rsrcmgr, self.device)
        self.startPage = startPage
        pageTotal = len(list(PDFPage.get_pages(fp)))
        self.endPage = endPage if 0 <= endPage <= pageTotal else pageTotal
        self.yOffset = yOffset
        self.height = height
        self.cols = columns
        self.multilineColumns = multilineColumns
        self.multilineSeparator = multilineSeparator


    def getColumns(self, objList, columns, page):
        """
        Generates a list containing all text included in each column.
        @params
            objList    - Required : list of objects being parsed
            columns    - Required : current list
            page       - Required : index of the page being parsed
        """
        for obj in objList:
            if isinstance(obj, pdfminer.layout.LTChar):
                for idx, col in enumerate(self.cols):
                    h = self.height.get(page, self.height[-1])
                    if (col['x'] <= obj.bbox[0] <= (col['x'] + col['W'])) and ((self.yOffset - h) <= obj.bbox[1] <= self.yOffset):
                        columns[idx][obj.bbox[1]] = columns[idx].get(obj.bbox[1], "") + obj.get_text()
            else:
                try:
                    self.getColumns(obj._objs, columns, page)
                except AttributeError:
                    continue
        return columns


    def generateRows(self, window):
        """
        Returns the rows.
        """
        rows = []
        rowIndex = 0
        peek = False
        for pageNumber, page in enumerate(PDFPage.create_pages(self.document)):
            if pageNumber > self.endPage: break
            if pageNumber < self.startPage: continue
            window.refreshProgress((pageNumber - self.startPage) + 1, self.endPage - self.startPage)            
            # read the page into a layout object
            self.interpreter.process_page(page)
            layout = self.device.get_result()
 
            columns = []
            columns = self.getColumns(layout._objs, [{} for i in range(len(self.cols))], pageNumber)
 
            # transforms each dict in the columns list into a sorted list of tuples
            for idx, col in enumerate(columns):
                columns[idx] = [(x, col[x]) for x in sorted(col, reverse=True)]
 
            # remove whitespace strings from the columns list
            for i, col in enumerate(columns):
                for j, (_, text) in enumerate(col):
                    if text.strip() == "":
                        columns[i].pop(j)
            for i, (_, text) in enumerate(columns[0]):
                if text.strip() == "":
                    columns[0].pop(i)
  
            try:
                firstY,_ = columns[0][0]
            except IndexError:
                firstY = -1
            
            # rows are generated here
            for idx, (rowY, txt) in enumerate(columns[0]):
                rows.append(["" for x in range(len(self.cols))])
                rowIndex = len(rows) - 1
                try:
                    (nextRowY, _) = columns[0][idx+1]
                except IndexError:
                    nextRowY = -1
                columnsCount = 0
                for i in range(len(self.cols)):
                    columnExists = False
                    for (lineY, line) in columns[i]:
                        if (lineY > firstY) and i in self.multilineColumns and peek:
                            rows[peek][i] += line + self.multilineSeparator
                        if (lineY == rowY) or ((nextRowY < lineY < rowY) and i in self.multilineColumns):
                            rows[rowIndex][i] += line + self.multilineSeparator
                            columnExists = True
                    if columnExists: columnsCount += 1
                if columnsCount != len(self.cols):
                    rows.pop(rowIndex)
                peek = False
            peek = rowIndex
        
        # format and encode the strings
        for rowIndex, row in enumerate(rows):
            for colIndex, text in enumerate(row):
                rows[rowIndex][colIndex] = text.strip(self.multilineSeparator).strip().replace("\n", " ").encode('utf-8')
        return rows