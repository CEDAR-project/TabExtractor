#!/usr/bin/env python

"""Extractor.py: A structure extractor for XLS files."""

import csv
from xlrd import open_workbook
from Levenshtein import ratio

class Extractor:
    """An XLS structure extractor"""

    def __init__(self):
        # Struct initialization
        self.municipalitiesPerSheet = {}
        self.similarityThreshold = 0.9

    def doExtraction(self, inputDataFile, col, row, lim):
        """Extract municipality names from specified location, column 
        top-bottom until lim"""

        # print "Harvesting data from input file {}...".format(inputDataFile)

        # Open workbook for input data
        self.inputDataFile = inputDataFile
        self.sourceData = open_workbook(self.inputDataFile, formatting_info=True)
        self.sourceSheet = self.sourceData.sheet_by_index(0)

        self.municipalitiesPerSheet[self.inputDataFile] = set()

        for i in range(row-1, lim):
            self.cell = self.sourceSheet.cell(i, col)
            if self.cell.value.strip() != "":
                self.municipalitiesPerSheet[self.inputDataFile].add(self.cell.value.strip())

    def serialize(self, outputDataFile):
        """Write municipalities found in specified output"""

        # print "Serializing to output file..."

        temp = set()

        for key in self.municipalitiesPerSheet.keys():
            
            for m in list(self.municipalitiesPerSheet[key]):
                temp.add(m.encode('utf8'))
                
        for m in sorted(list(temp)):                
            print '"'+m+'"',',',
            if m in self.municipalitiesPerSheet['data/VT_1859_01_H3A.xls']:
                print 'x,',
            else:
                print ',',
            if m in self.municipalitiesPerSheet['data/VT_1869_01_H1.xls']:
                print 'x,',
            else:
                print ',',
            if m in self.municipalitiesPerSheet['data/VT_1879_10_H1.xls']:
                print 'x,',
            else:
                print ',',
            if m in self.municipalitiesPerSheet['data/VT_1889_12_H1.xls']:
                print 'x,',
            else:
                print ',',
            if m in self.municipalitiesPerSheet['data/VT_1899_01_H1.xls']:
                print 'x,',
            else:
                print ',',
            if m in self.municipalitiesPerSheet['data/BRT_1909_01_T.xls']:
                print 'x'
            else:
                print ''

if __name__ == "__main__":
    munExtractorInstance = Extractor()
    munExtractorInstance.doExtraction('data/VT_1859_01_H3A.xls', 2, 9, 1216)
    munExtractorInstance.doExtraction('data/VT_1869_01_H1.xls', 0, 6, 1327)
    munExtractorInstance.doExtraction('data/VT_1879_10_H1.xls', 0, 12, 1920)
    munExtractorInstance.doExtraction('data/VT_1889_12_H1.xls', 0, 6, 2164)
    munExtractorInstance.doExtraction('data/VT_1899_01_H1.xls', 0, 5, 2796)
    munExtractorInstance.doExtraction('data/BRT_1909_01_T.xls', 0, 8, 36935)
    munExtractorInstance.serialize('output/output.csv')

__author__ = "Albert Meronyo-Penyuela"
__copyright__ = "Copyright 2012, VU University Amsterdam"
__credits__ = ["Albert Meronyo-Penyuela"]
__license__ = "LGPL v3.0"
__version__ = "0.1"
__maintainer__ = "Albert Meronyo-Penyuela"
__email__ = "albert.merono@vu.nl"
__status__ = "Prototype"

