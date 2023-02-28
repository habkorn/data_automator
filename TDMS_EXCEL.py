# class for TDMS data and Excel output
import logging
import numpy as np
import pandas as pd
import shutil

from Util import Const
from Util import InvalidFilePathLengthException

import subprocess
import os
from openpyxl import Workbook
from openpyxl import load_workbook
import time
from pandas import ExcelWriter
import openpyxl
import xlwings as xw
from pathlib import Path

class TDMS_EXCEL():

    def __init__(self)-> None:
    # body of the constructor
        self.title=Const.OPTIONS[0]
        self.df_load=None



        self.colNames=[]
        self.tdmsProperties={}


        logging.info('TDMS to Excel data procedure selected')


    def convert_to_csv(self,featureName, selectedDir,tdms_fileName, tdms_file,excelTemplateFilePath):

        # # 1. copy the template excel file into the dirctory with the tdms files

        excelDestPath=selectedDir + "/"+ tdms_fileName.split(".tdms")[0] + "--" + featureName + ".xlsx"
        excelDestPath.replace("/","\\")

        if len(excelDestPath)>Const.MAX_PATHLENGTH_DOS: raise InvalidFilePathLengthException
        


        shutil.copyfile(excelTemplateFilePath, excelDestPath)

        # 2. rename the columns of the TDMS data
        
        self.df_load = tdms_file.as_dataframe()


        for group in tdms_file.groups():
            for channel in group.channels():
                self.colNames.append(group.name + Const.TDMS_LIST_SEP  + channel.name)
                self.tdmsProperties.update({group.name + Const.TDMS_LIST_SEP  + channel.name:channel.properties})

        self.df_load.columns=self.colNames

        # self.df_load= self.df_load.astype("float64") 

        
        # 3. create the csv file

        self.df_load.to_csv(selectedDir + "/"+ tdms_fileName.split(".tdms")[0] + "--" + featureName +".txt", index=False, na_rep='')

        self.df_load=None
        self.colNames=[]
        self.tdmsProperties={}


    def run_excel_macro(self, selectedDir):
       
        excelSrcPath=(os.getcwd() + "/" + Const.EXCEL_TEMPLATEFOLDER + "/CSV_Automator.xlsm").replace("/","\\")
        excelDestPath=(selectedDir + "/" + Const.EXCEL_CSV_AUTOMATOR_FILENAME).replace("/","\\")


       # 0. delete the result excel file (if it exists)

        try:
            os.remove((selectedDir  + "/" + Const.EXCEL_RESULT_FILENAME).replace("/","\\"))
        except OSError:
            pass

       
       # 1. copy the excel macro automator file to the data directory
        if len(excelDestPath)>Const.MAX_PATHLENGTH_DOS: raise InvalidFilePathLengthException
        
        shutil.copyfile(excelSrcPath, excelDestPath)
        
        # 2.  run the excel macro

        logging.info("Run the excel macro to fill the templates with data..")
        wb=xw.Book(excelDestPath)
        macro=wb.macro("run_csv_to_excel")
        macro()
        
        wb.close()
        
        # 3.  delete the excel macro file
        os.remove(excelDestPath)
        