# class for TDMS data and Excel output
import logging
import numpy as np
import pandas as pd
import shutil
import csv
from Util import Const
from Util import Functions as fnc
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
        self.tdmsChannelProperties={}

        self.tdmsProperties={}


        logging.info('TDMS to Excel data procedure selected')


    def copy_template_excel_file(self,selectedDir,tdms_fileName,featureName,excelTemplateFilePath):
        # # 1. copy the template excel file into the dirctory with the tdms files

        excelDestPath=selectedDir + "/"+ tdms_fileName.split(".tdms")[0] + "--" + featureName + ".xlsx"
        excelDestPath.replace("/","\\")

        if len(excelDestPath)>Const.MAX_PATHLENGTH_DOS: raise InvalidFilePathLengthException

        shutil.copyfile(excelTemplateFilePath, excelDestPath)

        return excelDestPath
    

    
    def convert_data_to_csv(self,featureName, selectedDir,tdms_fileName, tdms_file):



        # 2. rename the columns of the TDMS data
        
        self.df_load = tdms_file.as_dataframe()


        for group in tdms_file.groups():
            for channel in group.channels():
                self.colNames.append(group.name + Const.TDMS_LIST_SEP  + channel.name)
                self.tdmsChannelProperties.update({group.name + Const.TDMS_LIST_SEP  + channel.name:channel.properties})

        self.df_load.columns=self.colNames

        # self.df_load= self.df_load.astype("float64") 


        
        # 3. create the csv file

        csvFilepath=selectedDir + "/"+ tdms_fileName.split(".tdms")[0] + "--" + featureName +".txt"
        self.df_load.to_csv(csvFilepath, index=False, na_rep='')

        self.df_load=None
        self.colNames=[]
        self.tdmsChannelProperties={}

        return csvFilepath


    def open_csv_file(self,csv_file_path):
        """
        Open and read data from a csv file without headers (skipping the first row)
        :param csv_file_path: path of the csv file to process
        :return: a list with the csv content
        """
        with open(csv_file_path, 'r', encoding='utf-8') as csv_file:
            reader = csv.reader(csv_file)

            # Skip header row
            # next(reader)

            # Add csv content to a list
            data = list()
            for row in reader:
                data.append(row)

            return data
        

    def write_list_to_excel(self, template_file, data_to_insert):
        """
        Inserting data to an existing Excel data table
        :param template_file: path of the Excel template file
        :param data_to_insert: data to insert (list)
        :return: None
        """

        
        # Start Visible Excel
        xl_app = xw.App(visible=False, add_book=False)

        # Open template file
        wb = xl_app.books.open(template_file)

        # Assign the sheet holding the template table to a variable
        ws = wb.sheets('Source')

        # First cell of the template (blank) table
        row = 1
        column = 1

        # Insert data
        ws.range((row, column)).value = data_to_insert

        # Save and Close the Excel template file
        wb.save()
        wb.close()

        # Close Excel
        xl_app.quit()



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
        # os.remove(excelDestPath)
        