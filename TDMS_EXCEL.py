# class for TDMS data and Excel output
import logging,traceback,sys
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
        self.resultDict={}

        logging.info('TDMS to Excel data procedure selected')


        self.startRowInResultsFile=3
        self.startColumnInResultsFile=2


    def copy_template_excel_file(self,excelDestPath,excelTemplateFilePath):
        # # 1. copy the data template excel file into the dirctory with the tdms files

        excelDestPath.replace("/","\\")

        if len(excelDestPath)>Const.MAX_PATHLENGTH_DOS: raise InvalidFilePathLengthException

        try:
            shutil.copyfile(excelTemplateFilePath, excelDestPath)
        except PermissionError:
            msg= traceback.format_exc()
            traceback.print_exc(file=sys.stdout)
            logging.warning(msg)
            pass #skip this error, will be handled later

        return excelDestPath
    

    
    def convert_data_to_csv(self,featureName, selectedDir,tdms_fileName, tdms_file):



        # 2. rename the columns of the TDMS data
        
        self.df_load = tdms_file.as_dataframe()


        for group in tdms_file.groups():
            for channel in group.channels():
                self.colNames.append(group.name + Const.TDMS_LIST_SEP  + channel.name)
                self.tdmsChannelProperties.update({group.name + Const.TDMS_LIST_SEP  + channel.name:channel.properties})

        self.df_load.columns=self.colNames

        # re-index the columns necessary. "Analog-IO -->" should always be first
        if ("TC -->" in self.df_load.columns[0]): 
            correct_list= [item for item in self.df_load.columns if "Analog-IO -->" in item]
            [correct_list.append(item) for item in self.df_load.columns if "TC -->" in item]
            # Reorder dataframe columns in alphabetical order. 
            self.df_load = self.df_load.reindex(columns=correct_list)

        # self.df_load= self.df_load.astype("float64") 


        
        # 3. create the csv file

        csvFilepath=selectedDir + "/"+ tdms_fileName.split(".tdms")[0] + "--" + featureName +".txt"
        self.df_load.to_csv(csvFilepath, index=False, na_rep='')

        self.df_load=None
        self.colNames=[]
        self.tdmsChannelProperties={}

        return csvFilepath


    def get_csv_data(self,csv_file_path):
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
        


    def write_result_to_excel_template(self, excelresultDestPath):     
        # Start  Excel
        xl_app = xw.App(visible=False, add_book=False)

        
        try:
            # Open template file
            wb = xl_app.books.open(excelresultDestPath)

            # Assign the sheet holding the template table to a variable
            ws = wb.sheets('Result')

            row = self.startRowInResultsFile
            column = self.startColumnInResultsFile
            # 1. Insert data to the Result Worksheet
            
            ws.range((row, column-1)).value="Link"

            for item in self.resultDict.keys():

                for iitem in self.resultDict[item].keys():
                    # write the label
                    if row==self.startRowInResultsFile:ws.range((row, column)).value = iitem
                    # write the value
                    ws.range((row+1, column)).value = self.resultDict[item][iitem]

   
                    if row==self.startRowInResultsFile+1: ws.autofit(axis="columns")

                    # ws.range((row, column+1)).api.WrapText = True
                    # ws.range((row-1, column+1)).api.WrapText = True
                    # ws.range((row, column+1)).column_width = 40
                    # ws.range((row, column+1)).row_height = 40
                    column=column+1

                column = self.startColumnInResultsFile
                
                ws.range((row+1, column-1)).add_hyperlink(item)  
                ws.range((row+1, column-1)).api.WrapText = True  
                ws.range((row+1, column-1)).column_width = 26
                ws.range((row+1, column-1)).row_height = 21


                row=row+1
                


            # ws.range("A:A").column_width = 26

            # ws.autofit(axis="rows")
            
           

            # Save and Close the Excel template file
            wb.save()
            wb.close()

            self.resultDict={}
            # Close Excel
            xl_app.quit()
        except:

            logging.warning("write_result_to_excel_template: Access to Excel went bad. Check if excel instance is in zombie state (Task Manager).")
            msg= traceback.format_exc()
            traceback.print_exc(file=sys.stdout)
            logging.error(msg)
            xl_app.quit()
            raise Exception("location: write_result_to_excel_template")



    def write_data_to_excel_template(self, template_file, data_to_insert, featureName,tdms_file):
        """
        Inserting data to an existing Excel data table
        :param template_file: path of the Excel template file
        :param data_to_insert: data to insert (list)
        :return: None
        """
        
            
        # Start  Excel
        xl_app = xw.App(visible=False, add_book=False)

        try:
            # Open template file
           
            wb = xl_app.books.open(template_file)

            # Assign the sheet holding the template table to a variable
            ws = wb.sheets('Source')
            row = 1
            column = 1
            # 1. Insert data to the Source Worksheet
            ws.range((row, column)).value = data_to_insert
            ws.autofit(axis="columns")

            # 2. do the same for the secified worksheet
            ws = wb.sheets(featureName)
            row = 1
            column = 1
            # however this time, rename the columns
            # Insert data
            search_str=["Analog-IO --> ","TC --> "]
            newHeaderList=[]
            for str in search_str:
                [newHeaderList.append(item.replace(str,"")) for item in data_to_insert[0] if str in item]
        
            data_to_insert[0]=newHeaderList
    
            ws.range((row, column)).value = data_to_insert
            ws.autofit(axis="columns")

            # 3. collect the result data to be used later

            # find the numbers of columns and rows in the sheet
            num_col = 54
            num_row = ws.range('BB2').end('down').row
            self.resultLabels=ws.range('BB:BB')[1:].value
            # collect data

            custom_content_list=[item for item in tdms_file.properties.keys()]

            content_list = ws.range((2,num_col),(num_row,num_col)).value
            result_list=   ws.range((2,num_col+1),(num_row,num_col+1)).value

            content_list=  custom_content_list + ["TestName"] + content_list 
         
            prop_list=[]
            for item in custom_content_list:
                prop_list.append(tdms_file.properties[item])

            
            result_list=prop_list+ [featureName] + result_list
      
            resultDict={}

            for num in range(0,len(result_list)-1):
                resultDict.update({content_list[num]:result_list[num]})
            

            self.resultDict.update({template_file:resultDict})

            # Save and Close the Excel template file
            wb.save()
            wb.close()

            # Close Excel
            xl_app.quit()
            # subprocess.call(["taskkill", "/f", "/im", "EXCEL.EXE"])
            return self.resultDict

        except: # close the started excel to not pose a problem later on
            # Close Excel
            msg= traceback.format_exc()
            traceback.print_exc(file=sys.stdout)
            logging.warning(msg)
            logging.warning("write_data_to_excel_template: Access to Excel went bad. Check if excel instance is in zombie state (Task Manager).")
            xl_app.quit()
            raise Exception("location: write_data_to_excel_template")




    