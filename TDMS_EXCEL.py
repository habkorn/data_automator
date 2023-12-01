# class for TDMS data and Excel output

from Util import InvalidFilePathLengthException, Const
import logging,traceback,sys
import numpy as np
import pandas as pd
import shutil
import csv
# from Util import Const




# from openpyxl import load_workbook

import xlwings as xw
from xlwings.utils import rgb_to_int


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
        self.startColumnInResultsFile=1


    def copy_template_excel_file(self,excelDestPath,excelTemplateFilePath):
        # # 1. copy the data template excel file into the dirctory with the tdms files

        excelDestPath.replace("/","\\")

        

        try:
            if len(excelDestPath)>Const.MAX_PATHLENGTH_DOS: raise InvalidFilePathLengthException
            shutil.copyfile(excelTemplateFilePath, excelDestPath)
        except PermissionError:
            msg= traceback.format_exc()
            traceback.print_exc(file=sys.stdout)
            logging.warning(msg)
            pass #skip this error, will be handled later
        except InvalidFilePathLengthException:
            logging.critical("The max file path length is exceeded.")
            raise
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

        if len(self.df_load)>=Const.EXCEL_MAX_ROWS-10:
            logging.critical("The number of rows in current TDMS file exceeds the maximum rows allowable in Excel (" + str(Const.EXCEL_MAX_ROWS) + ")")
            logging.critical("The program will save the current Excel file anyway, DISCARDING the excess rows.")
            
            # cut off the excess rows (and some)
            self.df_load=self.df_load.head(Const.EXCEL_MAX_ROWS - 10)
        
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
            ws_res = wb.sheets('Result')

            row = self.startRowInResultsFile
            column = self.startColumnInResultsFile

            # find the result file with the most entries and write the first line with captions
            lenEnt=-1
            savKey=[]
            for item in self.resultDict.keys():

                for iitem in self.resultDict[item].keys():

                    if lenEnt<len(self.resultDict[item].keys()): 
                        lenEnt=len(self.resultDict[item].keys())
                        savKey=item

            
            
            ws_res.range((row, 1)).value="Link"

            num=0
            for iitem in self.resultDict[savKey].keys():

                ws_res.range((row, 2+num)).value=iitem
                num=num+1


            # 1. Insert data to the Result Worksheet and find the correct cell

            for item in self.resultDict.keys():

                for iitem in self.resultDict[item].keys():
                    
                    # search for the label in the excel file. If it is not found, write new column
                    column=-1

                    for col in range(1, lenEnt+10):
                        if ws_res.range((3,col)).value == iitem:
                            column=col

                    if column==-1:
                        print("!!")
                    # write the value
                    ws_res.range((row+1, column)).value = self.resultDict[item][iitem]



                    # ws.range((row, column+1)).api.WrapText = True
                    # ws.range((row-1, column+1)).api.WrapText = True
                    # ws.range((row, column+1)).column_width = 40
                    # ws.range((row, column+1)).row_height = 40

                    # column=column+1

                # column = self.startColumnInResultsFile
                
                try:
                    ws_res.range((row+1, 1)).add_hyperlink(item)  

                except:
                    ws_res.range((row+1, 1)).value=str(item).replace("/","\\")
                    pass 

                if column>1:
                    ws_res.range((row+1, column-1)).api.WrapText = True  
                    ws_res.range((row+1, column-1)).column_width = 26
                    ws_res.range((row+1, column-1)).row_height = 21



                row=row+1
                
            ws_res.autofit(axis="columns")
            ws_res.range((1, 1)).column_width = 10
            # 2. transpose the data onto a seperate  worksheet

            # collect data
            content_list = ws_res.range((1,1),(100,200)).value

            ws_res_t = wb.sheets('Result_transponiert')
            ws_res_t.range('A1').options(transpose=True).value=content_list

            ws_res_t.range((16, 3),(29, 3)).api.Font.Color = rgb_to_int((0, 0, 255))

            for num_col in range(0, len(self.resultDict.keys())):
                try:
                    val=ws_res_t.range((1, num_col+4)).value
                    ws_res_t.range((1, num_col+4)).add_hyperlink(val)  
                # 16-29
                    ws_res_t.range((16, num_col+4),(29, num_col+4)).api.Font.Color = rgb_to_int((0, 0, 255))

                except:
                    # do nothing
                    pass     

            
            ws_res_t.autofit(axis="rows")
            ws_res_t.autofit(axis="columns")
            ws_res_t.range("A:B").column_width = 5
            ws_res_t.range("C:C").column_width = 45
            ws_res_t.range("D:CZ").column_width = 16
            ws_res.range('D2').WrapText = True  

            
            wraptext_t = wb.macro('wraptext_t')
            wraptext_t()
           
            # ws_res.range((1,3),(2,200)).WrapText = True  
            
            # for r,c in zip(range(3, 6),range(3, 6)):
            #     wb.sheets('Result_transponiert').range((r,c)).options(transpose=True).value = wb.sheets('Result_transponiert').range((r,c)).value

            # ws_t.range("A:DA").column_width = 52

            # ws.autofit(axis="rows")

            # sht.range('A1').options(transpose=True).value = [1,2,3,4]
            # Result_transponiert
           

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

    def dump_large_array_to_excel(self, ws, wb, startrow, startcolumn, data,transp):

        row=startrow
        col=startcolumn
        
        if len(data) <= (Const.EXCEL_MAX_CHUNK_SIZE + 1):

            if transp: 
                ws.range((row,col)).options(transpose=True).value = data
                # print(row)
            else: 
                ws.range((row,col)).value = data

        else:
            for chunk in (data[rw:rw + Const.EXCEL_MAX_CHUNK_SIZE] 
                for rw in range(0, len(data), Const.EXCEL_MAX_CHUNK_SIZE)):
                    # ws.range('Source', row + str(column), index = False, header = False).value = chunk
                    
                    if transp: 
                        ws.range((row,col)).options(transpose=True).value = chunk
                        # print(row)
                    else: 
                        ws.range((row,col)).value = chunk

                    row += Const.EXCEL_MAX_CHUNK_SIZE

        # print("!!")
        
    def filterData(self,featureName,data_from_csv):
        
        data=data_from_csv.copy() # default
        filteredData=None

        ########### general filter
        # rename the columns, exclude some annoying parts

        search_str=["Analog-IO --> ","TC --> "]  # default search Strings

        newHeaderList=[]
        for _str in search_str:
            [newHeaderList.append(item.replace(_str,"")) for item in data[0] if _str in item]

        data[0]=newHeaderList

        filteredData=np.array(data).transpose().tolist()

        ########### feature specific filter

        if featureName=="F2-F3 RPM": # reduce the columns and and limit the data
            
            # search_str=["Drehzahl","Strom_LD_Ebene_2","Strom_LD_Ebene_1"]

            CURRENT_SIG_NAME="Strom_LD_Ebene_1"

            search_str=["Drehzahl",CURRENT_SIG_NAME]
            newHeaderList=[]
            temp_data_to_insert=[]

            for _str in search_str:
                kk=0
                for item in data_from_csv[0]:
                    if  _str in item:
                        idx_LLZ_GV=[]
                        [idx_LLZ_GV.append(data[kkk][kk]) for kkk in range(0,len(data)-1)]
                        idx_LLZ_GV=[idx_LLZ_GV]
                        idx_LLZ_GV=[list(i) for i in zip(*idx_LLZ_GV)]
                        # del data_to_insert[:][kk]
                        temp_data_to_insert.append(idx_LLZ_GV)
                    kk=kk+1
            
            temp_data_to_insert=[list(i) for i in zip(*temp_data_to_insert)]

            flat_list=[]
            for sublist in temp_data_to_insert:
                flat_sublist = []
                
                for item in sublist: 
                    flat_sublist.append(item[0])
                
                flat_list.append(flat_sublist)
            

            np_arr = np.array(flat_list[1:]) 
            np_arr=np_arr.astype(np.float)
        
            tp_arr = np_arr.transpose()
            rpm_sig=tp_arr[0]
            current_sig=tp_arr[1]

            # rpm_sig = rpm_sig > 4
            # current_sig = current_sig > 1

            RPM_THRESHOLD=5.

            # idx_LD_AGD_level_up_flank=np.where(np.logical_and(rpm_sig > RPM_THRESHOLD, current_sig > 1))[0]
            idx_LD_AGD_level_up_flank=np.where(rpm_sig >= RPM_THRESHOLD)[0]
            
            
            partitions=idx_LD_AGD_level_up_flank[np.where(np.diff(idx_LD_AGD_level_up_flank)!=1)]
            
            if len(partitions)==0 or len(partitions)==1: 
                partitions=[idx_LD_AGD_level_up_flank.min(),idx_LD_AGD_level_up_flank.max()]

            else: 
                partitions = np.insert(partitions, 0, idx_LD_AGD_level_up_flank[0])


            # lastlen=0
            # while not lastlen==len(partitions):
            #     lastlen=len(partitions)
            #     partitions = np.delete(partitions, np.where(np.diff(partitions) < 1000))

            rpm_part=[]
            current_part=[]

            for i in range(1,len(partitions)):
                rpm_t=rpm_sig[partitions[i-1]:partitions[i]]
                current_t=current_sig[partitions[i-1]:partitions[i]]

                rpm_part.append(rpm_t)
                current_part.append(current_t)

            for k,p in enumerate(rpm_part):
                # remove outliers, i.e. anything below rpm threshold
                idx_del=np.where(p < RPM_THRESHOLD)

        
                rpm_part[k]=np.delete(rpm_part[k],idx_del)
                current_part[k]=np.delete(current_part[k],idx_del)

            
            rpm_tt=[]
            current_tt=[]

            for k,p in enumerate(rpm_part):
                if len(rpm_part[k])>=100: 
                    rpm_tt.append(rpm_part[k])
                    current_tt.append(current_part[k])

            
            rpm_part=rpm_tt
            current_part=current_tt

            for k,p in enumerate(rpm_part):
                # remove outliers, i.e. anything below rpm threshold
                idx_del=np.where(p<=np.percentile(p, 1))

                rpm_part[k]=np.delete(rpm_part[k],idx_del)
                current_part[k]=np.delete(current_part[k],idx_del)


            rpm_k=[]
            current_k=[]
            for k,p in enumerate(rpm_part):
                    if np.std(rpm_part[k])>=0.02:
                        rpm_k.append(rpm_part[k]) 
                        current_k.append(current_part[k]) 

            rpm_part=rpm_k
            current_part=current_k
            
            
            revs=[]
            lieferzeiten=[]

            
            for k,p in enumerate(rpm_part):
                if np.mean(p)==0. or len(p)==0.:
                    rpm_part.remove(p)
                    current_part.pop(k)
                else:
                    revs.append(np.mean(p)*1000./60.*len(p)/1000.)
                    lieferzeiten.append(len(p)/1000.)

            
                    #    'handle the case when only one signal is found'
            RPM_LLZ_THRESHOLD=4000.
            
            if len(revs)==1:
                if revs[0]<RPM_LLZ_THRESHOLD: 
                    'Gewichtsverstellung'
                    revs.append(revs[0])
                    revs[0]=0

                    lieferzeiten.append(lieferzeiten[0])
                    lieferzeiten[0]=0

                    rpm_part.append(rpm_part[0].copy())
                    for i,e in enumerate(rpm_part[0]): rpm_part[0][i]=0 

                    current_part.append(current_part[0].copy())
                    for i,e in enumerate(current_part[0]): current_part[0][i]=0 

                else:
                    'Lieferleistung'
                    revs.append(0)
                    lieferzeiten.append(0)

                    rpm_part.append(rpm_part[0].copy())
                    for i,e in enumerate(rpm_part[1]): rpm_part[1][i]=0

                    current_part.append(current_part[0].copy())
                    for i,e in enumerate(current_part[1]): current_part[1][i]=0 

            # for k,p in enumerate(rpm_part):

            #     xpoints = np.array(range(0,len(p)))
            #     ypoints = np.array(p)
            #     plt.subplot(1, len(rpm_part), k+1)
            #     plt.plot(xpoints, ypoints)
                
            #     print(f'{k+1}' + ": " +  f'{np.std(p)}')
  
            # plt.show()


            tp_flat_list = np.array(flat_list[1:]).transpose()
            

            temp_arr=[]

            for i,p in enumerate(tp_flat_list):
                temp_arr.append(tp_flat_list[i].tolist())

            temp_arr[0]=["Drehzahl"]+temp_arr[0]
            temp_arr[1]=[CURRENT_SIG_NAME]+temp_arr[1]
            # temp_arr[2]=["Strom_LD_Ebene_1"]+temp_arr[2]
            
            
            
            for i,sublist in enumerate(rpm_part):
                rp=rpm_part[i].astype(str)
                rp=np.insert(rp,0," ")
                rp=np.insert(rp,0,str(revs[i]))
                rp=np.insert(rp,0,str(lieferzeiten[i]))
                rp=np.insert(rp,0,str(pd.Series(current_part[i][2000:]).rolling(100).mean().max()))  # discard the first 2000 ms
        
                if revs[i]>RPM_LLZ_THRESHOLD: 
                    temp_arr.append(["Umdrehungen, Filter LLZ " + str(int(i/2))]+rp.tolist())
                else: 
                    temp_arr.append(["Umdrehungen, Filter GV " + str(int(i/2))]+rp.tolist())
            

            for i,sublist in enumerate(current_part):
                    if revs[i]>RPM_LLZ_THRESHOLD: 
                        temp_arr.append([CURRENT_SIG_NAME+", Filter LLZ "+ str(int(i/2))]+current_part[i].astype(str).tolist())
                    else: 
                        temp_arr.append([CURRENT_SIG_NAME+", Filter GV "+ str(int(i/2))]+current_part[i].astype(str).tolist())
              
   
            # flat_list=[]
            # # flat_list = [list(i) for i in zip(*temp_arr)]
            # c=0
            # for i in zip(*temp_arr):
            #     if c>0 :
            #         flat_list.append([float(ii) for ii in i])
            #     else:
            #         flat_list.append(list(i))
                
            #     c=c+1
                                

            filteredData=temp_arr

        return filteredData


    def write_data_to_excel_template_start_macro(self, template_file, data_to_insert, filtered_data,featureName,tdms_file):
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
            # 1. Insert ALL data to the Source Worksheet
            self.dump_large_array_to_excel(ws, wb,row, column, data_to_insert,False)

            ws.autofit(axis="columns")

            # 2. do the same for the secified worksheet, except only the columns with useful data
            ws = wb.sheets(featureName)
            row = 1
            column = 1
           

            # wb.sheets('Source').range('A1').end('down').end('right')

            # ws.range((row, column)).value=wb.sheets('Source').used_range.value
    
            for i,p in enumerate(filtered_data):
                self.dump_large_array_to_excel(ws, wb,row, column+i,  p,True)

            ws.autofit(axis="columns")
             # erase unuseful range in excel, e.g. exclude columns "F" to "P"
            # search fur the column names with only 1 letter

            discarded_columns=[]
            shift_alphabet=3
            
           
            newHeaderList=ws.range((1,1),(1,20))
            for item in newHeaderList:
                
                if len(str(item.value))==1: 
                    excel_column=chr(ord(str(item.value))+shift_alphabet)
                    discarded_columns.append(excel_column + ":"+ excel_column)

            # delete the discarded columns
            for item in discarded_columns:
                ws.range(item)[1:].clear_contents()
            

            ws.autofit(axis="columns")

            # 3. collect the result data to be used later

            # find the numbers of columns and rows in the sheet


            vb_macro = wb.macro("vbMacro")



            if vb_macro()==False: 
                logging.error("The TDMS File "+ (template_file.split("/")[-1]).replace(".xlsm","")+".tdms" +" is corrupted in some way. Please check the data carefully.")
            else: logging.info("vbMacro successful.")
            
            # collect result data

            num_col = 54
            num_row = ws.range('BB2').end('down').row
            if num_row>100:num_row=100

            self.resultLabels=ws.range('BB:BB')[1:num_row].value

            custom_content_list=[item for item in tdms_file.properties.keys()]

            content_list = ws.range((2,num_col),(num_row,num_col)).value
            result_list=   ws.range((2,num_col+1),(num_row,num_col+1)).value

            content_list=  custom_content_list + ["TestName"] + content_list 
            ws.autofit(axis="columns")
            prop_list=[]
            for item in custom_content_list:
                prop_list.append(tdms_file.properties[item])

            
            result_list=prop_list+ [featureName] + result_list
      
            resultDict={}

            for num in range(0,len(result_list)):
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




    