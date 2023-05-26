import sys, traceback
import os
import types
import subprocess



from TDMS_EXCEL import TDMS_EXCEL

from PyQt5.QtWidgets import QApplication, QWidget, QComboBox, QPushButton, QFileDialog, QVBoxLayout,QHBoxLayout, QLabel,QGridLayout,QRadioButton,QSpacerItem, QSizePolicy
from PyQt5 import QtCore, QtGui, QtWidgets
from nptdms import TdmsFile
import glob


from PyQt5.QtCore import *
from PyQt5.QtGui import *

import logging
import time, json
from Util import Const,InvalidFilePathLengthException, ProxyModel
import os.path



# Uncomment below for terminal log messages
# logging.basicConfig(level=logging.DEBUG, format=' %(asctime)s - %(name)s - %(levelname)s - %(message)s')

class QTextEditLogger(logging.Handler):
    def __init__(self, parent):
        super().__init__()
        self.widget = QtWidgets.QPlainTextEdit(parent)
        # self.widget.setReadOnly(True)

    def emit(self, record):
        msg = self.format(record)
        self.widget.appendPlainText(msg)
        # Process events between short sleep periods
        QtWidgets.QApplication.processEvents()
        # time.sleep(0.1)


class CSI_AUTOMATOR(QWidget):
    def __init__(self):
        super().__init__()
        # font
        smallFont = QFont('Arial', 14)
        bigFont = QFont('Arial', 20)

        self.workingDir=""

        self.tdms_excel=TDMS_EXCEL()
        self.functions = [self.procTDMSDataforCSI, self.procEmpty]
        self.jsonDict={}

        self.excelTemplateFilesPath=None
        self.radioButtons=[]
        self.optionsLayout = QHBoxLayout()

        self.window_width, self.window_height = 1000, 200
        self.setMinimumSize(self.window_width, self.window_height)

        self.label2=QLabel("Found Templates:")
        self.label2.setFont(smallFont)

        self.logTextBox = QTextEditLogger(self)
        # You can format what is printed to text box

        self.log_formatter=logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')

        self.logTextBox.setFormatter(self.log_formatter)

        logging.basicConfig(format='%(asctime)s - %(levelname)s - %(message)s', filemode='a') 

        logging.getLogger().addHandler(self.logTextBox)


        # You can control the logging level
        logging.getLogger().setLevel(logging.INFO)
        
        # timer = QtCore.QTimer(self)
        # timer.timeout.connect(self.doTimer)
        # timer.start(1000)
 
        self.selectedDir=None
        # self.resize(737, 596)

        self.setWindowTitle("Data Automator V1.3.1")
        self.setWindowIcon(QtGui.QIcon("icon.png"))

        layout = QVBoxLayout()
        self.setLayout(layout)

        self.options = (Const.OPTIONS)

        self.label1=QLabel("Choose action:")
        self.label1.setFont(smallFont)
        layout.addWidget(self.label1)
        
        self.combobox = QComboBox()
        self.combobox.currentIndexChanged.connect(self.selectionchange)
        

        # adding action to combo box
        
        self.combobox.addItems(self.options)
        self.combobox.setFont(smallFont)

        # self.combobox.setCurrentIndex(0) # trigger the signal

        layout.addWidget(self.combobox)



        self.launch_btn = QPushButton('Launch')
        self.launch_btn.setFont(bigFont)
        self.launch_btn.clicked.connect(self.launchButton)
        self.launch_btn.setEnabled(False)
        
        


        layout.addWidget(self.label2)


        layout.addLayout(self.optionsLayout)

        layout.addWidget(self.launch_btn)

        layout.addWidget(self.logTextBox.widget)

        logging.info('Start of program')
        # logging.info('something to remember')
        # logging.warning('that\'s not right')
        # logging.error('foobar')


    # def doTimer(self):
    #     text=self.logTextBox.widget.toPlainText()
    #     self.logTextBox.widget.setPlainText(text)




    def radioClicked(self):
        radioButton = self.sender()
        if radioButton.isChecked():
            self.launch_btn.setEnabled(True)
            print("feature is %s" % (radioButton.feature))
        

    def selectionchange(self,i):

        if self.combobox.currentText()==self.tdms_excel.title:
           

            # hSpacer = QSpacerItem(350, 40, QSizePolicy.Maximum, QSizePolicy.Expanding)
            # self.optionsLayout.addItem(hSpacer)

            self.workingDir=os.getcwd() + '/'
            
            xlsxTemplateFiles = glob.glob(self.workingDir + Const.EXCEL_TEMPLATEFOLDER + r'/*.xlsm')

            if len(xlsxTemplateFiles)==0:
                self.workingDir=os.getcwd() + '/dist/'
                xlsxTemplateFiles = glob.glob(self.workingDir + Const.EXCEL_TEMPLATEFOLDER + r'/*.xlsm')

            self.excelTemplateFilesPath = [item.replace("/","\\") for item in xlsxTemplateFiles if not "~" in item ]
            self.excelTemplateFilesPath = [item for item in self.excelTemplateFilesPath if not "Result_Collection_Template" in item]

            self.label2.setText("Found Templates (" + str(len (self.excelTemplateFilesPath)) + ")")
            
            for file in self.excelTemplateFilesPath:
                featureName=file.rsplit('\\')[-1].split(".")[0]
                radiobutton = QRadioButton(featureName)
                # radiobutton.setChecked(True)
                radiobutton.feature = featureName
                self.radioButtons.append(radiobutton)
                radiobutton.toggled.connect(self.radioClicked)
                self.optionsLayout.addWidget(radiobutton)

        else: 
            print("else")
            for i in reversed(range(self.optionsLayout.count())): 
                self.optionsLayout.itemAt(i).widget().setParent(None)
                self.radioButtons=[]

            self.label2.setText("Found Templates (0)")


    def launchButton(self):
        curIndex = self.options.index(self.combobox.currentText())
        response = self.functions[curIndex]()
        


    def procEmpty(self):
        print('Got Nothing')

        
    def procTDMSDataforCSI(self):

        # dialog = QFileDialog(self, 'Select a folder containing TDMS files', os.getcwd())


        # dialog.setFileMode(QFileDialog.DirectoryOnly)
        # dialog.setOption(QFileDialog.ShowDirsOnly, False)

        
       
        settings_file_path=(self.workingDir + "settings.json").replace("/","\\")
        
        if not os.path.isfile(settings_file_path):
            settings_file = open(settings_file_path, "w")
            self.jsonDict={"lastDir": os.getcwd()}
            json.dump(self.jsonDict, settings_file, indent = 6)
            settings_file.close()

        settings_file = open(settings_file_path)
        settings_data = json.load(settings_file)
        settings_file.close()


        dialog = QFileDialog(self, 'Select a folder containing TDMS files', settings_data["lastDir"])
        dialog.setFileMode(QFileDialog.Directory)
        dialog.setOption(QFileDialog.DontUseNativeDialog, True)

        proxy = ProxyModel(dialog)
        dialog.setProxyModel(proxy)

        # sidebar links
        urls = []
        urls.append(QUrl.fromLocalFile(os.path.expanduser('~')))
        urls.append(QUrl.fromLocalFile(settings_data["lastDir"]))

        urls.append(QUrl.fromLocalFile(QStandardPaths.writableLocation(QStandardPaths.DesktopLocation)))
        urls.append(QUrl.fromLocalFile(QStandardPaths.writableLocation(QStandardPaths.HomeLocation)))
        urls.append(QUrl.fromLocalFile(QStandardPaths.writableLocation(QStandardPaths.DownloadLocation)))

        urls.append(QUrl.fromLocalFile("F:\\Entwicklung"))
        urls.append(QUrl.fromLocalFile("K:\\"))

        dialog.setSidebarUrls(urls)


        # dialog.setFileMode(QFileDialog.DirectoryOnly)
        # dialog.selectNameFilter("TDMS Files (*.tdms)")

        if dialog.exec_() == QFileDialog.Accepted:

            

            self.selectedDir = dialog.selectedFiles()[0]

            # write the selected in a JSON file

            self.jsonDict={"lastDir": self.selectedDir}
            settings_file = open(os.getcwd() + "/settings.json".replace("/","\\"), "w")
            json.dump(self.jsonDict, settings_file, indent = 6)
            settings_file.close()

            fh = logging.FileHandler(filename=self.selectedDir + "/"+ "Log.txt", mode="a")
           
            fh.setFormatter(self.log_formatter)
            
            logging.getLogger().addHandler(fh)

            # self.selectedDir="F:/ENTWICKLUNG/SSC019 CSI 3 Modul/04 Erprobungen/02 Interne Prüfberichte/MST2022081100_SSC018-SSC019_DV2_CSI3_KompressorModul/Ergebnisse/L1 - Lebensdauerprüfung/12V/Thomas DV2 tdms/Run1"
    
            logging.info("Selected Dir: " + self.selectedDir)



            # search for all TDMS files within selected directory
            tdmsFiles = glob.glob(self.selectedDir + r'/*.tdms')

            if len(tdmsFiles)==0:
                logging.critical("No TDMS Files found in selected folder. Please retry.")
            else:
                self.launch_btn.setEnabled(False)
                # get the mst name from the first file
                mst_name="MST" + (tdmsFiles[0].split("MST")[1]).split("_")[0]

                logging.info("TDMS found:")
                for tdmsFilePath in tdmsFiles:
                    logging.info(str(tdmsFilePath))

                logging.info(str(len(tdmsFiles)) + " TDMS File(s) found.")    

                    
                logging.info("EXCEL Templates found in: " + str(os.getcwd()))
                for excelTemplateFile in self.excelTemplateFilesPath:
                    logging.info(str(excelTemplateFile))
                logging.info("and ./" + Const.EXCEL_RESULT_FILENAME )
                
                logging.info(str(len(self.excelTemplateFilesPath)+1) + " EXCEL Templates File(s) found.")   
                
                featureName=""
                for rb in self.radioButtons:
                    if rb.isChecked(): 
                        featureName=rb.feature
                        logging.info("Template for " + str(rb.feature) + " will be processed.")

                if featureName=="": 
                    logging.warning("No Template was selected. Please select one.")
                    return


                excelTemplateFilePath = [x for x in self.excelTemplateFilesPath if featureName+".xlsm" == x.split("\\")[-1]]
                
                # for x in self.excelTemplateFilesPath:
                #     if featureName+".xlsm" == x.split("\\")[-1]: 
                #         excelTemplateFilePath=x
                
                excelTemplateFilePath=excelTemplateFilePath[0]


                # if all conditions are met: do the process
                            
                try:  

                    # 0. write tdms properties to a text file
                    num=0
                
                    # delete the Result_Collection file (if it exists)
                    try:
                        resFiles = glob.glob(self.selectedDir + r'/Result_Collection' + mst_name + r'.xlsm')
                        if not len(resFiles)==0: os.remove((resFiles[0]).replace("/","\\"))
                    except OSError:
                        logging.warning("delete went bad on the Result_Collection file")
                        msg= traceback.format_exc()
                        traceback.print_exc(file=sys.stdout)
                        logging.error(msg)
                        
                            
                    # 1. convert_data_to_csv
                    logging.info("Processing started, please wait...")
                    num=1
                    for tdmsFilePath in tdmsFiles: 


                        startTimeLoadFile = time.time()

                        tdmsFileName=tdmsFilePath.rsplit('\\')[-1]
                        tdmsFilePath=tdmsFilePath.replace("/","\\")

                        logging.info(" ")
                        logging.info("-------------------")
                        logging.info("File " + str(num)+ "/" + str(len(tdmsFiles)) + ". Reading TDMS: " + tdmsFileName)


                        with TdmsFile.read(tdmsFilePath, memmap_dir=os.getcwd()) as tdms_file:
                            

                            csvFilepath=self.tdms_excel.convert_data_to_csv(featureName,self.selectedDir,tdmsFileName, tdms_file)

                            logging.info(" CSV File created in " + str(round(time.time()-startTimeLoadFile,1)) + "s : " + tdmsFileName.split(".tdms")[0] + "--" + featureName + ".txt ")
                            QtWidgets.QApplication.processEvents()

                            excelDestPath=self.selectedDir + "/"+ featureName +  "--" + tdmsFileName.split(".tdms")[0]  + ".xlsm"
                            
                            
                            exceldataDestPath=self.tdms_excel.copy_template_excel_file(excelDestPath, excelTemplateFilePath.replace("/","\\"))
                            
                            # Process events between short sleep periods
                            QtWidgets.QApplication.processEvents()
                            # time.sleep(0.1)
                            logging.info(str(num)+ "/" + str(len(tdmsFiles)) + ": Create Excel file...")
                            
                            startTimeLoadFile = time.time()

                            data_from_csv = self.tdms_excel.get_csv_data(csvFilepath)

                            # delete the csv file (if it exists)
                            os.remove(csvFilepath)

                            result_dict=self.tdms_excel.write_data_to_excel_template(exceldataDestPath, data_from_csv,featureName,tdms_file)

                            logging.info("...done in "+ str(round(time.time()-startTimeLoadFile,1)) +"s : Filename: " + exceldataDestPath)

                            num=num+1

                    
                    # 2. run the result collection 

                    logging.info(" ")
                    logging.info("==============================")

                    excelDestPath=(self.selectedDir + "/"+ "Result_Collection" +  "--" + mst_name + " -- "  + featureName + ".xlsm").replace("/","\\")
                    
                    excelresultDestPath=self.tdms_excel.copy_template_excel_file(excelDestPath,(self.workingDir + Const.EXCEL_TEMPLATEFOLDER + '/' + Const.EXCEL_RESULT_FILENAME).replace("/","\\"))
                    
                    logging.info("Result Generation started. File: " + excelresultDestPath)       
                    self.tdms_excel.write_result_to_excel_template(excelresultDestPath)
                    # logging.info("Starting Excel Macro Template execution. Please wait, this could take some time..")
                    # self.tdms_excel.run_excel_macro(self.selectedDir)
                    

                    logging.info("FINISHED and DONE.")

                except InvalidFilePathLengthException:
                    msg="Maximum file path length was exceeded. "
                    traceback.print_exc(file=sys.stdout)

                    self.launch_btn.setEnabled(True)
                    logging.error(msg)

                except:
                    # exc_type, exc_obj, exc_tb = sys.exc_info()
                    # fname = os.path.split(exc_tb.tb_frame.f_code.co_filename)[1]
                    # print(exc_type, fname, exc_tb.tb_lineno)
                    # logging.error(str(exc_type) + ", File: " + str(fname) + ", Line: " + str(exc_tb.tb_lineno))
                    
                    msg= traceback.format_exc()
                    traceback.print_exc(file=sys.stdout)

                    self.launch_btn.setEnabled(True)
                    logging.error(msg)

                    return    

        else: self.selectedDir=None

        self.launch_btn.setEnabled(True)


if __name__ == '__main__':
    app = QApplication(sys.argv)
    # app.setStyleSheet('''
    #     QWidget {
    #         font-size: 35px;
    #     }
    # ''')
    
    csiAuto = CSI_AUTOMATOR()
    csiAuto.show()

    try:
        sys.exit(app.exec_())
    except SystemExit:
        print('Closing Window...')
