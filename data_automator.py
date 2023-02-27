import sys
import os
import types



from TDMS_EXCEL import TDMS_EXCEL

from PyQt5.QtWidgets import QApplication, QWidget, QComboBox, QPushButton, QFileDialog, QVBoxLayout,QHBoxLayout, QLabel,QGridLayout,QRadioButton,QSpacerItem, QSizePolicy
from PyQt5 import QtCore, QtGui, QtWidgets
from nptdms import TdmsFile
import glob


from PyQt5.QtCore import *
from PyQt5.QtGui import *

import logging
import time
from Util import Const
from Util import InvalidFilePathLengthException


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



        self.tdms_excel=TDMS_EXCEL()
        self.functions = [self.procTDMSDataforCSI, self.procEmpty]


        self.excelTemplateFilesPath=None
        self.radioButtons=[]
        self.optionsLayout = QHBoxLayout()

        self.window_width, self.window_height = 1000, 200
        self.setMinimumSize(self.window_width, self.window_height)

        self.logTextBox = QTextEditLogger(self)
        # You can format what is printed to text box
        self.logTextBox.setFormatter(logging.Formatter('%(asctime)s - %(levelname)s - %(message)s'))
        logging.getLogger().addHandler(self.logTextBox)
        # You can control the logging level
        logging.getLogger().setLevel(logging.INFO)
        
        # timer = QtCore.QTimer(self)
        # timer.timeout.connect(self.doTimer)
        # timer.start(1000)
        
        
        # font
        smallFont = QFont('Arial', 14)
        bigFont = QFont('Arial', 20)

        self.selectedDir=None
        # self.resize(737, 596)

        self.setWindowTitle("Data Automator V1.0.2")
        self.setWindowIcon(QtGui.QIcon("icon.png"))

        layout = QVBoxLayout()
        self.setLayout(layout)

        self.options = (Const.OPTIONS)

        self.label=QLabel("Choose action:")
        self.label.setFont(smallFont)
        layout.addWidget(self.label)
        
        self.combobox = QComboBox()
        self.combobox.currentIndexChanged.connect(self.selectionchange)
        

        # adding action to combo box
        
        self.combobox.addItems(self.options)
        self.combobox.setFont(smallFont)

        # self.combobox.setCurrentIndex(0) # trigger the signal

        layout.addWidget(self.combobox)



        self.btn = QPushButton('Launch')
        self.btn.setFont(bigFont)
        self.btn.clicked.connect(self.launchButton)
        self.btn.setEnabled(False)
        

        layout.addLayout(self.optionsLayout)

        layout.addWidget(self.btn)

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
            self.btn.setEnabled(True)
            print("feature is %s" % (radioButton.feature))
        

    def selectionchange(self,i):


        if self.combobox.currentText()==self.tdms_excel.title:
           

            # hSpacer = QSpacerItem(350, 40, QSizePolicy.Maximum, QSizePolicy.Expanding)
            # self.optionsLayout.addItem(hSpacer)

            xlsxTemplateFiles = glob.glob(os.getcwd() + '/' + Const.EXCEL_TEMPLATEFOLDER + r'/*.xlsx')

            self.excelTemplateFilesPath = [item for item in xlsxTemplateFiles if not "~" in item]
            
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


    def launchButton(self):
        curIndex = self.options.index(self.combobox.currentText())
        response = self.functions[curIndex]()
        


    def procEmpty(self):
        print('Got Nothing')

        
    def procTDMSDataforCSI(self):

        dialog = QFileDialog(self, 'Select a folder containing TDMS files', os.getcwd())


        dialog.setFileMode(QFileDialog.DirectoryOnly)
        dialog.setOption(QFileDialog.ShowDirsOnly, False)

        # dialog.setFileMode(QFileDialog.DirectoryOnly)
        # dialog.selectNameFilter("TDMS Files (*.tdms)")

        if dialog.exec_() == QFileDialog.Accepted:
            self.selectedDir = dialog.selectedFiles()[0]


            # self.selectedDir="F:/ENTWICKLUNG/SSC019 CSI 3 Modul/04 Erprobungen/02 Interne Prüfberichte/MST2022081100_SSC018-SSC019_DV2_CSI3_KompressorModul/Ergebnisse/L1 - Lebensdauerprüfung/12V/Thomas DV2 tdms/Run1"
    
            logging.info("Selected Dir: " + self.selectedDir)

            # search for all TDMS files within selected directory
            tdmsFiles = glob.glob(self.selectedDir + r'/*.tdms')
     

            logging.info("TDMS found:")
            for tdmsFile in tdmsFiles:
                logging.info(str(tdmsFile))

            logging.info(str(len(tdmsFiles)) + " TDMS File(s) found.")    

                
            logging.info("EXCEL Templates found in: " + str(os.getcwd()))
            for excelTemplateFile in self.excelTemplateFilesPath:
                logging.info(str(excelTemplateFile))

            logging.info(str(len(self.excelTemplateFilesPath)) + " EXCEL Templates File(s) found.")   
            
            featureSelected=""
            for rb in self.radioButtons:
                if rb.isChecked(): 
                    featureSelected=rb.feature
                    logging.info("Template for " + str(rb.feature) + " will be processed.")

            if featureSelected=="": 
                logging.warning("No Template was selected. Please select one.")
                return

            # for item in self.excelTemplateFilesPath:
            #     if  rb.feature in item: 
            
            # excelTemplateFilePath=(item for item in self.excelTemplateFilesPath: rb.feature in item)

            excelTemplateFilePath = [x for x in self.excelTemplateFilesPath if featureSelected in x]
            excelTemplateFilePath=excelTemplateFilePath[0]
            # if all conditions are met: do the process
            
             
            try:

                # 1. convert_to_csv
                for tdmsFile in tdmsFiles: 

                    startTimeLoadFile = time.time()
                    logging.info("Processing started, please wait...")

                    with TdmsFile.read(tdmsFile, memmap_dir=os.getcwd()) as tdms_file:
                        endTimeLoadTime = time.time()
                            
                        tdmsFileName=tdmsFile.rsplit('\\')[-1]
                        self.tdms_excel.convert_to_csv(featureSelected,self.selectedDir,tdmsFileName, tdms_file, excelTemplateFilePath)

                        logging.info("CSV File created in "+str(round(time.time()-startTimeLoadFile,1)) +"s : " + tdmsFileName.split(".tdms")[0] + "--" + featureSelected + ".txt ")
        
                        # Process events between short sleep periods
                        QtWidgets.QApplication.processEvents()
                        # time.sleep(0.1)
                # 2. run the excel macro

                logging.info("Starting Excel Macro Template execution. Please wait, this could take some time..")
                self.tdms_excel.run_excel_macro(self.selectedDir)
                

                logging.info("FINISHED and DONE.")

            except InvalidFilePathLengthException:
                msg="Maximum file path length (259 characters) was exceeded. "
                return

            except:
                exc_type, exc_obj, exc_tb = sys.exc_info()
                fname = os.path.split(exc_tb.tb_frame.f_code.co_filename)[1]
                print(exc_type, fname, exc_tb.tb_lineno)
                logging.error(str(exc_type) + ", File: " + str(fname) + ", Line: " + str(exc_tb.tb_lineno))
                return    

        else: self.selectedDir=None




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