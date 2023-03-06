from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtWidgets import QFileDialog, QWidget
import sys,os

class Const():
    COLORS=[
    "#FF0000",#"red"
    "#006400",#,"darkgreen"
    "#0000CD",#,"mediumblue"
    "#000080",#,"navy" 
    "#FF8C00",#,"darkorange"
    "#8B0000",#,"darkred"
    "#B0E0E6",#,"powderblue"
    "#FFD700",#,"gold"
    "#FF00FF",#,"fuchsia"
    "#800080",#,"purple"
    "#D2691E",#,"chocolate"
    "#FFE4C4",#,"bisque"
    "#DEB887",#,"burlywood"
    "#BC8F8F",#,"rosybrown"
    "#F4A460",#,"sandybrown"
    "#00FF00",#,"lime" 
    "#8B4513",#,"saddlebrown"
    "#4169E1",#,"royalblue"
    "#BDB76B",#,"darkkhaki"
    "#00BFFF",#"deepskyblue":
    "#FF7F50",#,"coral"
    "#00FFFF",#,"aqua"
    "#FF4500",#,"orangered"
    "#FFA500",#,"orange"
    "#FFA07A",#,"lightsalmon"
    "#C0C0C0",#,"silver"
    "#808080",#,"gray"
    "#A52A2A",#"brown"
    "#808000",#,"olive"
    "#CD5C5C",#,"indianred"
    "#00FA9A",#,"mediumspringgreen"
    "#1E90FF",#,"dodgerblue"
    ]

    
    OPTIONS = ['Process TDMS Data for CSI','foobar']

    # Groups and Channels in TDMS files for CSI:
    TDMS_CSI_DATA_DISCR={"Analog-IO":["Strom_Ebene_1","Spannung_Ebene_1","Druckspeicher_1","Maximaldruck_Leitung","Drehzahl","C","D","E","F","G","H","I","J","K","L","M"],
                         "TC":["TC_Zylinder","TC_Motor","TC_Gehaeuse","TC_Druckspeicher"]}

    MAX_PATHLENGTH_DOS=259

    TDMS_LIST_SEP=" --> "

    EXCEL_SHEETNAME="Source"

    EXCEL_TEMPLATEFOLDER="TDMS_CSI_ExcelTemplates"




class Functions:
    def display_properties(tdms_object, level):
        if tdms_object.properties:
            print("properties:", level)
            for prop, val in tdms_object.properties.items():
                print("%s: %s" % (prop, val), level)

# define Python user-defined exceptions
class InvalidFilePathLengthException(Exception):
    "Raised when the max file path length is exceeded"
    pass



