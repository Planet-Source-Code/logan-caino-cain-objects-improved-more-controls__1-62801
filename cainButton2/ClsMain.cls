VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public Enum MyOrientation
    pbHorizontal = 1
    pbVertical = 2
End Enum

Public Enum eItemType
    
    itNormal = 1
    itCheck = 2
    itPlaceholder = 3
    itTitle = 4

End Enum

Public Enum eMenuItemType
    
    mitNormalButton = 1
    mitCheckButton = 2
    mitMenu = 3
    mitMenu2 = 4
    mitPlaceholder = 5
    mitSeparator = 6
    mitCaption = 7

End Enum

Public Enum eSortDirection

    sd_Ascending = 0
    sd_Descending = 1

End Enum

Public Enum eComboType

    ctDateBox = 1
    ctColorPicker = 2
    ctPrinterSelect = 3
    ctDriveSelect = 4
    ctFontSelect = 5

End Enum
