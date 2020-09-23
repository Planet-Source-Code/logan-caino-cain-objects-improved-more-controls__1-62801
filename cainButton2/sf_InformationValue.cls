VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "sf_InformationValue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Option Explicit

'lokale Variable(n) zum Zuweisen der Eigenschaft(en)
Private mvarPath As String 'lokale Kopie
Private mvarFilename As String 'lokale Kopie
Private mvarDateTime As String 'lokale Kopie
Private mvarAttrib As VbFileAttribute 'lokale Kopie
Private mvarSize As String 'lokale Kopie

Public Property Let Size(ByVal vData As String)
'wird beim Zuweisen eines Werts zu der Eigenschaft auf der linken Seite einer Zuweisung verwendet.
'Syntax: X.Size = 5
    mvarSize = vData
End Property


Public Property Get Size() As String
'wird beim Ermitteln eines Eigenschaftswertes auf der rechten Seite einer Zuweisung verwendet.
'Syntax: Debug.Print X.Size
    Size = mvarSize
End Property

Public Property Let Attrib(ByVal vData As VbFileAttribute)
'wird beim Zuweisen eines Objekts zu der Eigenschaft auf der linken Seite einer Set-Anweisung verwendet.
'Syntax: Set x.Attrib = Form1
    mvarAttrib = vData
End Property


Public Property Get Attri() As VbFileAttribute
'wird beim Ermitteln eines Eigenschaftswertes auf der rechten Seite einer Zuweisung verwendet.
'Syntax: Debug.Print X.Attrib
    Attri = mvarAttrib
End Property



Public Property Let DateTime(ByVal vData As String)
'wird beim Zuweisen eines Werts zu der Eigenschaft auf der linken Seite einer Zuweisung verwendet.
'Syntax: X.DateTime = 5
    mvarDateTime = vData
End Property


Public Property Get DateTime() As String
'wird beim Ermitteln eines Eigenschaftswertes auf der rechten Seite einer Zuweisung verwendet.
'Syntax: Debug.Print X.DateTime
    DateTime = mvarDateTime
End Property



Public Property Let Filename(ByVal vData As String)
'wird beim Zuweisen eines Werts zu der Eigenschaft auf der linken Seite einer Zuweisung verwendet.
'Syntax: X.Filename = 5
    mvarFilename = vData
End Property


Public Property Get Filename() As String
Attribute Filename.VB_UserMemId = 0
'wird beim Ermitteln eines Eigenschaftswertes auf der rechten Seite einer Zuweisung verwendet.
'Syntax: Debug.Print X.Filename
    Filename = mvarFilename
End Property



Public Property Let Path(ByVal vData As String)
'wird beim Zuweisen eines Werts zu der Eigenschaft auf der linken Seite einer Zuweisung verwendet.
'Syntax: X.Path = 5
    mvarPath = vData
End Property


Public Property Get Path() As String
'wird beim Ermitteln eines Eigenschaftswertes auf der rechten Seite einer Zuweisung verwendet.
'Syntax: Debug.Print X.Path
    Path = mvarPath
End Property



