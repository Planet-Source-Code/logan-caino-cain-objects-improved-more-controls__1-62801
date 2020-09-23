VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ComboData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Option Explicit

'lokale Variable(n) zum Zuweisen der Eigenschaft(en)
Private mvarCaption As String 'lokale Kopie
Private mvarValue As String 'lokale Kopie
Private mvarIcon As Integer 'lokale Kopie
'lokale Variable(n) zum Zuweisen der Eigenschaft(en)
Private mvarIndex As Integer 'lokale Kopie
'lokale Variable(n) zum Zuweisen der Eigenschaft(en)
Private mvarComment As String 'lokale Kopie
Private mvarKey As String 'lokale Kopie

Public Property Let Comment(ByVal vData As String)
'wird beim Zuweisen eines Werts zu der Eigenschaft auf der linken Seite einer Zuweisung verwendet.
'Syntax: X.Comment = 5
    mvarComment = vData
End Property


Public Property Get Comment() As String
'wird beim Ermitteln eines Eigenschaftswertes auf der rechten Seite einer Zuweisung verwendet.
'Syntax: Debug.Print X.Comment
    Comment = mvarComment
End Property

Public Property Let Key(ByVal vKey As String)
'wird beim Zuweisen eines Werts zu der Eigenschaft auf der linken Seite einer Zuweisung verwendet.
'Syntax: X.Comment = 5
    mvarKey = vKey
End Property


Public Property Get Key() As String
'wird beim Ermitteln eines Eigenschaftswertes auf der rechten Seite einer Zuweisung verwendet.
'Syntax: Debug.Print X.Comment
    Key = mvarKey
End Property


Public Property Let Index(ByVal vData As Integer)
'wird beim Zuweisen eines Werts zu der Eigenschaft auf der linken Seite einer Zuweisung verwendet.
'Syntax: X.Index = 5
    mvarIndex = vData
End Property


Public Property Get Index() As Integer
'wird beim Ermitteln eines Eigenschaftswertes auf der rechten Seite einer Zuweisung verwendet.
'Syntax: Debug.Print X.Index
    Index = mvarIndex
End Property



Public Property Let Icon(ByVal vData As Integer)
'wird beim Zuweisen eines Werts zu der Eigenschaft auf der linken Seite einer Zuweisung verwendet.
'Syntax: X.Icon = 5
    mvarIcon = vData
End Property


Public Property Get Icon() As Integer
'wird beim Ermitteln eines Eigenschaftswertes auf der rechten Seite einer Zuweisung verwendet.
'Syntax: Debug.Print X.Icon
    Icon = mvarIcon
End Property



Public Property Let Value(ByVal vData As String)
'wird beim Zuweisen eines Werts zu der Eigenschaft auf der linken Seite einer Zuweisung verwendet.
'Syntax: X.Value = 5
    mvarValue = vData
End Property


Public Property Get Value() As String
'wird beim Ermitteln eines Eigenschaftswertes auf der rechten Seite einer Zuweisung verwendet.
'Syntax: Debug.Print X.Value
    Value = mvarValue
End Property



Public Property Let Caption(ByVal vData As String)
'wird beim Zuweisen eines Werts zu der Eigenschaft auf der linken Seite einer Zuweisung verwendet.
'Syntax: X.Caption = 5
    mvarCaption = vData
End Property


Public Property Get Caption() As String
'wird beim Ermitteln eines Eigenschaftswertes auf der rechten Seite einer Zuweisung verwendet.
'Syntax: Debug.Print X.Caption
    Caption = mvarCaption
End Property




