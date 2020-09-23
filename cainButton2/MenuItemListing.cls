VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MenuItemListing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Option Explicit

'lokale Variable(n) zum Zuweisen der Eigenschaft(en)
Private mvarIndex As Integer 'lokale Kopie
Private mvarTop As Long 'lokale Kopie
Private mvarHeight As Long 'lokale Kopie
'lokale Variable(n) zum Zuweisen der Eigenschaft(en)
Private mvarSelectable As Boolean 'lokale Kopie
Private mvarKey As String 'lokale Kopie

Public Property Let Key(ByVal vData As String)
'wird beim Zuweisen eines Werts zu der Eigenschaft auf der linken Seite einer Zuweisung verwendet.
'Syntax: X.Caption = 5
    mvarKey = vData
End Property


Public Property Get Key() As String
'wird beim Ermitteln eines Eigenschaftswertes auf der rechten Seite einer Zuweisung verwendet.
'Syntax: Debug.Print X.Caption
    Key = mvarKey
End Property


Public Property Let Selectable(ByVal vData As Boolean)
'wird beim Zuweisen eines Werts zu der Eigenschaft auf der linken Seite einer Zuweisung verwendet.
'Syntax: X.Selectable = 5
    mvarSelectable = vData
End Property


Public Property Get Selectable() As Boolean
'wird beim Ermitteln eines Eigenschaftswertes auf der rechten Seite einer Zuweisung verwendet.
'Syntax: Debug.Print X.Selectable
    Selectable = mvarSelectable
End Property



Public Property Let Height(ByVal vData As Long)
'wird beim Zuweisen eines Werts zu der Eigenschaft auf der linken Seite einer Zuweisung verwendet.
'Syntax: X.Height = 5
    mvarHeight = vData
End Property


Public Property Get Height() As Long
'wird beim Ermitteln eines Eigenschaftswertes auf der rechten Seite einer Zuweisung verwendet.
'Syntax: Debug.Print X.Height
    Height = mvarHeight
End Property



Public Property Let Top(ByVal vData As Long)
'wird beim Zuweisen eines Werts zu der Eigenschaft auf der linken Seite einer Zuweisung verwendet.
'Syntax: X.Top = 5
    mvarTop = vData
End Property


Public Property Get Top() As Long
'wird beim Ermitteln eines Eigenschaftswertes auf der rechten Seite einer Zuweisung verwendet.
'Syntax: Debug.Print X.Top
    Top = mvarTop
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



