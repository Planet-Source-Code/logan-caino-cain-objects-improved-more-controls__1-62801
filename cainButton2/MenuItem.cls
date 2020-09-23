VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MenuItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Option Explicit

'lokale Variable(n) zum Zuweisen der Eigenschaft(en)
Private mvarIcon As Integer 'lokale Kopie
Private mvarCaption As String 'lokale Kopie
Private mvarEnabled As Boolean 'lokale Kopie
Private mvarChild As String 'lokale Kopie
Private mvarGroup As String 'lokale Kopie
Private mvarItemType As eItemType   'lokale Kopie
Private mvarChecked As Boolean 'lokale Kopie
Private mvarBold As Boolean 'lokale Kopie
Private mvarKey As String 'lokale Kopie

Public Property Let Bold(ByVal vData As Boolean)
'wird beim Zuweisen eines Werts zu der Eigenschaft auf der linken Seite einer Zuweisung verwendet.
'Syntax: X.Bold = 5
    mvarBold = vData
End Property


Public Property Get Bold() As Boolean
'wird beim Ermitteln eines Eigenschaftswertes auf der rechten Seite einer Zuweisung verwendet.
'Syntax: Debug.Print X.Bold
    Bold = mvarBold
End Property



Public Property Let Checked(ByVal vData As Boolean)
'wird beim Zuweisen eines Werts zu der Eigenschaft auf der linken Seite einer Zuweisung verwendet.
'Syntax: X.Checked = 5
    mvarChecked = vData
End Property


Public Property Get Checked() As Boolean
'wird beim Ermitteln eines Eigenschaftswertes auf der rechten Seite einer Zuweisung verwendet.
'Syntax: Debug.Print X.Checked
    Checked = mvarChecked
End Property



Public Property Let ItemType(ByVal vData As eItemType)
'wird beim Zuweisen eines Werts zu der Eigenschaft auf der linken Seite einer Zuweisung verwendet.
'Syntax: X.ItemType = 5
    mvarItemType = vData
End Property


Public Property Get ItemType() As eItemType
'wird beim Ermitteln eines Eigenschaftswertes auf der rechten Seite einer Zuweisung verwendet.
'Syntax: Debug.Print X.ItemType
    ItemType = mvarItemType
End Property



Public Property Let Group(ByVal vData As String)
'wird beim Zuweisen eines Werts zu der Eigenschaft auf der linken Seite einer Zuweisung verwendet.
'Syntax: X.Group = 5
    mvarGroup = vData
End Property


Public Property Get Group() As String
'wird beim Ermitteln eines Eigenschaftswertes auf der rechten Seite einer Zuweisung verwendet.
'Syntax: Debug.Print X.Group
    Group = mvarGroup
End Property



Public Property Let Child(ByVal vData As String)
'wird beim Zuweisen eines Werts zu der Eigenschaft auf der linken Seite einer Zuweisung verwendet.
'Syntax: X.Child = 5
    mvarChild = vData
End Property


Public Property Get Child() As String
'wird beim Ermitteln eines Eigenschaftswertes auf der rechten Seite einer Zuweisung verwendet.
'Syntax: Debug.Print X.Child
    Child = mvarChild
End Property



Public Property Let Enabled(ByVal vData As Boolean)
'wird beim Zuweisen eines Werts zu der Eigenschaft auf der linken Seite einer Zuweisung verwendet.
'Syntax: X.Enabled = 5
    mvarEnabled = vData
End Property


Public Property Get Enabled() As Boolean
'wird beim Ermitteln eines Eigenschaftswertes auf der rechten Seite einer Zuweisung verwendet.
'Syntax: Debug.Print X.Enabled
    Enabled = mvarEnabled
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



