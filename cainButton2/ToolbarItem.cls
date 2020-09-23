VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ToolbarItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Option Explicit

'lokale Variable(n) zum Zuweisen der Eigenschaft(en)
Private mvarCaption As String 'lokale Kopie
Private mvarIconIndex As Integer 'lokale Kopie
Private mvarChecked As Boolean  'lokale Kopie
Private mvarEnabled As Boolean  'lokale Kopie
Private mvarStyle As eMenuItemType 'lokale Kopie
Private mvarKey As String 'lokale Kopie
Public ToolbarMenuItems As MenuItems

Public Property Let Style(ByVal vData As eMenuItemType)
'wird beim Zuweisen eines Werts zu der Eigenschaft auf der linken Seite einer Zuweisung verwendet.
'Syntax: X.Style = 5
    mvarStyle = vData
End Property

Public Property Get Style() As eMenuItemType
'wird beim Ermitteln eines Eigenschaftswertes auf der rechten Seite einer Zuweisung verwendet.
'Syntax: Debug.Print X.Style
    Style = mvarStyle
End Property

Public Property Let IconIndex(ByVal vData As Integer)
'wird beim Zuweisen eines Werts zu der Eigenschaft auf der linken Seite einer Zuweisung verwendet.
'Syntax: X.IconIndex = 5
    mvarIconIndex = vData
End Property


Public Property Get IconIndex() As Integer
'wird beim Ermitteln eines Eigenschaftswertes auf der rechten Seite einer Zuweisung verwendet.
'Syntax: Debug.Print X.IconIndex
    IconIndex = mvarIconIndex
End Property

Public Property Let Checked(ByVal vData As Boolean)
'wird beim Zuweisen eines Werts zu der Eigenschaft auf der linken Seite einer Zuweisung verwendet.
'Syntax: X.IconIndex = 5
    mvarChecked = vData
End Property


Public Property Get Checked() As Boolean
'wird beim Ermitteln eines Eigenschaftswertes auf der rechten Seite einer Zuweisung verwendet.
'Syntax: Debug.Print X.IconIndex
    Checked = mvarChecked
End Property

Public Property Let Enabled(ByVal vData As Boolean)
'wird beim Zuweisen eines Werts zu der Eigenschaft auf der linken Seite einer Zuweisung verwendet.
'Syntax: X.IconIndex = 5
    mvarEnabled = vData
End Property


Public Property Get Enabled() As Boolean
'wird beim Ermitteln eines Eigenschaftswertes auf der rechten Seite einer Zuweisung verwendet.
'Syntax: Debug.Print X.IconIndex
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

Public Property Let skey(ByVal vData As String)
'wird beim Zuweisen eines Werts zu der Eigenschaft auf der linken Seite einer Zuweisung verwendet.
'Syntax: X.Caption = 5
    mvarKey = vData
End Property


Public Property Get skey() As String
'wird beim Ermitteln eines Eigenschaftswertes auf der rechten Seite einer Zuweisung verwendet.
'Syntax: Debug.Print X.Caption
    skey = mvarKey
End Property

Private Sub Class_Initialize()
    Set ToolbarMenuItems = New MenuItems
End Sub

Private Sub Class_Terminate()
    If Not (ToolbarMenuItems Is Nothing) Then Set ToolbarMenuItems = Nothing
End Sub

