VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ToolbarItems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"ToolbarItem"
Attribute VB_Ext_KEY = "Member0" ,"ToolbarItem"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'Lokale Variable für eine Auflistung.
Private mCol As Collection

Public Event ItemAdded()
Public Event ItemChanged()

Public Function Add(Key As String, Optional Caption As String = "", Optional IconIndex As Integer = 0, Optional Style As eMenuItemType = 1, Optional Checked As Boolean = False, Optional Enabled As Boolean = True) As ToolbarItem
    'neues Objekt erstellen
    Dim objNewMember As ToolbarItem
    Set objNewMember = New ToolbarItem


    'die Eigenschaften festlegen, die der Methode übergeben werden
    objNewMember.Caption = Caption
    objNewMember.IconIndex = IconIndex
    objNewMember.Style = Style
    objNewMember.skey = Key
    objNewMember.Checked = Checked
    objNewMember.Enabled = Enabled
    
    If Len(Key) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, Key
    End If

    'das erstellte Objekt zurückgeben
    Set Add = objNewMember
    Set objNewMember = Nothing
    
    RaiseEvent ItemAdded
    
End Function

Public Sub Clear()
    'Wird verwendet, wenn ein Element aus der Auflistung
    'entfernt werden soll. vntIndexKey enthält entweder
    'den Index oder den Schlüssel. Daher wird der Parameter
    'als Variant deklariert.
    'Syntax: x.Remove(xyz)
    
    Dim i
    
    For i = 1 To mCol.Count
        mCol.Remove 1
    Next i
    
    RaiseEvent ItemChanged
    
End Sub

Public Property Get Item(vntIndexKey As Variant) As ToolbarItem
    'Wird verwendet, wenn auf ein Element der Auflistung zugegriffen
    'wird. vntIndexKey enthält entweder den Index oder den Schlüssel.
    'Daher wird der Parameter als Variant deklariert.
    'Syntax: Set foo = x.Item(xyz) or Set foo = x.Item(5)
    Set Item = mCol(vntIndexKey)
    
End Property



Public Property Get Count() As Long
    'Wird verwendet, wenn die Anzahl der Elemente
    'in der Auflistung ermittelt werden soll.
    'Syntax: Debug.Print x.Count
    Count = mCol.Count
End Property

Public Sub Refresh()
    RaiseEvent ItemChanged
End Sub

Public Sub Remove(vntIndexKey As Variant)
    'Wird verwendet, wenn ein Element aus der Auflistung
    'entfernt werden soll. vntIndexKey enthält entweder
    'den Index oder den Schlüssel. Daher wird der Parameter
    'als Variant deklariert.
    'Syntax: x.Remove(xyz)


    mCol.Remove vntIndexKey
    
    RaiseEvent ItemChanged
    
End Sub


Public Property Get NewEnum() As IUnknown
    'Diese Eigenschaft ermöglicht, die Elemente der
    'Auflistung mit For...Each zu durchlaufen.
    Set NewEnum = mCol.[_NewEnum]
End Property


Private Sub Class_Initialize()
    'Erstellt die Auflistung, wenn die Klasse erstellt wird.
    Set mCol = New Collection
End Sub


Private Sub Class_Terminate()
    'Entfernt die Auflistung aus dem Speicher, wenn die Klasse beendet wird.
    Set mCol = Nothing
End Sub



