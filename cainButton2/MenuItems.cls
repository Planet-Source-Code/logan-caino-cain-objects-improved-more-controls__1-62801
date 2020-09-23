VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MenuItems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"MenuItem"
Attribute VB_Ext_KEY = "Member0" ,"MenuItem"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'Lokale Variable für eine Auflistung.
Private mCol As Collection

Public Function Add(Key As String, Optional Caption As String = "", Optional Icon As Integer = 0, Optional Enabled As Boolean = True, Optional Child As String = "", Optional Group As String = "", Optional ItemType As eItemType = 1, Optional Checked As Boolean = False, Optional Bold As Boolean = False) As MenuItem
    'neues Objekt erstellen
    Dim objNewMember As MenuItem
    Set objNewMember = New MenuItem


    'die Eigenschaften festlegen, die der Methode übergeben werden
    objNewMember.Icon = Icon
    objNewMember.Caption = Caption
    objNewMember.Enabled = Enabled
    objNewMember.Child = Child
    objNewMember.Group = Group
    objNewMember.ItemType = ItemType
    objNewMember.Checked = Checked
    objNewMember.Bold = Bold
    objNewMember.Key = Key
    If Len(Key) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, Key
    End If


    'das erstellte Objekt zurückgeben
    Set Add = objNewMember
    Set objNewMember = Nothing


End Function

Public Property Get Item(vntIndexKey As Variant) As MenuItem
Attribute Item.VB_UserMemId = 0
    If Len(vntIndexKey) = 0 Then Exit Property
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


Public Sub Remove(vntIndexKey As Variant)
    'Wird verwendet, wenn ein Element aus der Auflistung
    'entfernt werden soll. vntIndexKey enthält entweder
    'den Index oder den Schlüssel. Daher wird der Parameter
    'als Variant deklariert.
    'Syntax: x.Remove(xyz)


    mCol.Remove vntIndexKey
End Sub


Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
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

