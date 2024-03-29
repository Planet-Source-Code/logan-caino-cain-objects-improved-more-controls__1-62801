VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MenuItemListings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"MenuItemListing"
Attribute VB_Ext_KEY = "Member0" ,"MenuItemListing"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'Lokale Variable für eine Auflistung.
Private mCol As Collection

Public Function Add(Key As String, Index As Integer, Top As Long, Height As Long, Selectable As Boolean) As MenuItemListing
    'neues Objekt erstellen
    Dim objNewMember As MenuItemListing
    Set objNewMember = New MenuItemListing


    'die Eigenschaften festlegen, die der Methode übergeben werden
    objNewMember.Index = Index
    objNewMember.Top = Top
    objNewMember.Key = Key
    objNewMember.Height = Height
    objNewMember.Selectable = Selectable
    If Len(Key) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, Key
    End If


    'das erstellte Objekt zurückgeben
    Set Add = objNewMember
    Set objNewMember = Nothing


End Function

Public Property Get Item(vntIndexKey As Variant) As MenuItemListing
Attribute Item.VB_UserMemId = 0
    'Wird verwendet, wenn auf ein Element der Auflistung zugegriffen
    'wird. vntIndexKey enthält entweder den Index oder den Schlüssel.
    'Daher wird der Parameter als Variant deklariert.
    'Syntax: Set foo = x.Item(xyz) or Set foo = x.Item(5)
  Set Item = mCol(vntIndexKey)
End Property


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
    
End Sub

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

