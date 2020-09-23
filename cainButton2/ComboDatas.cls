VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ComboDatas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"ComboData"
Attribute VB_Ext_KEY = "Member0" ,"ComboData"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'Lokale Variable für eine Auflistung.
Private mCol As Collection

Public Event ItemAdded()
Public Event ItemChanged()

'lokale Variable(n) zum Zuweisen der Eigenschaft(en)
Private mvarSorted As Boolean 'lokale Kopie
'lokale Variable(n) zum Zuweisen der Eigenschaft(en)
Private mvarSorting As eSortDirection 'lokale Kopie

Public Property Let Sorting(ByVal vData As eSortDirection)
'wird beim Zuweisen eines Objekts zu der Eigenschaft auf der linken Seite einer Set-Anweisung verwendet.
'Syntax: Set x.Sorting = Form1
    mvarSorting = vData
 End Property


Public Property Get Sorting() As eSortDirection
'wird beim Ermitteln eines Eigenschaftswertes auf der rechten Seite einer Zuweisung verwendet.
'Syntax: Debug.Print X.Sorting
    Sorting = mvarSorting
End Property

Public Property Let Sorted(ByVal vData As Boolean)
'wird beim Zuweisen eines Werts zu der Eigenschaft auf der linken Seite einer Zuweisung verwendet.
'Syntax: X.Sorted = 5
    mvarSorted = vData
End Property


Public Property Get Sorted() As Boolean
'wird beim Ermitteln eines Eigenschaftswertes auf der rechten Seite einer Zuweisung verwendet.
'Syntax: Debug.Print X.Sorted
    Sorted = mvarSorted
End Property

Public Function Add(Key As String, Optional Caption As String = "", Optional Value As String = "", Optional Icon As Integer = 0, Optional sComment As String = "") As ComboData
    'neues Objekt erstellen
    Dim objNewMember As ComboData
    Set objNewMember = New ComboData
    
    Dim i As Integer
    Dim iRes As Integer
    Dim iPos As Integer


    'die Eigenschaften festlegen, die der Methode übergeben werden
    objNewMember.Key = Key
    objNewMember.Caption = Caption
    objNewMember.Value = Value
    objNewMember.Icon = Icon
    objNewMember.Comment = sComment
    
    If mCol.Count <> 0 And mvarSorted = True Then
        
        For i = 1 To mCol.Count
            
            If mvarSorting = sd_Ascending Then
                iRes = StrComp(mCol(mCol.Count - i + 1).Caption, Caption, vbTextCompare)
                
                '-1 = before
                '0 = equal
                '1 = after
                
                If iRes = -1 Then
                    Exit For
                ElseIf iRes = 0 Then
                    Exit For
                ElseIf iRes = 1 Then
                    'do nothing
                End If
                
            ElseIf mvarSorting = sd_Descending Then
                iRes = StrComp(mCol(i).Caption, Caption, vbTextCompare)
                
                If iRes = -1 Then
                    Exit For
                ElseIf iRes = 0 Then
                    Exit For
                ElseIf iRes = 1 Then
                    'do nothing
                End If
            
            End If
            
        Next i
        
        'debug.Print mCol.Count
        
    ElseIf mvarSorted = True Then
        i = 1
    End If
    
    'ipos = 0 append
    'ipos = 1 before
    'ipos = 2 after
    
    If mvarSorted = False Or mCol.Count = 0 Then
        iPos = 0

    ElseIf mvarSorting = sd_Ascending And i > mCol.Count Then
        iPos = 1
        i = 1
    
    ElseIf mvarSorting = sd_Descending And i > mCol.Count Then
        iPos = 0
        'i = mCol.Count
    
    ElseIf mvarSorting = sd_Ascending Then
        iPos = 2
        i = mCol.Count - i + 1
    
    ElseIf mvarSorting = sd_Descending Then
        iPos = 1

    End If
    
    If iPos = 0 Then
        If Len(Key) = 0 Then
            mCol.Add objNewMember
        Else
            mCol.Add objNewMember, Key
        End If
        
    ElseIf iPos = 1 Then
        If Len(Key) = 0 Then
            mCol.Add objNewMember, , i
        Else
            mCol.Add objNewMember, Key, i
        End If

    ElseIf iPos = 2 Then
        If Len(Key) = 0 Then
            mCol.Add objNewMember, , , i
        Else
            mCol.Add objNewMember, Key, , i
        End If

    End If

    'das erstellte Objekt zurückgeben
    Set Add = objNewMember
    Set objNewMember = Nothing
    
    RaiseEvent ItemAdded

End Function

Public Property Get Item(vntIndexKey As Variant) As ComboData
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
    RaiseEvent ItemChanged
End Sub

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


