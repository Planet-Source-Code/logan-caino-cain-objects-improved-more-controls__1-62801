VERSION 5.00
Begin VB.UserControl cainListbox 
   AutoRedraw      =   -1  'True
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   Begin VB.Timer MousePos 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   0
      Top             =   0
   End
   Begin cainObjects.cainScrollBar cainScrollBar1 
      Height          =   1575
      Left            =   4320
      Top             =   120
      Width           =   225
      _ExtentX        =   397
      _ExtentY        =   2778
      Enabled         =   0   'False
   End
End
Attribute VB_Name = "cainListbox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'Standard-Eigenschaftswerte:
Const m_def_HoverSelect = 0
Const m_def_SelectedItem = 0
Const m_def_BackColor = &HFFFFFF
Const m_def_ForeColor = &HC00000
Const m_def_SelectionColor = &H80FF&
'Eigenschaftsvariablen:
Dim m_HoverSelect As Boolean
Dim m_SelectedItem As Integer
Dim m_BackColor As OLE_COLOR
Dim m_ForeColor As OLE_COLOR
Dim m_SelectionColor As OLE_COLOR
Dim m_ImageList As Object
'Ereignisdeklarationen:
Event Click()
Attribute Click.VB_Description = "Tritt auf, wenn der Benutzer eine Maustaste über einem Objekt drückt und wieder losläßt."
Event DblClick()
Attribute DblClick.VB_Description = "Tritt auf, wenn der Benutzer eine Maustaste über einem Objekt drückt und wieder losläßt und anschließend erneut drückt und wieder losläßt."
Event KeyDown(KeyCode As Integer, Shift As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyDown
Attribute KeyDown.VB_Description = "Tritt auf, wenn der Benutzer eine Taste drückt, während ein Objekt den Fokus besitzt."
Event KeyPress(KeyAscii As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyPress
Attribute KeyPress.VB_Description = "Tritt auf, wenn der Benutzer eine ANSI-Taste drückt und losläßt."
Event KeyUp(KeyCode As Integer, Shift As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyUp
Attribute KeyUp.VB_Description = "Tritt auf, wenn der Benutzer eine Taste losläßt, während ein Objekt den Fokus hat."
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseDown
Attribute MouseDown.VB_Description = "Tritt auf, wenn der Benutzer die Maustaste drückt, während ein Objekt den Fokus hat."
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseMove
Attribute MouseMove.VB_Description = "Tritt auf, wenn der Benutzer die Maus bewegt."
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseUp
Attribute MouseUp.VB_Description = "Tritt auf, wenn der Benutzer die Maustaste losläßt, während ein Objekt den Fokus hat."
Event ItemClicked(iIndex As Integer, sKey As String)

Dim tColorSet As ColorSet
Public WithEvents ListboxItems As ComboDatas
Attribute ListboxItems.VB_VarHelpID = -1
Public TempListboxItems As MenuItemListings
Attribute TempListboxItems.VB_VarHelpID = -1

Dim MouseClick As Boolean
Dim iSelection As Integer
Dim bHasFocus As Boolean

Dim Mouse As POINTAPI
Dim MouseOverMe As POINTAPI

Public ValueFont As Boolean

Private Const ITEM_GAP = 4
Private Const TOP_GAP = 1

Public Function ID() As Integer
    ID = eControlIDs.id_Listbox
End Function

Private Sub cainScrollBar1_Click()
    On Error Resume Next
    
    UserControl.SetFocus
    
End Sub

Private Sub cainScrollBar1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub cainScrollBar1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub cainScrollBar1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Private Sub cainScrollBar1_ValueChanged()
    SetVisibleData
    DrawFace True
End Sub

Public Sub Refresh()
    ListboxItems_ItemAdded
End Sub

Private Sub ListboxItems_ItemAdded()
    
    If ListboxItems.Count <> 0 Then
        cainScrollBar1.Min = 1
        cainScrollBar1.Max = ListboxItems.Count - GetMaxSlot + 1
        cainScrollBar1.Value = 0
    End If
    
    If GetMaxSlot > ListboxItems.Count Then
        cainScrollBar1.Enabled = False
    Else
        cainScrollBar1.Enabled = True
    End If
    
    SetVisibleData
    
    If m_SelectedItem = 0 Then m_SelectedItem = 1
    
    DrawFace
    
End Sub

Private Sub DrawSelection(iIndex As Integer)
    
    If bHasFocus = True Or m_HoverSelect = True Then
        DrawGradient UserControl.hdc, 3, TempListboxItems(iIndex).Top + 2 + TOP_GAP, UserControl.ScaleWidth - cainScrollBar1.Width - 5, TempListboxItems(iIndex).Height - 2, GetRGBColors(tColorSet.csColor1(9)), GetRGBColors(tColorSet.csColor1(8)), 0
    Else
        DrawGradient UserControl.hdc, 3, TempListboxItems(iIndex).Top + 2 + TOP_GAP, UserControl.ScaleWidth - cainScrollBar1.Width - 7, TempListboxItems(iIndex).Height - 2, GetRGBColors(tColorSet.csColor1(1)), GetRGBColors(tColorSet.csColor1(3)), 0
    End If
    
    UserControl.Line (2, TempListboxItems(iIndex).Top + 1 + TOP_GAP)-(UserControl.ScaleWidth - cainScrollBar1.Width - 4, TempListboxItems(iIndex).Top + 1 + TOP_GAP), tColorSet.csColor1(6)
    UserControl.Line (2, TempListboxItems(iIndex).Top + 1 + TOP_GAP)-(2, TempListboxItems(iIndex).Top + TempListboxItems(iIndex).Height - 1 + TOP_GAP), tColorSet.csColor1(6)
    UserControl.Line (UserControl.ScaleWidth - cainScrollBar1.Width - 4, TempListboxItems(iIndex).Top + 1 + TOP_GAP)-(UserControl.ScaleWidth - cainScrollBar1.Width - 4, TempListboxItems(iIndex).Top + TempListboxItems(iIndex).Height + TOP_GAP), tColorSet.csColor1(6)
    UserControl.Line (2, TempListboxItems(iIndex).Top + TempListboxItems(iIndex).Height - 1 + TOP_GAP)-(UserControl.ScaleWidth - cainScrollBar1.Width - 4, TempListboxItems(iIndex).Top + TempListboxItems(iIndex).Height - 1 + TOP_GAP), tColorSet.csColor1(6)

End Sub

Private Sub RefreshColor()

    tColorSet = GetColorSetNormal(m_BackColor, m_ForeColor, m_SelectionColor)
    cainScrollBar1.BackColor = tColorSet.csBackColor
    cainScrollBar1.ForeColor = tColorSet.csFrontColor
    
    UserControl.BackColor = tColorSet.csBackColor

End Sub

Private Sub SetVisibleData()
    
    Dim i As Integer
    Dim iMaxSlot As Integer
    Dim tHeight As Long
    Dim tTop As Long
    
    TempListboxItems.Clear
    
    iMaxSlot = GetMaxSlot + cainScrollBar1.Value
    If iMaxSlot + cainScrollBar1.Value > ListboxItems.Count Then iMaxSlot = ListboxItems.Count
    
    If m_ImageList Is Nothing Then
        tHeight = UserControl.TextHeight("I") + ITEM_GAP
    Else
        If m_ImageList.ImageHeight > UserControl.TextHeight("I") Then
            tHeight = m_ImageList.ImageHeight + ITEM_GAP
        Else
            tHeight = UserControl.TextHeight("I") + ITEM_GAP
        End If
    
    End If
    
    tTop = 0
    
    For i = cainScrollBar1.Value To iMaxSlot
        TempListboxItems.Add "", i, tTop, tHeight, True
        tTop = tTop + tHeight
    Next i
    
End Sub

Private Function GetMaxSlot() As Integer
    
    If m_ImageList Is Nothing Then
        GetMaxSlot = Fix((UserControl.ScaleHeight - TOP_GAP) / (UserControl.TextHeight("I") + ITEM_GAP)) '+ 1
    Else
        If m_ImageList.ImageHeight > UserControl.TextHeight("I") Then
            GetMaxSlot = Fix((UserControl.ScaleHeight - TOP_GAP) / (m_ImageList.ImageHeight + ITEM_GAP + TOP_GAP)) ' + 1
        Else
            GetMaxSlot = Fix((UserControl.ScaleHeight - TOP_GAP) / (UserControl.TextHeight("I") + ITEM_GAP + TOP_GAP)) '+ 1
        End If
    End If
    
End Function

Private Sub DrawFace(Optional bSelected As Boolean = False)

    
    Dim i As Integer
    Dim strTmp As String
    Dim tmpX As Integer
    Dim Special_Gap As Integer
    Dim sFont As String
    
    UserControl.Cls
    
    If TempListboxItems.Count <> 0 Then
        If ListboxItems.Count = 0 Then Exit Sub
        
            
        If ValueFont = True Then _
            sFont = UserControl.FontName
        
        For i = 1 To TempListboxItems.Count
            strTmp = ListboxItems.Item(TempListboxItems.Item(i).Index).Caption
            Special_Gap = CountTabs(strTmp) * 9
           
            If TempListboxItems.Item(i).Index = iSelection Then DrawSelection i
            
            If m_ImageList Is Nothing Then
                tmpX = ITEM_GAP + Special_Gap
            Else
                tmpX = (ITEM_GAP * 2) + m_ImageList.ImageWidth + Special_Gap
                DrawImage i, Special_Gap
            End If
            
            If ValueFont = True Then UserControl.FontName = ListboxItems.Item(TempListboxItems.Item(i).Index).Value
             
            UserControl.CurrentX = tmpX
            UserControl.CurrentY = TempListboxItems.Item(i).Top + ((TempListboxItems.Item(i).Height / 2) - (UserControl.TextHeight("I") / 2)) + TOP_GAP
            
            If UserControl.Enabled = True Then
                UserControl.ForeColor = tColorSet.csColor1(7)
            Else
                UserControl.ForeColor = tColorSet.csColor1(6)
            End If
            UserControl.Print ShortWord(Replace(strTmp, vbTab, ""), UserControl.ScaleWidth - tmpX - cainScrollBar1.Width - ITEM_GAP)
            
        Next i
        
        If ValueFont = True Then UserControl.FontName = sFont
        
    End If

    UserControl.Line (0, 0)-(0, UserControl.ScaleHeight), tColorSet.csColor1(6)
    UserControl.Line (0, 0)-(UserControl.ScaleWidth, 0), tColorSet.csColor1(6)
    UserControl.Line (UserControl.ScaleWidth - 1, 0)-(UserControl.ScaleWidth - 1, UserControl.ScaleHeight), tColorSet.csColor1(6)
    UserControl.Line (0, UserControl.ScaleHeight - 1)-(UserControl.ScaleWidth - 1, UserControl.ScaleHeight - 1), tColorSet.csColor1(6)
        
    If m_HoverSelect = False And bSelected = True Then
        UserControl.Line (1, 1)-(1, UserControl.ScaleHeight - 2), tColorSet.csColor1(10)
        UserControl.Line (1, 1)-(UserControl.ScaleWidth - 2, 1), tColorSet.csColor1(10)
        UserControl.Line (UserControl.ScaleWidth - 2, 1)-(UserControl.ScaleWidth - 2, UserControl.ScaleHeight - 1), tColorSet.csColor1(10)
        UserControl.Line (1, UserControl.ScaleHeight - 2)-(UserControl.ScaleWidth - 2, UserControl.ScaleHeight - 2), tColorSet.csColor1(10)
    End If
    
End Sub

Private Function CountTabs(strString As String) As Integer
    
    Dim i As Integer
    Dim i2 As Integer
    Dim c As Integer
    
    c = 0
    i = 1
    Do
    
        i2 = InStr(i, strString, vbTab)
        If i2 = 0 Then Exit Do
        
        c = c + 1
        i = i2 + 1
    
    Loop
    
    CountTabs = c
    
End Function

Private Sub DrawImage(iIndex As Integer, iSpecialGap As Integer)

    If ListboxItems.Item(TempListboxItems.Item(iIndex).Index).Icon = 0 Then Exit Sub
    If m_ImageList.ListImages.Count = 0 Then Exit Sub
    
    If UserControl.Enabled = False Then
        Call DrawState(UserControl.hdc, 0, 0, _
        m_ImageList.ListImages(ListboxItems.Item(TempListboxItems.Item(iIndex).Index).Icon).ExtractIcon, 0, _
        ITEM_GAP + iSpecialGap, TempListboxItems.Item(iIndex).Top + (ITEM_GAP / 2) + TOP_GAP, _
        m_ImageList.ImageWidth, m_ImageList.ImageHeight, DST_ICON Or DSS_MONO)
    
    Else
        Call DrawState(UserControl.hdc, 0, 0, _
        m_ImageList.ListImages(ListboxItems.Item(TempListboxItems.Item(iIndex).Index).Icon).ExtractIcon, 0, _
        ITEM_GAP + iSpecialGap, TempListboxItems.Item(iIndex).Top + (ITEM_GAP / 2) + TOP_GAP, _
        m_ImageList.ImageWidth, m_ImageList.ImageHeight, DST_ICON Or DSS_NORMAL)
    End If

End Sub

Private Function ShortWord(sString As String, RefLenght As Long) As String
    
    Dim tmpString As String
    
    tmpString = sString
    
    If UserControl.TextWidth(tmpString) > RefLenght Then
    
        Do Until UserControl.TextWidth(tmpString) < RefLenght
        
            If Len(tmpString) <= 3 Then Exit Do
            tmpString = Left(tmpString, Len(tmpString) - 4) & "..."
            
        Loop
        
    End If
    
    ShortWord = tmpString

End Function

'ACHTUNG! DIE FOLGENDEN KOMMENTIERTEN ZEILEN NICHT ENTFERNEN ODER VERÄNDERN!
'MemberInfo=10,0,0,0
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Gibt die Hintergrundfarbe zurück, die verwendet wird, um Text und Grafik in einem Objekt anzuzeigen, oder legt diese fest."
Attribute BackColor.VB_ProcData.VB_Invoke_Property = ";Darstellung"
    BackColor = m_BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    m_BackColor = New_BackColor
    PropertyChanged "BackColor"
    RefreshColor
    DrawFace
End Property

'ACHTUNG! DIE FOLGENDEN KOMMENTIERTEN ZEILEN NICHT ENTFERNEN ODER VERÄNDERN!
'MemberInfo=10,0,0,0
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Gibt die Vordergrundfarbe zurück, die zum Anzeigen von Text und Grafiken in einem Objekt verwendet wird, oder legt diese fest."
Attribute ForeColor.VB_ProcData.VB_Invoke_Property = ";Darstellung"
    ForeColor = m_ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    m_ForeColor = New_ForeColor
    PropertyChanged "ForeColor"
    RefreshColor
    DrawFace
End Property

'ACHTUNG! DIE FOLGENDEN KOMMENTIERTEN ZEILEN NICHT ENTFERNEN ODER VERÄNDERN!
'MappingInfo=UserControl,UserControl,-1,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Gibt einen Wert zurück, der bestimmt, ob ein Objekt auf vom Benutzer erzeugte Ereignisse reagieren kann, oder legt diesen fest."
Attribute Enabled.VB_ProcData.VB_Invoke_Property = ";Verhalten"
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    PropertyChanged "Enabled"
    DrawFace
End Property

'ACHTUNG! DIE FOLGENDEN KOMMENTIERTEN ZEILEN NICHT ENTFERNEN ODER VERÄNDERN!
'MappingInfo=UserControl,UserControl,-1,Font
Public Property Get Font() As Font
Attribute Font.VB_Description = "Gibt ein Font-Objekt zurück."
Attribute Font.VB_ProcData.VB_Invoke_Property = ";Schriftart"
Attribute Font.VB_UserMemId = -512
    Set Font = UserControl.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set UserControl.Font = New_Font
    PropertyChanged "Font"
    DrawFace
End Property

Private Sub ListboxItems_ItemChanged()
    
    If GetMaxSlot > ListboxItems.Count Then
        cainScrollBar1.Enabled = False
    Else
        cainScrollBar1.Enabled = True
    End If

    DrawFace
End Sub

Private Sub MousePos_Timer()

    GetCursorPos Mouse

    If ((MouseOverMe.X > Mouse.X - 2) And (MouseOverMe.X < Mouse.X + 2)) And ((MouseOverMe.Y > Mouse.Y - 2) And (MouseOverMe.Y < Mouse.Y + 2)) Then
        DrawFace True
    Else
        DrawFace
        MousePos.Enabled = False
    End If

    DoEvents
    
    'RaiseEvent MouseOver

End Sub

Private Sub UserControl_EnterFocus()
    bHasFocus = True
End Sub

Private Sub UserControl_ExitFocus()
    bHasFocus = False
    DrawFace
End Sub

Private Sub UserControl_Initialize()

    Set ListboxItems = New ComboDatas
    Set TempListboxItems = New MenuItemListings
    bHasFocus = False
    ValueFont = False
    
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Dim i As Long
    Select Case KeyCode
     
        Case 38
            If m_SelectedItem > 1 Then _
                SelectedItem = m_SelectedItem - 1
            
        Case 40
            If m_SelectedItem < ListboxItems.Count Then _
                SelectedItem = m_SelectedItem + 1
        
    End Select
    
    RaiseEvent KeyDown(KeyCode, Shift)
    
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub UserControl_LostFocus()
    UserControl_ExitFocus
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> 1 Then Exit Sub
    MouseClick = True
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    GetCursorPos MouseOverMe
        If MousePos.Enabled = False Then MousePos.Enabled = True
    
    Dim i As Integer
    
    If m_HoverSelect = True Then
        For i = 1 To TempListboxItems.Count
            If (Y >= TempListboxItems(i).Top And Y <= (TempListboxItems(i).Top + TempListboxItems(i).Height)) And (X < (UserControl.ScaleWidth - cainScrollBar1.Width) - 4) Then
                iSelection = TempListboxItems(i).Index
                m_SelectedItem = iSelection
            End If
        Next i
        DrawFace

    End If
    
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If Button <> 1 Then Exit Sub
    If MouseClick = True Then RaiseEvent Click
    MouseClick = False
    
    Dim i As Integer
    
    'iSelection = 0
    If m_HoverSelect = False Then
        For i = 1 To TempListboxItems.Count
            If (Y >= TempListboxItems(i).Top And Y <= (TempListboxItems(i).Top + TempListboxItems(i).Height)) Then
                iSelection = TempListboxItems(i).Index
                m_SelectedItem = iSelection
                RaiseEvent ItemClicked(i, ListboxItems.Item(TempListboxItems(i).Index).Key)
            End If
        Next i
    Else
        If iSelection <> 0 Then _
        RaiseEvent ItemClicked(iSelection, ListboxItems.Item(iSelection).Key)
    End If
    
    DrawFace
    RaiseEvent MouseUp(Button, Shift, X, Y)

End Sub

'ACHTUNG! DIE FOLGENDEN KOMMENTIERTEN ZEILEN NICHT ENTFERNEN ODER VERÄNDERN!
'MemberInfo=10,0,0,0
Public Property Get SelectionColor() As OLE_COLOR
Attribute SelectionColor.VB_ProcData.VB_Invoke_Property = ";Darstellung"
    SelectionColor = m_SelectionColor
End Property

Public Property Let SelectionColor(ByVal New_SelectionColor As OLE_COLOR)
    m_SelectionColor = New_SelectionColor
    PropertyChanged "SelectionColor"
    RefreshColor
    DrawFace
End Property

'ACHTUNG! DIE FOLGENDEN KOMMENTIERTEN ZEILEN NICHT ENTFERNEN ODER VERÄNDERN!
'MemberInfo=9,0,0,0
Public Property Get ImageList() As Object
Attribute ImageList.VB_ProcData.VB_Invoke_Property = ";Darstellung"
    Set ImageList = m_ImageList
End Property

Public Property Set ImageList(ByVal New_ImageList As Object)
    Set m_ImageList = New_ImageList
    PropertyChanged "ImageList"
    DrawFace
End Property

'Eigenschaften für Benutzersteuerelement initialisieren
Private Sub UserControl_InitProperties()
    m_BackColor = m_def_BackColor
    m_ForeColor = m_def_ForeColor
    Set UserControl.Font = Ambient.Font
    m_SelectionColor = m_def_SelectionColor
    m_SelectedItem = m_def_SelectedItem
    m_HoverSelect = m_def_HoverSelect
    
    RefreshColor
    DrawFace
End Sub

'Eigenschaftenwerte vom Speicher laden
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    m_BackColor = PropBag.ReadProperty("BackColor", m_def_BackColor)
    m_ForeColor = PropBag.ReadProperty("ForeColor", m_def_ForeColor)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    Set UserControl.Font = PropBag.ReadProperty("Font", Ambient.Font)
    m_SelectionColor = PropBag.ReadProperty("SelectionColor", m_def_SelectionColor)
    Set m_ImageList = PropBag.ReadProperty("ImageList", Nothing)
    m_SelectedItem = PropBag.ReadProperty("SelectedItem", m_def_SelectedItem)
    m_HoverSelect = PropBag.ReadProperty("HoverSelect", m_def_HoverSelect)
    
    RefreshColor
    DrawFace
End Sub

Private Sub UserControl_Resize()

    On Error Resume Next

    cainScrollBar1.Top = 2
    cainScrollBar1.Left = UserControl.ScaleWidth - cainScrollBar1.Width - 2
    cainScrollBar1.Height = UserControl.ScaleHeight - 4
    
    DrawFace

End Sub

Private Sub UserControl_Terminate()

    If Not (ListboxItems Is Nothing) Then Set ListboxItems = Nothing
    If Not (TempListboxItems Is Nothing) Then Set TempListboxItems = Nothing
    
End Sub

'Eigenschaftenwerte in den Speicher schreiben
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("BackColor", m_BackColor, m_def_BackColor)
    Call PropBag.WriteProperty("ForeColor", m_ForeColor, m_def_ForeColor)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("Font", UserControl.Font, Ambient.Font)
    Call PropBag.WriteProperty("SelectionColor", m_SelectionColor, m_def_SelectionColor)
    Call PropBag.WriteProperty("ImageList", m_ImageList, Nothing)
    Call PropBag.WriteProperty("SelectedItem", m_SelectedItem, m_def_SelectedItem)
    Call PropBag.WriteProperty("HoverSelect", m_HoverSelect, m_def_HoverSelect)
    
    RefreshColor
    DrawFace
End Sub

'ACHTUNG! DIE FOLGENDEN KOMMENTIERTEN ZEILEN NICHT ENTFERNEN ODER VERÄNDERN!
'MemberInfo=7,0,0,0
Public Property Get SelectedItem() As Integer
Attribute SelectedItem.VB_ProcData.VB_Invoke_Property = ";Daten"
Attribute SelectedItem.VB_MemberFlags = "400"
    SelectedItem = m_SelectedItem
End Property

Public Property Let SelectedItem(ByVal New_SelectedItem As Integer)
    m_SelectedItem = New_SelectedItem
    PropertyChanged "SelectedItem"
    
    iSelection = New_SelectedItem
    cainScrollBar1.Value = iSelection
    
    RaiseEvent ItemClicked(New_SelectedItem, ListboxItems.Item(New_SelectedItem).Key)
    
    DrawFace
    
End Property

Public Sub SelectItem(Optional mKey As String = "", Optional mValue As String = "", Optional mCaption As String = "")

    If ListboxItems.Count = 0 Then Exit Sub

    Dim i As Integer
        
    For i = 1 To ListboxItems.Count
        
        If mKey <> "" Then
            If ListboxItems.Item(i).Key = mKey Then
                SelectedItem = i
                Exit Sub
            End If
        ElseIf mValue <> "" Then
            If ListboxItems.Item(i).Value = mValue Then
                SelectedItem = i
                Exit Sub
            End If
        ElseIf mCaption <> "" Then
            If ListboxItems.Item(i).Caption = mCaption Then
                SelectedItem = i
                Exit Sub
            End If
        End If
        
    Next i
    
    SelectedItem = 1

End Sub

'ACHTUNG! DIE FOLGENDEN KOMMENTIERTEN ZEILEN NICHT ENTFERNEN ODER VERÄNDERN!
'MemberInfo=0,0,0,0
Public Property Get HoverSelect() As Boolean
Attribute HoverSelect.VB_ProcData.VB_Invoke_Property = ";Verhalten"
    HoverSelect = m_HoverSelect
End Property

Public Property Let HoverSelect(ByVal New_HoverSelect As Boolean)
    m_HoverSelect = New_HoverSelect
    PropertyChanged "HoverSelect"
End Property

