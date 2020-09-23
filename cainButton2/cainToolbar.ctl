VERSION 5.00
Begin VB.UserControl cainToolbar 
   Alignable       =   -1  'True
   AutoRedraw      =   -1  'True
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   ToolboxBitmap   =   "cainToolbar.ctx":0000
   Begin cainObjects.cainPUMenu cainPUMenu1 
      Height          =   705
      Left            =   720
      TabIndex        =   0
      Top             =   480
      Width           =   705
      _ExtentX        =   1244
      _ExtentY        =   1244
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Timer MousePos 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   0
      Top             =   0
   End
End
Attribute VB_Name = "cainToolbar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'Standard-Eigenschaftswerte:
Const m_def_BackColor = &HFFFFFF
Const m_def_ForeColor = &HC00000
'Const m_def_Enabled = 0
Const m_def_SelectionColor = &H80FF&
'Eigenschaftsvariablen:
Dim m_BackColor As OLE_COLOR
Dim m_ForeColor As OLE_COLOR
'Dim m_Enabled As Boolean
'Dim m_Font As Font
Dim m_SelectionColor As OLE_COLOR
Dim m_MenuImageList As Object
Dim m_ImageList As Object
'Ereignisdeklarationen:
Event ItemClicked(ItemIndex As Integer, ItemKey As String)
Event MenuItemClicked(ItemIndex As Integer, ItemKey As String, MenuItemIndex As Integer, MenuItemKey As String)

Dim MyColorSet As ColorSet
Dim iSelection As Integer

Dim Mouse As POINTAPI
Dim MouseOverMe As POINTAPI
Dim MouseClick As Boolean
Dim iWLeft As Integer
Dim iWTop As Integer

Dim iWElementHeight As Integer
Dim bBU_Selection As Integer

Private Const ITEM_GAP As Integer = 15
Private Const TOP_GAP As Integer = 6
Private Const ICON_GAP As Integer = 3

Public WithEvents ToolBarItems As ToolBarItems
Attribute ToolBarItems.VB_VarHelpID = -1

Public Function ID() As Integer
    ID = eControlIDs.id_Toolbar
End Function

Private Sub RefreshColor()

    MyColorSet = GetColorSetNormal(m_BackColor, m_ForeColor, m_SelectionColor)
    UserControl.BackColor = MyColorSet.csBackColor
    cainPUMenu1.BackColor = m_BackColor
    cainPUMenu1.ForeColor = m_ForeColor
    cainPUMenu1.SelectionColor = m_SelectionColor
    cainPUMenu1.Font = UserControl.Font

End Sub

Private Sub DrawIcon(iTop As Integer, iLeft As Integer, iHeight As Integer, iImageIndex As Integer, Optional bMonocrom As Boolean = False)

    'On Error Resume Next
    
    If m_ImageList Is Nothing Then Exit Sub
    If iImageIndex = 0 Then Exit Sub
    
    Dim PicTop As Integer
    Dim PicLeft As Integer
    
    PicTop = iTop + (Fix(iHeight / 2) - Fix(m_ImageList.ImageHeight / 2))
    PicLeft = iLeft
    
    If bMonocrom = True Then
        Call DrawState(UserControl.hdc, 0, 0, _
        m_ImageList.ListImages(iImageIndex).ExtractIcon, 0, _
        PicLeft, PicTop, _
        m_ImageList.ImageWidth, m_ImageList.ImageHeight, DST_ICON Or DSS_MONO)
    
    Else
        Call DrawState(UserControl.hdc, 0, 0, _
        m_ImageList.ListImages(iImageIndex).ExtractIcon, 0, _
        PicLeft, PicTop, _
        m_ImageList.ImageWidth, m_ImageList.ImageHeight, DST_ICON Or DSS_NORMAL)
    End If

End Sub

Private Sub DrawSelection(iTop As Integer, iLeft As Integer, iHeight As Integer, iWidth As Integer)

    If iSelection = 0 Then Exit Sub
    
    If (ToolBarItems.Item(iSelection).Style = eMenuItemType.mitMenu Or ToolBarItems.Item(iSelection).Style = eMenuItemType.mitMenu2) And MouseClick = True Then
        DrawGradient UserControl.hdc, iLeft * 1, iTop * 1, iWidth * 1, iHeight * 1, GetRGBColors(MyColorSet.csColor1(1)), GetRGBColors(MyColorSet.csColor1(6)), 1
        UserControl.Line (iLeft, iTop)-(iLeft + iWidth, iTop), MyColorSet.csColor1(6)
        UserControl.Line (iLeft, iTop)-(iLeft, iTop + iHeight), MyColorSet.csColor1(6)
        UserControl.Line (iLeft + iWidth, iTop)-(iLeft + iWidth, iTop + iHeight), MyColorSet.csColor1(6)
        UserControl.Line (iLeft, iTop + iHeight)-(iLeft + iWidth + 1, iTop + iHeight), MyColorSet.csColor1(6)
    
    ElseIf MouseClick = True Then
        DrawGradient UserControl.hdc, iLeft * 1, iTop * 1, iWidth * 1, iHeight * 1, GetRGBColors(MyColorSet.csColor1(10)), GetRGBColors(MyColorSet.csColor1(8)), 1
        UserControl.Line (iLeft, iTop)-(iLeft + iWidth, iTop), MyColorSet.csColor1(6)
        UserControl.Line (iLeft, iTop)-(iLeft, iTop + iHeight), MyColorSet.csColor1(6)
        UserControl.Line (iLeft + iWidth, iTop)-(iLeft + iWidth, iTop + iHeight), MyColorSet.csColor1(6)
        UserControl.Line (iLeft, iTop + iHeight)-(iLeft + iWidth + 1, iTop + iHeight), MyColorSet.csColor1(6)
    
    ElseIf ToolBarItems.Item(iSelection).Style = eMenuItemType.mitCheckButton And ToolBarItems.Item(iSelection).Checked = True Then
        DrawGradient UserControl.hdc, iLeft * 1, iTop * 1, iWidth * 1, iHeight * 1, GetRGBColors(MyColorSet.csColor1(8)), GetRGBColors(MyColorSet.csColor1(10)), 1
        UserControl.Line (iLeft, iTop)-(iLeft + iWidth, iTop), MyColorSet.csColor1(6)
        UserControl.Line (iLeft, iTop)-(iLeft, iTop + iHeight), MyColorSet.csColor1(6)
        UserControl.Line (iLeft + iWidth, iTop)-(iLeft + iWidth, iTop + iHeight), MyColorSet.csColor1(6)
        UserControl.Line (iLeft, iTop + iHeight)-(iLeft + iWidth + 1, iTop + iHeight), MyColorSet.csColor1(6)
    
    Else
        DrawGradient UserControl.hdc, iLeft * 1, iTop * 1, iWidth * 1, iHeight * 1, GetRGBColors(MyColorSet.csColor1(8)), GetRGBColors(MyColorSet.csColor1(9)), 1
        UserControl.Line (iLeft, iTop)-(iLeft + iWidth, iTop), MyColorSet.csColor1(6)
        UserControl.Line (iLeft, iTop)-(iLeft, iTop + iHeight), MyColorSet.csColor1(6)
        UserControl.Line (iLeft + iWidth, iTop)-(iLeft + iWidth, iTop + iHeight), MyColorSet.csColor1(6)
        UserControl.Line (iLeft, iTop + iHeight)-(iLeft + iWidth + 1, iTop + iHeight), MyColorSet.csColor1(6)
    
    End If

End Sub

Private Sub DrawChecked(iTop As Integer, iLeft As Integer, iHeight As Integer, iWidth As Integer)
    
    DrawGradient UserControl.hdc, iLeft * 1, iTop * 1, iWidth * 1, iHeight * 1, GetRGBColors(MyColorSet.csColor1(8)), GetRGBColors(MyColorSet.csColor1(9)), 1
    UserControl.Line (iLeft, iTop)-(iLeft + iWidth, iTop), MyColorSet.csColor1(6)
    UserControl.Line (iLeft, iTop)-(iLeft, iTop + iHeight), MyColorSet.csColor1(6)
    UserControl.Line (iLeft + iWidth, iTop)-(iLeft + iWidth, iTop + iHeight), MyColorSet.csColor1(6)
    UserControl.Line (iLeft, iTop + iHeight)-(iLeft + iWidth + 1, iTop + iHeight), MyColorSet.csColor1(6)

End Sub

Private Sub DrawFace()

    Dim i As Integer
    Dim X As Integer
    Dim iHeight As Integer
    Dim iWidth As Integer
    Dim iLeft As Integer
    Dim iTop As Integer
    Dim iTextLeft As Integer
    Dim iTextTop As Integer
    
    UserControl.Cls
    
    X = ITEM_GAP
    
    
    If m_ImageList Is Nothing Then
        iWElementHeight = (TOP_GAP * 2) + UserControl.TextHeight("I") + 2
    Else
        
        If m_ImageList.ImageHeight > UserControl.TextHeight("I") Then
            iWElementHeight = (TOP_GAP * 2) + m_ImageList.ImageHeight + 2
        Else
            iWElementHeight = (TOP_GAP * 2) + UserControl.TextHeight("I") + 2
        End If
    
    End If
    
    
    If UserControl.ScaleHeight <> iWElementHeight Then UserControl.Height = iWElementHeight * Screen.TwipsPerPixelY
    
    'background
    GradientCy UserControl.hdc, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight, MyColorSet.csColor1(1), MyColorSet.csColor1(2), MyColorSet.csColor1(3), pbHorizontal

    'Borders
    GradientLine UserControl.hdc, 1, 0, UserControl.ScaleWidth - 2, pbHorizontal, MyColorSet.csColor1(4), MyColorSet.csColor1(5)
    GradientLine UserControl.hdc, 0, 1, UserControl.ScaleHeight - 2, pbVertical, MyColorSet.csColor1(4), MyColorSet.csColor1(6)
    UserControl.Line (1, UserControl.ScaleHeight - 1)-(UserControl.ScaleWidth - 1, UserControl.ScaleHeight - 1), MyColorSet.csColor1(6)
    GradientLine UserControl.hdc, UserControl.ScaleWidth - 1, 1, UserControl.ScaleHeight - 2, pbVertical, MyColorSet.csColor1(5), MyColorSet.csColor1(6)
    
    If ToolBarItems.Count = 0 Then Exit Sub
    
    For i = 1 To ToolBarItems.Count
        
        
        UserControl.ForeColor = MyColorSet.csColor1(7)
        
        iTop = TOP_GAP
        
        If m_ImageList Is Nothing Then
        
            iLeft = X
            iTextLeft = iLeft + 2
            iTextTop = iTop
            
            iWidth = UserControl.TextWidth(ToolBarItems.Item(i).Caption) + ITEM_GAP
            X = X + UserControl.TextWidth(ToolBarItems.Item(i).Caption) + ITEM_GAP
            iHeight = UserControl.TextHeight("I")
            
        Else
            
            If m_ImageList.ImageHeight > UserControl.TextHeight("I") Then
                iTextTop = iTop + ((m_ImageList.ImageHeight / 2) - (UserControl.TextHeight("I") / 2))
                iHeight = m_ImageList.ImageHeight
            Else
                iTextTop = iTop
                iHeight = UserControl.TextHeight("I")
            End If
            
        
            If ToolBarItems.Item(i).IconIndex <> 0 Then
                iLeft = X
                iTextLeft = iLeft + m_ImageList.ImageWidth + ICON_GAP
                iWidth = UserControl.TextWidth(ToolBarItems.Item(i).Caption) + ITEM_GAP + m_ImageList.ImageWidth
                X = X + UserControl.TextWidth(ToolBarItems.Item(i).Caption) + ITEM_GAP + m_ImageList.ImageWidth + ICON_GAP
                
            Else
                iLeft = X
                iTextLeft = iLeft
                iWidth = UserControl.TextWidth(ToolBarItems.Item(i).Caption) + ITEM_GAP
                X = X + UserControl.TextWidth(ToolBarItems.Item(i).Caption) + ITEM_GAP
                
            End If
            
        End If
        
        If ToolBarItems.Item(i).Style = eMenuItemType.mitMenu2 Then iWidth = iWidth + ICON_GAP + UserControl.TextWidth("6")
        
        If iSelection = i Then DrawSelection iTop - Fix(TOP_GAP / 2), iLeft - Fix(ITEM_GAP / 2), iHeight + TOP_GAP, iWidth
        If iSelection <> i And ToolBarItems.Item(i).Style = mitCheckButton And ToolBarItems.Item(i).Checked = True Then DrawChecked iTop - Fix(TOP_GAP / 2), iLeft - Fix(ITEM_GAP / 2), iHeight + TOP_GAP, iWidth
                
        Select Case ToolBarItems.Item(i).Style
        
        Case eMenuItemType.mitPlaceholder
            If ToolBarItems.Item(i).IconIndex <> 0 Then DrawIcon iTop, iLeft, iHeight, ToolBarItems.Item(i).IconIndex, False
            
        Case eMenuItemType.mitSeparator
            UserControl.Line (iLeft + Fix(iWidth / 4), iTop)-(iLeft + Fix(iWidth / 4), iTop + iHeight), MyColorSet.csColor1(6)
            UserControl.Line (iLeft + Fix(iWidth / 4) + 1, iTop + 1)-(iLeft + Fix(iWidth / 4) + 1, iTop + iHeight + 1), MyColorSet.csColor1(1)
        
        Case Else
            If ToolBarItems.Item(i).IconIndex <> 0 Then DrawIcon iTop, iLeft, iHeight, ToolBarItems.Item(i).IconIndex, False
            
            UserControl.CurrentY = iTextTop
            UserControl.CurrentX = iTextLeft
            
            If ToolBarItems.Item(i).Enabled = False Then
                UserControl.ForeColor = MyColorSet.csColor1(6)
            Else
                UserControl.ForeColor = MyColorSet.csColor1(7)
            End If
            
            UserControl.Print ToolBarItems.Item(i).Caption
            
            If ToolBarItems.Item(i).Style = eMenuItemType.mitMenu2 Then
            
                Dim OldFont As String
                
                OldFont = UserControl.FontName
                UserControl.CurrentY = iTextTop - 3
                UserControl.CurrentX = iTextLeft + UserControl.TextWidth(ToolBarItems.Item(i).Caption) + ICON_GAP
                UserControl.FontName = "WebDings"
                UserControl.Print "6"
                UserControl.FontName = OldFont
           
            End If
        
        End Select
        
    Next i
    
End Sub


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
    
    cainPUMenu1.BackColor = m_BackColor
   
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
    
    cainPUMenu1.ForeColor = m_ForeColor
    
    RefreshColor
    DrawFace
    
End Property
'
''ACHTUNG! DIE FOLGENDEN KOMMENTIERTEN ZEILEN NICHT ENTFERNEN ODER VERÄNDERN!
''MemberInfo=0,0,0,0
'Public Property Get Enabled() As Boolean
'    Enabled = m_Enabled
'End Property
'
'Public Property Let Enabled(ByVal New_Enabled As Boolean)
'    m_Enabled = New_Enabled
'    PropertyChanged "Enabled"
'End Property
'
''ACHTUNG! DIE FOLGENDEN KOMMENTIERTEN ZEILEN NICHT ENTFERNEN ODER VERÄNDERN!
''MemberInfo=6,0,0,0
'Public Property Get Font() As Font
'    Set Font = m_Font
'End Property
'
'Public Property Set Font(ByVal New_Font As Font)
'    Set m_Font = New_Font
'    PropertyChanged "Font"
'
'    RefreshColor
'    DrawFace
'
'End Property

'ACHTUNG! DIE FOLGENDEN KOMMENTIERTEN ZEILEN NICHT ENTFERNEN ODER VERÄNDERN!
'MemberInfo=10,0,0,0
Public Property Get SelectionColor() As OLE_COLOR
Attribute SelectionColor.VB_ProcData.VB_Invoke_Property = ";Darstellung"
    SelectionColor = m_SelectionColor
End Property

Public Property Let SelectionColor(ByVal New_SelectionColor As OLE_COLOR)
    m_SelectionColor = New_SelectionColor
    PropertyChanged "SelectionColor"
    
    cainPUMenu1.SelectionColor = m_SelectionColor
    
    RefreshColor
    DrawFace
    
End Property

'ACHTUNG! DIE FOLGENDEN KOMMENTIERTEN ZEILEN NICHT ENTFERNEN ODER VERÄNDERN!
'MemberInfo=9,0,0,0
Public Property Get MenuImageList() As Object
Attribute MenuImageList.VB_ProcData.VB_Invoke_Property = ";Darstellung"
    Set MenuImageList = m_MenuImageList
End Property

Public Property Set MenuImageList(ByVal New_MenuImageList As Object)
    Set m_MenuImageList = New_MenuImageList
    PropertyChanged "MenuImageList"
    
    Set cainPUMenu1.ImageList = New_MenuImageList
    
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
    PropertyChanged "Imagelist"
    
    RefreshColor
    DrawFace
    
End Property

Private Sub cainPUMenu1_Closed()

    MouseClick = False
    iSelection = 0
    DrawFace

End Sub

Private Sub cainPUMenu1_ItemClick(ItemIndex As Integer, ItemKey As String)
    RaiseEvent MenuItemClicked(bBU_Selection, ToolBarItems.Item(bBU_Selection).sKey, ItemIndex, ItemKey)
End Sub

Private Sub UserControl_Initialize()

    Set ToolBarItems = New ToolBarItems
    
    iSelection = 0
    bBU_Selection = 0
    iWElementHeight = ITEM_GAP + TOP_GAP + UserControl.TextHeight("I")
    
    cainPUMenu1.Top = -cainPUMenu1.Top - 10
    cainPUMenu1.Left = -cainPUMenu1.Left - 10
    
    
    
End Sub

Private Sub ToolBarItems_ItemAdded()
    
    'RefreshColor
    DrawFace
    
End Sub

Private Sub ToolBarItems_ItemChanged()
    
    'RefreshColor
    DrawFace
    'Debug.Print Second(Time)
    
End Sub

Private Sub MousePos_Timer()

    GetCursorPos Mouse

    If ((MouseOverMe.X > Mouse.X - 2) And (MouseOverMe.X < Mouse.X + 2)) And ((MouseOverMe.Y > Mouse.Y - 2) And (MouseOverMe.Y < Mouse.Y + 2)) Then
        'If MouseClick = True Then
        '    DrawFace
        'Else
            DrawFace
        'End If
    Else
        DrawFace
        MousePos.Enabled = False
    End If

    DoEvents
    
    'RaiseEvent MouseOver

End Sub

Private Sub Create_Menu(Optional mGapWidth As Integer = 0)

    If iSelection = 0 Then Exit Sub
    If ToolBarItems.Item(iSelection).Style <> eMenuItemType.mitMenu And ToolBarItems.Item(iSelection).Style <> eMenuItemType.mitMenu2 Then
        cainPUMenu1.KillMenu
        Exit Sub
    End If

    If iSelection <> 0 Then
        
        cainPUMenu1.ClearItems
        If ToolBarItems.Item(iSelection).ToolbarMenuItems.Count <> 0 Then _
        Set cainPUMenu1.MenuItem = ToolBarItems.Item(iSelection).ToolbarMenuItems
        
        Dim oRect As RECT
        Call GetWindowRect(hwnd, oRect)
        
        cainPUMenu1.CreateMenu (oRect.Left + iWLeft + 1) * Screen.TwipsPerPixelX, (oRect.Top + iWTop + 3) * Screen.TwipsPerPixelY, , mGapWidth
        
    End If

End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    
    If Button <> 1 Then Exit Sub
    MouseClick = True
    MenuItemEvents X, Y
    'Create_Menu
    

End Sub

Private Sub MenuItemEvents(X As Single, Y As Single)
    
    Dim i As Integer
    Dim iX As Integer
    Dim iHeight As Integer
    Dim iWidth As Integer
    Dim iLeft As Integer
    Dim iTop As Integer

    GetCursorPos MouseOverMe
    If MousePos.Enabled = False Then MousePos.Enabled = True
    If cainPUMenu1.MenuVisible = False Then iSelection = 0
    
    iX = Fix(ITEM_GAP / 2)
    
    For i = 1 To ToolBarItems.Count
        
        iTop = TOP_GAP
        
        If m_ImageList Is Nothing Then
        
            iLeft = iX
            iWidth = UserControl.TextWidth(ToolBarItems.Item(i).Caption) + ITEM_GAP
            iX = iX + UserControl.TextWidth(ToolBarItems.Item(i).Caption) + ITEM_GAP
            iHeight = UserControl.TextHeight("I")
            
        Else
            
            If m_ImageList.ImageHeight > UserControl.TextHeight("I") Then
                iHeight = m_ImageList.ImageHeight
            Else
                iHeight = UserControl.TextHeight("I")
            End If
        
            If ToolBarItems.Item(i).IconIndex <> 0 Then
                iLeft = iX
                iWidth = UserControl.TextWidth(ToolBarItems.Item(i).Caption) + ITEM_GAP + m_ImageList.ImageWidth
                iX = iX + UserControl.TextWidth(ToolBarItems.Item(i).Caption) + ITEM_GAP + m_ImageList.ImageWidth + ICON_GAP
            Else
                iLeft = iX
                iWidth = UserControl.TextWidth(ToolBarItems.Item(i).Caption) + ITEM_GAP
                iX = iX + UserControl.TextWidth(ToolBarItems.Item(i).Caption) + ITEM_GAP
            End If
            
        End If
        If ToolBarItems.Item(i).Style = eMenuItemType.mitMenu2 Then iWidth = iWidth + ICON_GAP + UserControl.TextWidth("6")
        
        If ((X > iLeft) And (X < (iLeft + iWidth))) And ((Y > iTop) And (Y < (iTop + iHeight))) Then
                
            Select Case ToolBarItems.Item(i).Style
            
            Case eMenuItemType.mitMenu, eMenuItemType.mitMenu2
                
                If i <> iSelection Then
                    iSelection = i
                    bBU_Selection = i
                    iWLeft = iLeft
                    iWTop = iTop + iHeight
                
                    If MouseClick = True Then
                        Create_Menu iWidth
                    End If
                
                End If
                
                'Debug.Print "iSelection = " & iSelection
                'Debug.Print "cainPUMenu1.MenuVisible = " & cainPUMenu1.MenuVisible
                
            Case eMenuItemType.mitNormalButton, eMenuItemType.mitCheckButton
                
                If i <> iSelection And ToolBarItems.Item(i).Enabled = True Then
                    iSelection = i
                    If cainPUMenu1.MenuVisible = True Then cainPUMenu1.KillMenu
'                    iWLeft = iLeft
'                    iWTop = iTop + iHeight

                ElseIf ToolBarItems.Item(i).Enabled = True Then
                    iSelection = 0
                    If cainPUMenu1.MenuVisible = True Then cainPUMenu1.KillMenu
                End If
            
            Case eMenuItemType.mitPlaceholder, eMenuItemType.mitSeparator, eMenuItemType.mitCaption
                If cainPUMenu1.MenuVisible = True Then cainPUMenu1.KillMenu
                iSelection = 0
            End Select
            
            Exit Sub
        End If
    
    Next i
    
    'DrawFace

End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MenuItemEvents X, Y
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If iSelection = 0 Then: _
        MouseClick = False: _
        Exit Sub
    
    'If Button <> 1 Then Exit Sub
    If MouseClick = True Then RaiseEvent ItemClicked(iSelection, ToolBarItems.Item(iSelection).sKey)
    
    Select Case ToolBarItems.Item(iSelection).Style
        Case eMenuItemType.mitMenu, eMenuItemType.mitMenu2
            'Do nothing
        Case Else
            MouseClick = False
    End Select
    
End Sub

'Eigenschaften für Benutzersteuerelement initialisieren
Private Sub UserControl_InitProperties()
    m_BackColor = m_def_BackColor
    m_ForeColor = m_def_ForeColor
'    m_Enabled = m_def_Enabled
'    Set m_Font = Ambient.Font
    m_SelectionColor = m_def_SelectionColor
    Set UserControl.Font = Ambient.Font
    
    RefreshColor
    DrawFace
    
End Sub

'Eigenschaftenwerte vom Speicher laden
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    m_BackColor = PropBag.ReadProperty("BackColor", m_def_BackColor)
    m_ForeColor = PropBag.ReadProperty("ForeColor", m_def_ForeColor)
'    m_Enabled = PropBag.ReadProperty("Enabled", m_def_Enabled)
'    Set m_Font = PropBag.ReadProperty("Font", Ambient.Font)
    m_SelectionColor = PropBag.ReadProperty("SelectionColor", m_def_SelectionColor)
    Set m_MenuImageList = PropBag.ReadProperty("MenuImageList", Nothing)
    Set m_ImageList = PropBag.ReadProperty("Imagelist", Nothing)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    Set UserControl.Font = PropBag.ReadProperty("Font", Ambient.Font)
    
    RefreshColor
    DrawFace
    
End Sub

Private Sub UserControl_Resize()
        
    UserControl.Height = iWElementHeight * Screen.TwipsPerPixelY
        
    RefreshColor
    DrawFace

End Sub

Private Sub UserControl_Terminate()
    If Not (ToolBarItems Is Nothing) Then Set ToolBarItems = Nothing
End Sub

'Eigenschaftenwerte in den Speicher schreiben
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("BackColor", m_BackColor, m_def_BackColor)
    Call PropBag.WriteProperty("ForeColor", m_ForeColor, m_def_ForeColor)
'    Call PropBag.WriteProperty("Enabled", m_Enabled, m_def_Enabled)
'    Call PropBag.WriteProperty("Font", m_Font, Ambient.Font)
    Call PropBag.WriteProperty("SelectionColor", m_SelectionColor, m_def_SelectionColor)
    Call PropBag.WriteProperty("MenuImageList", m_MenuImageList, Nothing)
    Call PropBag.WriteProperty("Imagelist", m_ImageList, Nothing)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("Font", UserControl.Font, Ambient.Font)
    
    RefreshColor
    DrawFace
    
End Sub

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
    
    RefreshColor
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
    
    cainPUMenu1.Font = UserControl.Font
    
    RefreshColor
    DrawFace
    
End Property

