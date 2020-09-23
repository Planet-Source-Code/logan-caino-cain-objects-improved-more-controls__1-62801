VERSION 5.00
Begin VB.UserControl cainPUMenu 
   AutoRedraw      =   -1  'True
   ClientHeight    =   705
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   705
   Picture         =   "cainPUMenu.ctx":0000
   ScaleHeight     =   705
   ScaleWidth      =   705
   ToolboxBitmap   =   "cainPUMenu.ctx":628A
   Begin VB.Timer tmrFocus 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   0
      Top             =   0
   End
   Begin VB.PictureBox PUMenu 
      Appearance      =   0  '2D
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'Kein
      ForeColor       =   &H80000008&
      Height          =   2175
      Left            =   120
      ScaleHeight     =   145
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   201
      TabIndex        =   0
      Top             =   720
      Visible         =   0   'False
      Width           =   3015
   End
End
Attribute VB_Name = "cainPUMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'Standard-Eigenschaftswerte:
Const m_def_SelectionColor = &H80FF&
Const m_def_BackColor = &HFFFFFF
Const m_def_ForeColor = &HC00000
'Eigenschaftsvariablen:
Dim m_SelectionColor As OLE_COLOR
Dim m_ImageList As Object
Dim m_BackColor As OLE_COLOR
Dim m_ForeColor As OLE_COLOR
Dim m_Font As Font
Public MenuItem As MenuItems
Attribute MenuItem.VB_VarProcData = ";Verschiedenes"
Public MenuVisible As Boolean

Dim mItemListing As MenuItemListings

Dim MyColorSet As ColorSet
Dim objMenu() As Object
Dim iSelection As Integer

'Gap Width
Dim mGapWidth As Integer

Event ItemClick(ItemIndex As Integer, ItemKey As String)
Event Closed()

Private Const GAP_SIZE As Integer = 8
Private Const TOP_START As Integer = 2

Public Function ID() As Integer
    ID = eControlIDs.id_PUMenu
End Function

Private Sub RefreshColor()
    
    MyColorSet = GetColorSetNormal(m_BackColor, m_ForeColor, m_SelectionColor)
    PUMenu.BackColor = MyColorSet.csBackColor
    
End Sub

Private Sub DrawSelection()

    If iSelection = 0 Then Exit Sub
    
    DrawGradient PUMenu.hdc, 3, mItemListing(iSelection).Top + 1, PUMenu.ScaleWidth - 6, mItemListing(iSelection).Height - 2, GetRGBColors(MyColorSet.csColor1(9)), GetRGBColors(MyColorSet.csColor1(8)), 0
    PUMenu.Line (2, mItemListing(iSelection).Top)-(PUMenu.ScaleWidth - 2, mItemListing(iSelection).Top), MyColorSet.csColor1(6)
    PUMenu.Line (2, mItemListing(iSelection).Top)-(2, mItemListing(iSelection).Top + mItemListing(iSelection).Height - 1), MyColorSet.csColor1(6)
    PUMenu.Line (PUMenu.ScaleWidth - 3, mItemListing(iSelection).Top)-(PUMenu.ScaleWidth - 3, mItemListing(iSelection).Top + mItemListing(iSelection).Height), MyColorSet.csColor1(6)
    PUMenu.Line (2, mItemListing(iSelection).Top + mItemListing(iSelection).Height - 1)-(PUMenu.ScaleWidth - 2, mItemListing(iSelection).Top + mItemListing(iSelection).Height - 1), MyColorSet.csColor1(6)

End Sub

Private Sub DrawIcon(iTop As Integer, iLeft As Integer, iHeight As Integer, iImageIndex As Integer, Optional bMonocrom As Boolean = False)

    'On Error Resume Next
    
    If m_ImageList Is Nothing Then Exit Sub
    If iImageIndex = 0 Then Exit Sub
    
    Dim PicTop As Integer
    Dim PicLeft As Integer
    
    PicTop = iTop + (Fix(iHeight / 2) - Fix(m_ImageList.ImageHeight / 2))
    PicLeft = Fix((iLeft - 4) / 2) - Fix(m_ImageList.ImageWidth / 2)
    
    If bMonocrom = True Then
        Call DrawState(PUMenu.hdc, 0, 0, _
        m_ImageList.ListImages(iImageIndex).ExtractIcon, 0, _
        PicLeft, PicTop, _
        m_ImageList.ImageWidth, m_ImageList.ImageHeight, DST_ICON Or DSS_MONO)
    
    Else
        Call DrawState(PUMenu.hdc, 0, 0, _
        m_ImageList.ListImages(iImageIndex).ExtractIcon, 0, _
        PicLeft, PicTop, _
        m_ImageList.ImageWidth, m_ImageList.ImageHeight, DST_ICON Or DSS_NORMAL)
    End If

End Sub

Private Sub DrawFace(Optional sGroup As String = "")

    Dim i As Integer
    Dim iTop As Long
    Dim iHeight As Long
    Dim iLeft As Long
    
    PUMenu.Cls
    iLeft = GetLeft
    
    'GradientCy PUMenu.hDC, 0, 0, iLeft - 4, PUMenu.ScaleHeight, MyColorSet.csColor1(1), MyColorSet.csColor1(2), MyColorSet.csColor1(3), pbVertical

    PUMenu.Line (0, 0)-(0, PUMenu.ScaleHeight), MyColorSet.csColor1(6)
    PUMenu.Line (mGapWidth, 0)-(PUMenu.ScaleWidth, 0), MyColorSet.csColor1(6)
    PUMenu.Line (PUMenu.ScaleWidth - 1, 0)-(PUMenu.ScaleWidth - 1, PUMenu.ScaleHeight), MyColorSet.csColor1(6)
    PUMenu.Line (0, PUMenu.ScaleHeight - 1)-(PUMenu.ScaleWidth - 1, PUMenu.ScaleHeight - 1), MyColorSet.csColor1(6)
    
    iTop = TOP_START
    
    DrawSelection
    
    For i = 1 To MenuItem.Count
    
        If MenuItem(i).Group = sGroup Then
        
            
            iHeight = GetHeight(i)
            iTop = iTop + iHeight
            
            If MenuItem(i).Bold = True Then
                If PUMenu.FontBold = False Then PUMenu.FontBold = True
            Else
                If PUMenu.FontBold = True Then PUMenu.FontBold = False
            End If
            
                
            Select Case MenuItem(i).ItemType
            
            Case eItemType.itNormal, eItemType.itCheck
                 
                
                PUMenu.CurrentX = iLeft
                PUMenu.CurrentY = iTop - iHeight + (Fix(iHeight / 2) - Fix(PUMenu.TextHeight("I") / 2))
                
                PUMenu.ForeColor = MyColorSet.csColor1(7)
                PUMenu.Print MenuItem(i).Caption
                
                If (MenuItem(i).ItemType = itCheck) And (MenuItem(i).Checked = True) Then
                    
                    Dim iColorIndex As Integer
                    
                    If iSelection <> 0 Then
                        If mItemListing(iSelection).Index = i Then
                            iColorIndex = 10
                        Else
                            iColorIndex = 8
                        End If
                    Else
                        iColorIndex = 8
                    End If
                    
                    DrawGradient PUMenu.hdc, 5, iTop - iHeight + 2, iLeft - 13, iHeight - 3, GetRGBColors(MyColorSet.csColor1(iColorIndex)), GetRGBColors(MyColorSet.csColor1(iColorIndex)), 0
                    PUMenu.Line (4, iTop - iHeight + 1)-(iLeft - 9, iTop - iHeight + 1), MyColorSet.csColor1(6)
                    PUMenu.Line (4, iTop - iHeight + 1)-(4, iTop - 1), MyColorSet.csColor1(6)
                    PUMenu.Line (4, iTop - 2)-(iLeft - 9, iTop - 2), MyColorSet.csColor1(6)
                    PUMenu.Line (iLeft - 9, iTop - 2)-(iLeft - 9, iTop - iHeight), MyColorSet.csColor1(6)
                
                End If
                
                DrawIcon iTop - iHeight, iLeft * 1, iHeight * 1, MenuItem(i).Icon
                
            Case eItemType.itPlaceholder
                GradientLine PUMenu.hdc, 1, iTop - iHeight + 1, PUMenu.ScaleWidth, pbHorizontal, MyColorSet.csColor1(1), MyColorSet.csColor1(6)
            
            Case eItemType.itTitle
                'GradientCy PUMenu.hDC, 1, iTop - iHeight + 1, PUMenu.ScaleWidth - 2, iHeight, MyColorSet.csColor1(1), MyColorSet.csColor1(2), MyColorSet.csColor1(3), pbHorizontal
                DrawGradient PUMenu.hdc, TOP_START + 1, iTop - iHeight + 1, PUMenu.ScaleWidth - (TOP_START * 2) - TOP_START, iHeight - TOP_START, GetRGBColors(MyColorSet.csColor1(3)), GetRGBColors(MyColorSet.csColor1(1)), 0
                 
                PUMenu.CurrentX = iLeft
                PUMenu.CurrentY = iTop - iHeight + GAP_SIZE
                'PUMenu.FontBold = True
                
                PUMenu.ForeColor = MyColorSet.csColor1(6)
                PUMenu.Print MenuItem(i).Caption
                
                DrawIcon iTop - iHeight, iLeft * 1, iHeight * 1, MenuItem(i).Icon
                
            End Select
            
        Else
            'Later!
        End If
        
        'PUMenu.Line (iLeft, iTop)-(PUMenu.ScaleWidth - iLeft, iTop)
    
    Next i
    

End Sub

'Private Sub cainButton1_Click()
'
'    Dim oRect As RECT
'    Call GetWindowRect(hwnd, oRect)
'    CreateMenu oRect.Left * Screen.TwipsPerPixelX, oRect.Bottom * Screen.TwipsPerPixelY + 21
'
'End Sub

Private Function GetHeight(Index As Integer)

    Dim i As Integer

    If m_ImageList Is Nothing Then
        
        i = PUMenu.TextHeight("I")
        i = i + GAP_SIZE
        
        Select Case MenuItem(Index).ItemType
        
        Case eItemType.itNormal, eItemType.itCheck
            'Do nothing
        Case eItemType.itPlaceholder
            i = 3
        Case eItemType.itTitle
            i = i * 2
        End Select
        
    Else
        
        i = PUMenu.TextHeight("I")
        i = i + GAP_SIZE
        
        Select Case MenuItem(Index).ItemType
        
        Case eItemType.itNormal, eItemType.itCheck
            If m_ImageList.ImageHeight > i Then i = m_ImageList.ImageHeight + 4
            
        Case eItemType.itPlaceholder
            i = 3
        Case eItemType.itTitle
            If m_ImageList.ImageHeight > i Then
                i = m_ImageList.ImageHeight + (GAP_SIZE * 2)
            Else
                i = i * 2
            End If
        End Select
    
    End If
    
    GetHeight = i

End Function

Private Function GetLeft() As Integer

    If m_ImageList Is Nothing Then
        GetLeft = 18
    Else
        GetLeft = m_ImageList.ImageWidth + 18
    End If

End Function

Private Function GetTotalWidth() As Integer

    Dim i As Integer
    Dim i2 As Integer
    Dim lWidth As Integer
    
    lWidth = 0
    For i = 1 To MenuItem.Count
        
        If MenuItem(i).Bold = True Then
            If PUMenu.FontBold = False Then PUMenu.FontBold = True
        Else
            If PUMenu.FontBold = True Then PUMenu.FontBold = False
        End If
        
        i2 = PUMenu.TextWidth(MenuItem(i).Caption)
        If lWidth < i2 Then lWidth = i2
        
    Next i

    GetTotalWidth = (lWidth + GetLeft + 32) * Screen.TwipsPerPixelX

End Function

Public Sub ClearItems()

    Set MenuItem = New MenuItems
    iSelection = 0

End Sub

Public Sub CreateMenu(Optional X As Long = 0, Optional Y As Long = 0, Optional sGroup As String = "", Optional M_Width As Integer = 0)

    On Error Resume Next
    
    'redim preserve
    Dim mPos As POINTAPI
    
    If X = 0 And Y = 0 Then
        Call GetCursorPos(mPos)
        X = mPos.X * Screen.TwipsPerPixelX
        Y = mPos.Y * Screen.TwipsPerPixelY
    End If
    
    Dim i As Integer
    Dim lHeight As Integer
    Dim iHeight As Long
    Dim iTop As Long
    
    lHeight = 0
    iTop = TOP_START
    
    Set mItemListing = New MenuItemListings
    Set PUMenu.Font = m_Font
    
    For i = 1 To MenuItem.Count
    
        If MenuItem(i).Group = sGroup Then
        
            iHeight = GetHeight(i)
            iTop = iTop + iHeight
        
            lHeight = lHeight + iHeight
            
            Select Case MenuItem(i).ItemType
            Case eItemType.itNormal, eItemType.itCheck
                mItemListing.Add "", i, iTop - iHeight, iHeight, True
            Case Else
                mItemListing.Add "", i, iTop - iHeight, iHeight, False
            End Select
        Else
            'Later!
        End If
    
    Next i
    
    PUMenu.Width = GetTotalWidth
    PUMenu.Height = (lHeight + 2 + TOP_START) * Screen.TwipsPerPixelY
    
    If X + PUMenu.Width > Screen.Width Then X = X - PUMenu.Width
    If Y + PUMenu.Height > Screen.Height Then
        Y = Y - PUMenu.Height
        mGapWidth = 0
    Else
        mGapWidth = M_Width
    End If
    
    DrawFace sGroup
    
    
    Call PUMenu.Move(X, Y)
    Call SetWindowPos(PUMenu.hwnd, HWND_TOPMOST, 0, 0, 0, 0, TOPMOST_FLAGS)
    
    DrawFace
    PUMenu.Visible = True
    Transparency PUMenu.hwnd, 230
    
    'We are visible now!
    MenuVisible = True
        
    'Delayer!!
    '/////////////////////////////////
    Dim KeyLeft As Boolean
    Dim KeyRight As Boolean
    
    KeyLeft = True
    KeyRight = True
    
    Do Until (KeyLeft = False) And (KeyRight = False)
        DoEvents
        KeyLeft = GetAsyncKeyState(VK_LBUTTON)
        KeyRight = GetAsyncKeyState(VK_RBUTTON)
    Loop
    '/////////////////////////////////
    
    tmrFocus.Enabled = True

End Sub

Public Sub KillMenu()
    PUMenu.Visible = False
End Sub

Private Sub tmrFocus_Timer()

    If (InFocusControl(UserControl.hwnd) = True) And (PUMenu.Visible = False) Then
    ElseIf PUMenu.Visible = False Then
        tmrFocus.Enabled = False
        MenuVisible = False
        RaiseEvent Closed
        Exit Sub
    End If
    
    If MenuVisible = False Then MenuVisible = True
    
End Sub

Private Function InFocusControl(ByVal ObjecthWnd As Long) As Boolean

    Dim mPos As POINTAPI
    Dim KeyLeft As Boolean
    Dim KeyRight As Boolean
    Dim oRect As RECT
    Dim oRect2 As RECT
    
    Call GetCursorPos(mPos)
    'Call GetWindowRect(ObjecthWnd, oRect)
    Call GetWindowRect(PUMenu.hwnd, oRect)
        
    KeyLeft = GetAsyncKeyState(VK_LBUTTON)
    KeyRight = GetAsyncKeyState(VK_RBUTTON)
    
    If (mPos.X >= oRect.Left) And (mPos.X <= oRect.Right) And (mPos.Y >= oRect.Top) And (mPos.Y <= oRect.Bottom) Then
        InFocusControl = True
    ElseIf (KeyLeft = True) Or (KeyRight = True) Then
        If (mPos.X < oRect.Left) Or (mPos.X > oRect.Right) Or (mPos.Y < oRect.Top) Or (mPos.Y > oRect.Bottom) Then
            
'        If m_Clicked = True Then
'            picList.Visible = False
'            InFocusControl = False
'            m_Clicked = False
'        End If
        
            'If (mPos.X < oRect2.Left) Or (mPos.X > oRect2.Right) Or (mPos.Y < oRect2.Top) Or (mPos.Y > oRect2.Bottom) Then
                PUMenu.Visible = False
                InFocusControl = False
                'm_Clicked = False
                
            'End If
        
        End If
        
    End If
    
    InFocusControl = False
    
End Function

Private Sub PUMenu_Click()
    If iSelection <> 0 Then RaiseEvent ItemClick(mItemListing(iSelection).Index, MenuItem(mItemListing(iSelection).Index).Key)
    PUMenu.Visible = False
    iSelection = 0
End Sub

Private Sub puMenu_LostFocus()
    tmrFocus.Enabled = True
    PUMenu.Visible = False

End Sub

Private Sub UserControl_Initialize()

    Dim lResult As Long

    On Error Resume Next
    lResult = GetWindowLong(PUMenu.hwnd, GWL_EXSTYLE)
    Call SetWindowLong(PUMenu.hwnd, GWL_EXSTYLE, lResult Or WS_EX_TOOLWINDOW)
    Call SetWindowPos(PUMenu.hwnd, PUMenu.hwnd, 0, 0, 0, 0, 39)
    Call SetWindowLong(PUMenu.hwnd, -8, Parent.hwnd)
    Call SetParent(PUMenu.hwnd, 0)
    
    MenuVisible = False
    
    Set MenuItem = New MenuItems
    DropShadow PUMenu.hwnd
    
End Sub

Private Sub PUMenu_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Dim i As Integer
    
    iSelection = 0
    For i = 1 To mItemListing.Count
        If (Y > mItemListing(i).Top And Y < (mItemListing(i).Top + mItemListing(i).Height)) And mItemListing(i).Selectable = True Then
            iSelection = mItemListing(i).Index
        End If
    Next i
    
    DrawFace
    
End Sub

Private Sub UserControl_Resize()
    
    UserControl.Height = 705
    UserControl.Width = 705
    
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
'MemberInfo=6,0,0,0
Public Property Get Font() As Font
Attribute Font.VB_Description = "Gibt ein Font-Objekt zurück."
Attribute Font.VB_ProcData.VB_Invoke_Property = ";Schriftart"
Attribute Font.VB_UserMemId = -512
    Set Font = m_Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set m_Font = New_Font
    PropertyChanged "Font"
End Property

''ACHTUNG! DIE FOLGENDEN KOMMENTIERTEN ZEILEN NICHT ENTFERNEN ODER VERÄNDERN!
''MemberInfo=14,0,0,0
''Public Property Get MenuItem() As MenuItems
''    Set MenuItem = m_MenuItem
''End Property
''
''Public Property Let MenuItem(ByVal New_MenuItem As MenuItems)
''    m_MenuItem = New_MenuItem
''    PropertyChanged "MenuItem"
''End Property

'Eigenschaften für Benutzersteuerelement initialisieren
Private Sub UserControl_InitProperties()
    m_BackColor = m_def_BackColor
    m_ForeColor = m_def_ForeColor
    Set m_Font = Ambient.Font
    m_SelectionColor = m_def_SelectionColor
    
    RefreshColor
    DrawFace
    
End Sub

'Eigenschaftenwerte vom Speicher laden
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    m_BackColor = PropBag.ReadProperty("BackColor", m_def_BackColor)
    m_ForeColor = PropBag.ReadProperty("ForeColor", m_def_ForeColor)
    Set m_Font = PropBag.ReadProperty("Font", Ambient.Font)
'    m_MenuItem = PropBag.ReadProperty("MenuItem", m_def_MenuItem)
    
    Set m_ImageList = PropBag.ReadProperty("ImageList", Nothing)
    m_SelectionColor = PropBag.ReadProperty("SelectionColor", m_def_SelectionColor)
    
    RefreshColor
    DrawFace
    
End Sub

Private Sub UserControl_Terminate()
    
    If Not (MenuItem Is Nothing) Then Set MenuItem = Nothing
    If Not (mItemListing Is Nothing) Then Set mItemListing = Nothing
    DisableAlpha PUMenu.hwnd
    
End Sub

'Eigenschaftenwerte in den Speicher schreiben
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("BackColor", m_BackColor, m_def_BackColor)
    Call PropBag.WriteProperty("ForeColor", m_ForeColor, m_def_ForeColor)
    Call PropBag.WriteProperty("Font", m_Font, Ambient.Font)
'    Call PropBag.WriteProperty("MenuItem", m_MenuItem, m_def_MenuItem)
    
    Call PropBag.WriteProperty("ImageList", m_ImageList, Nothing)
    Call PropBag.WriteProperty("SelectionColor", m_SelectionColor, m_def_SelectionColor)
    
    RefreshColor
    DrawFace
End Sub

'ACHTUNG! DIE FOLGENDEN KOMMENTIERTEN ZEILEN NICHT ENTFERNEN ODER VERÄNDERN!
'MemberInfo=9,0,0,0
Public Property Get ImageList() As Object
Attribute ImageList.VB_ProcData.VB_Invoke_Property = ";Verschiedenes"
    Set ImageList = m_ImageList
End Property

Public Property Set ImageList(ByVal New_ImageList As Object)
    Set m_ImageList = New_ImageList
    PropertyChanged "ImageList"
    DrawFace
End Property

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

