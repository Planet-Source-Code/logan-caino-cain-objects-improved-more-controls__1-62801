VERSION 5.00
Begin VB.UserControl cainMonthview 
   AutoRedraw      =   -1  'True
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   Begin cainObjects.cainPUMenu cainPUMenu1 
      Height          =   705
      Left            =   960
      TabIndex        =   0
      Top             =   960
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
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   0
      Top             =   480
   End
   Begin VB.Timer MousePos 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   0
      Top             =   0
   End
End
Attribute VB_Name = "cainMonthview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Type tDateHolder

    Top As Integer
    Left As Integer
    Height As Integer
    Width As Integer
    DayNum As Integer

End Type

Private Type tButtonHolder

    Top As Integer
    Left As Integer
    Height As Integer
    Width As Integer

End Type

Private Const DAYS_COL = 7
Private Const DAYS_ROW = 6
Private Const ITEM_GAPX = 3
Private Const ITEM_GAPY = 1
'Standard-Eigenschaftswerte:
Const m_def_WeekdayName = "Mo,Tu,We,Th,Fr,Sa,Su"
Const m_def_MonthName = "January,February,March,April,May,June,July,August,September,October,November,December"
Const m_def_BackColor = &HFFFFFF
Const m_def_ForeColor = &HC00000
Const m_def_SelectionColor = &H80FF&
Const m_def_Day = 1
Const m_def_Month = 2
Const m_def_Year = 2005
'Eigenschaftsvariablen:
Dim m_WeekdayName As String
Dim m_MonthName As String
Dim m_BackColor As OLE_COLOR
Dim m_ForeColor As OLE_COLOR
Dim m_SelectionColor As OLE_COLOR
Dim m_Day As Integer
Dim m_Month As Integer
Dim m_Year As Integer
'Ereignisdeklarationen:
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseDown
Attribute MouseDown.VB_Description = "Tritt auf, wenn der Benutzer die Maustaste drückt, während ein Objekt den Fokus hat."
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseMove
Attribute MouseMove.VB_Description = "Tritt auf, wenn der Benutzer die Maus bewegt."
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseUp
Attribute MouseUp.VB_Description = "Tritt auf, wenn der Benutzer die Maustaste losläßt, während ein Objekt den Fokus hat."
Event Click()
Event DblClick()

Dim Mouse As POINTAPI
Dim MouseOverMe As POINTAPI
Dim MouseClick As Boolean

Dim DHolder(1 To 42) As tDateHolder
Dim iHoveredItem As Integer
Dim tColorSet As ColorSet

Dim Side_Gap As Integer
Dim TOP_GAP  As Integer
Dim Title_Height As Long
Dim sMonths() As String
Dim sWDays() As String

Dim TitleButtons(1 To 5) As tButtonHolder
Dim iSelectedButton As Integer
Public bMenuOpen As Boolean

Public Function ID() As Integer
    ID = eControlIDs.id_Monthview
End Function

Private Sub RefreshColor()

    tColorSet = GetColorSetNormal(m_BackColor, m_ForeColor, m_SelectionColor)
    UserControl.BackColor = tColorSet.csBackColor
    cainPUMenu1.BackColor = m_BackColor
    cainPUMenu1.ForeColor = m_ForeColor
    cainPUMenu1.SelectionColor = m_SelectionColor
    cainPUMenu1.Font = UserControl.Font

End Sub

Private Sub CalculateBoxes()

    Dim ir As Integer
    Dim iX As Integer
    Dim iy As Integer
    Dim iCounter As Integer
    Dim Item_Height As Integer
    Dim Item_Width As Integer
    
    Side_Gap = UserControl.TextWidth("WWO")
    TOP_GAP = Side_Gap
    Title_Height = UserControl.TextHeight("I") * 2
    
    ir = 0
    iX = Side_Gap
    iy = Title_Height + TOP_GAP
    iCounter = 0
    Item_Width = UserControl.TextWidth("WW")
    Item_Height = UserControl.TextHeight("I")
    
    Do
        
        iCounter = iCounter + 1
        DHolder(iCounter).Top = iy
        DHolder(iCounter).Left = iX
        DHolder(iCounter).Height = Item_Height
        DHolder(iCounter).Width = Item_Width
        
        ir = ir + 1
        iX = iX + ITEM_GAPX + Item_Width
        
        If ir = DAYS_COL Then
            ir = 0
            iX = Side_Gap
            iy = iy + ITEM_GAPY + Item_Height
        End If
        
        If iCounter = 42 Then Exit Do
        
    Loop
    
    UserControl_Resize

End Sub

Private Sub CalcuButtons()

    Dim iCounter As Integer
    Dim iMaxWidth As Integer
    
    'Get Month and Day Names
    sMonths = Split(m_MonthName, ",")
    sWDays = Split(m_WeekdayName, ",")
    ReDim Preserve sMonths(11)
    ReDim Preserve sWDays(6)

    For iCounter = 0 To 11
        If UserControl.TextWidth(sMonths(iCounter)) > iMaxWidth Then
            iMaxWidth = UserControl.TextWidth(sMonths(iCounter) & " " & m_Year)
        End If
    Next iCounter
    
    TitleButtons(1).Height = UserControl.TextHeight("I") + 6
    TitleButtons(1).Width = iMaxWidth + 32
    TitleButtons(1).Top = (Title_Height / 2) - (TitleButtons(1).Height / 2)
    TitleButtons(1).Left = (UserControl.ScaleWidth / 2) - (TitleButtons(1).Width / 2)

    cainPUMenu1.ClearItems
    For iCounter = 0 To 11
        cainPUMenu1.MenuItem.Add sMonths(iCounter), sMonths(iCounter)
    Next iCounter
    
    TitleButtons(2).Height = TitleButtons(1).Height
    TitleButtons(2).Width = 16
    TitleButtons(2).Top = TitleButtons(1).Top
    TitleButtons(2).Left = TitleButtons(1).Left - TitleButtons(2).Width
    
    TitleButtons(3).Height = TitleButtons(2).Height
    TitleButtons(3).Width = 16
    TitleButtons(3).Top = TitleButtons(2).Top
    TitleButtons(3).Left = TitleButtons(2).Left - TitleButtons(3).Width
    
    TitleButtons(4).Height = Fix(TitleButtons(2).Height / 2)
    TitleButtons(4).Width = 16
    TitleButtons(4).Top = TitleButtons(2).Top
    TitleButtons(4).Left = TitleButtons(1).Left + TitleButtons(1).Width
    
    TitleButtons(5).Height = Fix(TitleButtons(2).Height / 2)
    TitleButtons(5).Width = 16
    TitleButtons(5).Top = TitleButtons(4).Top + TitleButtons(4).Height - 1
    TitleButtons(5).Left = TitleButtons(1).Left + TitleButtons(1).Width

End Sub

Private Sub DrawMonthView()

    Dim i As Integer
    UserControl.Cls
    DrawSelection iHoveredItem, 0
    
    If UserControl.Enabled = True Then
        UserControl.ForeColor = tColorSet.csColor1(7)
    Else
        UserControl.ForeColor = tColorSet.csColor1(6)
    End If
    For i = 1 To 42
        
        If Month(Date) = m_Month And Year(Date) = m_Year And DHolder(i).DayNum = Day(Date) Then
            If DHolder(i).DayNum = m_Day And iHoveredItem = i Then
            ElseIf DHolder(i).DayNum = m_Day Then
            Else
                DrawSelection i, 3
            End If

        End If
        
        If DHolder(i).DayNum = m_Day And iHoveredItem = i Then
            DrawSelection i, 2
        ElseIf DHolder(i).DayNum = m_Day Then
            DrawSelection i, 1
        End If
        
        UserControl.CurrentX = (DHolder(i).Left - 1) + (Fix(DHolder(i).Width / 2) - Fix(UserControl.TextWidth(Str(DHolder(i).DayNum)) / 2))
        UserControl.CurrentY = (DHolder(i).Top) + (Fix(DHolder(i).Height / 2) - Fix(UserControl.TextHeight("0") / 2))
        
        If DHolder(i).DayNum <> 0 Then UserControl.Print Str(DHolder(i).DayNum)
        
    Next i
    
    UserControl.Line (0, 0)-(0, UserControl.ScaleHeight), tColorSet.csColor1(6)
    UserControl.Line (0, 0)-(UserControl.ScaleWidth, 0), tColorSet.csColor1(6)
    UserControl.Line (UserControl.ScaleWidth - 1, 0)-(UserControl.ScaleWidth - 1, UserControl.ScaleHeight), tColorSet.csColor1(6)
    UserControl.Line (0, UserControl.ScaleHeight - 1)-(UserControl.ScaleWidth - 1, UserControl.ScaleHeight - 1), tColorSet.csColor1(6)
    
    DrawGradient UserControl.hdc, 2, 2, UserControl.ScaleWidth - 4, Title_Height, GetRGBColors(tColorSet.csColor1(3)), GetRGBColors(tColorSet.csColor1(1)), 0
    
    UserControl.Line (DHolder(1).Left, DHolder(1).Top - 5)-(DHolder(DAYS_COL).Left + DHolder(DAYS_COL).Width, DHolder(DAYS_COL).Top - 5), tColorSet.csColor1(3)
    UserControl.Line (DHolder(1).Left - 5, DHolder(1).Top)-(DHolder(DAYS_ROW * DAYS_ROW).Left - 5, DHolder(DAYS_ROW * DAYS_ROW).Top + DHolder(DAYS_ROW * DAYS_ROW).Height), tColorSet.csColor1(3)

    UserControl.ForeColor = tColorSet.csColor1(7)
    For i = 1 To DAYS_COL
        
        UserControl.CurrentX = (DHolder(i).Left - 1) + (Fix(DHolder(i).Width / 2) - Fix(UserControl.TextWidth(sWDays(i - 1)) / 2))
        UserControl.CurrentY = Title_Height + (Fix(TOP_GAP / 2) - Fix(UserControl.TextHeight("I") / 2))
        
        UserControl.Print sWDays(i - 1)
    
    Next i
    
    If (iSelectedButton <> 0 And MouseClick = False) And bMenuOpen = False Then
        ButtonSelection 0
    ElseIf (iSelectedButton <> 0 And MouseClick = True) Or bMenuOpen = True Then
        ButtonSelection 1
    End If
    
    UserControl.FontBold = True
    UserControl.ForeColor = tColorSet.csBackColor
    UserControl.CurrentY = TitleButtons(1).Top + ((TitleButtons(1).Height / 2) - (UserControl.TextHeight("I") / 2))
    UserControl.CurrentX = TitleButtons(1).Left + ((TitleButtons(1).Width / 2) - (UserControl.TextWidth(sMonths(m_Month - 1) & " " & m_Year) / 2))
    UserControl.Print sMonths(m_Month - 1) & " " & m_Year
    UserControl.FontBold = False
    
    Dim stFont As String
    stFont = UserControl.FontName
    i = UserControl.FontSize
    
    UserControl.FontName = "webdings"
    UserControl.FontSize = 8
    
    UserControl.ForeColor = tColorSet.csBackColor
    UserControl.CurrentY = TitleButtons(2).Top + ((TitleButtons(2).Height / 2) - (UserControl.TextHeight("4") / 2))
    UserControl.CurrentX = TitleButtons(2).Left + ((TitleButtons(2).Width / 2) - (UserControl.TextWidth("4") / 2))
    UserControl.Print "4"
    
    UserControl.CurrentY = TitleButtons(3).Top + ((TitleButtons(3).Height / 2) - (UserControl.TextHeight("3") / 2))
    UserControl.CurrentX = TitleButtons(3).Left + ((TitleButtons(3).Width / 2) - (UserControl.TextWidth("3") / 2))
    UserControl.Print "3"
    
    UserControl.CurrentY = TitleButtons(4).Top + ((TitleButtons(4).Height / 2) - (UserControl.TextHeight("5") / 2))
    UserControl.CurrentX = TitleButtons(4).Left + ((TitleButtons(4).Width / 2) - (UserControl.TextWidth("5") / 2)) + 1
    UserControl.Print "5"
    
    UserControl.CurrentY = TitleButtons(5).Top + ((TitleButtons(5).Height / 2) - (UserControl.TextHeight("6") / 2))
    UserControl.CurrentX = TitleButtons(5).Left + ((TitleButtons(5).Width / 2) - (UserControl.TextWidth("6") / 2)) + 1
    UserControl.Print "6"
    
    UserControl.FontName = stFont
    UserControl.FontSize = i

End Sub

Private Sub ButtonSelection(iIndex As Integer)
    
    If iSelectedButton = 0 Then Exit Sub
    If iIndex = 0 Then
        DrawGradient UserControl.hdc, TitleButtons(iSelectedButton).Left * 1, TitleButtons(iSelectedButton).Top * 1, TitleButtons(iSelectedButton).Width * 1, TitleButtons(iSelectedButton).Height * 1, GetRGBColors(tColorSet.csColor1(8)), GetRGBColors(tColorSet.csColor1(10)), 1
    Else
        DrawGradient UserControl.hdc, TitleButtons(iSelectedButton).Left * 1, TitleButtons(iSelectedButton).Top * 1, TitleButtons(iSelectedButton).Width * 1, TitleButtons(iSelectedButton).Height * 1, GetRGBColors(tColorSet.csColor1(1)), GetRGBColors(tColorSet.csColor1(3)), 1
    End If
    
    UserControl.Line (TitleButtons(iSelectedButton).Left, TitleButtons(iSelectedButton).Top)-(TitleButtons(iSelectedButton).Left, TitleButtons(iSelectedButton).Top + TitleButtons(iSelectedButton).Height), tColorSet.csColor1(6)
    UserControl.Line (TitleButtons(iSelectedButton).Left, TitleButtons(iSelectedButton).Top)-(TitleButtons(iSelectedButton).Left + TitleButtons(iSelectedButton).Width, TitleButtons(iSelectedButton).Top), tColorSet.csColor1(6)
    UserControl.Line (TitleButtons(iSelectedButton).Left + TitleButtons(iSelectedButton).Width, TitleButtons(iSelectedButton).Top)-(TitleButtons(iSelectedButton).Left + TitleButtons(iSelectedButton).Width, TitleButtons(iSelectedButton).Top + TitleButtons(iSelectedButton).Height), tColorSet.csColor1(6)
    UserControl.Line (TitleButtons(iSelectedButton).Left, TitleButtons(iSelectedButton).Top + TitleButtons(iSelectedButton).Height)-(TitleButtons(iSelectedButton).Left + TitleButtons(iSelectedButton).Width + 1, TitleButtons(iSelectedButton).Top + TitleButtons(iSelectedButton).Height), tColorSet.csColor1(6)

End Sub

Private Sub CalcuMonthDays()

    Dim iMonthMax As Integer
    Dim iRemain As Integer
    Dim i As Integer
    Dim iStartDay As Integer
    Dim iC As Integer

    If (m_Month = 4) Or (m_Month = 6) Or (m_Month = 9) Or (m_Month = 11) Then
        iMonthMax = 30
    ElseIf m_Month = 2 Then
    
        iRemain = GetRemainder(m_Year, 4)
        
        If iRemain = 0 Then
        
            iRemain = GetRemainder(m_Year, 100)
            
            If iRemain = 0 Then
            
                iRemain = GetRemainder(m_Year, 400)
                
                If iRemain = 0 Then
                     iMonthMax = 29
                Else
                     iMonthMax = 28
                End If
                
            Else
                iMonthMax = 29
            End If
            
        Else
            iMonthMax = 28
        End If
        
    Else
        iMonthMax = 31
    End If
    
    iStartDay = Weekday(m_Month & "." & m_Year, vbMonday)
    
    'Assign Day to Slots
    iC = 0
    For i = 1 To 42
    
        If i >= iStartDay And i <= iMonthMax + iStartDay - 1 Then
            iC = iC + 1
            DHolder(i).DayNum = iC
        Else
            DHolder(i).DayNum = 0
        End If
    
    Next i
    
End Sub

Private Function GetRemainder(iDividend As Integer, iDivisor As Integer) As Integer
    
    Dim i As Integer
    
    i = Int(iDividend / iDivisor)
    GetRemainder = iDividend - (i * iDivisor)
    
End Function

Private Sub DrawSelection(iIndex As Integer, iGradientColor As Integer)

    If iIndex = 0 Then Exit Sub
    
    If iGradientColor = 0 Then
        DrawGradient UserControl.hdc, DHolder(iIndex).Left * 1, DHolder(iIndex).Top * 1, DHolder(iIndex).Width * 1, DHolder(iIndex).Height * 1, GetRGBColors(tColorSet.csColor1(8)), GetRGBColors(tColorSet.csColor1(10)), 0
    ElseIf iGradientColor = 1 Then
        DrawGradient UserControl.hdc, DHolder(iIndex).Left * 1, DHolder(iIndex).Top * 1, DHolder(iIndex).Width * 1, DHolder(iIndex).Height * 1, GetRGBColors(tColorSet.csColor1(1)), GetRGBColors(tColorSet.csColor1(3)), 0
    ElseIf iGradientColor = 2 Then
        DrawGradient UserControl.hdc, DHolder(iIndex).Left * 1, DHolder(iIndex).Top * 1, DHolder(iIndex).Width * 1, DHolder(iIndex).Height * 1, GetRGBColors(BlendColor(tColorSet.csColor1(1), tColorSet.csColor1(8))), GetRGBColors(BlendColor(tColorSet.csColor1(10), tColorSet.csColor1(3))), 0
    ElseIf iGradientColor = 3 Then
        DrawGradient UserControl.hdc, DHolder(iIndex).Left * 1, DHolder(iIndex).Top * 1, DHolder(iIndex).Width * 1, DHolder(iIndex).Height * 1, GetRGBColors(tColorSet.csColor1(1)), GetRGBColors(tColorSet.csColor1(1)), 0
    End If
    
    UserControl.Line (DHolder(iIndex).Left, DHolder(iIndex).Top)-(DHolder(iIndex).Left, DHolder(iIndex).Top + DHolder(iIndex).Height), tColorSet.csColor1(6)
    UserControl.Line (DHolder(iIndex).Left, DHolder(iIndex).Top)-(DHolder(iIndex).Left + DHolder(iIndex).Width, DHolder(iIndex).Top), tColorSet.csColor1(6)
    UserControl.Line (DHolder(iIndex).Left + DHolder(iIndex).Width, DHolder(iIndex).Top)-(DHolder(iIndex).Left + DHolder(iIndex).Width, DHolder(iIndex).Top + DHolder(iIndex).Height), tColorSet.csColor1(6)
    UserControl.Line (DHolder(iIndex).Left, DHolder(iIndex).Top + DHolder(iIndex).Height)-(DHolder(iIndex).Left + DHolder(iIndex).Width + 1, DHolder(iIndex).Top + DHolder(iIndex).Height), tColorSet.csColor1(6)

End Sub

Private Sub cainPUMenu1_Closed()
    bMenuOpen = False
    iSelectedButton = 0
    DrawMonthView
End Sub

Private Sub cainPUMenu1_ItemClick(ItemIndex As Integer, ItemKey As String)
    iMonth = ItemIndex
End Sub

Private Sub MousePos_Timer()

    GetCursorPos Mouse
    
    If ((MouseOverMe.X > Mouse.X - 2) And (MouseOverMe.X < Mouse.X + 2)) And ((MouseOverMe.Y > Mouse.Y - 2) And (MouseOverMe.Y < Mouse.Y + 2)) Then
    Else
        iHoveredItem = 0
        If bMenuOpen = False Then iSelectedButton = 0
        DrawMonthView
        MousePos.Enabled = False
    End If

    DoEvents
    
End Sub

Private Sub Timer1_Timer()
    ButtonEvent
End Sub

Private Sub UserControl_DblClick()
    If iHoveredItem = 0 Then Exit Sub
    If m_Day = DHolder(iHoveredItem).DayNum Then
        RaiseEvent DblClick
    End If
End Sub

Private Sub UserControl_Initialize()
    
    cainPUMenu1.Top = -cainPUMenu1.Height - 10
    cainPUMenu1.Left = -cainPUMenu1.Width - 10
    bMenuOpen = False

End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If Button <> 1 Then Exit Sub
    MouseClick = True

    Dim i As Integer
    
    GetCursorPos MouseOverMe
    If MousePos.Enabled = False Then MousePos.Enabled = True
    
    For i = 1 To 42
        If ((X > DHolder(i).Left) And (X < DHolder(i).Left + DHolder(i).Width)) And ((Y > DHolder(i).Top) And (Y < DHolder(i).Top + DHolder(i).Height)) And DHolder(i).DayNum <> 0 Then
            If MouseClick = True Then: _
                m_Day = DHolder(i).DayNum: _
                RaiseEvent Click
            iHoveredItem = i
            Exit For
        End If
    Next i
    ButtonEvent
    Timer1.Enabled = True
    DrawMonthView
    RaiseEvent MouseDown(Button, Shift, X, Y)

End Sub

Private Sub ButtonEvent()
    
    If MouseClick <> True Then Exit Sub
    
    Select Case iSelectedButton
    Case 1
        
        Dim oRect As RECT
        Call GetWindowRect(hwnd, oRect)
        
        cainPUMenu1.CreateMenu (oRect.Left + TitleButtons(1).Left) * Screen.TwipsPerPixelX, (oRect.Top + TitleButtons(1).Top + TitleButtons(1).Height) * Screen.TwipsPerPixelY, , TitleButtons(1).Width
        bMenuOpen = True
        
    Case 2
        iMonth = m_Month + 1
        
    Case 3
        iMonth = m_Month - 1
        
    Case 4
        iYear = m_Year + 1
        
    Case 5
        iYear = m_Year - 1
    
    End Select
    
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Dim i As Integer
    
    GetCursorPos MouseOverMe
    If MousePos.Enabled = False Then MousePos.Enabled = True
    
    For i = 1 To 42
        If ((X > DHolder(i).Left) And (X < DHolder(i).Left + DHolder(i).Width)) And ((Y > DHolder(i).Top) And (Y < DHolder(i).Top + DHolder(i).Height)) And DHolder(i).DayNum <> 0 Then
            iHoveredItem = i
            DrawMonthView
            Exit For
        End If
    Next i
    
    For i = 1 To 5
        If X > TitleButtons(i).Left And X < TitleButtons(i).Left + TitleButtons(i).Width And Y > TitleButtons(i).Top And Y < TitleButtons(i).Top + TitleButtons(i).Height Then
            If iSelectedButton <> i Then
                iSelectedButton = i
                cainPUMenu1.KillMenu
            End If
            DrawMonthView
            Exit For
        End If
    Next i
    If i > 5 And bMenuOpen = False Then
        iSelectedButton = 0
        DrawMonthView
    End If
    
    RaiseEvent MouseMove(Button, Shift, X, Y)

End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If Button <> 1 Then Exit Sub
    MouseClick = False
    Timer1.Enabled = False
    RaiseEvent MouseUp(Button, Shift, X, Y)
    
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
    DrawMonthView
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
    DrawMonthView
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
    RefreshColor
    DrawMonthView
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
    
    RefreshColor
    CalculateBoxes
    CalcuButtons
    DrawMonthView
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
    DrawMonthView
End Property

'ACHTUNG! DIE FOLGENDEN KOMMENTIERTEN ZEILEN NICHT ENTFERNEN ODER VERÄNDERN!
'MemberInfo=7,0,0,1
Public Property Get iDay() As Integer
Attribute iDay.VB_ProcData.VB_Invoke_Property = ";Darstellung"
    iDay = m_Day
End Property

Public Property Let iDay(ByVal New_Day As Integer)
    m_Day = New_Day
    PropertyChanged "iDay"
    CalcuMonthDays
    DrawMonthView
End Property

'ACHTUNG! DIE FOLGENDEN KOMMENTIERTEN ZEILEN NICHT ENTFERNEN ODER VERÄNDERN!
'MemberInfo=7,0,0,1
Public Property Get iMonth() As Integer
Attribute iMonth.VB_ProcData.VB_Invoke_Property = ";Darstellung"
    iMonth = m_Month
End Property

Public Property Let iMonth(ByVal New_Month As Integer)
    If New_Month >= 13 Then New_Month = 12
    If New_Month <= 0 Then New_Month = 1
    m_Month = New_Month
    PropertyChanged "iMonth"
    CalcuMonthDays
    DrawMonthView
End Property

'ACHTUNG! DIE FOLGENDEN KOMMENTIERTEN ZEILEN NICHT ENTFERNEN ODER VERÄNDERN!
'MemberInfo=7,0,0,2005
Public Property Get iYear() As Integer
Attribute iYear.VB_ProcData.VB_Invoke_Property = ";Darstellung"
    iYear = m_Year
End Property

Public Property Let iYear(ByVal New_Year As Integer)
    m_Year = New_Year
    PropertyChanged "iYear"
    CalcuMonthDays
    DrawMonthView
End Property

'Eigenschaften für Benutzersteuerelement initialisieren
Private Sub UserControl_InitProperties()
    m_BackColor = m_def_BackColor
    m_ForeColor = m_def_ForeColor
    Set UserControl.Font = Ambient.Font
    m_SelectionColor = m_def_SelectionColor
    m_Day = m_def_Day
    m_Month = m_def_Month
    m_Year = m_def_Year
    m_WeekdayName = m_def_WeekdayName
    m_MonthName = m_def_MonthName
    
    CalculateBoxes
    CalcuButtons
    RefreshColor
    CalcuMonthDays
    DrawMonthView
    
End Sub

'Eigenschaftenwerte vom Speicher laden
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    m_BackColor = PropBag.ReadProperty("BackColor", m_def_BackColor)
    m_ForeColor = PropBag.ReadProperty("ForeColor", m_def_ForeColor)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    Set UserControl.Font = PropBag.ReadProperty("Font", Ambient.Font)
    m_SelectionColor = PropBag.ReadProperty("SelectionColor", m_def_SelectionColor)
    m_Day = PropBag.ReadProperty("Day", m_def_Day)
    m_Month = PropBag.ReadProperty("Month", m_def_Month)
    m_Year = PropBag.ReadProperty("Year", m_def_Year)
    m_WeekdayName = PropBag.ReadProperty("WeekdayName", m_def_WeekdayName)
    m_MonthName = PropBag.ReadProperty("MonthName", m_def_MonthName)
    
    CalculateBoxes
    CalcuButtons
    RefreshColor
    CalcuMonthDays
    DrawMonthView
    
End Sub

Private Sub UserControl_Resize()
    
    On Error Resume Next
    UserControl.Width = (DHolder(DAYS_COL).Left + DHolder(DAYS_COL).Width + Fix(Side_Gap / 2)) * Screen.TwipsPerPixelX
    UserControl.Height = (DHolder(DAYS_COL * DAYS_ROW).Top + DHolder(DAYS_COL * DAYS_ROW).Height + Fix(Side_Gap / 2)) * Screen.TwipsPerPixelY

End Sub

'Eigenschaftenwerte in den Speicher schreiben
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("BackColor", m_BackColor, m_def_BackColor)
    Call PropBag.WriteProperty("ForeColor", m_ForeColor, m_def_ForeColor)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("Font", UserControl.Font, Ambient.Font)
    Call PropBag.WriteProperty("SelectionColor", m_SelectionColor, m_def_SelectionColor)
    Call PropBag.WriteProperty("Day", m_Day, m_def_Day)
    Call PropBag.WriteProperty("Month", m_Month, m_def_Month)
    Call PropBag.WriteProperty("Year", m_Year, m_def_Year)
    Call PropBag.WriteProperty("WeekdayName", m_WeekdayName, m_def_WeekdayName)
    Call PropBag.WriteProperty("MonthName", m_MonthName, m_def_MonthName)
    
    CalculateBoxes
    CalcuButtons
    RefreshColor
    CalcuMonthDays
    DrawMonthView
    
End Sub

'ACHTUNG! DIE FOLGENDEN KOMMENTIERTEN ZEILEN NICHT ENTFERNEN ODER VERÄNDERN!
'MemberInfo=13,0,0,Mo,Tu,We,Th,Fr,Sa,Su
Public Property Get WeekdayName() As String
Attribute WeekdayName.VB_ProcData.VB_Invoke_Property = ";Darstellung"
    WeekdayName = m_WeekdayName
End Property

Public Property Let WeekdayName(ByVal New_WeekdayName As String)
    m_WeekdayName = New_WeekdayName
    PropertyChanged "WeekdayName"
    CalcuButtons
    CalcuMonthDays
    DrawMonthView
End Property

'ACHTUNG! DIE FOLGENDEN KOMMENTIERTEN ZEILEN NICHT ENTFERNEN ODER VERÄNDERN!
'MemberInfo=13,0,0,Jan,Feb,Mar,Apr,May,Jun,Jul,Aug,Sep,Oct,Nov,Dec
Public Property Get MonthName() As String
Attribute MonthName.VB_ProcData.VB_Invoke_Property = ";Darstellung"
    MonthName = m_MonthName
End Property

Public Property Let MonthName(ByVal New_MonthName As String)
    m_MonthName = New_MonthName
    PropertyChanged "MonthName"
    CalcuButtons
    CalcuMonthDays
    DrawMonthView
End Property

