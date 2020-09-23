VERSION 5.00
Begin VB.UserControl cainSlider 
   AutoRedraw      =   -1  'True
   ClientHeight    =   1815
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5265
   ScaleHeight     =   121
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   351
   Begin VB.Timer MousePos 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   0
      Top             =   0
   End
End
Attribute VB_Name = "cainSlider"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Enum TickStyles

    Tick_None = 0
    Tick_Half = 1
    Tick_Four = 2
    Tick_Half_Value = 3
    Tick_Four_Value = 4
    
End Enum

'Standard-Eigenschaftswerte:
Const m_def_ForeColor = &HC00000
Const m_def_BackColor = &HFFFFFF
Const m_def_TabSelectColor = &H96E7&
Const m_def_Orientation = 1
Const m_def_TickStyle = 0
Const m_def_Min = 0
Const m_def_Max = 10
Const m_def_Value = 5
'Eigenschaftsvariablen:
Dim m_ForeColor As OLE_COLOR
Dim m_BackColor As OLE_COLOR
Dim m_TabSelectColor As OLE_COLOR
Dim m_Orientation As MyOrientation
Dim m_TickStyle As TickStyles
Dim m_Min As Long
Dim m_Max As Long
Dim m_Value As Long
'Ereignisdeklarationen:
Event Click() 'MappingInfo=UserControl,UserControl,-1,Click
Attribute Click.VB_Description = "Tritt auf, wenn der Benutzer eine Maustaste über einem Objekt drückt und wieder losläßt."
Event KeyDown(KeyCode As Integer, Shift As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyDown
Attribute KeyDown.VB_Description = "Tritt auf, wenn der Benutzer eine Taste drückt, während ein Objekt den Fokus besitzt."
Event KeyPress(KeyAscii As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyPress
Attribute KeyPress.VB_Description = "Tritt auf, wenn der Benutzer eine ANSI-Taste drückt und losläßt."
Event KeyUp(KeyCode As Integer, Shift As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyUp
Attribute KeyUp.VB_Description = "Tritt auf, wenn der Benutzer eine Taste losläßt, während ein Objekt den Fokus hat."
Event ValueChange()
Event Slide()

Dim ColorVars(1 To 15) As Long
Dim MyColorSet As ColorSet
Dim MyState As SlideState
Dim MyTabState As TabSelected

Dim Mouse As POINTAPI
Dim MouseOverMe As POINTAPI
Dim MouseClick As Boolean

Dim OldValue As Long
Dim MyLeft As Long

Public Function ID() As Integer
    ID = eControlIDs.id_Slider
End Function

Private Function GetLeftFromValue(lValue As Long) As Long

    Dim i As Long
    Dim i2 As Long
    
    i = lValue - m_Min 'real Value
    If i = 0 Then
        i2 = 0 'if the value = 0 then we know it is 0 %
    Else
        i2 = (i / (m_Max - m_Min)) * 100 'Percent of Value from max
    End If
    If m_Orientation = pbHorizontal Then
        i = (UserControl.ScaleWidth - 20) 'Slider's sliding space in pixel
    ElseIf m_Orientation = pbVertical Then
        i = (UserControl.ScaleHeight - 20) 'Slider's sliding space in pixel
    End If
    MyLeft = Fix((i2 * i) / 100)
    GetLeftFromValue = MyLeft

End Function

Private Function GetValueFromLeft(SliderPos As Single) As Long

    Dim i2 As Long
    Dim i As Long
    
    If SliderPos = 0 Then
        i2 = 0 'if the value = 0 then we know it is 0 %
    Else
        If m_Orientation = pbHorizontal Then
            i2 = (SliderPos / (UserControl.ScaleWidth - 20)) * 100 'Percent of Value from max
        ElseIf m_Orientation = pbVertical Then
            i2 = (SliderPos / (UserControl.ScaleHeight - 20)) * 100 'Percent of Value from max
        End If
    End If
    
    i = Fix((i2 * (m_Max - m_Min)) / 100)
    i = i + m_Min
    If i < m_Min Then i = m_Min
    If i > m_Max Then i = m_Max
    
    GetValueFromLeft = i

End Function

Private Sub DrawSlider(Optional bForceRedraw As Boolean = True)
    
    
    If bForceRedraw = True Then
        UserControl.Cls
        DrawPath
        DrawTicks
    End If
    
    If UserControl.Enabled = True Then DrawSlide
    
End Sub

Private Sub RefreshColors()

    If UserControl.Enabled = True Then
        If MyState = Slide_Disabled Then MyState = Slide_Normal
    Else
         MyState = Slide_Disabled
    End If
           
    If MyState = Slide_Normal And (MyTabState = tbNormal Or MyTabState = tbTabed) Then
        MyColorSet = GetColorSetNormal(m_BackColor, m_ForeColor)
        
    ElseIf MyState = Slide_Hover And (MyTabState = tbNormal Or MyTabState = tbTabed) Then
        MyColorSet = GetColorSetHovered(m_BackColor, m_ForeColor)
        
    ElseIf MyState = Slide_Clicked And (MyTabState = tbNormal Or MyTabState = tbTabed) Then
        MyColorSet = GetColorSetNormal(m_BackColor, m_ForeColor)
        
    ElseIf MyState = Slide_Disabled And (MyTabState = tbNormal Or MyTabState = tbTabed) Then
        MyColorSet = GetColorSetDisabled(m_BackColor, m_ForeColor)
        
    ElseIf MyState = Slide_Unselected And MyTabState = tbNormal Then
        MyColorSet = GetColorSetNormal(m_BackColor, m_ForeColor)
        
    ElseIf MyState = Slide_Unselected And MyTabState = tbTabed Then
        MyColorSet = GetColorSetTabbed(m_BackColor, m_ForeColor, m_TabSelectColor)
        
    End If
        
End Sub

Private Sub DrawSlide()

    Dim X As Long
    
    X = GetLeftFromValue(m_Value) + 5 'temp
    
    If OldValue <> m_Value Then
        RaiseEvent ValueChange
        OldValue = m_Value
    End If
    
    If m_Orientation = pbHorizontal Then
    
        UserControl.Line (X + 10, 4)-(X + 10, 15), MyColorSet.csColor1(6)
        
        UserControl.Line (X + 1, 3)-(X + 10, 3), MyColorSet.csColor1(3)
        UserControl.Line (X, 4)-(X, 15), MyColorSet.csColor1(3)
        
        DrawGradient UserControl.hdc, X + 1, 4, 9, 3, GetRGBColors(MyColorSet.csColor1(2)), GetRGBColors(MyColorSet.csColor1(3))
        DrawGradient UserControl.hdc, X + 1, 7, 7, 8, GetRGBColors(MyColorSet.csColor1(1)), GetRGBColors(MyColorSet.csColor1(1)), 0

        UserControl.Line (X + 9, 6)-(X + 9, 14), MyColorSet.csColor1(3)
        UserControl.Line (X + 8, 7)-(X + 8, 14), MyColorSet.csColor1(2)

        UserControl.Line (X, 15)-(X + 3, 18), MyColorSet.csColor1(2)

        UserControl.Line (X + 3, 15)-(X + 7, 15), MyColorSet.csColor1(1)
        UserControl.PSet (X + 4, 16), MyColorSet.csColor1(1)
        UserControl.PSet (X + 5, 16), MyColorSet.csColor1(1)
        UserControl.PSet (X + 4, 17), MyColorSet.csColor1(1)

        UserControl.PSet (X + 10, 15), MyColorSet.csColor1(6)
        UserControl.PSet (X + 9, 16), MyColorSet.csColor1(6)

        UserControl.PSet (X + 8, 14), MyColorSet.csColor1(2) 'ColorVars(5)
        UserControl.PSet (X + 7, 14), MyColorSet.csColor1(5) 'ColorVars(7)
        UserControl.PSet (X + 7, 15), MyColorSet.csColor1(5) 'ColorVars(7)
        UserControl.PSet (X + 7, 16), MyColorSet.csColor1(2) 'ColorVars(5)
        UserControl.PSet (X + 6, 16), MyColorSet.csColor1(5) 'ColorVars(7)
        UserControl.PSet (X + 6, 17), MyColorSet.csColor1(2) 'ColorVars(5)
        UserControl.PSet (X + 5, 17), MyColorSet.csColor1(5) 'ColorVars(7)
        UserControl.PSet (X + 5, 18), MyColorSet.csColor1(2) 'ColorVars(5)

        UserControl.PSet (X + 9, 14), MyColorSet.csColor1(3)
        UserControl.PSet (X + 8, 15), MyColorSet.csColor1(3)
        UserControl.PSet (X + 7, 17), MyColorSet.csColor1(3)
        UserControl.PSet (X + 6, 18), MyColorSet.csColor1(3)

        UserControl.PSet (X + 8, 17), MyColorSet.csColor1(3)
        UserControl.PSet (X + 6, 19), MyColorSet.csColor1(3)

        UserControl.PSet (X + 9, 15), MyColorSet.csColor1(6)
        UserControl.PSet (X + 8, 16), MyColorSet.csColor1(6)
        UserControl.PSet (X + 7, 18), MyColorSet.csColor1(6)

        UserControl.Line (X + 1, 14)-(X + 4, 17), MyColorSet.csColor1(4)
        UserControl.Line (X + 1, 15)-(X + 4, 18), MyColorSet.csColor1(4)

        UserControl.PSet (X + 3, 18), MyColorSet.csColor1(3)
        UserControl.PSet (X + 4, 19), MyColorSet.csColor1(3)
        UserControl.PSet (X + 5, 20), MyColorSet.csColor1(3)

        UserControl.PSet (X + 4, 18), MyColorSet.csColor1(2)
        UserControl.PSet (X + 5, 19), MyColorSet.csColor1(2)
        
        UserControl.PSet (X + 4, 17), MyColorSet.csColor1(5)
    
    ElseIf m_Orientation = pbVertical Then
    
        UserControl.Line (4, X + 10)-(15, X + 10), MyColorSet.csColor1(6)
        
        UserControl.Line (3, X + 1)-(3, X + 10), MyColorSet.csColor1(3)
        UserControl.Line (4, X)-(15, X), MyColorSet.csColor1(3)
        
        DrawGradient UserControl.hdc, 4, X + 1, 3, 9, GetRGBColors(MyColorSet.csColor1(2)), GetRGBColors(MyColorSet.csColor1(3))
        DrawGradient UserControl.hdc, 7, X + 1, 8, 7, GetRGBColors(MyColorSet.csColor1(1)), GetRGBColors(MyColorSet.csColor1(1)), 0

        UserControl.Line (6, X + 9)-(14, X + 9), MyColorSet.csColor1(3)
        UserControl.Line (7, X + 8)-(14, X + 8), MyColorSet.csColor1(2)

        UserControl.Line (15, X)-(18, X + 3), MyColorSet.csColor1(2)

        UserControl.Line (15, X + 3)-(15, X + 7), MyColorSet.csColor1(1)
        UserControl.PSet (16, X + 4), MyColorSet.csColor1(1)
        UserControl.PSet (16, X + 5), MyColorSet.csColor1(1)
        UserControl.PSet (17, X + 4), MyColorSet.csColor1(1)

        UserControl.PSet (15, X + 10), MyColorSet.csColor1(6)
        UserControl.PSet (16, X + 9), MyColorSet.csColor1(6)

        UserControl.PSet (14, X + 8), MyColorSet.csColor1(2) 'ColorVars(5)
        UserControl.PSet (14, X + 7), MyColorSet.csColor1(5) 'ColorVars(7)
        UserControl.PSet (15, X + 7), MyColorSet.csColor1(5) 'ColorVars(7)
        UserControl.PSet (16, X + 7), MyColorSet.csColor1(2) 'ColorVars(5)
        UserControl.PSet (16, X + 6), MyColorSet.csColor1(5) 'ColorVars(7)
        UserControl.PSet (17, X + 6), MyColorSet.csColor1(2) 'ColorVars(5)
        UserControl.PSet (17, X + 5), MyColorSet.csColor1(5) 'ColorVars(7)
        UserControl.PSet (18, X + 5), MyColorSet.csColor1(2) 'ColorVars(5)

        UserControl.PSet (14, X + 9), MyColorSet.csColor1(3)
        UserControl.PSet (15, X + 8), MyColorSet.csColor1(3)
        UserControl.PSet (17, X + 7), MyColorSet.csColor1(3)
        UserControl.PSet (18, X + 6), MyColorSet.csColor1(3)

        UserControl.PSet (17, X + 8), MyColorSet.csColor1(3)
        UserControl.PSet (19, X + 6), MyColorSet.csColor1(3)

        UserControl.PSet (15, X + 9), MyColorSet.csColor1(6)
        UserControl.PSet (16, X + 8), MyColorSet.csColor1(6)
        UserControl.PSet (18, X + 7), MyColorSet.csColor1(6)

        UserControl.Line (14, X + 1)-(17, X + 4), MyColorSet.csColor1(4)
        UserControl.Line (15, X + 1)-(18, X + 4), MyColorSet.csColor1(4)

        UserControl.PSet (18, X + 3), MyColorSet.csColor1(3)
        UserControl.PSet (19, X + 4), MyColorSet.csColor1(3)
        UserControl.PSet (20, X + 5), MyColorSet.csColor1(3)

        UserControl.PSet (18, X + 4), MyColorSet.csColor1(2)
        UserControl.PSet (19, X + 5), MyColorSet.csColor1(2)
        
        UserControl.PSet (17, X + 4), MyColorSet.csColor1(5)
    
    End If
    
End Sub

Private Sub DrawPath()
    
    UserControl.BackColor = BlendColor(m_BackColor, m_ForeColor, 240) 'MyColorSet.csBackColor
    
    If m_Orientation = pbHorizontal Then
    
        UserControl.Line (8, 8)-(UserControl.ScaleWidth - 8, 8), MyColorSet.csColor1(1)
        UserControl.Line (8, 9)-(UserControl.ScaleWidth - 8, 9), MyColorSet.csColor1(2)
        UserControl.Line (8, 10)-(UserControl.ScaleWidth - 8, 10), MyColorSet.csColor1(2)
        UserControl.Line (8, 11)-(UserControl.ScaleWidth - 8, 11), MyColorSet.csColor1(3)
        
        UserControl.Line (7, 9)-(7, 11), MyColorSet.csColor1(1)
        UserControl.Line (UserControl.ScaleWidth - 8, 9)-(UserControl.ScaleWidth - 8, 11), MyColorSet.csColor1(3)
    
    ElseIf m_Orientation = pbVertical Then
    
        UserControl.Line (8, 8)-(8, UserControl.ScaleHeight - 8), MyColorSet.csColor1(1)
        UserControl.Line (9, 8)-(9, UserControl.ScaleHeight - 8), MyColorSet.csColor1(2)
        UserControl.Line (10, 8)-(10, UserControl.ScaleHeight - 8), MyColorSet.csColor1(2)
        UserControl.Line (11, 8)-(11, UserControl.ScaleHeight - 8), MyColorSet.csColor1(3)
        
        UserControl.Line (9, 7)-(11, 7), MyColorSet.csColor1(1)
        UserControl.Line (9, UserControl.ScaleHeight - 8)-(11, UserControl.ScaleHeight - 8), MyColorSet.csColor1(3)
        
    End If

End Sub

Private Sub DrawTicks()

    Dim i As Long

    Select Case m_TickStyle
        
        Case TickStyles.Tick_None
            'Do Nothing
        
        Case TickStyles.Tick_Half, TickStyles.Tick_Half_Value
            i = m_Max - m_Min
            i = GetLeftFromValue(Fix(i / 2) + m_Min) + 10
            
            If m_Orientation = pbHorizontal Then
                UserControl.Line (i, 15)-(i, 25), MyColorSet.csColor1(7)
                UserControl.Line (10, 15)-(10, 25), MyColorSet.csColor1(7)
                UserControl.Line (UserControl.ScaleWidth - 10, 15)-(UserControl.ScaleWidth - 10, 25), MyColorSet.csColor1(7)
            ElseIf Orientation = pbVertical Then
                UserControl.Line (15, i)-(25, i), MyColorSet.csColor1(7)
                UserControl.Line (15, 10)-(25, 10), MyColorSet.csColor1(7)
                UserControl.Line (15, UserControl.ScaleHeight - 10)-(25, UserControl.ScaleHeight - 10), MyColorSet.csColor1(7)
            End If
            
            DrawCaption Fix((m_Max - m_Min) / 2) + m_Min, ColorVars(1), i
            If m_Orientation = pbHorizontal Then
                DrawCaption Str(m_Max), ColorVars(1), UserControl.ScaleWidth - 10
            Else
                DrawCaption Str(m_Max), ColorVars(1), UserControl.ScaleHeight - 10
            End If
            DrawCaption Str(m_Min), ColorVars(1), 10
            
        Case TickStyles.Tick_Four, TickStyles.Tick_Four_Value
            i = m_Max - m_Min
            i = GetLeftFromValue(Fix(i / 2) + m_Min) + 10
            
            If m_Orientation = pbHorizontal Then
                UserControl.Line (i, 15)-(i, 25), MyColorSet.csColor1(7)
                UserControl.Line (10, 15)-(10, 25), MyColorSet.csColor1(7)
                UserControl.Line (UserControl.ScaleWidth - 10, 15)-(UserControl.ScaleWidth - 10, 25), MyColorSet.csColor1(7)
            ElseIf Orientation = pbVertical Then
                UserControl.Line (15, i)-(25, i), MyColorSet.csColor1(7)
                UserControl.Line (15, 10)-(25, 10), MyColorSet.csColor1(7)
                UserControl.Line (15, UserControl.ScaleHeight - 10)-(25, UserControl.ScaleHeight - 10), MyColorSet.csColor1(7)
            End If
            
            DrawCaption Fix((m_Max - m_Min) / 2) + m_Min, MyColorSet.csColor1(7), i
            
            i = m_Max - m_Min
            i = GetLeftFromValue(Fix(i / 4) + m_Min) + 10
            If m_Orientation = pbHorizontal Then
                UserControl.Line (i, 15)-(i, 20), MyColorSet.csColor1(3)
            ElseIf Orientation = pbVertical Then
                UserControl.Line (15, i)-(20, i), MyColorSet.csColor1(3)
            End If
            
            DrawCaption Fix((m_Max - m_Min) / 4) + m_Min, MyColorSet.csColor1(7), i
            
            i = m_Max - m_Min
            i = GetLeftFromValue(Fix((i / 4) * 3) + m_Min) + 10
            If m_Orientation = pbHorizontal Then
                UserControl.Line (i, 15)-(i, 20), MyColorSet.csColor1(3)
            ElseIf Orientation = pbVertical Then
                UserControl.Line (15, i)-(20, i), MyColorSet.csColor1(3)
            End If
                    
            DrawCaption Fix(((m_Max - m_Min) / 4) * 3) + m_Min, MyColorSet.csColor1(7), i
            
            If m_Orientation = pbHorizontal Then
                DrawCaption Str(m_Max), MyColorSet.csColor1(7), UserControl.ScaleWidth - 10
            Else
                DrawCaption Str(m_Max), MyColorSet.csColor1(7), UserControl.ScaleHeight - 10
            End If
            DrawCaption Str(m_Min), MyColorSet.csColor1(7), 10
            
    End Select

End Sub

Private Sub DrawCaption(strText As String, FntColor As Long, X As Long)
    
    Select Case m_TickStyle
    
    Case TickStyles.Tick_Half_Value, TickStyles.Tick_Four_Value
    
        If m_Orientation = pbHorizontal Then
            UserControl.CurrentX = Fix((X - (UserControl.TextWidth(strText) / 2)))
            UserControl.CurrentY = 26
        ElseIf m_Orientation = pbVertical Then
            UserControl.CurrentX = 26
            UserControl.CurrentY = Fix((X - (UserControl.TextWidth(strText) / 2)))
        End If
    
        UserControl.ForeColor = FntColor
        UserControl.Print strText
    
    End Select

End Sub

Private Sub UserControl_EnterFocus()
    
    MyTabState = tbTabed
    RefreshColors
    DrawSlider
    
End Sub

Private Sub UserControl_ExitFocus()

    MyTabState = tbNormal
    RefreshColors
    DrawSlider
    
End Sub

Private Sub UserControl_Initialize()

    MyState = Slide_Unselected
    MyTabState = tbNormal
    DrawSlider True

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
    RefreshColors
    DrawSlider True
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
    RefreshColors
    DrawSlider True
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
    RefreshColors
    DrawSlider True
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
    DrawSlider True
End Property

Private Sub UserControl_Click()
    RaiseEvent Click
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
    Dim i As Long
    Select Case KeyCode
    
        Case 37
            If m_Orientation = pbHorizontal Then
                i = Value
                i = i - 1
                If i < m_Min Then i = m_Min
                Value = i
            End If
            
        Case 39
            If m_Orientation = pbHorizontal Then
                i = Value
                i = i + 1
                If i > m_Max Then i = m_Max
                Value = i
            End If
    
        Case 38
            If m_Orientation = pbVertical Then
                i = Value
                i = i - 1
                If i < m_Min Then i = m_Min
                Value = i
            End If
            
        Case 40
            If m_Orientation = pbVertical Then
                i = Value
                i = i + 1
                If i > m_Max Then i = m_Max
                Value = i
            End If
        
    End Select
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

'ACHTUNG! DIE FOLGENDEN KOMMENTIERTEN ZEILEN NICHT ENTFERNEN ODER VERÄNDERN!
'MemberInfo=10,0,0,0
Public Property Get TabSelectColor() As OLE_COLOR
Attribute TabSelectColor.VB_ProcData.VB_Invoke_Property = ";Darstellung"
    TabSelectColor = m_TabSelectColor
End Property

Public Property Let TabSelectColor(ByVal New_TabSelectColor As OLE_COLOR)
    m_TabSelectColor = New_TabSelectColor
    PropertyChanged "TabSelectColor"
    RefreshColors
    DrawSlider True
End Property

'ACHTUNG! DIE FOLGENDEN KOMMENTIERTEN ZEILEN NICHT ENTFERNEN ODER VERÄNDERN!
'MemberInfo=7,0,0,0
Public Property Get Orientation() As MyOrientation
Attribute Orientation.VB_ProcData.VB_Invoke_Property = ";Darstellung"
    Orientation = m_Orientation
End Property

Public Property Let Orientation(ByVal New_Orientation As MyOrientation)
    m_Orientation = New_Orientation
    PropertyChanged "Orientation"
    UserControl_Resize
    DrawSlider True
End Property

'ACHTUNG! DIE FOLGENDEN KOMMENTIERTEN ZEILEN NICHT ENTFERNEN ODER VERÄNDERN!
'MemberInfo=7,0,0,0
Public Property Get TickStyle() As TickStyles
Attribute TickStyle.VB_ProcData.VB_Invoke_Property = ";Darstellung"
    TickStyle = m_TickStyle
End Property

Public Property Let TickStyle(ByVal New_TickStyle As TickStyles)
    m_TickStyle = New_TickStyle
    PropertyChanged "TickStyle"
    UserControl_Resize
    DrawSlider True
End Property

'ACHTUNG! DIE FOLGENDEN KOMMENTIERTEN ZEILEN NICHT ENTFERNEN ODER VERÄNDERN!
'MemberInfo=8,0,0,0
Public Property Get Min() As Long
Attribute Min.VB_ProcData.VB_Invoke_Property = ";Verschiedenes"
    Min = m_Min
End Property

Public Property Let Min(ByVal New_Min As Long)
    
    If New_Min > m_Max Then New_Min = m_Max - 1
    If New_Min = m_Max Then New_Min = m_Max - 1
    
    m_Min = New_Min
    PropertyChanged "Min"
    
    If m_Value < m_Min Then
        m_Value = New_Min
        PropertyChanged "Value"
    End If

    DrawSlider True
End Property

'ACHTUNG! DIE FOLGENDEN KOMMENTIERTEN ZEILEN NICHT ENTFERNEN ODER VERÄNDERN!
'MemberInfo=8,0,0,0
Public Property Get Max() As Long
Attribute Max.VB_ProcData.VB_Invoke_Property = ";Verschiedenes"
    Max = m_Max
End Property

Public Property Let Max(ByVal New_Max As Long)

    If New_Max < m_Min Then New_Max = m_Min + 1
    If New_Max = m_Min Then New_Max = m_Min + 1
   
    m_Max = New_Max
    PropertyChanged "Max"
    
    If m_Value > m_Max Then
        m_Value = New_Max
        PropertyChanged "Value"
    End If
    
    DrawSlider True
End Property

'ACHTUNG! DIE FOLGENDEN KOMMENTIERTEN ZEILEN NICHT ENTFERNEN ODER VERÄNDERN!
'MemberInfo=8,0,0,0
Public Property Get Value() As Long
Attribute Value.VB_ProcData.VB_Invoke_Property = ";Verschiedenes"
    Value = m_Value
End Property

Public Property Let Value(ByVal New_Value As Long)
    
    'Prevents Slider from executing this line again and again
    If OldValue = New_Value Then Exit Property
    OldValue = m_Value
    
    m_Value = New_Value
    PropertyChanged "Value"
    DrawSlider True
        
End Property

'Eigenschaften für Benutzersteuerelement initialisieren
Private Sub UserControl_InitProperties()
    m_ForeColor = m_def_ForeColor
    Set UserControl.Font = Ambient.Font
    m_TabSelectColor = m_def_TabSelectColor
    m_Orientation = m_def_Orientation
    m_TickStyle = m_def_TickStyle
    m_Min = m_def_Min
    m_Max = m_def_Max
    m_Value = m_def_Value
    m_BackColor = m_def_BackColor
    
    RefreshColors
    
End Sub

'Eigenschaftenwerte vom Speicher laden
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    m_BackColor = PropBag.ReadProperty("BackColor", m_def_BackColor)
    m_ForeColor = PropBag.ReadProperty("ForeColor", m_def_ForeColor)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    Set UserControl.Font = PropBag.ReadProperty("Font", Ambient.Font)
    m_TabSelectColor = PropBag.ReadProperty("TabSelectColor", m_def_TabSelectColor)
    m_Orientation = PropBag.ReadProperty("Orientation", m_def_Orientation)
    m_TickStyle = PropBag.ReadProperty("TickStyle", m_def_TickStyle)
    m_Min = PropBag.ReadProperty("Min", m_def_Min)
    m_Max = PropBag.ReadProperty("Max", m_def_Max)
    m_Value = PropBag.ReadProperty("Value", m_def_Value)
    
    RefreshColors
    DrawSlider True
    
End Sub

Private Sub UserControl_Resize()
    If m_TickStyle = Tick_None Or m_TickStyle = Tick_Four Or m_TickStyle = Tick_Half Then
        If m_Orientation = pbHorizontal Then
            UserControl.Height = 25 * Screen.TwipsPerPixelY
        ElseIf Orientation = pbVertical Then
            UserControl.Width = 25 * Screen.TwipsPerPixelX
        End If
    Else
        If m_Orientation = pbHorizontal Then
            UserControl.Height = 41 * Screen.TwipsPerPixelY
        ElseIf Orientation = pbVertical Then
            UserControl.Width = (33 + UserControl.TextWidth(m_Max) + UserControl.TextWidth(m_Min)) * Screen.TwipsPerPixelX
        End If
    End If
    DrawSlider True
End Sub

'Eigenschaftenwerte in den Speicher schreiben
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, m_def_BackColor)
    Call PropBag.WriteProperty("ForeColor", m_ForeColor, m_def_ForeColor)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("Font", UserControl.Font, Ambient.Font)
'    Call PropBag.WriteProperty("HoverColor", m_HoverColor, m_def_HoverColor)
    Call PropBag.WriteProperty("TabSelectColor", m_TabSelectColor, m_def_TabSelectColor)
    Call PropBag.WriteProperty("Orientation", m_Orientation, m_def_Orientation)
    Call PropBag.WriteProperty("TickStyle", m_TickStyle, m_def_TickStyle)
    Call PropBag.WriteProperty("Min", m_Min, m_def_Min)
    Call PropBag.WriteProperty("Max", m_Max, m_def_Max)
    Call PropBag.WriteProperty("Value", m_Value, m_def_Value)
    
    RefreshColors
    DrawSlider True
End Sub

Private Sub MousePos_Timer()

    GetCursorPos Mouse

    If ((MouseOverMe.X > Mouse.X - 2) And (MouseOverMe.X < Mouse.X + 2)) And ((MouseOverMe.Y > Mouse.Y - 2) And (MouseOverMe.Y < Mouse.Y + 2)) Then
        If MouseClick = True Then
            MyState = Slide_Clicked
            RefreshColors
            DrawSlider
            RaiseEvent Click
        Else
            MyState = Slide_Hover
            RefreshColors
            DrawSlider
        End If
    Else
        MyState = Slide_Unselected
        RefreshColors
        DrawSlider
        MousePos.Enabled = False
    End If

    DoEvents
    
    'RaiseEvent MouseOver

End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> 1 Then Exit Sub
    MouseClick = True
    If Button = 1 Then
        If m_Orientation = pbHorizontal Then
            Value = GetValueFromLeft(X - 10)
        ElseIf m_Orientation = pbVertical Then
            Value = GetValueFromLeft(Y - 10)
        End If
    End If
    
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If m_Orientation = pbHorizontal Then
        If (((X > MyLeft + 5) And (X < (MyLeft + 15)) And ((Y > 3) And (Y < 20)))) Or MouseClick = True Then
            GetCursorPos MouseOverMe
        
                If MousePos.Enabled = False Then MousePos.Enabled = True
            
            If Button = 1 Then
                Value = GetValueFromLeft(X - 10)
                RaiseEvent Slide
            End If
        End If
    ElseIf m_Orientation = pbVertical Then
        If (((Y > MyLeft + 5) And (Y < (MyLeft + 15)) And ((X > 3) And (X < 20)))) Or MouseClick = True Then
            GetCursorPos MouseOverMe
        
                If MousePos.Enabled = False Then MousePos.Enabled = True
            
            If Button = 1 Then
                Value = GetValueFromLeft(Y - 10)
                RaiseEvent Slide
            End If
        End If
    End If

End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If Button <> 1 Then Exit Sub
    If MouseClick = True Then RaiseEvent Click
    MouseClick = False
    
End Sub
