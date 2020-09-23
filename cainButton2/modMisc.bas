Attribute VB_Name = "modMisc"
Option Explicit

Public Sub RGBRev(R As Integer, G As Integer, B As Integer, lColor As Long)

    Dim i As Double
    
    On Error GoTo ErrorTrap
    
    i = lColor / 65536
    B = Int(i)
    
    i = (i - Int(i)) * 256
    G = Int(i)
    
    i = (i - Int(i)) * 256
    R = Int(i)
    
    Exit Sub

ErrorTrap:
    R = 255
    G = 255
    B = 255
    
End Sub

Public Function GetColor(rgbtColor As RGB) As Long
    GetColor = RGB(rgbtColor.R, rgbtColor.G, rgbtColor.B)
End Function

Public Function GetColorSetNormal(m_BackColor As OLE_COLOR, m_ForeColor As OLE_COLOR, Optional m_SelectionColor As OLE_COLOR = &HFFFFFF) As ColorSet

    Dim tColorSet As ColorSet
    
    tColorSet.csBackColor = m_BackColor
    tColorSet.csFrontColor = m_ForeColor
    
    tColorSet.csColor1(1) = BlendColor(m_BackColor, m_ForeColor, 230)
    tColorSet.csColor1(2) = BlendColor(m_BackColor, m_ForeColor, 190)
    tColorSet.csColor1(3) = BlendColor(m_BackColor, m_ForeColor, 170)
    tColorSet.csColor1(4) = BlendColor(m_BackColor, m_ForeColor, 220)
    tColorSet.csColor1(5) = BlendColor(m_BackColor, m_ForeColor, 200)
    tColorSet.csColor1(6) = BlendColor(m_BackColor, m_ForeColor, 120)
    tColorSet.csColor1(7) = BlendColor(vbBlack, m_ForeColor, 140)
    
    tColorSet.csColor1(8) = BlendColor(m_BackColor, m_SelectionColor, 230)
    tColorSet.csColor1(9) = BlendColor(m_BackColor, m_SelectionColor, 180)
    tColorSet.csColor1(10) = BlendColor(m_BackColor, m_SelectionColor, 130)

    GetColorSetNormal = tColorSet
    
End Function

Public Function GetColorSetDisabled(m_BackColor As OLE_COLOR, m_ForeColor As OLE_COLOR, Optional m_SelectionColor As OLE_COLOR = &HFFFFFF) As ColorSet

    Dim tColorSet As ColorSet
    
    tColorSet.csBackColor = m_BackColor
    tColorSet.csFrontColor = m_ForeColor
    
    tColorSet.csColor1(1) = BlendColor(m_BackColor, m_ForeColor, 200)
    tColorSet.csColor1(2) = BlendColor(m_BackColor, m_ForeColor, 190)
    tColorSet.csColor1(3) = BlendColor(m_BackColor, m_ForeColor, 180)
    tColorSet.csColor1(4) = BlendColor(m_BackColor, m_ForeColor, 220)
    tColorSet.csColor1(5) = BlendColor(m_BackColor, m_ForeColor, 200)
    tColorSet.csColor1(6) = BlendColor(m_BackColor, m_ForeColor, 120)
    tColorSet.csColor1(7) = BlendColor(m_BackColor, m_ForeColor, 140)
    
    tColorSet.csColor1(8) = BlendColor(m_BackColor, m_SelectionColor, 230)
    tColorSet.csColor1(9) = BlendColor(m_BackColor, m_SelectionColor, 180)
    tColorSet.csColor1(10) = BlendColor(m_BackColor, m_SelectionColor, 130)

    GetColorSetDisabled = tColorSet
    
End Function

Public Function GetColorSetHovered(m_BackColor As OLE_COLOR, m_ForeColor As OLE_COLOR, Optional m_SelectionColor As OLE_COLOR = &HFFFFFF) As ColorSet

    Dim tColorSet As ColorSet
    
    tColorSet.csBackColor = m_BackColor
    tColorSet.csFrontColor = m_ForeColor
    
    tColorSet.csColor1(1) = BlendColor(m_BackColor, m_ForeColor, 245)
    tColorSet.csColor1(2) = BlendColor(m_BackColor, m_ForeColor, 205)
    tColorSet.csColor1(3) = BlendColor(m_BackColor, m_ForeColor, 185)
    tColorSet.csColor1(4) = BlendColor(m_BackColor, m_ForeColor, 220)
    tColorSet.csColor1(5) = BlendColor(m_BackColor, m_ForeColor, 200)
    tColorSet.csColor1(6) = BlendColor(m_BackColor, m_ForeColor, 120)
    tColorSet.csColor1(7) = BlendColor(vbBlack, m_ForeColor, 140)
    
    tColorSet.csColor1(8) = BlendColor(m_BackColor, m_SelectionColor, 230)
    tColorSet.csColor1(9) = BlendColor(m_BackColor, m_SelectionColor, 180)
    tColorSet.csColor1(10) = BlendColor(m_BackColor, m_SelectionColor, 130)

    GetColorSetHovered = tColorSet
    
End Function

Public Function GetColorSetHovered2(m_BackColor As OLE_COLOR, m_ForeColor As OLE_COLOR, m_SelectionColor As OLE_COLOR) As ColorSet

    Dim tColorSet As ColorSet
    
    tColorSet.csBackColor = m_BackColor
    tColorSet.csFrontColor = m_ForeColor
    
    tColorSet.csColor1(1) = BlendColor(m_BackColor, m_SelectionColor, 230)
    tColorSet.csColor1(2) = BlendColor(m_BackColor, m_SelectionColor, 190)
    tColorSet.csColor1(3) = BlendColor(m_BackColor, m_SelectionColor, 170)
    tColorSet.csColor1(4) = BlendColor(m_BackColor, m_ForeColor, 220)
    tColorSet.csColor1(5) = BlendColor(m_BackColor, m_ForeColor, 200)
    tColorSet.csColor1(6) = BlendColor(m_BackColor, m_ForeColor, 120)
    tColorSet.csColor1(7) = BlendColor(vbBlack, m_ForeColor, 140)
    
    tColorSet.csColor1(8) = BlendColor(m_BackColor, m_SelectionColor, 230)
    tColorSet.csColor1(9) = BlendColor(m_BackColor, m_SelectionColor, 180)
    tColorSet.csColor1(10) = BlendColor(m_BackColor, m_SelectionColor, 130)

    GetColorSetHovered2 = tColorSet
    
End Function

Public Function GetColorSetTabbed(m_BackColor As OLE_COLOR, m_ForeColor As OLE_COLOR, m_TabbedColor As OLE_COLOR) As ColorSet

    Dim tColorSet As ColorSet
    Dim lRealColor As Long
    
    tColorSet.csBackColor = m_BackColor
    tColorSet.csFrontColor = m_ForeColor
    
    lRealColor = BlendColor(m_ForeColor, m_TabbedColor)
    
    tColorSet.csColor1(1) = BlendColor(m_BackColor, lRealColor, 230)
    tColorSet.csColor1(2) = BlendColor(m_BackColor, lRealColor, 190)
    tColorSet.csColor1(3) = BlendColor(m_BackColor, lRealColor, 170)
    tColorSet.csColor1(4) = BlendColor(m_BackColor, lRealColor, 220)
    tColorSet.csColor1(5) = BlendColor(m_BackColor, lRealColor, 200)
    tColorSet.csColor1(6) = BlendColor(m_BackColor, lRealColor, 120)
    tColorSet.csColor1(7) = BlendColor(vbBlack, lRealColor, 140)
    
    tColorSet.csColor1(8) = BlendColor(m_BackColor, m_TabbedColor, 230)
    tColorSet.csColor1(9) = BlendColor(m_BackColor, m_TabbedColor, 180)
    tColorSet.csColor1(10) = BlendColor(m_BackColor, m_TabbedColor, 130)

    GetColorSetTabbed = tColorSet
    
End Function

Public Function GetColorSetClicked(m_BackColor As OLE_COLOR, m_ForeColor As OLE_COLOR, Optional m_SelectionColor As OLE_COLOR = &HFFFFFF) As ColorSet

    Dim tColorSet As ColorSet
    
    tColorSet.csBackColor = m_BackColor
    tColorSet.csFrontColor = m_ForeColor
    
    tColorSet.csColor1(1) = BlendColor(m_BackColor, m_ForeColor, 170)
    tColorSet.csColor1(2) = BlendColor(m_BackColor, m_ForeColor, 190)
    tColorSet.csColor1(3) = BlendColor(m_BackColor, m_ForeColor, 230)
    
    tColorSet.csColor1(4) = BlendColor(m_BackColor, m_ForeColor, 120)
    tColorSet.csColor1(5) = BlendColor(m_BackColor, m_ForeColor, 120)
    tColorSet.csColor1(6) = BlendColor(m_BackColor, m_ForeColor, 220)
    
    tColorSet.csColor1(7) = BlendColor(vbBlack, m_ForeColor, 140)
    
    tColorSet.csColor1(8) = BlendColor(m_BackColor, m_SelectionColor, 230)
    tColorSet.csColor1(9) = BlendColor(m_BackColor, m_SelectionColor, 180)
    tColorSet.csColor1(10) = BlendColor(m_BackColor, m_SelectionColor, 130)

    GetColorSetClicked = tColorSet
    
End Function

Public Sub DropShadow(hwnd As Long)
    SetClassLong hwnd, GCL_STYLE, GetClassLong(hwnd, GCL_STYLE) Or CS_DROPSHADOW
End Sub

Public Function BlendColor(ByVal oColorFrom As OLE_COLOR, ByVal oColorTo As OLE_COLOR, Optional ByVal Alpha As Long = 128) As Long
    
    Dim lCFrom As Long
    Dim lCTo As Long
    Dim lSrcR As Long
    Dim lSrcG As Long
    Dim lSrcB As Long
    Dim lDstR As Long
    Dim lDstG As Long
    Dim lDstB As Long
   
    lCFrom = GetLngColor(oColorFrom)
    lCTo = GetLngColor(oColorTo)
    
    lSrcR = lCFrom And &HFF
    lSrcG = (lCFrom And &HFF00&) \ &H100&
    lSrcB = (lCFrom And &HFF0000) \ &H10000
    lDstR = lCTo And &HFF
    lDstG = (lCTo And &HFF00&) \ &H100&
    lDstB = (lCTo And &HFF0000) \ &H10000
   
    BlendColor = RGB( _
      ((lSrcR * Alpha) / 255) + ((lDstR * (255 - Alpha)) / 255), _
      ((lSrcG * Alpha) / 255) + ((lDstG * (255 - Alpha)) / 255), _
      ((lSrcB * Alpha) / 255) + ((lDstB * (255 - Alpha)) / 255))
      
End Function

Public Sub DrawGradient(cHdc As Long, X As Long, Y As Long, X2 As Long, Y2 As Long, Color1 As RGB, Color2 As RGB, Optional Direction As Integer = 1)

    Dim Vert(1) As TRIVERTEX   '2 Colors
    Dim gRect As GRADIENT_RECT

   
    With Vert(0)
        .X = X
        .Y = Y
        .Red = Color1.R
        .Green = Color1.G
        .Blue = Color1.B
        .Alpha = 0&
    End With

    With Vert(1)
        .X = Vert(0).X + X2
        .Y = Vert(0).Y + Y2
        .Red = Color2.R
        .Green = Color2.G
        .Blue = Color2.B
        .Alpha = 0&
    End With

    gRect.UpperLeft = 1
    gRect.LowerRight = 0

    GradientFill cHdc, Vert(0), 2, gRect, 1, Direction

End Sub

Public Function GetLngColor(Color As Long) As Long
    
    If (Color And &H80000000) Then
        GetLngColor = GetSysColor(Color And &H7FFFFFFF)
    Else
        GetLngColor = Color
    End If
End Function

Public Function SetColor(lColor As Long) As RGB
   
    Dim tmpColorHolder As RGB
    Dim R As Integer, B As Integer, G As Integer
    
    On Error GoTo ErrorTrap
    RGBRev R, G, B, lColor
    
    tmpColorHolder.R = R
    tmpColorHolder.G = G
    tmpColorHolder.B = B
    
    SetColor = tmpColorHolder
    
    Exit Function
ErrorTrap:

    tmpColorHolder.R = 255
    tmpColorHolder.G = 255
    tmpColorHolder.B = 255

End Function

Public Function GetRGBColors(lColor As Long) As RGB

    Dim HexColor As String

    HexColor = String(6 - Len(Hex(lColor)), "0") & Hex(lColor)
    GetRGBColors.R = "&H" & Mid(HexColor, 5, 2) & "00"
    GetRGBColors.G = "&H" & Mid(HexColor, 3, 2) & "00"
    GetRGBColors.B = "&H" & Mid(HexColor, 1, 2) & "00"

End Function

Public Function PercentOf(iTotal As Long, iPercent As Integer) As Long
    On Error Resume Next
    PercentOf = Fix(((iTotal / 10) * (iPercent / 10) * 100) / 100)
End Function

Public Sub GradientCy(cHdc As Long, X As Long, Y As Long, X2 As Long, Y2 As Long, Color1 As Long, Color2 As Long, Color3 As Long, tOrientation As MyOrientation)

    If tOrientation = pbHorizontal Then
        DrawGradient cHdc, X, Y, X2, Fix(Y2 / 2) + 1, GetRGBColors(Color1), GetRGBColors(Color2)
        DrawGradient cHdc, X, Y2 - (Fix(Y2 / 2)), X2, Fix(Y2 / 2), GetRGBColors(Color2), GetRGBColors(Color3)
    ElseIf tOrientation = pbVertical Then
        DrawGradient cHdc, X, Y, Fix(X2 / 2) + 1, Y2, GetRGBColors(Color1), GetRGBColors(Color2), 0
        DrawGradient cHdc, X2 - (Fix(X2 / 2)), Y, Fix(X2 / 2), Y2, GetRGBColors(Color2), GetRGBColors(Color3), 0
    End If

End Sub

Public Sub GradientLine(hdc As Long, X As Long, Y As Long, Lenght As Long, iOrientation As MyOrientation, Color1 As Long, Color2 As Long)
    If iOrientation = pbVertical Then
        DrawGradient hdc, X, Y, 1, Lenght, GetRGBColors(Color1), GetRGBColors(Color2)
    Else
        DrawGradient hdc, X, Y, Lenght, 1, GetRGBColors(Color1), GetRGBColors(Color2), 0
    End If
End Sub

Public Sub Transparency(ByVal hwnd As Long, ByVal bAlpha As Byte)
    
    Dim nullRect As RECT
    Dim lret As Long
    
    lret = GetWindowLong(hwnd, GWL_EXSTYLE)
    lret = lret Or WS_EX_LAYERED
    SetWindowLong hwnd, GWL_EXSTYLE, lret
    SetLayeredWindowAttributes hwnd, 0, bAlpha, LWA_ALPHA
        
End Sub

Public Sub DisableAlpha(ByVal hwnd As Long)
    
    Dim nullRect As RECT
    Dim lret As Long
    
    lret = GetWindowLong(hwnd, GWL_EXSTYLE)
    lret = lret And Not WS_EX_LAYERED
    SetWindowLong hwnd, GWL_EXSTYLE, lret
    RedrawWindow hwnd, nullRect, 0&, RDW_ALLCHILDREN Or RDW_ERASE Or RDW_FRAME Or RDW_INVALIDATE

End Sub

Public Function ShellDocument(sDocName As String, Optional ByVal Action As String = "Open", Optional ByVal Parameters As String = vbNullString, Optional ByVal Directory As String = vbNullString, Optional ByVal WindowState As StartWindowState) As Boolean
    
    Dim Response
    
    Response = ShellExecute(&O0, Action, sDocName, Parameters, Directory, WindowState)
    
    Select Case Response
        Case Is < 33
            ShellDocument = False
        Case Else
            ShellDocument = True
    End Select
    
End Function

Public Function Get_WindowsDir() As String
    
    Dim WindowsDir As String
    
    WindowsDir = Space(256)
    WindowsDir = Trim(Left$(WindowsDir, GetWindowsDirectory(WindowsDir, 256&)))

    If Right(WindowsDir, 1) <> "\" Then
        Get_WindowsDir = WindowsDir & "\"
    Else
        Get_WindowsDir = WindowsDir
    End If

End Function

Public Function GetSystemDir() As String
    
    Dim SystemDir As String
    
    SystemDir = Space(256)
    SystemDir = Trim(Left$(SystemDir, GetSystemDirectory(SystemDir, 256&)))

    If Right(SystemDir, 1) <> "\" Then
        GetSystemDir = SystemDir & "\"
    Else
        GetSystemDir = SystemDir
    End If

End Function

Public Function HexEx(dnumber As Long) As String

    Dim TmpHex As String
    Dim tmpLong As Long
    Dim TmpOut As Long
    
    tmpLong = dnumber
    
    Do
        TmpOut = NtHC(tmpLong)
        If TmpOut > 15 Then
            TmpHex = Hex(tmpLong) & TmpHex
        Else
            TmpHex = Hex(TmpOut) & Hex(tmpLong) & TmpHex
            Exit Do
        End If
        tmpLong = TmpOut
    Loop

    HexEx = FormatLenght(TmpHex)

End Function

Private Function NtHC(l_long As Long) As Long
   
    Dim ValueHolder1 As Long
    Dim ValueHolder2 As Long
    Dim FullNumber As Long
    
    FullNumber = l_long
    
    ValueHolder1 = Int(FullNumber / 16)
    
    FullNumber = FullNumber - (ValueHolder1 * 16)
    
    ValueHolder2 = FullNumber
    l_long = ValueHolder2
    
    
    NtHC = ValueHolder1

End Function

Public Function FormatLenght(sText As String, Optional iLenght As Integer = 6) As String
    
    Dim cTemp As String
    cTemp = Trim(sText)
    
    If Len(cTemp) < iLenght Then
        Do Until Len(cTemp) = iLenght
            cTemp = "0" & cTemp
        Loop
    End If
    FormatLenght = cTemp
    
End Function

Public Function Get_FileNameFromPNF(strPathFilename As String) As String
    
    Dim tmpInt As Integer
    
    tmpInt = InStrRev(strPathFilename, "\")
    If tmpInt = 0 Then
        Get_FileNameFromPNF = Trim(strPathFilename):
    Else
        Get_FileNameFromPNF = Mid(strPathFilename, tmpInt + 1, Len(strPathFilename) - tmpInt + 1)
    End If

End Function

Public Function Get_PathFromPNF(strPathFilename As String) As String
    
    Dim tmpInt As Integer
    
    tmpInt = InStrRev(strPathFilename, "\")
    If tmpInt = 0 Then
        Get_PathFromPNF = Trim(strPathFilename)
    Else
        Get_PathFromPNF = Left(strPathFilename, tmpInt)
    End If
    
End Function
