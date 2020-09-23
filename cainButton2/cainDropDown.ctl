VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.UserControl cainDropDown 
   AutoRedraw      =   -1  'True
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   Begin cainObjects.cainXButton cainXButton1 
      Height          =   375
      Left            =   960
      Top             =   120
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Webdings"
         Size            =   8.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3240
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin VB.Timer tmrFocus 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   2760
      Top             =   0
   End
   Begin VB.Timer MousePos 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   2280
      Top             =   0
   End
   Begin VB.PictureBox picList 
      Appearance      =   0  '2D
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FCF8F0&
      BorderStyle     =   0  'Kein
      ForeColor       =   &H80000008&
      Height          =   1695
      Left            =   360
      ScaleHeight     =   113
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   177
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   1560
      Visible         =   0   'False
      Width           =   2655
      Begin cainObjects.cainListbox cainListbox1 
         Height          =   1215
         Left            =   360
         TabIndex        =   4
         Top             =   360
         Visible         =   0   'False
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   2143
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         HoverSelect     =   -1  'True
      End
      Begin cainObjects.cainMonthview mw 
         Height          =   2310
         Left            =   1320
         TabIndex        =   3
         Top             =   120
         Visible         =   0   'False
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   4075
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
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  '2D
      BorderStyle     =   0  'Kein
      Height          =   255
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   960
      Width           =   2175
   End
   Begin cainObjects.cainPUMenu cainPUMenu1 
      Height          =   705
      Left            =   120
      TabIndex        =   0
      Top             =   120
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
End
Attribute VB_Name = "cainDropDown"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'Standard-Eigenschaftswerte:
Const m_def_FontName = "Arial"
Const m_def_Day = 1
Const m_def_Year = 2005
Const m_def_Month = 1
Const m_def_PrinterName = ""
Const m_def_DriveLetter = "C:"
Const m_def_DateFormat = "dd.mm.yyyy"
Const m_def_ComboType = 1
Const m_def_TabSelectColor = &H96E7&
Const m_def_BackColor = &HFFFFFF
Const m_def_ForeColor = &HC00000
Const m_def_Enabled = 0
Const m_def_SelectionColor = &H80FF&
'Eigenschaftsvariablen:
Dim m_FontName As String
Dim m_Day As Integer
Dim m_Year As Integer
Dim m_Month As Integer
Dim m_PrinterName As String
Dim m_DriveLetter As String
Dim m_DateFormat As String
Dim m_ComboType As eComboType
Dim m_TabSelectColor As OLE_COLOR
Dim m_BackColor As OLE_COLOR
Dim m_ForeColor As OLE_COLOR
Dim m_Enabled As Boolean
Dim m_Font As Font
Dim m_SelectionColor As OLE_COLOR
'Ereignisdeklarationen:
Event ValueChanged()
Event KeyUp(KeyCode As Integer, Shift As Integer)
Attribute KeyUp.VB_Description = "Tritt auf, wenn der Benutzer eine Taste losläßt, während ein Objekt den Fokus hat."
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute MouseDown.VB_Description = "Tritt auf, wenn der Benutzer die Maustaste drückt, während ein Objekt den Fokus hat."
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute MouseMove.VB_Description = "Tritt auf, wenn der Benutzer die Maus bewegt."
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute MouseUp.VB_Description = "Tritt auf, wenn der Benutzer die Maustaste losläßt, während ein Objekt den Fokus hat."
Event KeyDown(KeyCode As Integer, Shift As Integer)
Attribute KeyDown.VB_Description = "Tritt auf, wenn der Benutzer eine Taste drückt, während ein Objekt den Fokus besitzt."
Event KeyPress(KeyAscii As Integer)
Attribute KeyPress.VB_Description = "Tritt auf, wenn der Benutzer eine ANSI-Taste drückt und losläßt."

Private Const ICON_SIZE As Integer = 16

Dim tColorSet As ColorSet

Dim MyState As ButtonState
Dim MyTabState As TabSelected

Dim OldComboType As eComboType

Dim Mouse As POINTAPI
Dim MouseOverMe As POINTAPI
Dim MouseClick As Boolean
Dim MenuOpened As Boolean
Dim iIconIndex As Integer

Dim dClicked As Boolean

Public Function ID() As Integer
    ID = eControlIDs.id_DropDown
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
    
    DrawDropDown
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
    
    DrawDropDown
End Property

'ACHTUNG! DIE FOLGENDEN KOMMENTIERTEN ZEILEN NICHT ENTFERNEN ODER VERÄNDERN!
'MemberInfo=0,0,0,0
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Gibt einen Wert zurück, der bestimmt, ob ein Objekt auf vom Benutzer erzeugte Ereignisse reagieren kann, oder legt diesen fest."
Attribute Enabled.VB_ProcData.VB_Invoke_Property = ";Verhalten"
    Enabled = m_Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    m_Enabled = New_Enabled
    PropertyChanged "Enabled"
    
    DrawDropDown
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
    
    SetFont
    DrawDropDown
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
    
    DrawDropDown
End Property

'ACHTUNG! DIE FOLGENDEN KOMMENTIERTEN ZEILEN NICHT ENTFERNEN ODER VERÄNDERN!
'MemberInfo=10,0,0,0
Public Property Get TabSelectColor() As OLE_COLOR
Attribute TabSelectColor.VB_ProcData.VB_Invoke_Property = ";Darstellung"
    TabSelectColor = m_TabSelectColor
End Property

Public Property Let TabSelectColor(ByVal New_TabSelectColor As OLE_COLOR)
    m_TabSelectColor = New_TabSelectColor
    PropertyChanged "TabSelectColor"
    DrawDropDown
End Property

'ACHTUNG! DIE FOLGENDEN KOMMENTIERTEN ZEILEN NICHT ENTFERNEN ODER VERÄNDERN!
'MemberInfo=7,0,0,0
Public Property Get DateFormat() As String
    DateFormat = m_DateFormat
End Property

Public Property Let DateFormat(ByVal New_DateFormat As String)
    m_DateFormat = New_DateFormat
    PropertyChanged "DateFormat"
    
    SetComboType
End Property

'ACHTUNG! DIE FOLGENDEN KOMMENTIERTEN ZEILEN NICHT ENTFERNEN ODER VERÄNDERN!
'MemberInfo=7,0,0,0
Public Property Get ComboType() As eComboType
Attribute ComboType.VB_ProcData.VB_Invoke_Property = ";Verhalten"
    ComboType = m_ComboType
End Property

Public Property Let ComboType(ByVal New_ComboType As eComboType)
    
    m_ComboType = New_ComboType
    PropertyChanged "ComboType"
    
    AddIcons
    SetComboType
    SetDisplay
    DrawDropDown
    
End Property

'ACHTUNG! DIE FOLGENDEN KOMMENTIERTEN ZEILEN NICHT ENTFERNEN ODER VERÄNDERN!
'MemberInfo=7,0,0,1
Public Property Get Day() As Integer
    Day = m_Day
End Property

Public Property Let Day(ByVal New_Day As Integer)
    m_Day = New_Day
    PropertyChanged "Day"
    SetDisplay
End Property

'ACHTUNG! DIE FOLGENDEN KOMMENTIERTEN ZEILEN NICHT ENTFERNEN ODER VERÄNDERN!
'MemberInfo=7,0,0,2005
Public Property Get Year() As Integer
    Year = m_Year
End Property

Public Property Let Year(ByVal New_Year As Integer)
    m_Year = New_Year
    PropertyChanged "Year"
    SetDisplay
End Property

'ACHTUNG! DIE FOLGENDEN KOMMENTIERTEN ZEILEN NICHT ENTFERNEN ODER VERÄNDERN!
'MemberInfo=7,0,0,1
Public Property Get Month() As Integer
    Month = m_Month
End Property

Public Property Let Month(ByVal New_Month As Integer)
    m_Month = New_Month
    PropertyChanged "Month"
    SetDisplay
End Property

'ACHTUNG! DIE FOLGENDEN KOMMENTIERTEN ZEILEN NICHT ENTFERNEN ODER VERÄNDERN!
'MemberInfo=13,0,0,0
Public Property Get PrinterName() As String
    PrinterName = m_PrinterName
End Property

Public Property Let PrinterName(ByVal New_PrinterName As String)
    m_PrinterName = New_PrinterName
    PropertyChanged "PrinterName"
    
    SetDisplay
End Property
'ACHTUNG! DIE FOLGENDEN KOMMENTIERTEN ZEILEN NICHT ENTFERNEN ODER VERÄNDERN!
'MemberInfo=13,0,0,0
Public Property Get FontName() As String
    FontName = m_FontName
End Property

Public Property Let FontName(ByVal New_FontName As String)
    m_FontName = New_FontName
    PropertyChanged "FontName"
    SetDisplay
End Property

'ACHTUNG! DIE FOLGENDEN KOMMENTIERTEN ZEILEN NICHT ENTFERNEN ODER VERÄNDERN!
'MemberInfo=13,0,0,0
Public Property Get DriveLetter() As String
    DriveLetter = m_DriveLetter
End Property

Public Property Let DriveLetter(ByVal New_DriveLetter As String)
    m_DriveLetter = New_DriveLetter
    PropertyChanged "DriveLetter"
    
    SetDisplay
End Property

'===================================================================================================================
'===================================================================================================================
'===================================================================================================================
'===================================================================================================================
'===================================================================================================================
'===================================================================================================================
'===================================================================================================================
'===================================================================================================================

Private Sub RefreshColor()
    
    cainListbox1.BackColor = m_BackColor
    cainListbox1.ForeColor = m_ForeColor
    cainListbox1.SelectionColor = m_SelectionColor
    
    mw.BackColor = m_BackColor
    mw.ForeColor = m_ForeColor
    mw.SelectionColor = m_SelectionColor
    
End Sub

Private Sub SetFont()
    
    Text1.FontBold = m_Font.Bold
    Text1.FontItalic = m_Font.Italic
    Text1.FontName = m_Font.Name
    Text1.FontSize = m_Font.Size
    Text1.FontStrikethru = m_Font.Strikethrough
    Text1.FontUnderline = m_Font.Underline
    
    UserControl_Resize

End Sub

Private Sub DrawDropDown()

    UserControl.Cls

    If UserControl.Enabled = True Then
        If MyState = bDisabled Then MyState = bNormal
    Else
        MyState = bDisabled
    End If
           
    If MyState = bNormal And (MyTabState = tbNormal Or MyTabState = tbTabed) Then
        tColorSet = GetColorSetNormal(m_BackColor, m_ForeColor)
        
    ElseIf MyState = bHovered And (MyTabState = tbNormal Or MyTabState = tbTabed) Then
        tColorSet = GetColorSetHovered(m_BackColor, m_ForeColor)
        
    ElseIf MyState = bPressed And (MyTabState = tbNormal Or MyTabState = tbTabed) Then
        tColorSet = GetColorSetHovered(m_BackColor, m_ForeColor)
        
    ElseIf MyState = bDisabled And (MyTabState = tbNormal Or MyTabState = tbTabed) Then
        tColorSet = GetColorSetDisabled(m_BackColor, m_ForeColor)
        
    ElseIf MyState = bUnselected And MyTabState = tbNormal Then
        tColorSet = GetColorSetNormal(m_BackColor, m_ForeColor)
        
    ElseIf MyState = bUnselected And MyTabState = tbTabed Then
        tColorSet = GetColorSetTabbed(m_BackColor, m_TabSelectColor, m_ForeColor)
        
    End If
   
    UserControl.BackColor = tColorSet.csColor1(1)
    Text1.BackColor = tColorSet.csColor1(1)
    Text1.ForeColor = tColorSet.csColor1(7)
    cainXButton1.BackColor = tColorSet.csBackColor
    cainXButton1.ForeColor = tColorSet.csFrontColor
    picList.BackColor = tColorSet.csBackColor
    
    UserControl.Cls
    DrawIcon
    
    'background
    'GradientCy UserControl.hDC, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight, tColorSet.csColor1(1), tColorSet.csColor1(2), tColorSet.csColor1(3), pbHorizontal

    'Borders
    GradientLine UserControl.hdc, 1, 0, UserControl.ScaleWidth - 2, pbHorizontal, tColorSet.csColor1(4), tColorSet.csColor1(5)
    GradientLine UserControl.hdc, 0, 1, UserControl.ScaleHeight - 2, pbVertical, tColorSet.csColor1(4), tColorSet.csColor1(6)
    UserControl.Line (1, UserControl.ScaleHeight - 1)-(UserControl.ScaleWidth - 1, UserControl.ScaleHeight - 1), tColorSet.csColor1(6)
    GradientLine UserControl.hdc, UserControl.ScaleWidth - 1, 1, UserControl.ScaleHeight - 2, pbVertical, tColorSet.csColor1(5), tColorSet.csColor1(6)

    'UserControl.Line (cainXButton1.Left - 2, 1)-(cainXButton1.Left - 2, UserControl.ScaleHeight - 1), tColorSet.csColor1(6)
    GradientLine UserControl.hdc, cainXButton1.Left - 2, 1, UserControl.ScaleHeight - 1, pbVertical, tColorSet.csColor1(4), tColorSet.csColor1(6)
    
    
    'piclist borders
'    If m_ComboType = ctColorPicker Or m_ComboType = ctDateBox Then
'        picList.Line (0, 0)-(0, picList.ScaleHeight), tColorSet.csColor1(6)
'        picList.Line (0, 0)-(picList.ScaleWidth, 0), tColorSet.csColor1(6)
'        picList.Line (picList.ScaleWidth - 1, 0)-(picList.ScaleWidth - 1, picList.ScaleHeight), tColorSet.csColor1(6)
'        picList.Line (0, picList.ScaleHeight - 1)-(picList.ScaleWidth - 1, picList.ScaleHeight - 1), tColorSet.csColor1(6)
'    End If

End Sub

Private Sub DrawIcon()

    If iIconIndex = 0 Then Exit Sub
    If ImageList1.ListImages.Count = 0 Then Exit Sub

    If UserControl.Enabled = False Then
        Call DrawState(UserControl.hdc, 0, 0, _
        ImageList1.ListImages(iIconIndex).ExtractIcon, 0, _
        4, (UserControl.ScaleHeight / 2) - 8, _
        16, 16, DST_ICON Or DSS_MONO)
    
    Else
        Call DrawState(UserControl.hdc, 0, 0, _
        ImageList1.ListImages(iIconIndex).ExtractIcon, 0, _
        4, (UserControl.ScaleHeight / 2) - 8, _
        16, 16, DST_ICON Or DSS_NORMAL)
        
    End If

End Sub

Private Sub AddIcons()
    
    Dim cExtractor As IconExtractor
    Set cExtractor = New IconExtractor

    Select Case m_ComboType
    
        Case eComboType.ctColorPicker
            
        Case eComboType.ctDriveSelect
            cainListbox1.ListboxItems.Clear
            ImageList1.ListImages.Clear
            cExtractor.MaskColor = ImageList1.MaskColor
            ImageList1.ListImages.Add , , cExtractor.ExtractIcon(GetSystemDir & "shell32.dll", 8, SmallIcon) 'fixed
            ImageList1.ListImages.Add , , cExtractor.ExtractIcon(GetSystemDir & "shell32.dll", 6, SmallIcon) 'removable
            ImageList1.ListImages.Add , , cExtractor.ExtractIcon(GetSystemDir & "shell32.dll", 11, SmallIcon) 'cd rom
            ImageList1.ListImages.Add , , cExtractor.ExtractIcon(GetSystemDir & "shell32.dll", 9, SmallIcon) 'network
            ImageList1.ListImages.Add , , cExtractor.ExtractIcon(GetSystemDir & "shell32.dll", 34, SmallIcon) 'desktop
            ImageList1.ListImages.Add , , cExtractor.ExtractIcon(GetSystemDir & "shell32.dll", 126, SmallIcon) 'Personal
            ImageList1.ListImages.Add , , cExtractor.ExtractIcon(GetSystemDir & "shell32.dll", 15, SmallIcon) 'Personal
            ImageList1.ListImages.Add , , cExtractor.ExtractIcon(GetSystemDir & "shell32.dll", 18, SmallIcon) 'Personal
            ImageList1.ListImages.Add , , cExtractor.ExtractIcon(GetSystemDir & "shell32.dll", 3, SmallIcon) 'Personal
            Set cainListbox1.ImageList = ImageList1
        
        Case eComboType.ctPrinterSelect
            cainListbox1.ListboxItems.Clear
            ImageList1.ListImages.Clear
            cExtractor.MaskColor = ImageList1.MaskColor
            ImageList1.ListImages.Add , , cExtractor.ExtractIcon(GetSystemDir & "shell32.dll", 16, SmallIcon)
            ImageList1.ListImages.Add , , cExtractor.ExtractIcon(GetSystemDir & "shell32.dll", 60, SmallIcon)
            ImageList1.ListImages.Add , , cExtractor.ExtractIcon(GetSystemDir & "shell32.dll", 81, SmallIcon)
            ImageList1.ListImages.Add , , cExtractor.ExtractIcon(GetSystemDir & "shell32.dll", 82, SmallIcon)
            Set cainListbox1.ImageList = ImageList1
        
        Case eComboType.ctDateBox
            cainListbox1.ListboxItems.Clear
            ImageList1.ListImages.Clear
            cExtractor.MaskColor = ImageList1.MaskColor
            ImageList1.ListImages.Add , , cExtractor.ExtractIcon(GetSystemDir & "shell32.dll", 167, SmallIcon)
            Set cainListbox1.ImageList = ImageList1
            
        Case eComboType.ctFontSelect
            cainListbox1.ListboxItems.Clear
            ImageList1.ListImages.Clear
            'Set cainListbox1.ImageList = ImageList1
    
    End Select
    
End Sub

Private Sub AddDrive()

    Dim FSO As FileSystemObject: Set FSO = New FileSystemObject
    Dim Drive As Drive
    Dim Drives As Drives
    Dim intIcon As Integer
    Dim sName As String
    Dim cDirs As sf_Information
    Dim cSearch As Search: Set cSearch = New Search
    Dim i As Integer
    
    Set Drives = FSO.Drives
    cainListbox1.ListboxItems.Clear
    cainListbox1.ListboxItems.Sorted = False
    cainListbox1.ValueFont = False
    
    'Add Desktop path
    sName = GetSettingString(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders", "Desktop")
    cainListbox1.ListboxItems.Add "X_Desktop", "Desktop", sName, 5
    cSearch.Search_Directory sName, "*", cDirs
    
    'Add Arbeitsplatz path
    sName = GetSettingString(HKEY_CURRENT_USER, "Software\Microsoft\Windows\ShellNoRoam\MUICache", "@C:\WINDOWS\system32\SHELL32.dll,-9216")
    cainListbox1.ListboxItems.Add "X_MyComputer", vbTab & sName, "X_MyComputer", 7
    
    For Each Drive In Drives
        
        
        If Drive.IsReady = True Then
            'intIcon
        Else
            'intIcon
        End If
        
        Select Case Drive.DriveType
            Case Remote
                If Drive.IsReady = True Then
                    sName = "(" & Drive.DriveLetter & ":) " & Get_FileNameFromPNF(Drive.ShareName) & " [" & Replace(Get_PathFromPNF(Drive.ShareName), "\\", "") & "]"
                Else
                    sName = "(" & Drive.DriveLetter & ":)"
                End If
                intIcon = 4
            
            Case Removable, Fixed, CDRom
                If Drive.IsReady = True Then
                    sName = "(" & Drive.DriveLetter & ":) " & Drive.VolumeName
                Else
                    sName = "(" & Drive.DriveLetter & ":)"
                End If
                
                Select Case Drive.DriveType
                    
                    Case Removable
                        intIcon = 2
                        
                    Case Fixed
                        intIcon = 1
                    
                    Case CDRom
                        intIcon = 3
                    
                End Select
                
                
            
        End Select
        
        cainListbox1.ListboxItems.Add sName, vbTab & vbTab & sName, Drive.DriveLetter & ":", intIcon ', prThis.DeviceName, intIcons
        
     Next Drive
    
    'Eigene Dateien
    sName = GetSettingString(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders", "Personal")
    cainListbox1.ListboxItems.Add "X_Personal", Get_FileNameFromPNF(sName), sName, 6
    
    'Network neighborhood
    sName = GetSettingString(HKEY_CURRENT_USER, "Software\Microsoft\Windows\ShellNoRoam\MUICache", "@C:\WINDOWS\system32\SHELL32.dll,-9217")
    cainListbox1.ListboxItems.Add "X_Network", sName, "X_Network", 8
    
    For i = 1 To cDirs.Count
        cainListbox1.ListboxItems.Add cDirs(i).Path & cDirs(i).Filename, cDirs(i).Filename, cDirs(i).Path & cDirs(i).Filename, 9
    Next i
    
    If Not (FSO Is Nothing) Then Set FSO = Nothing
    If Not (cSearch Is Nothing) Then Set cSearch = Nothing

End Sub

Private Sub AddFonts()
    
    Dim i As Integer
    Dim tS As String
    
    
    cainListbox1.ListboxItems.Clear
    cainListbox1.ListboxItems.Sorted = True
    cainListbox1.ValueFont = True

    For i = 1 To Screen.FontCount
        tS = Screen.Fonts(i)
        If Trim(tS) <> "" Then _
            cainListbox1.ListboxItems.Add tS, tS, tS ', 2  ', intIcons
    Next i

End Sub

Private Sub AddPrinter()
    Dim prThis As Printer
    Dim intIcons As Integer
    
    cainListbox1.ListboxItems.Clear
    cainListbox1.ListboxItems.Sorted = False
    cainListbox1.ValueFont = False

    If Printers.Count > 0 Then
        For Each prThis In Printers
            
            If Left(prThis.DeviceName, 2) = "\\" Then
                intIcons = 2
            Else
                intIcons = 1
            End If
            
            If Printer.DeviceName = prThis.DeviceName Then
                If Left(prThis.DeviceName, 2) = "\\" Then
                    intIcons = 4
                Else
                    intIcons = 3
                End If
            End If
            cainListbox1.ListboxItems.Add prThis.DeviceName, Replace(prThis.DeviceName, "\\", ""), prThis.DeviceName, intIcons

        Next prThis
    End If

End Sub

Public Sub SetComboType()

    If m_ComboType <> OldComboType Then
    
        cainListbox1.Visible = False
        mw.Visible = False
        iIconIndex = 1
    
        Select Case m_ComboType
        
            Case eComboType.ctColorPicker
            
            Case eComboType.ctDateBox
                picList.Height = (mw.Height) * Screen.TwipsPerPixelY
                picList.Width = (mw.Width) * Screen.TwipsPerPixelX
                mw.Top = 0
                mw.Left = 0
                mw.Visible = True
                Text1.Alignment = 2
                
                'MW.Value = Text1.Text
                    
            Case eComboType.ctDriveSelect, eComboType.ctPrinterSelect, eComboType.ctFontSelect
                picList.Height = 113 * Screen.TwipsPerPixelY
                cainListbox1.Height = picList.ScaleHeight
                cainListbox1.Width = picList.ScaleWidth
                cainListbox1.Top = 0
                cainListbox1.Left = 0
                cainListbox1.Visible = True
                Text1.Alignment = 0
                
                Select Case m_ComboType
                    Case eComboType.ctDriveSelect:   AddDrive
                    Case eComboType.ctPrinterSelect: AddPrinter
                    Case eComboType.ctFontSelect: AddFonts
                End Select
                
        
        End Select
        
        OldComboType = m_ComboType
    
    End If

End Sub

Private Function InFocusControl(ByVal ObjecthWnd As Long) As Boolean

    Dim mPos As POINTAPI
    Dim KeyLeft As Boolean
    Dim KeyRight As Boolean
    Dim oRect As RECT
    Dim oRect2 As RECT
    
    Call GetCursorPos(mPos)
    Call GetWindowRect(ObjecthWnd, oRect)
    Call GetWindowRect(picList.hwnd, oRect2)
    KeyLeft = GetAsyncKeyState(VK_LBUTTON)
    KeyRight = GetAsyncKeyState(VK_RBUTTON)
    
    If ((mPos.X >= oRect.Left) And (mPos.X <= oRect.Right) And (mPos.Y >= oRect.Top) And (mPos.Y <= oRect.Bottom)) Or mw.bMenuOpen = True Then
        InFocusControl = True
    ElseIf ((KeyLeft = True) Or (KeyRight = True)) And dClicked = False Then
        If (mPos.X < oRect.Left) Or (mPos.X > oRect.Right) Or (mPos.Y < oRect.Top) Or (mPos.Y > oRect.Bottom) Then
        
            If (mPos.X < oRect2.Left) Or (mPos.X > oRect2.Right) Or (mPos.Y < oRect2.Top) Or (mPos.Y > oRect2.Bottom) Then
                picList.Visible = False
                InFocusControl = False
                
            End If
        
        End If
    
    ElseIf dClicked = True Then
        InFocusControl = True
    
    End If
    
End Function

Public Sub CreateMenu(X As Long, Y As Long)

    On Error Resume Next
    
    Dim mPos As POINTAPI
    
    Dim i As Integer
    Dim lHeight As Integer
    Dim iHeight As Long
    Dim iTop As Long
    
    Call picList.Move(X, Y)
    Call SetWindowPos(picList.hwnd, HWND_TOPMOST, 0, 0, 0, 0, TOPMOST_FLAGS)
    
    RefreshColor
'    If m_ComboType = ctColorPicker Or m_ComboType = ctDateBox Then
'        picList.Line (0, 0)-(0, picList.ScaleHeight), tColorSet.csColor1(6)
'        picList.Line (0, 0)-(picList.ScaleWidth, 0), tColorSet.csColor1(6)
'        picList.Line (picList.ScaleWidth - 1, 0)-(picList.ScaleWidth - 1, picList.ScaleHeight), tColorSet.csColor1(6)
'        picList.Line (0, picList.ScaleHeight - 1)-(picList.ScaleWidth - 1, picList.ScaleHeight - 1), tColorSet.csColor1(6)
'    End If
    
    Select Case m_ComboType
        Case eComboType.ctDriveSelect
            cainListbox1.SelectItem , m_DriveLetter
            
        Case eComboType.ctPrinterSelect
            cainListbox1.SelectItem , m_PrinterName
            
        Case eComboType.ctFontSelect
            cainListbox1.SelectItem , m_FontName
        
        Case eComboType.ctDateBox
            mw.iDay = m_Day
            mw.iMonth = m_Month
            mw.iYear = m_Year
            
    End Select
                
    picList.Visible = True
    MenuOpened = True
    Transparency picList.hwnd, 230
        
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

Private Sub Format_DateText(sDay As Integer, sMonth As Integer, sYear As Integer)

    Dim sDate As String
    sDate = m_DateFormat
    
    'double
    sDate = Replace(sDate, "dd", FormatLenght(Str(sDay), 2))
    sDate = Replace(sDate, "mm", FormatLenght(Str(sMonth), 2))
    sDate = Replace(sDate, "yyyy", sYear)
    
    'single
    sDate = Replace(sDate, "d", sDay)
    sDate = Replace(sDate, "m", sMonth)
    sDate = Replace(sDate, "yy", Right(sYear, 2))

    Text1.Text = sDate

End Sub

Private Sub SetDisplay()

    Dim i As Integer
    
    Select Case m_ComboType

        Case eComboType.ctColorPicker

        Case eComboType.ctDateBox
            Format_DateText m_Day, m_Month, m_Year
            
        Case eComboType.ctDriveSelect
            cainListbox1.SelectItem , m_DriveLetter

        Case eComboType.ctPrinterSelect
            cainListbox1.SelectItem , m_PrinterName

        Case eComboType.ctFontSelect
            cainListbox1.SelectItem , m_FontName

    End Select

    DrawDropDown

End Sub

'===================================================================================================================
'===================================================================================================================
'===================================================================================================================
'===================================================================================================================
'===================================================================================================================
'===================================================================================================================
'===================================================================================================================
'===================================================================================================================
' EVENTS!!
Private Sub MousePos_Timer()

    GetCursorPos Mouse

    If (((MouseOverMe.X > Mouse.X - 2) And (MouseOverMe.X < Mouse.X + 2)) And ((MouseOverMe.Y > Mouse.Y - 2) And (MouseOverMe.Y < Mouse.Y + 2))) Or MenuOpened = True Then
        If MouseClick = True Then
            MyState = bPressed
            DrawDropDown
        Else
            MyState = bHovered
            DrawDropDown
        End If
    Else
        MyState = bUnselected
        DrawDropDown
        MousePos.Enabled = False
    End If

    DoEvents
    
    'RaiseEvent MouseOver

End Sub

Private Sub cainListbox1_ItemClicked(iIndex As Integer, sKey As String)
    
    picList.Visible = False
    Text1.Text = Replace(cainListbox1.ListboxItems.Item(iIndex).Caption, vbTab, "")
    
    Select Case m_ComboType
    
    Case eComboType.ctDriveSelect
        m_DriveLetter = cainListbox1.ListboxItems.Item(iIndex).Value
        
    Case eComboType.ctPrinterSelect
        m_PrinterName = cainListbox1.ListboxItems.Item(iIndex).Value
        
    Case eComboType.ctFontSelect
        m_FontName = cainListbox1.ListboxItems.Item(iIndex).Value
    
    Case eComboType.ctColorPicker
    
    End Select
    
    iIconIndex = cainListbox1.ListboxItems.Item(iIndex).Icon
    RaiseEvent ValueChanged
    
End Sub

Private Sub cainListbox1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    dClicked = True
End Sub

Private Sub cainListbox1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    dClicked = False
End Sub

Private Sub cainXButton1_Click()
    On Error Resume Next
    
    Dim oRect As RECT
    Call GetWindowRect(UserControl.hwnd, oRect)
    
    CreateMenu (oRect.Left) * Screen.TwipsPerPixelX, (oRect.Top + UserControl.ScaleHeight) * Screen.TwipsPerPixelY
    UserControl.SetFocus
    
End Sub

Private Sub cainXButton1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    UserControl_MouseDown Button, Shift, X, Y
End Sub

Private Sub cainXButton1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    UserControl_MouseMove Button, Shift, X, Y
End Sub

Private Sub cainXButton1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    UserControl_MouseUp Button, Shift, X, Y
End Sub

Private Sub MW_DblClick()
    
    picList.Visible = False
            
    m_Day = mw.iDay
    m_Month = mw.iMonth
    m_Year = mw.iYear
    Format_DateText m_Day, m_Month, m_Year
        
    RaiseEvent ValueChanged

End Sub

Private Sub picList_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    dClicked = True
End Sub

Private Sub picList_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    dClicked = False
End Sub

Private Sub Text1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    UserControl_MouseDown Button, Shift, X, Y
End Sub

Private Sub Text1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    UserControl_MouseMove Button, Shift, X, Y
End Sub

Private Sub Text1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    UserControl_MouseUp Button, Shift, X, Y
End Sub

Private Sub tmrFocus_Timer()

    If (InFocusControl(UserControl.hwnd) = True) And (picList.Visible = False) Then
    ElseIf picList.Visible = False Then
        MenuOpened = False
        tmrFocus.Enabled = False
        'RaiseEvent Closed
        Exit Sub
    End If
    
End Sub

Private Sub UserControl_Initialize()
    
    Dim lResult As Long
    
    cainPUMenu1.Top = -cainPUMenu1.Height - 10
    cainPUMenu1.Left = -cainPUMenu1.Width - 10
    cainXButton1.Font.Name = "webdings"
    cainXButton1.Font.Size = 8
    cainXButton1.Caption = 6
    
    'Set ListboxItems = New ComboDatas

    dClicked = False
    MenuOpened = False
    iIconIndex = 0
    

    On Error Resume Next
    lResult = GetWindowLong(picList.hwnd, GWL_EXSTYLE)
    Call SetWindowLong(picList.hwnd, GWL_EXSTYLE, lResult Or WS_EX_TOOLWINDOW)
    Call SetWindowPos(picList.hwnd, picList.hwnd, 0, 0, 0, 0, 39)
    Call SetWindowLong(picList.hwnd, -8, Parent.hwnd)
    Call SetParent(picList.hwnd, 0)
    
    DropShadow picList.hwnd
    
    UserControl_Resize

End Sub

'Eigenschaften für Benutzersteuerelement initialisieren
Private Sub UserControl_InitProperties()
    m_BackColor = m_def_BackColor
    m_ForeColor = m_def_ForeColor
    m_Enabled = m_def_Enabled
    Set m_Font = Ambient.Font
    m_SelectionColor = m_def_SelectionColor
    m_TabSelectColor = m_def_TabSelectColor
    m_ComboType = m_def_ComboType
    m_DateFormat = m_def_DateFormat
    m_Day = m_def_Day
    m_Year = m_def_Year
    m_Month = m_def_Month
    m_PrinterName = m_def_PrinterName
    m_DriveLetter = m_def_DriveLetter
    m_FontName = m_def_FontName

    SetFont
    DrawDropDown
    
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> 1 Then Exit Sub
    MouseClick = True
    RaiseEvent MouseDown(Button, Shift, X, Y)

End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    GetCursorPos MouseOverMe
    If MousePos.Enabled = False Then MousePos.Enabled = True
    RaiseEvent MouseMove(Button, Shift, X, Y)

End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If Button <> 1 Then Exit Sub
    'If MouseClick = True Then RaiseEvent Click
    MouseClick = False
    RaiseEvent MouseUp(Button, Shift, X, Y)
    
End Sub
'Eigenschaftenwerte vom Speicher laden
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    m_BackColor = PropBag.ReadProperty("BackColor", m_def_BackColor)
    m_ForeColor = PropBag.ReadProperty("ForeColor", m_def_ForeColor)
    m_Enabled = PropBag.ReadProperty("Enabled", m_def_Enabled)
    Set m_Font = PropBag.ReadProperty("Font", Ambient.Font)
    m_SelectionColor = PropBag.ReadProperty("SelectionColor", m_def_SelectionColor)
    m_TabSelectColor = PropBag.ReadProperty("TabSelectColor", m_def_TabSelectColor)
    m_ComboType = PropBag.ReadProperty("ComboType", m_def_ComboType)
    m_DateFormat = PropBag.ReadProperty("DateFormat", m_def_DateFormat)
    m_Day = PropBag.ReadProperty("Day", m_def_Day)
    m_Year = PropBag.ReadProperty("Year", m_def_Year)
    m_Month = PropBag.ReadProperty("Month", m_def_Month)
    m_PrinterName = PropBag.ReadProperty("PrinterName", m_def_PrinterName)
    m_DriveLetter = PropBag.ReadProperty("DriveLetter", m_def_DriveLetter)
    m_FontName = PropBag.ReadProperty("FontName", m_def_FontName)
    
    SetFont
    AddIcons
    SetComboType
    SetDisplay
    DrawDropDown
    
End Sub

Private Sub UserControl_Resize()
    
    On Error Resume Next
    
    Text1.Left = ICON_SIZE + 6
    Text1.Top = 2
    Text1.Height = UserControl.TextHeight("I")
    UserControl.Height = (Text1.Height + 4) * Screen.TwipsPerPixelY
    
    cainXButton1.Height = UserControl.ScaleHeight - 4
    cainXButton1.Width = cainXButton1.Height
    cainXButton1.Top = 2
    cainXButton1.Left = UserControl.ScaleWidth - cainXButton1.Width - 2
    
    Text1.Width = UserControl.ScaleWidth - 13 - ICON_SIZE - cainXButton1.Width
    
    If UserControl.ScaleWidth < 177 Then
        picList.Width = 177 * Screen.TwipsPerPixelX
    Else
        picList.Width = UserControl.ScaleWidth * Screen.TwipsPerPixelX
    End If
    
    DrawDropDown
    
End Sub

'Eigenschaftenwerte in den Speicher schreiben
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("BackColor", m_BackColor, m_def_BackColor)
    Call PropBag.WriteProperty("ForeColor", m_ForeColor, m_def_ForeColor)
    Call PropBag.WriteProperty("Enabled", m_Enabled, m_def_Enabled)
    Call PropBag.WriteProperty("Font", m_Font, Ambient.Font)
    Call PropBag.WriteProperty("SelectionColor", m_SelectionColor, m_def_SelectionColor)
    Call PropBag.WriteProperty("TabSelectColor", m_TabSelectColor, m_def_TabSelectColor)
    Call PropBag.WriteProperty("ComboType", m_ComboType, m_def_ComboType)
    Call PropBag.WriteProperty("DateFormat", m_DateFormat, m_def_DateFormat)
    
    Call PropBag.WriteProperty("Day", m_Day, m_def_Day)
    Call PropBag.WriteProperty("Year", m_Year, m_def_Year)
    Call PropBag.WriteProperty("Month", m_Month, m_def_Month)
    Call PropBag.WriteProperty("PrinterName", m_PrinterName, m_def_PrinterName)
    Call PropBag.WriteProperty("DriveLetter", m_DriveLetter, m_def_DriveLetter)
    Call PropBag.WriteProperty("FontName", m_FontName, m_def_FontName)
    
    DrawDropDown
    
End Sub

Private Sub UserControl_Terminate()
    
'    If Not (ComboData Is Nothing) Then Set ComboData = Nothing
'    If Not (ListData Is Nothing) Then Set ListData = Nothing
    
End Sub

