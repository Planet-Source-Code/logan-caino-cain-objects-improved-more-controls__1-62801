VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IconExtractor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"

'Written by Caino



Option Explicit

Public Enum IconSize

    LargeIcon = 0
    SmallIcon = 1
    
End Enum

Private mvarMaskColor As Long 'lokale Kopie

Public Function GetIconCount(sIconPath As String) As Long
    GetIconCount = ExtractIconEx(sIconPath, -1, 0, 0, 0)
End Function

Public Function ExtractIcon(sIconPath As String, iIconIndex As Integer, Optional iIconSize As IconSize) As IPictureDisp
    
    Dim lLargeIcon As Long, lSmallIcon As Long
    
    With frmEdit.Picture1
        .AutoRedraw = True
        .ScaleMode = 1
        .BackColor = mvarMaskColor
        .Picture = LoadPicture("")
    
        Call ExtractIconEx(sIconPath, iIconIndex, lLargeIcon, lSmallIcon, 1)
        
        If iIconSize = SmallIcon Then
            .Width = 16
            .Height = 16
            Call DrawIconEx(.hdc, 0, 0, lSmallIcon, 16, 16, 0, 0, 3)
        Else
            .Width = 32
            .Height = 32
            Call DrawIconEx(.hdc, 0, 0, lLargeIcon, 32, 32, 0, 0, 3)
        End If
        
        Set ExtractIcon = .Image
    
    End With
    
End Function

Public Function ExtractIconOfFileType(sFileType As String, Optional iIconSize As IconSize) As IPictureDisp
    
    Dim lLargeIcon As Long, lSmallIcon As Long
    Dim iIconIndex As Long
    Dim tmpsIconPath As String
    Dim sIconPath As String
    Dim iDividerPosition As Integer
    
    With frmEdit.Picture1
        .AutoRedraw = True
        .ScaleMode = 1
        .BackColor = mvarMaskColor
        .Width = 16
        .Height = 16
        .Picture = LoadPicture("")
        
        tmpsIconPath = GetSettingString(HKEY_CLASSES_ROOT, sFileType, "", "")
        
        If tmpsIconPath = "" Then: _
            tmpsIconPath = "Unknown"
        
        sIconPath = GetSettingString(HKEY_CLASSES_ROOT, tmpsIconPath & "\DefaultIcon", "", "")
        
        If sIconPath = "" Or InStr(1, sIconPath, WIN_ICO_ME) <> 0 Then: _
            sIconPath = WIN_WINDOWSROOT_TAG & "system32\shell32.dll,0"
        
        iDividerPosition = InStr(1, sIconPath, ",")
        iIconIndex = Mid(sIconPath, iDividerPosition + 1, Len(sIconPath) - iDividerPosition) * 1
        sIconPath = Trim(Replace(Left(sIconPath, iDividerPosition - 1), WIN_WINDOWSROOT_TAG, Get_WindowsDir))
        
        Call ExtractIconEx(sIconPath, iIconIndex, lLargeIcon, lSmallIcon, 1)
        
        If iIconSize = SmallIcon Then
            .Width = 16
            .Height = 16
            Call DrawIconEx(.hdc, 0, 0, lSmallIcon, 16, 16, 0, 0, 3)
        Else
            .Width = 32
            .Height = 32
            Call DrawIconEx(.hdc, 0, 0, lLargeIcon, 32, 32, 0, 0, 3)
        End If
        
        Set ExtractIconOfFileType = .Image
    End With
    
End Function

'lokale Variable(n) zum Zuweisen der Eigenschaft(en)
Public Property Let MaskColor(ByVal vData As Long)
'wird beim Zuweisen eines Werts zu der Eigenschaft auf der linken Seite einer Zuweisung verwendet.
'Syntax: X.MaskColor = 5
    mvarMaskColor = vData
End Property

Public Property Get MaskColor() As Long
'wird beim Ermitteln eines Eigenschaftswertes auf der rechten Seite einer Zuweisung verwendet.
'Syntax: Debug.Print X.MaskColor
    MaskColor = mvarMaskColor
End Property

Private Sub Class_Terminate()
    On Error Resume Next
    Unload frmEdit
End Sub
