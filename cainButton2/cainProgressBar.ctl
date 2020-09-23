VERSION 5.00
Begin VB.UserControl cainProgressBar 
   AutoRedraw      =   -1  'True
   CanGetFocus     =   0   'False
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
End
Attribute VB_Name = "cainProgressBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Enum MyType
    tNormal = 0
    tSearch = 1
End Enum

'Standard-Eigenschaftswerte:
Const m_def_ForeColor = &HC00000
Const m_def_BackColor = &HFFFFFF
Const m_def_Caption = "%"
Const m_def_Value = 50
Const m_def_Orientation = 1
Const m_def_Style = 0
'Eigenschaftsvariablen:
Dim m_ForeColor As OLE_COLOR
Dim m_BackColor As OLE_COLOR
Dim m_Caption As String
Dim m_Value As Integer
Dim m_OldValue As Integer
Dim m_Orientation As MyOrientation
Dim m_Style As MyType

Dim ProgChunks As Integer
Dim ChunkIndex As Integer

Private Const ChunkWidth As Integer = 6
Private Const ChunkGap As Integer = 2

Public Function ID() As Integer
    ID = eControlIDs.id_Progressbar
End Function

Private Function GetParentBackcolor() As Long
    GetParentBackcolor = UserControl.Parent.BackColor
End Function

Private Sub DrawFace(tColorSet As ColorSet)

    UserControl.BackColor = m_BackColor
    
    'background
    GradientCy UserControl.hdc, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight, tColorSet.csColor1(1), tColorSet.csColor1(2), tColorSet.csColor1(3), m_Orientation

    'Borders
    GradientLine UserControl.hdc, 1, 0, UserControl.ScaleWidth - 2, pbHorizontal, tColorSet.csColor1(4), tColorSet.csColor1(5)
    GradientLine UserControl.hdc, 0, 1, UserControl.ScaleHeight - 2, pbVertical, tColorSet.csColor1(4), tColorSet.csColor1(6)
    UserControl.Line (1, UserControl.ScaleHeight - 1)-(UserControl.ScaleWidth - 1, UserControl.ScaleHeight - 1), tColorSet.csColor1(6)
    GradientLine UserControl.hdc, UserControl.ScaleWidth - 1, 1, UserControl.ScaleHeight - 2, pbVertical, tColorSet.csColor1(5), tColorSet.csColor1(6)

End Sub

Private Sub DrawBorder(Optional bRefresh As Boolean = False)

    Dim Bordercolor(1 To 3) As Long
    Dim iWidth As Long
    Dim iHeight As Long
    Dim iLeft As Long
    Dim iPos As Long
    Dim i As Integer
    
    If (m_OldValue < m_Value) Or (m_Value = 0) Or (bRefresh = True) Or (m_Caption <> "") Or (m_Style = tSearch) Or (UserControl.Enabled = False) Then
    
        UserControl.Cls
                
        If UserControl.Enabled = False Then
            DrawFace GetColorSetDisabled(m_BackColor, m_ForeColor)
        Else
            DrawFace GetColorSetNormal(m_BackColor, m_ForeColor)
        End If
    
    End If
    
    If UserControl.Enabled = False Then Exit Sub
    
    If m_Value > 0 Then
        
        If m_Style = tNormal Then
       
            Bordercolor(1) = BlendColor(m_BackColor, m_ForeColor, 180)
            Bordercolor(2) = BlendColor(m_BackColor, m_ForeColor, 140)
            Bordercolor(3) = BlendColor(m_BackColor, m_ForeColor, 80)
        
            
            If m_Orientation = pbHorizontal Then
            
                iWidth = UserControl.ScaleWidth - 6
                ProgChunks = Fix(iWidth / (ChunkWidth + ChunkGap))
                If (ProgChunks * (ChunkWidth + ChunkGap)) > iWidth Then ProgChunks = ProgChunks - 1
                
                iPos = PercentOf(iWidth, m_Value)
                ChunkIndex = Get_ChunkIndex(iPos)
            
                iLeft = 4
                
                For i = 1 To ChunkIndex
                    GradientCy UserControl.hdc, iLeft, 3, ChunkWidth, UserControl.ScaleHeight - 3, Bordercolor(1), Bordercolor(2), Bordercolor(3), pbHorizontal
                    iLeft = iLeft + ChunkWidth + ChunkGap
                Next i
            
            ElseIf m_Orientation = pbVertical Then
            
                iHeight = UserControl.ScaleHeight - 6
                ProgChunks = Fix(iHeight / (ChunkWidth + ChunkGap))
                If (ProgChunks * (ChunkWidth + ChunkGap)) > iHeight Then ProgChunks = ProgChunks - 1
                
                iPos = PercentOf(iHeight, m_Value)
                ChunkIndex = Get_ChunkIndex(iPos)
            
                iLeft = 4
                
                For i = 1 To ChunkIndex
                    GradientCy UserControl.hdc, 3, (UserControl.ScaleHeight - 6) - iLeft, UserControl.ScaleWidth - 3, ChunkWidth, Bordercolor(1), Bordercolor(2), Bordercolor(3), pbVertical
                    iLeft = iLeft + ChunkWidth + ChunkGap
                Next i
            
            End If
            
        ElseIf m_Style = tSearch Then
            
            Bordercolor(1) = BlendColor(m_BackColor, m_ForeColor, 180)
            Bordercolor(2) = BlendColor(m_BackColor, m_ForeColor, 140)
            Bordercolor(3) = BlendColor(m_BackColor, m_ForeColor, 80)
            
            If m_Orientation = pbHorizontal Then
            
                iWidth = (UserControl.ScaleWidth - 6) + ((ChunkWidth + ChunkGap) * 3) + 1
                iPos = PercentOf(iWidth, m_Value)
                
                iLeft = iPos
                iWidth = ChunkWidth
                If iLeft - 3 < iWidth Then iWidth = iLeft - 3
                If iLeft > UserControl.ScaleWidth - 3 Then
                    iLeft = UserControl.ScaleWidth - 3
                    iWidth = ChunkWidth - (iPos - (UserControl.ScaleWidth - 3))
                End If
                
                If iWidth >= 1 Then
                    GradientCy UserControl.hdc, iLeft - iWidth, 3, iWidth, UserControl.ScaleHeight - 3, Bordercolor(1), Bordercolor(2), Bordercolor(3), pbHorizontal
                End If
                
                If iPos > ChunkWidth + ChunkGap Then
    
                    iLeft = iPos - (ChunkWidth + ChunkGap)
                    iWidth = ChunkWidth
                    If iLeft - 3 < iWidth Then iWidth = iLeft - 3
                    If iLeft > UserControl.ScaleWidth - 3 Then
                        iLeft = UserControl.ScaleWidth - 3
                        iWidth = ChunkWidth - ((iPos - (ChunkWidth + ChunkGap)) - (UserControl.ScaleWidth - 3))
                    End If
                
                    Bordercolor(1) = BlendColor(m_BackColor, m_ForeColor, 198)
                    Bordercolor(2) = BlendColor(m_BackColor, m_ForeColor, 165)
                    Bordercolor(3) = BlendColor(m_BackColor, m_ForeColor, 120)
    
                    If iWidth >= 1 Then
                        GradientCy UserControl.hdc, iLeft - iWidth, 3, iWidth, UserControl.ScaleHeight - 3, Bordercolor(1), Bordercolor(2), Bordercolor(3), pbHorizontal
                    End If
    
                End If
                
                If iPos > (ChunkWidth + ChunkGap) * 2 Then
    
                    iLeft = iPos - ((ChunkWidth + ChunkGap) * 2)
                    iWidth = ChunkWidth
                    If iLeft - 3 < iWidth Then iWidth = iLeft - 3
                    If iLeft > UserControl.ScaleWidth - 3 Then
                        iLeft = UserControl.ScaleWidth - 3
                        iWidth = ChunkWidth - ((iPos - ((ChunkWidth + ChunkGap) * 2)) - (UserControl.ScaleWidth - 3))
                    End If
                
                    Bordercolor(1) = BlendColor(m_BackColor, m_ForeColor, 210)
                    Bordercolor(2) = BlendColor(m_BackColor, m_ForeColor, 180)
                    Bordercolor(3) = BlendColor(m_BackColor, m_ForeColor, 130)
   
                    If iWidth >= 1 Then
                        GradientCy UserControl.hdc, iLeft - iWidth, 3, iWidth, UserControl.ScaleHeight - 3, Bordercolor(1), Bordercolor(2), Bordercolor(3), pbHorizontal
                    End If
    
                End If
                
            ElseIf m_Orientation = pbVertical Then
            
                iHeight = (UserControl.ScaleHeight - 6) + ((ChunkWidth + ChunkGap) * 3) + 1
                iPos = PercentOf(iHeight, m_Value)
                
                iLeft = UserControl.ScaleHeight - iPos
                iHeight = ChunkWidth
                
                If iLeft <= 3 Then
                
                    iLeft = 3
                    If (UserControl.ScaleHeight - iPos) < 3 Then
                        iHeight = ChunkWidth - (3 - (UserControl.ScaleHeight - iPos))
                    End If
                    
                ElseIf iLeft < UserControl.ScaleHeight - 3 Then
                
                    If (UserControl.ScaleHeight - 3) - iLeft < ChunkWidth Then
                        iHeight = ((UserControl.ScaleHeight - 3) - iLeft)
                    Else
                        iHeight = ChunkWidth
                    
                    End If
                
                Else
                    iHeight = 0
                End If
                
                If iHeight >= 1 Then
                    GradientCy UserControl.hdc, 3, iLeft, UserControl.ScaleWidth - 3, iHeight, Bordercolor(1), Bordercolor(2), Bordercolor(3), pbVertical
                End If
                
                If iPos > ChunkWidth + ChunkGap Then
                
                    iLeft = (UserControl.ScaleHeight - iPos) + (ChunkWidth + ChunkGap)
                    iHeight = ChunkWidth
                    
                    If iLeft <= 3 Then
                    
                        iLeft = 3
                        If (UserControl.ScaleHeight - iPos) + (ChunkWidth + ChunkGap) < 3 Then
                            iHeight = ChunkWidth - (3 - ((UserControl.ScaleHeight - iPos) + (ChunkWidth + ChunkGap)))
                        End If
                        
                    ElseIf iLeft < UserControl.ScaleHeight - 3 Then
                    
                        If (UserControl.ScaleHeight - 3) - iLeft < ChunkWidth Then
                            iHeight = ((UserControl.ScaleHeight - 3) - iLeft)
                        Else
                            iHeight = ChunkWidth
                        End If
                    
                    Else
                        iHeight = 0
                    End If
                    
                    Bordercolor(1) = BlendColor(m_BackColor, m_ForeColor, 198)
                    Bordercolor(2) = BlendColor(m_BackColor, m_ForeColor, 165)
                    Bordercolor(3) = BlendColor(m_BackColor, m_ForeColor, 120)
                    
                    If iHeight >= 1 Then
                        GradientCy UserControl.hdc, 3, iLeft, UserControl.ScaleWidth - 3, iHeight, Bordercolor(1), Bordercolor(2), Bordercolor(3), pbVertical
                    End If
                    
                End If
                
                If iPos > (ChunkWidth + ChunkGap) * 2 Then
                
                    iLeft = (UserControl.ScaleHeight - iPos) + ((ChunkWidth + ChunkGap) * 2)
                    iHeight = ChunkWidth
                    
                    If iLeft <= 3 Then
                    
                        iLeft = 3
                        If (UserControl.ScaleHeight - iPos) + ((ChunkWidth + ChunkGap) * 2) < 3 Then
                            iHeight = ChunkWidth - (3 - ((UserControl.ScaleHeight - iPos) + ((ChunkWidth + ChunkGap) * 2)))
                        End If
                        
                    ElseIf iLeft < UserControl.ScaleHeight - 3 Then
                    
                        If (UserControl.ScaleHeight - 3) - iLeft < ChunkWidth Then
                            iHeight = ((UserControl.ScaleHeight - 3) - iLeft)
                        Else
                            iHeight = ChunkWidth
                        End If
                    
                    Else
                        iHeight = 0
                    End If
                    
                    Bordercolor(1) = BlendColor(m_BackColor, m_ForeColor, 210)
                    Bordercolor(2) = BlendColor(m_BackColor, m_ForeColor, 180)
                    Bordercolor(3) = BlendColor(m_BackColor, m_ForeColor, 130)
                                        
                    If iHeight >= 1 Then
                        GradientCy UserControl.hdc, 3, iLeft, UserControl.ScaleWidth - 3, iHeight, Bordercolor(1), Bordercolor(2), Bordercolor(3), pbVertical
                    End If
                    
                End If
            
            End If
            
        End If
        
    End If
    
    If m_Caption <> "" Then
        
        If m_Caption = "%" And m_Style = tNormal Then
            DrawCaption UserControl.Font, m_Value & " %", m_ForeColor
        ElseIf m_Caption = "%" And m_Style = tSearch Then
            DrawCaption UserControl.Font, "Searching...", m_ForeColor
        Else
            DrawCaption UserControl.Font, m_Caption, m_ForeColor
        End If
        
    End If
    
End Sub

Private Sub DrawCaption(fntFont As Font, strText As String, FntColor As Long)
    
    If m_Orientation = pbVertical Then Exit Sub
    
    Set UserControl.Font = fntFont
    UserControl.ForeColor = m_BackColor
      
    '==== FIRST LAYER
    UserControl.CurrentX = Fix(((UserControl.ScaleWidth / 2) - (UserControl.TextWidth(strText) / 2))) - 1
    UserControl.CurrentY = Fix(((UserControl.ScaleHeight / 2) - (UserControl.TextHeight("I") / 2)))

    UserControl.Print strText
        
    UserControl.CurrentX = Fix(((UserControl.ScaleWidth / 2) - (UserControl.TextWidth(strText) / 2)))
    UserControl.CurrentY = Fix(((UserControl.ScaleHeight / 2) - (UserControl.TextHeight("I") / 2))) - 1

    'UserControl.ForeColor = m_BackColor
    UserControl.Print strText
        
    UserControl.CurrentX = Fix(((UserControl.ScaleWidth / 2) - (UserControl.TextWidth(strText) / 2))) + 1
    UserControl.CurrentY = Fix(((UserControl.ScaleHeight / 2) - (UserControl.TextHeight("I") / 2)))

    'UserControl.ForeColor = m_BackColor
    UserControl.Print strText
        
    UserControl.CurrentX = Fix(((UserControl.ScaleWidth / 2) - (UserControl.TextWidth(strText) / 2)))
    UserControl.CurrentY = Fix(((UserControl.ScaleHeight / 2) - (UserControl.TextHeight("I") / 2))) + 1

    'UserControl.ForeColor = m_BackColor
    UserControl.Print strText
    
    UserControl.CurrentX = Fix(((UserControl.ScaleWidth / 2) - (UserControl.TextWidth(strText) / 2))) - 1
    UserControl.CurrentY = Fix(((UserControl.ScaleHeight / 2) - (UserControl.TextHeight("I") / 2))) - 1

    'UserControl.ForeColor = m_BackColor
    UserControl.Print strText
        
    UserControl.CurrentX = Fix(((UserControl.ScaleWidth / 2) - (UserControl.TextWidth(strText) / 2))) + 1
    UserControl.CurrentY = Fix(((UserControl.ScaleHeight / 2) - (UserControl.TextHeight("I") / 2))) - 1

    'UserControl.ForeColor = m_BackColor
    UserControl.Print strText
        
    UserControl.CurrentX = Fix(((UserControl.ScaleWidth / 2) - (UserControl.TextWidth(strText) / 2))) + 1
    UserControl.CurrentY = Fix(((UserControl.ScaleHeight / 2) - (UserControl.TextHeight("I") / 2))) + 1

    'UserControl.ForeColor = m_BackColor
    UserControl.Print strText
        
    UserControl.CurrentX = Fix(((UserControl.ScaleWidth / 2) - (UserControl.TextWidth(strText) / 2))) - 1
    UserControl.CurrentY = Fix(((UserControl.ScaleHeight / 2) - (UserControl.TextHeight("I") / 2))) + 1

    'UserControl.ForeColor = m_BackColor
    UserControl.Print strText
    
    '==== SECOND LAYER
    UserControl.CurrentX = Fix(((UserControl.ScaleWidth / 2) - (UserControl.TextWidth(strText) / 2))) - 2
    UserControl.CurrentY = Fix(((UserControl.ScaleHeight / 2) - (UserControl.TextHeight("I") / 2)))

    'UserControl.ForeColor = m_BackColor
    UserControl.Print strText
        
    UserControl.CurrentX = Fix(((UserControl.ScaleWidth / 2) - (UserControl.TextWidth(strText) / 2)))
    UserControl.CurrentY = Fix(((UserControl.ScaleHeight / 2) - (UserControl.TextHeight("I") / 2))) - 2

    'UserControl.ForeColor = m_BackColor
    UserControl.Print strText
        
    UserControl.CurrentX = Fix(((UserControl.ScaleWidth / 2) - (UserControl.TextWidth(strText) / 2))) + 2
    UserControl.CurrentY = Fix(((UserControl.ScaleHeight / 2) - (UserControl.TextHeight("I") / 2)))

    'UserControl.ForeColor = m_BackColor
    UserControl.Print strText
        
    UserControl.CurrentX = Fix(((UserControl.ScaleWidth / 2) - (UserControl.TextWidth(strText) / 2)))
    UserControl.CurrentY = Fix(((UserControl.ScaleHeight / 2) - (UserControl.TextHeight("I") / 2))) + 2

    'UserControl.ForeColor = m_BackColor
    UserControl.Print strText
    
    UserControl.CurrentX = Fix(((UserControl.ScaleWidth / 2) - (UserControl.TextWidth(strText) / 2))) - 2
    UserControl.CurrentY = Fix(((UserControl.ScaleHeight / 2) - (UserControl.TextHeight("I") / 2))) - 2

    'UserControl.ForeColor = m_BackColor
    UserControl.Print strText
        
    UserControl.CurrentX = Fix(((UserControl.ScaleWidth / 2) - (UserControl.TextWidth(strText) / 2))) + 2
    UserControl.CurrentY = Fix(((UserControl.ScaleHeight / 2) - (UserControl.TextHeight("I") / 2))) - 2

    'UserControl.ForeColor = m_BackColor
    UserControl.Print strText
        
    UserControl.CurrentX = Fix(((UserControl.ScaleWidth / 2) - (UserControl.TextWidth(strText) / 2))) + 2
    UserControl.CurrentY = Fix(((UserControl.ScaleHeight / 2) - (UserControl.TextHeight("I") / 2))) + 2

    'UserControl.ForeColor = m_BackColor
    UserControl.Print strText
        
    UserControl.CurrentX = Fix(((UserControl.ScaleWidth / 2) - (UserControl.TextWidth(strText) / 2))) - 2
    UserControl.CurrentY = Fix(((UserControl.ScaleHeight / 2) - (UserControl.TextHeight("I") / 2))) + 2

    'UserControl.ForeColor = m_BackColor
    UserControl.Print strText
    
    '=============================================================
    UserControl.CurrentX = Fix(((UserControl.ScaleWidth / 2) - (UserControl.TextWidth(strText) / 2)))
    UserControl.CurrentY = Fix(((UserControl.ScaleHeight / 2) - (UserControl.TextHeight("I") / 2)))

    UserControl.ForeColor = FntColor
    UserControl.Print strText
    
End Sub

Private Function Get_ChunkIndex(iPosition As Long) As Integer
    
    Dim i As Integer
    
    i = Fix(iPosition / (ChunkWidth + ChunkGap)) + 1
    If i > ProgChunks Then _
        i = ProgChunks
    
    Get_ChunkIndex = i

End Function

Private Sub UserControl_Initialize()
    'DrawBorder True
End Sub
'ACHTUNG! DIE FOLGENDEN KOMMENTIERTEN ZEILEN NICHT ENTFERNEN ODER VERÄNDERN!
'MappingInfo=UserControl,UserControl,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Gibt die Hintergrundfarbe zurück, die verwendet wird, um Text und Grafik in einem Objekt anzuzeigen, oder legt diese fest."
Attribute BackColor.VB_ProcData.VB_Invoke_Property = ";Darstellung"
    BackColor = m_BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    m_BackColor = New_BackColor
    PropertyChanged "BackColor"
    DrawBorder
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
    DrawBorder
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
    DrawBorder
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
    DrawBorder
End Property

'ACHTUNG! DIE FOLGENDEN KOMMENTIERTEN ZEILEN NICHT ENTFERNEN ODER VERÄNDERN!
'MemberInfo=7,0,0,0
Public Property Get Style() As MyType
Attribute Style.VB_ProcData.VB_Invoke_Property = ";Darstellung"
    Style = m_Style
End Property

Public Property Let Style(ByVal New_Style As MyType)
    m_Style = New_Style
    PropertyChanged "Style"
    
    DrawBorder
End Property

'Eigenschaften für Benutzersteuerelement initialisieren
Private Sub UserControl_InitProperties()

    Set UserControl.Font = Ambient.Font
    m_Orientation = m_def_Orientation
    m_Style = m_def_Style
    m_Value = m_def_Value
    
    m_Caption = m_def_Caption
    m_ForeColor = m_def_ForeColor
    m_BackColor = m_def_BackColor
End Sub

'Eigenschaftenwerte vom Speicher laden
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    m_BackColor = PropBag.ReadProperty("BackColor", m_def_BackColor)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    Set UserControl.Font = PropBag.ReadProperty("Font", Ambient.Font)
    m_Orientation = PropBag.ReadProperty("Orientation", m_def_Orientation)
    m_Style = PropBag.ReadProperty("Style", m_def_Style)
    m_Value = PropBag.ReadProperty("Value", m_def_Value)
    m_Caption = PropBag.ReadProperty("Caption", m_def_Caption)
    m_ForeColor = PropBag.ReadProperty("ForeColor", m_def_ForeColor)
    
    DrawBorder
    
End Sub

Private Sub UserControl_Resize()
    DrawBorder True
End Sub

'Eigenschaftenwerte in den Speicher schreiben
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("BackColor", m_BackColor, m_def_BackColor)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("Font", UserControl.Font, Ambient.Font)
    Call PropBag.WriteProperty("Orientation", m_Orientation, m_def_Orientation)
    Call PropBag.WriteProperty("Style", m_Style, m_def_Style)
    Call PropBag.WriteProperty("Value", m_Value, m_def_Value)
    Call PropBag.WriteProperty("Caption", m_Caption, m_def_Caption)
    Call PropBag.WriteProperty("ForeColor", m_ForeColor, m_def_ForeColor)
    
    DrawBorder
    
End Sub

'ACHTUNG! DIE FOLGENDEN KOMMENTIERTEN ZEILEN NICHT ENTFERNEN ODER VERÄNDERN!
'MemberInfo=7,0,0,0
Public Property Get Value() As Integer
Attribute Value.VB_ProcData.VB_Invoke_Property = ";Verschiedenes"
    Value = m_Value
End Property

Public Property Let Value(ByVal New_Value As Integer)
    
    If New_Value > 100 Then New_Value = 100
    If New_Value < 0 Then New_Value = 0
    
    m_Value = New_Value
    PropertyChanged "Value"
    DrawBorder
    
End Property

'ACHTUNG! DIE FOLGENDEN KOMMENTIERTEN ZEILEN NICHT ENTFERNEN ODER VERÄNDERN!
'MemberInfo=13,0,0,%
Public Property Get Caption() As String
Attribute Caption.VB_ProcData.VB_Invoke_Property = ";Darstellung"
    Caption = m_Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    m_Caption = New_Caption
    PropertyChanged "Caption"
    DrawBorder
End Property

'ACHTUNG! DIE FOLGENDEN KOMMENTIERTEN ZEILEN NICHT ENTFERNEN ODER VERÄNDERN!
'MemberInfo=10,0,0,255
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Gibt die Vordergrundfarbe zurück, die zum Anzeigen von Text und Grafiken in einem Objekt verwendet wird, oder legt diese fest."
Attribute ForeColor.VB_ProcData.VB_Invoke_Property = ";Darstellung"
    ForeColor = m_ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    m_ForeColor = New_ForeColor
    PropertyChanged "ForeColor"
    DrawBorder
End Property

