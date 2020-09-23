VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Search"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'By: Caino MDK
'loganx@t-online.de
'2004
'CainClasses.dll

'lokale Variable(n) zum Zuweisen der Eigenschaft(en)
Private mvarSeparator As Byte 'lokale Kopie
Private sFilelist As New sf_Information

Public Property Let Separator(ByVal vData As Byte)
'wird beim Zuweisen eines Werts zu der Eigenschaft auf der linken Seite einer Zuweisung verwendet.
'Syntax: X.Separator = 5
    mvarSeparator = vData
End Property

Public Property Get Separator() As Byte
'wird beim Ermitteln eines Eigenschaftswertes auf der rechten Seite einer Zuweisung verwendet.
'Syntax: Debug.Print X.Separator
    Separator = mvarSeparator
End Property

Public Sub Search_File(ByVal sDirectory As String, sFileToSearch As String, FileAttributes As VbFileAttribute, cFoundFiles As sf_Information)
  
    Dim DirPaths() As String
    Dim DirPathCount As Integer
    Dim tmpInt As Integer
    Dim TmpInt2 As Integer
    Dim tmpString As String
    Dim FoundFile As String
    Dim tmpFoundFiles As New sf_Information
    Dim tmpAttrib As VbFileAttribute
    Dim ChopedFileToSearch() As String
    Dim ChopedFileToSearchCount As Integer
    
    DirPathCount = SplitStrings(sDirectory, mvarSeparator, DirPaths())
    ChopedFileToSearchCount = SplitStrings(sFileToSearch, mvarSeparator, ChopedFileToSearch())
  
    On Error Resume Next
    If ChopedFileToSearchCount = 0 Then
        Set cFoundFiles = tmpFoundFiles
        Exit Sub
    End If
    
    For TmpInt2 = 1 To ChopedFileToSearchCount
    
        For tmpInt = 1 To DirPathCount
        
          tmpString = DirPaths(tmpInt)
          If Right(tmpString, 1) <> "\" Then _
              tmpString = tmpString & "\"
          
          FoundFile = Dir(tmpString & ChopedFileToSearch(TmpInt2), FileAttributes)
          
          Do
            If Trim(FoundFile) = "" Then Exit Do
            'DoEvents
            
                tmpAttrib = GetAttr(tmpString & FoundFile)
                If ((tmpAttrib And vbDirectory) <> vbDirectory) And ((tmpAttrib And FileAttributes) = FileAttributes) Then
                    tmpFoundFiles.Add tmpString, FoundFile, FileDateTime(tmpString & FoundFile), tmpAttrib, FileLen(tmpString & FoundFile), FoundFile
                End If
            
            FoundFile = Dir
          Loop
        
        Next tmpInt
    Next TmpInt2
    
    Set cFoundFiles = tmpFoundFiles
  
End Sub

Public Sub Search_Directory(ByVal sDirectory As String, sDirToSearch As String, cFoundDirs As sf_Information)
  
    Dim DirPaths() As String
    Dim DirPathCount As Integer
    Dim tmpInt As Integer
    Dim TmpInt2 As Integer
    Dim tmpString As String
    Dim FoundDir As String
    Dim tmpFoundDirs As New sf_Information
    Dim tmpAttrib As VbFileAttribute
    Dim ChopedDirToSearch() As String
    Dim ChopedDirToSearchCount As Integer
  
    DirPathCount = SplitStrings(sDirectory, mvarSeparator, DirPaths())
    ChopedDirToSearchCount = SplitStrings(sDirToSearch, mvarSeparator, ChopedDirToSearch())
    
    On Error Resume Next
    If ChopedDirToSearchCount = 0 Then: _
        Set cFoundDirs = tmpFoundDirs: _
        Exit Sub
        
    For TmpInt2 = 1 To ChopedDirToSearchCount
    
        For tmpInt = 1 To DirPathCount
        
            tmpString = DirPaths(tmpInt)
            If Right(tmpString, 1) <> "\" Then _
                tmpString = tmpString & "\"
            
            FoundDir = Dir(tmpString & ChopedDirToSearch(TmpInt2), vbDirectory + vbNormal + vbHidden)
            
            Do
                'DoEvents
                If Trim(FoundDir) = "" Then Exit Do
                If FoundDir <> ".." And FoundDir <> "." And FoundDir <> "" Then
                
                    tmpAttrib = GetAttr(tmpString & FoundDir)
                    If (tmpAttrib And vbDirectory) = vbDirectory Then
                        tmpFoundDirs.Add tmpString, FoundDir, FileDateTime(tmpString & FoundDir), tmpAttrib, 0, FoundDir
                    End If
                
                
                End If
                
                FoundDir = Dir
            Loop
        
        Next tmpInt
    Next TmpInt2
    
    Set cFoundDirs = tmpFoundDirs
    
    End Sub

Private Function SplitStrings(ByVal sPaths As String, bSeparator As Byte, sSplitedPaths() As String) As Integer
    
    Dim lPosition As Long
    Dim tmpSplitedPaths() As String
    Dim tmpInt As Integer
    Dim TmpInt2 As Integer
    
    If bSeparator = 0 Then
      bSeparator = mvarSeparator
    End If
    
    tmpInt = 1
    TmpInt2 = 1
    
    Do
            
        lPosition = InStr(tmpInt, sPaths & Chr(bSeparator), Chr(bSeparator))
        If lPosition = 0 Then
            
            If tmpInt = 1 Then: _
                ReDim Preserve tmpSplitedPaths(1): _
                tmpSplitedPaths(TmpInt2) = Trim(sPaths)
                
            Exit Do
            
        End If
        
        ReDim Preserve tmpSplitedPaths(TmpInt2)
        tmpSplitedPaths(TmpInt2) = Trim(Mid(sPaths & Chr(bSeparator), tmpInt, (lPosition - tmpInt)))
        
        tmpInt = lPosition + 1
        TmpInt2 = TmpInt2 + 1
            
    Loop
    
    sSplitedPaths() = tmpSplitedPaths()
    SplitStrings = TmpInt2 - 1
    
End Function

Private Function CorrectPath(sPath As String) As String
    
    If Right(sPath, 1) = "\" Or Right(sPath, 1) = "/" Then
        CorrectPath = sPath
    Else
        CorrectPath = sPath & "\"
    End If
    
End Function

Public Sub ScanPathforFiles(SearchPath As String, sFilter As String, cFoundFiles As sf_Information)

    subScanPath SearchPath, sFilter
    Set cFoundFiles = sFilelist
    
    If Not (sFilelist Is Nothing) Then Set sFilelist = Nothing
    
End Sub

Private Sub subScanPath(SearchPath As String, sFilter As String)

    Dim cSupPaths As sf_Information
    Dim cFiles As sf_Information
    Dim i As Integer
    Dim i2 As Integer
    
    On Error Resume Next
    
    Search_Directory SearchPath, "*.*", cSupPaths
    Search_File SearchPath, sFilter, vbNormal, cFiles

    If cFiles.Count <> 0 Then
        For i = 1 To cFiles.Count
            sFilelist.Add CorrectPath(cFiles.Item(i).Path), cFiles.Item(i).Filename, cFiles(i).DateTime, cFiles(i).Attri, cFiles(i).Size
        Next i
    End If
    
    If cSupPaths.Count <> 0 Then
        For i = 1 To cSupPaths.Count
            
            DoEvents
            Search_File CorrectPath(cSupPaths.Item(i).Path) & cSupPaths.Item(i).Filename, sFilter, vbNormal, cFiles
    
            If cFiles.Count <> 0 Then
                For i2 = 1 To cFiles.Count
                    DoEvents
                    sFilelist.Add CorrectPath(cFiles.Item(i).Path), cFiles.Item(i).Filename, cFiles(i).DateTime, cFiles(i).Attri, cFiles(i).Size
                Next i2
            End If
            
            subScanPath CorrectPath(cSupPaths.Item(i).Path) & cSupPaths.Item(i).Filename, sFilter
        Next i
    End If

End Sub

Public Sub ScanPathforPaths(SearchPath As String, cFoundFiles As sf_Information)

    subScanPathPath SearchPath
    Set cFoundFiles = sFilelist
    
    If Not (sFilelist Is Nothing) Then Set sFilelist = Nothing
    
End Sub

Private Sub subScanPathPath(SearchPath As String)

    Dim cSupPaths As sf_Information
    Dim i As Integer
    Dim i2 As Integer
    
    On Error Resume Next
    
    Search_Directory SearchPath, "*.*", cSupPaths

    If cSupPaths.Count <> 0 Then
        For i = 1 To cSupPaths.Count
            sFilelist.Add CorrectPath(cSupPaths.Item(i).Path), cSupPaths.Item(i).Filename, cSupPaths(i).DateTime, cSupPaths(i).Attri, cSupPaths(i).Size
        Next i
    End If
    
    If cSupPaths.Count <> 0 Then
        For i = 1 To cSupPaths.Count
            
            DoEvents
            Search_Directory CorrectPath(cSupPaths.Item(i).Path) & cSupPaths.Item(i).Filename, "*.*", cSupPaths
    
            If cSupPaths.Count <> 0 Then
                For i2 = 1 To cSupPaths.Count
                    DoEvents
                    sFilelist.Add CorrectPath(cSupPaths.Item(i).Path), cSupPaths.Item(i).Filename, cSupPaths(i).DateTime, cSupPaths(i).Attri, cSupPaths(i).Size
                Next i2
            End If
            
            subScanPathPath CorrectPath(cSupPaths.Item(i).Path) & cSupPaths.Item(i).Filename
        Next i
    End If

End Sub

Private Sub Class_Initialize()
    mvarSeparator = Asc(";")
End Sub

Private Sub Class_Terminate()
    'MsgBox "Terminated"
End Sub


