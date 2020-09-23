Attribute VB_Name = "modReg"
Option Explicit

'Original from: Unknown
'Edited by: Caino

Public Enum HKEY_KeyWords

    HKEY_CLASSES_ROOT = &H80000000
    HKEY_CURRENT_USER = &H80000001
    HKEY_LOCAL_MACHINE = &H80000002
    HKEY_USERS = &H80000003
    HKEY_PERFORMANCE_DATA = &H80000004
    HKEY_CURRENT_CONFIG = &H80000005
    HKEY_DYN_DATA = &H80000006

End Enum

Public Function GetSettingString(Hkey As HKEY_KeyWords, strPath As String, strValue As String, Optional Default As String) As String
    Dim hCurKey As Long
    Dim lValueType As Long
    Dim strBuffer As String
    Dim lDataBufferSize As Long
    Dim intZeroPos As Integer
    Dim lRegResult As Long
    
    ' Set up default value
    If Not IsEmpty(Default) Then
      GetSettingString = Default
    Else
      GetSettingString = ""
    End If
    
    ' Open the key and get length of string
    lRegResult = RegOpenKey(Hkey, strPath, hCurKey)
    lRegResult = RegQueryValueEx(hCurKey, strValue, 0&, lValueType, ByVal 0&, lDataBufferSize)
    
    If lRegResult = ERROR_SUCCESS Then
    
      If lValueType = REG_SZ Or lValueType = REG_EXPAND_SZ Or lValueType = REG_MULTI_SZ Then
        ' initialise string buffer and retrieve string
        strBuffer = String(lDataBufferSize, " ")
        lRegResult = RegQueryValueEx(hCurKey, strValue, 0&, 0&, ByVal strBuffer, lDataBufferSize)
        
        ' format string
        intZeroPos = InStr(strBuffer, Chr$(0))
        If intZeroPos > 0 Then
          GetSettingString = Left$(strBuffer, intZeroPos - 1)
        Else
          GetSettingString = strBuffer
        End If
    
      End If
    
    Else
      ' there is a problem
    End If
    
    lRegResult = RegCloseKey(hCurKey)
End Function

Public Function GetSettingLong(ByVal Hkey As HKEY_KeyWords, ByVal strPath As String, ByVal strValue As String, Optional Default As Long) As Long

    Dim lRegResult As Long
    Dim lValueType As Long
    Dim lBuffer As Long
    Dim lDataBufferSize As Long
    Dim hCurKey As Long
    
    ' Set up default value
    If Not IsEmpty(Default) Then
      GetSettingLong = Default
    Else
      GetSettingLong = 0
    End If
    
    lRegResult = RegOpenKey(Hkey, strPath, hCurKey)
    lDataBufferSize = 4       ' 4 bytes = 32 bits = long
    
    lRegResult = RegQueryValueEx(hCurKey, strValue, 0&, lValueType, lBuffer, lDataBufferSize)
    
    If lRegResult = ERROR_SUCCESS Then
    
      If lValueType = REG_DWORD Then
        GetSettingLong = lBuffer
      End If
    
    Else
      'there is a problem
    End If
    
    lRegResult = RegCloseKey(hCurKey)

End Function

Public Function GetSettingByte(ByVal Hkey As HKEY_KeyWords, ByVal strPath As String, ByVal strValueName As String, Optional Default As Variant) As Variant
    Dim lValueType As Long
    Dim byBuffer() As Byte
    Dim lDataBufferSize As Long
    Dim lRegResult As Long
    Dim hCurKey As Long
    
    ' setup default value
    If Not IsEmpty(Default) Then
      If VarType(Default) = vbArray + vbByte Then
        GetSettingByte = Default
      Else
        GetSettingByte = 0
      End If
    
    Else
      GetSettingByte = 0
    End If
    
    ' Open the key and get number of bytes
    lRegResult = RegOpenKey(Hkey, strPath, hCurKey)
    lRegResult = RegQueryValueEx(hCurKey, strValueName, 0&, lValueType, ByVal 0&, lDataBufferSize)
    
    If lRegResult = ERROR_SUCCESS Then
    
      If lValueType = REG_BINARY Then
      
        ' initialise buffers and retrieve value
        ReDim byBuffer(lDataBufferSize - 1) As Byte
        lRegResult = RegQueryValueEx(hCurKey, strValueName, 0&, lValueType, byBuffer(0), lDataBufferSize)
        
        GetSettingByte = byBuffer
    
      End If
    
    Else
      'there is a problem
    End If
    
    lRegResult = RegCloseKey(hCurKey)

End Function


