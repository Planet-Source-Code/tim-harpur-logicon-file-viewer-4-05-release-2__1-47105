Attribute VB_Name = "modRegistry"
Option Explicit

Private Const REG_SZ As Long = 1                   ' string data
Private Const REG_EXPAND_SZ = 2                  ' Unicode nul terminated string
Private Const REG_BINARY = 3                     ' Free form binary
Private Const REG_DWORD As Long = 4         ' long data
Private Const REG_DWORD_LITTLE_ENDIAN = 4        ' 32-bit number (same as REG_DWORD)
Private Const REG_DWORD_BIG_ENDIAN = 5           ' 32-bit number
Private Const REG_LINK = 6                       ' Symbolic Link (unicode)
Private Const REG_MULTI_SZ = 7                   ' Multiple Unicode strings
Private Const REG_RESOURCE_LIST = 8              ' Resource list in the resource map
Private Const REG_FULL_RESOURCE_DESCRIPTOR = 9   ' Resource list in the hardware description
Private Const REG_RESOURCE_REQUIREMENTS_LIST = 10

Public Enum REGKEY_ROOT
  HKEY_CLASSES_ROOT = &H80000000
  HKEY_CURRENT_USER = &H80000001
  HKEY_LOCAL_MACHINE = &H80000002
  HKEY_USERS = &H80000003
  HKEY_CURRENT_CONFIG = &H80000005
  HKEY_DYN_DATA = &H80000006
End Enum

Private Type FILETIME
  dwLowDateTime As Long
  dwHighDateTime As Long
End Type

Private Const KEY_ALL_ACCESS = &H3F
Private Const REG_OPTION_NON_VOLATILE = 0       ' Key is preserved when system is rebooted

Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long

Private Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, lpSecurityAttributes As Long, phkResult As Long, lpdwDisposition As Long) As Long
Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long

Private Declare Function RegQueryInfoKey Lib "advapi32.dll" Alias "RegQueryInfoKeyA" (ByVal hKey As Long, ByVal lpClass As Long, lpcbClass As Long, ByVal lpReserved As Long, lpcSubKeys As Long, lpcbMaxSubKeyLen As Long, lpcbMaxClassLen As Long, lpcValues As Long, lpcbMaxValueNameLen As Long, lpcbMaxValueLen As Long, lpcbSecurityDescriptor As Long, lpftLastWriteTime As FILETIME) As Long
Private Declare Function RegQueryValueExString Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long         ' Note that if you declare the lpData parameter as String, you must pass it By Value.
Private Declare Function RegQueryValueExBinary Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long                     ' Note that if you declare the lpData parameter as any (for an array), you must pass the first value of the array
Private Declare Function RegQueryValueExLong Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Long, lpcbData As Long) As Long
Private Declare Function RegQueryValueExNull Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As Long, lpcbData As Long) As Long         ' Note that if you declare the lpData parameter as NULL, you must pass it By Value.
Private Declare Function RegSetValueExString Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpData As String, ByVal cbData As Long) As Long         ' Note that if you declare the lpData parameter as String, you must pass it By Value.
Private Declare Function RegSetValueExLong Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Long, ByVal cbData As Long) As Long
Private Declare Function RegSetValueExBinary Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long            ' Note that if you declare the lpData parameter as any (for an array), you must pass the first value of the array
Private Declare Function RegEnumKeyEx Lib "advapi32.dll" Alias "RegEnumKeyExA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, lpcbName As Long, ByVal lpReserved As Long, ByVal lpClass As Long, lpcbClass As Long, lpftLastWriteTime As FILETIME) As Long
Private Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, lpcbValueName As Long, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long

Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long

Public Function Create_RegistryKey(ByVal RegistrySection As REGKEY_ROOT, ByVal KeyName As String) As Long
  Dim hKey As Long, rValue As Long
  
  On Error GoTo badKey
  
  If RegCreateKeyEx(RegistrySection, KeyName, 0&, vbNullString, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, 0&, hKey, rValue) = 0 Then
    RegCloseKey hKey
  Else
    Create_RegistryKey = -1
  End If
  
  Exit Function
  
badKey:
  Create_RegistryKey = -1
End Function

Public Sub Delete_RegistryKey(ByVal RegistrySection As REGKEY_ROOT, ByVal ParentKeyName As String, ByVal SubKeyName As String)
  Dim hKey As Long
  
  On Error GoTo badKey
  
  If RegOpenKeyEx(RegistrySection, ParentKeyName, 0, KEY_ALL_ACCESS, hKey) = 0 Then
    RegDeleteKey hKey, SubKeyName
    
    RegCloseKey hKey
  End If
  
badKey:
End Sub

Public Sub Delete_RegistryValue(ByVal RegistrySection As REGKEY_ROOT, ByVal KeyName As String, ByVal ValueName As String)
  Dim hKey As Long
  
  On Error GoTo badKey
  
  If RegOpenKeyEx(RegistrySection, KeyName, 0, KEY_ALL_ACCESS, hKey) = 0 Then
    RegDeleteValue hKey, ValueName
    
    RegCloseKey hKey
  End If
  
badKey:
End Sub

Public Function Get_RegistryLongValue(ByVal RegistrySection As REGKEY_ROOT, ByVal KeyName As String, ByVal ValueName As String) As Long
  Dim hKey As Long, lValue As Long, vSize As Long, vType As Long
  
  On Error GoTo badKey
  
  If RegOpenKeyEx(RegistrySection, KeyName, 0, KEY_ALL_ACCESS, hKey) = 0 Then
    If RegQueryValueExNull(hKey, ValueName, 0&, vType, 0&, vSize) = 0 Then
      If vType = REG_DWORD Then
        If RegQueryValueExLong(hKey, ValueName, 0&, vType, lValue, vSize) = 0 Then Get_RegistryLongValue = lValue
      End If
    End If
    
    RegCloseKey hKey
  End If
  
badKey:
End Function

Public Sub Set_RegistryLongValue(ByVal RegistrySection As REGKEY_ROOT, ByVal KeyName As String, ByVal ValueName As String, ByVal Value As Long)
  Dim hKey As Long
  
  On Error GoTo badKey
  
  If RegOpenKeyEx(RegistrySection, KeyName, 0, KEY_ALL_ACCESS, hKey) = 0 Then
    RegSetValueExLong hKey, ValueName, 0&, REG_DWORD, Value, 4
    
    RegCloseKey hKey
  End If
  
badKey:
End Sub

'specifying a ValueName of "" accesses the default string value for that key
Public Function Get_RegistryStringValue(ByVal RegistrySection As REGKEY_ROOT, ByVal KeyName As String, ByVal ValueName As String) As String
  Dim hKey As Long, sValue As String, vSize As Long, vType As Long
  
  On Error GoTo badKey
  
  If RegOpenKeyEx(RegistrySection, KeyName, 0, KEY_ALL_ACCESS, hKey) = 0 Then
    If RegQueryValueExNull(hKey, ValueName, 0&, vType, 0&, vSize) = 0 Then
      If vType = REG_SZ Then
        sValue = String(vSize, 0)
        
        If RegQueryValueExString(hKey, ValueName, 0&, vType, sValue, vSize) = 0 Then Get_RegistryStringValue = Left$(sValue, vSize - 1)
      End If
    End If
    
    RegCloseKey hKey
  End If
  
badKey:
End Function

'specifying a ValueName of "" accesses the default string value for that key
Public Sub Set_RegistryStringValue(ByVal RegistrySection As REGKEY_ROOT, ByVal KeyName As String, ByVal ValueName As String, ByVal Value As String)
  Dim hKey As Long, sValue As String
  
  On Error GoTo badKey
  
  If RegOpenKeyEx(RegistrySection, KeyName, 0, KEY_ALL_ACCESS, hKey) = 0 Then
    sValue = Value & Chr$(0)
    
    RegSetValueExString hKey, ValueName, 0&, REG_SZ, sValue, Len(sValue)
    
    RegCloseKey hKey
  End If
  
badKey:
End Sub

Public Function Get_RegistryBinaryValue(ByVal RegistrySection As REGKEY_ROOT, ByVal KeyName As String, ByVal ValueName As String, Value() As Byte) As Long
  Dim hKey As Long, vSize As Long, vType As Long
  
  On Error GoTo badKey
  
  If RegOpenKeyEx(RegistrySection, KeyName, 0, KEY_ALL_ACCESS, hKey) = 0 Then
    If RegQueryValueExNull(hKey, ValueName, 0&, vType, 0&, vSize) = 0 Then
      If vType = REG_BINARY Then
        ReDim Value(1 To vSize)
        
        If RegQueryValueExBinary(hKey, ValueName, 0&, vType, Value(1), vSize) = 0 Then Get_RegistryBinaryValue = vSize
      End If
    End If
    
    RegCloseKey hKey
  End If
  
badKey:
End Function

Public Sub Set_RegistryBinaryValue(ByVal RegistrySection As REGKEY_ROOT, ByVal KeyName As String, ByVal ValueName As String, Value() As Byte)
  Dim hKey As Long, sValue As String, lBase As Long, uBase As Long, aLength As Long
  
  On Error GoTo badKey
  
  lBase = LBound(Value)
  uBase = UBound(Value)
  
  aLength = uBase - lBase + 1
  
  If aLength > 0 Then
    If RegOpenKeyEx(RegistrySection, KeyName, 0, KEY_ALL_ACCESS, hKey) = 0 Then
      RegSetValueExBinary hKey, ValueName, 0&, REG_BINARY, Value(lBase), aLength
      
      RegCloseKey hKey
    End If
  End If
  
badKey:
End Sub

Public Function Get_RegistrySection(ByVal RegistrySection As REGKEY_ROOT, ByVal KeyName As String, SubKeyList() As String, SubKeyCount As Long, _
            StringValueList() As String, StringValueCount As Long, LongValueList() As String, LongValueCount As Long, BinaryValueList() As String, BinaryValueCount As Long) As Long
  
  Dim hKey As Long, maxSubKeyLength As Long, maxValues As Long, maxValueLength As Long
  Dim loop1 As Long, stringBuffer As String, validLength As Long, fileT As FILETIME, valueType As Long
  Dim dudBuffer() As Byte, dudLength As Long
  
  On Error GoTo badKey
  
  If RegOpenKeyEx(RegistrySection, KeyName, 0, KEY_ALL_ACCESS, hKey) = 0 Then
    If RegQueryInfoKey(hKey, 0&, 0&, 0&, SubKeyCount, maxSubKeyLength, 0&, maxValues, maxValueLength, 0&, 0&, fileT) = 0 Then
      If SubKeyCount > 0 Then
        ReDim SubKeyList(1 To SubKeyCount)
        
        For loop1 = 1 To SubKeyCount
          validLength = maxSubKeyLength + 1
          stringBuffer = String(validLength, 0)
          
          If RegEnumKeyEx(hKey, loop1 - 1, stringBuffer, validLength, 0&, 0&, 0&, fileT) = 0 Then SubKeyList(loop1) = Left$(stringBuffer, validLength)
        Next loop1
      End If
      
      For loop1 = 1 To maxValues
        validLength = maxValueLength + 1
        stringBuffer = String(validLength, 0)
        
        dudLength = 2000
        ReDim dudBuffer(1 To 2000)
        
        If RegEnumValue(hKey, loop1 - 1, stringBuffer, validLength, 0&, valueType, dudBuffer(1), dudLength) = 0 Then
          Select Case valueType
            Case REG_SZ
              StringValueCount = StringValueCount + 1
              ReDim Preserve StringValueList(1 To StringValueCount)
              
              StringValueList(StringValueCount) = Left$(stringBuffer, validLength)
            Case REG_BINARY
              BinaryValueCount = BinaryValueCount + 1
              ReDim Preserve BinaryValueList(1 To BinaryValueCount)
        
              BinaryValueList(BinaryValueCount) = Left$(stringBuffer, validLength)
            Case REG_DWORD
              LongValueCount = LongValueCount + 1
              ReDim Preserve LongValueList(1 To LongValueCount)
        
              LongValueList(LongValueCount) = Left$(stringBuffer, validLength)
          End Select
        End If
      Next loop1
    End If
  End If
  
badKey:
End Function

