Attribute VB_Name = "Registry"
'KPD-Team 2000
'URL: http://www.allapi.net
'E-Mail: KPDTeam@Allapi.com

Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_CURRENT_CONFIG = &H80000005
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_DYN_DATA = &H80000006
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_PERFORMANCE_DATA = &H80000004
Public Const HKEY_USERS = &H80000003
Public Const ERROR_SUCCESS = 0&

' Registry API prototypes
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
Public Const REG_SZ = 1                         ' Unicode nul terminated string
Public Const REG_DWORD = 4                      ' 32-bit number
Public Const REG_BINARY = 3                     ' Free form binary
Public Sub SaveKey(hKey As Long, strPath As String)
    Dim Keyhand&
    r = RegCreateKey(hKey, strPath, Keyhand&)
    r = RegCloseKey(Keyhand&)
End Sub
Public Function GetString(hKey As Long, strPath As String, strValue As String)
    Dim Keyhand As Long, lResult As Long
    Dim strBuf As String, lDataBufSize As Long, intZeroPos As Integer
    r = RegOpenKey(hKey, strPath, Keyhand)
    lResult = RegQueryValueEx(Keyhand, strValue, 0&, lValueType, ByVal 0&, lDataBufSize)
    If lValueType = REG_SZ Then
        strBuf = String(lDataBufSize, " ")
        lResult = RegQueryValueEx(Keyhand, strValue, 0&, 0&, ByVal strBuf, lDataBufSize)
        If lResult = ERROR_SUCCESS Then
            intZeroPos = InStr(strBuf, Chr$(0))
            If intZeroPos > 0 Then
                GetString = Left$(strBuf, intZeroPos - 1)
            Else
                GetString = strBuf
            End If
        End If
    End If
End Function
Public Sub SaveString(hKey As Long, strPath As String, strValue As String, strdata As String)
    Dim Keyhand As Long, r As Long
    r = RegCreateKey(hKey, strPath, Keyhand)
    r = RegSetValueEx(Keyhand, strValue, 0, REG_SZ, ByVal strdata, Len(strdata))
    r = RegCloseKey(Keyhand)
End Sub

Public Function DeleteKey(ByVal hKey As Long, ByVal strKey As String)
    Dim r As Long
    r = RegDeleteKey(hKey, strKey)
End Function
Public Function DeleteValue(ByVal hKey As Long, ByVal strPath As String, ByVal strValue As String)
    Dim Keyhand As Long
    r = RegOpenKey(hKey, strPath, Keyhand)
    r = RegDeleteValue(Keyhand, strValue)
    r = RegCloseKey(Keyhand)
End Function

Public Function StripTerminator(sInput As String) As String
    Dim ZeroPos As Integer
    ZeroPos = InStr(1, sInput, vbNullChar)
    If ZeroPos > 0 Then
        StripTerminator = Left$(sInput, ZeroPos - 1)
    Else
        StripTerminator = sInput
    End If
End Function




Public Function KeyExists(hKey As Long, strPath As String) As Boolean
    Dim lngKeyHandle As Long
    
    If RegOpenKeyEx(hKey, strPath, 0, 1, lngKeyHandle) = ERROR_SUCCESS Then
        KeyExists = True
        RegCloseKey lngKeyHandle
    Else
        KeyExists = False
    End If
    
End Function
