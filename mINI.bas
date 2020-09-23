Attribute VB_Name = "mINI"
'---------------------------------------------------------------------------------------
' Module    : modINIFunctions
' Author    : Ruturaj
'
' Purpose   : This Module contains methods to manipulate INI files.
'
'---------------------------------------------------------------------------------------
Option Explicit


Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

Dim m_File As String, m_Buffer As Long

Public Sub INISetup(Filename As String, BufferSize As Long)

'---------------------------------------------------------------------------------------
' Procedure : INISetup
' Author    : Ruturaj
'
'
' Purpose   :
'---------------------------------------------------------------------------------------

    m_Buffer = BufferSize
    m_File = Filename
End Sub

Public Function Read_Ini(iSection As String, iKeyName As String, Optional iDefault As String)

'---------------------------------------------------------------------------------------
' Procedure : Read_Ini
' Author    : Ruturaj
'
'
' Purpose   : This Function Reads the content of Spcified Key from Specified Section.
'
'---------------------------------------------------------------------------------------

    If m_Buffer = 0 Then
        Err.Raise "670", "INIReadWrite", "Buffer size not set"
    ElseIf m_File = "" Then
        Err.Raise "670", "INIReadWrite", "Filename not set"
    End If
    Dim Ret As String, NC As Long
    
    'Create the buffer
    Ret = String(m_Buffer, 0)
    
    'Retrieve the string
    NC = GetPrivateProfileString(iSection, iKeyName, iDefault, Ret, m_Buffer, m_File)
    
    'NC is the number of characters copied to the buffer
    If NC <> 0 Then
        Ret = Left$(Ret, NC)
    Else
        'Make sure to cut it down to number of char's returned
        Ret = ""
    End If
    
    'Turn the funky vbcrlf string into VBCRLFs
    Ret = Replace(Ret, "%%&&Chr(13)&&%%", vbCrLf)
    
    'Return the setting
    Read_Ini = Ret
End Function

Public Sub Write_Ini(iSection As String, iKeyName As String, iValue As Variant)

'---------------------------------------------------------------------------------------
' Procedure : Write_Ini
' Author    : Ruturaj
'
'
' Purpose   : This Sub Writes the content to specified Key in specified Section.
'
'---------------------------------------------------------------------------------------

    If m_Buffer = 0 Then
        Err.Raise "670", "INIReadWrite", "Buffer size not set"
    ElseIf m_File = "" Then
        Err.Raise "670", "INIReadWrite", "Filename not set"
    End If
    
    'Make sure to change it to a String
    iValue = CStr(iValue)
    
    'Turn all vbcrlf's into that funky string
    iValue = Replace(iValue, vbCrLf, "%%&&Chr(13)&&%%")
    WritePrivateProfileString iSection, iKeyName, CStr(iValue), m_File
End Sub
