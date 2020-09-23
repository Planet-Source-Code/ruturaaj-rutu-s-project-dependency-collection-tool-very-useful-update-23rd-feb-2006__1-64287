Attribute VB_Name = "mGeneralFunctions"
Option Explicit
Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOMOVE = &H2
Private Const SWP_SHOWWINDOW = &H40
Private Const SW_SHOWNORMAL = 1

Private Type SHELLEXECUTEINFO
    cbSize As Long
    fMask As Long
    hwnd As Long
    lpVerb As String
    lpFile As String
    lpParameters As String
    lpDirectory As String
    nShow As Long
    hInstApp As Long
    lpIDList As Long 'Optional
    lpClass As String 'Optional
    hkeyClass As Long 'Optional
    dwHotKey As Long 'Optional
    hIcon As Long 'Optional
    hProcess As Long 'Optional
    End Type
    Private Const SEE_MASK_INVOKEIDLIST = &HC
    Private Const SEE_MASK_NOCLOSEPROCESS = &H40
    Private Const SEE_MASK_FLAG_NO_UI = &H400

Private Declare Function ShellExecuteEx Lib "shell32" Alias "ShellExecuteExA" (SEI As SHELLEXECUTEINFO) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Sub SetWindowPos Lib "User32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)

Public Function OpenURL(ByVal sURL As String, ByVal lngParentHWND As String)
    ShellExecute lngParentHWND, vbNullString, sURL, vbNullString, "C:\", SW_SHOWNORMAL
End Function

Public Function SetAlwaysOnTopMode(ByVal H_Wnd As Long, Optional ByVal OnTop As Boolean = True)
    ' get the hWnd of the form to be move on top
    SetWindowPos H_Wnd, IIf(OnTop, HWND_TOPMOST, HWND_NOTOPMOST), 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_SHOWWINDOW
End Function

Public Function ApplicationBuild(Optional bIncludeRevision As Boolean = True) As String
    Dim sVersion As String
    
    On Error GoTo ApplicationBuild_Error

    sVersion = App.Major & "." & App.Minor
    
    If bIncludeRevision Then
        sVersion = sVersion & App.Revision
    End If
    
    ApplicationBuild = sVersion

    'This will avoid empty error window to appear.
    Exit Function

ApplicationBuild_Error:

    'Show the Error Message with Error Number and its Description.
    MsgBox Err.Number & " : " & vbCrLf & vbCrLf & Err.Description, vbInformation, "Error ! (Source : ApplicationBuild)"

    'Safe Exit from ApplicationBuild
    Exit Function

End Function
Public Function IsStringEmpty(ByVal sInput As String) As Boolean
    
'---------------------------------------------------------------------------------------
' Procedure : IsStringEmpty
' DateTime  : 10 February 2004 07:21
' Author    : Ruturaj
'
' CopyRight : This code is issued under GNU-GPL. You are free to use it provided _
'             you keep this head information intact. You may use this code for Personal as _
'             well as Commertial purpose. Visit our website to get more useful code for your _
'             projects. Also feel free share your valuable code with us ...
'
'
' Purpose   : This functions checks if the argument string is Empty string or not and
'             accordingly returns TRUE or FALSE.
'
'---------------------------------------------------------------------------------------

    Dim sTest As String
    
    sTest = Trim(sInput)
    If Len(sTest) = 0 Then
        IsStringEmpty = True
    Else
        IsStringEmpty = False
    End If
End Function
Public Function ReadFile(ByVal sFilename As String) As String
    On Error GoTo ReadFile_Error
    
    Dim sTemp As String
    Dim F As Long

    F = FreeFile
    sTemp = ""
    Open sFilename For Binary As #F        ' Open file.(can be text or image)
    sTemp = Input(FileLen(sFilename), #F) ' Get entire Files data
    Close #F
    ReadFile = sTemp
    
'This will avoid empty error window to appear.
     Exit Function
    
ReadFile_Error:
    
'Show the Error Message with Error Number and its Description.
     MsgBox Err.Number & " : " & vbCrLf & vbCrLf & Err.Description, vbInformation, "Error !"
    
'Safe Exit from ReadFile
     Exit Function

End Function
Public Function GetFileFromRes(ByVal vResID As Variant, ByVal vResSection As Variant, ByVal sSaveAs As String) As Boolean

'---------------------------------------------------------------------------------------
' Author     : iNova Creations
' Website    : www.inovacreations.com
' Email      : support@inovacreations.com
'
' Procedure  : GetFileFromRes
' Type       : Function
' ReturnType : Boolean
'
' Arguments  : [1] vResID     : The Resource ID at which File Data is stored.
'              [2] vResSection  : The Resource File Section Name under which
'                                 File data is saved.
'              [3] sSaveAs      : After extracting the File , where to save it ?
'
' Purpose    : This Function is written to extract the File from Resource File and then
'              save to specified location on User's Hard Drive.
'
'---------------------------------------------------------------------------------------

    On Error GoTo GetFileFromRes_Error

    Dim bytFileData() As Byte
    
    bytFileData = LoadResData(vResID, vResSection)
    
    Open sSaveAs For Binary Access Write As #1
        Put #1, , bytFileData
    Close #1
    
    If Dir(sSaveAs, vbHidden Or vbNormal Or vbReadOnly Or vbSystem) <> "" Then
        GetFileFromRes = True
    Else
        GetFileFromRes = False
    End If

    
'This will avoid empty error window to appear.
    Exit Function

GetFileFromRes_Error:

'Show the Error Message with Error Number and its Description.
    MsgBox Err.Number & " : " & vbCrLf & vbCrLf & Err.Description, vbInformation, "Error ! "

'Safe Exit from GetFileFromRes
    Exit Function

        
End Function


Public Function FileExists(ByVal sPath As String) As Boolean

'---------------------------------------------------------------------------------------
' Author     : iNova Creations
' Website    : www.inovacreations.com
' Email      : support@inovacreations.com
'
' Procedure  : FileExists
' Type       : Function
' ReturnType : Boolean
'
' Arguments  : [1] sPath : Complete Path including File Name.
'
' Purpose    : This Function checks whether the specified File exists at
'              given Path Location.
'
'---------------------------------------------------------------------------------------

    On Error GoTo FileExists_Error
    
    If Len(Trim$(sPath)) = 0 Then FileExists = False: Exit Function
    
    If Dir(sPath, vbHidden Or vbNormal Or vbReadOnly Or vbSystem) <> "" Then FileExists = True _
    Else FileExists = False

'This will avoid empty error window to appear.
    Exit Function

FileExists_Error:

'Show the Error Message with Error Number and its Description.
    MsgBox Err.Number & " : " & vbCrLf & vbCrLf & Err.Description, vbInformation, "Error ! (Source : FileExists)"

'Safe Exit from FileExists
    Exit Function

End Function
Public Function FolderExists(ByVal sPath As String) As Boolean
    On Error GoTo FolderExists_Error

    If Dir(sPath, vbDirectory) <> "" Then
        FolderExists = True
    Else
        FolderExists = False
    End If

'This will avoid empty error window to appear.
    Exit Function

FolderExists_Error:

'Show the Error Message with Error Number and its Description.
    MsgBox Err.Number & " : " & vbCrLf & vbCrLf & Err.Description, vbInformation, "Error !"

'Safe Exit from FolderExists
    Exit Function

End Function
Public Function GetParentFolderPath(ByVal sFilePath As String, ByVal bAppendSlash As Boolean) As String

'---------------------------------------------------------------------------------------
' Author     : iNova Creations
' Website    : www.inovacreations.com
' Email      : support@inovacreations.com
'
' Procedure  : GetParentFolderPath
' Type       : Function
' ReturnType : String
' Arguments  : [1] sFilePath    : Complete Path including File Name.
'              [2] bAppendSlash : If set to TRUE , this will add "\" to the
'                  end of return string.
'
' Purpose    : This method will return the complete Folder Path in which the File is.
'              In other words , it returns the Parent Folder Path for a given complete
'              Path of File. Optionally , it can return with "\" character at the end.
'
'---------------------------------------------------------------------------------------

    On Error GoTo GetParentFolderPath_Error

    Dim iPos As Integer
    Dim sFolder As String
    
    iPos = InStrRev(sFilePath, "\")
    If iPos > 0 Then
        sFolder = Mid$(sFilePath, 1, iPos - 1)
    Else
        sFolder = ""
    End If
    
    If bAppendSlash = True And Not (IsStringEmpty(sFolder)) Then
        sFolder = sFolder & "\"
    End If
    
    GetParentFolderPath = sFolder

'This will avoid empty error window to appear.
    Exit Function

GetParentFolderPath_Error:

'Show the Error Message with Error Number and its Description.
    MsgBox Err.Number & " : " & vbCrLf & vbCrLf & Err.Description, vbInformation, "Error ! (Source : GetParentFolderPath)"

'Safe Exit from GetParentFolderPath
    Exit Function

End Function


Public Function FileNameFromPath(ByVal sPath As String, Optional ByVal bWithExt As Boolean = True) As String

'---------------------------------------------------------------------------------------
' Procedure : FileNameFromPath
' DateTime  : 10 February 2004 07:22
' Author    : Ruturaj
'
' CopyRight : This code is issued under GNU-GPL. You are free to use it provided _
'             you keep this head information intact. You may use this code for Personal as _
'             well as Commertial purpose. Visit our website to get more useful code for your _
'             projects. Also feel free share your valuable code with us ...
'
'
' Purpose   : This function extracts File Name from given complete path. It can return the
'             File Name with or without File Extension.
'
'---------------------------------------------------------------------------------------

    On Error GoTo FileNameFromPath_Error
    
    Dim iPos1 As Integer
    Dim iPos2 As Integer
    Dim sFilename As String
    
    iPos1 = InStrRev(sPath, "\")
    iPos2 = InStrRev(sPath, ".")
    
    If iPos1 <> 0 And iPos2 <> 0 Then
        If bWithExt = True Then
            sFilename = Mid$(sPath, iPos1 + 1, Len(sPath))
        ElseIf bWithExt = False Then
            sFilename = Mid$(sPath, iPos1 + 1, iPos2 - iPos1 - 1)
        End If
        
        FileNameFromPath = sFilename
    Else
        FileNameFromPath = ""
    End If
    
'This will avoid empty error window to appear.
    Exit Function

FileNameFromPath_Error:

'Show the Error Message with Error Number and its Description.
    MsgBox Err.Number & " : " & vbCrLf & vbCrLf & Err.Description, vbInformation, "Error !"

'Safe Exit from FileNameFromPath
    Exit Function

End Function
Public Sub ShowPropertiesDlg(sFilename As String, lHwnd As Long)
    Dim tpShellExec As SHELLEXECUTEINFO
    
    On Error GoTo ShowPropertiesDlg_Error

    With tpShellExec
        .cbSize = Len(tpShellExec)
        .fMask = SEE_MASK_NOCLOSEPROCESS Or _
        SEE_MASK_INVOKEIDLIST Or _
        SEE_MASK_FLAG_NO_UI
        .hwnd = lHwnd
        .lpVerb = "properties"
        .lpFile = sFilename
        .lpParameters = vbNullChar
        .lpDirectory = vbNullChar
        .nShow = 0
        .hInstApp = 0
        .lpIDList = 0
    End With

    
    Call ShellExecuteEx(tpShellExec)

    'This will avoid empty error window to appear.
    Exit Sub

ShowPropertiesDlg_Error:

    'Show the Error Message with Error Number and its Description.
    MsgBox Err.Number & " : " & vbCrLf & vbCrLf & Err.Description, vbInformation, "Error::ShowPropertiesDlg)"

    'Safe Exit from ShowPropertiesDlg
    Exit Sub

    
End Sub



