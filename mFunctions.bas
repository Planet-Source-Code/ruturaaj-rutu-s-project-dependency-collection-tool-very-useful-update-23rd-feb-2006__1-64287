Attribute VB_Name = "mFunctions"
Option Explicit

'HKEY_CLASSES_ROOT\VisualBasic.Project
Private Const Scan_Key = "\shell\RPDCT_Scan"
Private Const Collect_Key = "\shell\RPDCT_Collect"
Private Const ScanCommand_Key = "\shell\RPDCT_Scan\command"
Private Const CollectCommand_Key = "\shell\RPDCT_Collect\command"
Private Const Scan_MenuCaption = "Scan for Dependencies"
Private Const Collect_MenuCaption = "Collect Depedencies"

Private Enum RefType_Const
    [DLL]
    [OCX]
    [NONE]
End Enum
Private Function ImageIndexByType(sFileNamePath As String) As Integer
    
    If Right(LCase(sFileNamePath), 3) = "tlb" Then
        ImageIndexByType = 3
    ElseIf Right(LCase(sFileNamePath), 3) = "ocx" Then
        ImageIndexByType = 2
    ElseIf Right(LCase(sFileNamePath), 3) = "dll" Then
        ImageIndexByType = 1
    End If
    
End Function

Private Function ReferenceType(sVBPLine As String) As RefType_Const
    Dim sWord As String
    
    On Error GoTo ReferenceType_Error

    If InStr(sVBPLine, "=") Then sWord = Mid$(sVBPLine, 1, InStr(sVBPLine, "=") - 1)
    
    If sWord = "Reference" Then
        ReferenceType = DLL
    ElseIf sWord = "Object" And Right(LCase(sVBPLine), 3) = "ocx" Then
        ReferenceType = OCX
    Else
        ReferenceType = NONE
    End If

    'This will avoid empty error window to appear.
    Exit Function

ReferenceType_Error:

    'Show the Error Message with Error Number and its Description.
    MsgBox Err.Number & " : " & vbCrLf & vbCrLf & Err.Description, vbInformation, "Error::ReferenceType)"

    'Safe Exit from ReferenceType
    Exit Function

End Function
Public Function AddToList(sFilePathName As String)
    Dim iCnt As Integer
    Dim oListItem As ListItem

    For iCnt = 1 To frmMain.lstDep.ListItems.Count
        If frmMain.lstDep.ListItems(iCnt).SubItems(4) = sFilePathName Then
            MsgBox "Selected component is already in List; so need not to add it again to Dependency List.", vbInformation, "Component already Listed!"
            Exit Function
        End If
    Next iCnt
        
    Set oListItem = frmMain.lstDep.ListItems.Add(, , FileNameFromPath(sFilePathName, False) & " " & UCase(Right(sFilePathName, 3)), ImageIndexByType(sFilePathName), ImageIndexByType(sFilePathName))
    With oListItem
        .SubItems(1) = Right(UCase(sFilePathName), 3)
        .SubItems(2) = Round(FileLen(sFilePathName) / 1024, 0) & " KB"
        If FileExists(sFilePathName) Then
            .SubItems(3) = "File Exists."
        Else
            .SubItems(3) = "File Missing!"
            .Bold = True
            .ForeColor = vbRed
        End If
        .SubItems(4) = sFilePathName
    End With
    
            
End Function
Public Function ProcessVBP(sFilePathName As String, Optional bCopyFiles As Boolean = False, Optional sDestFolder As String = "")
    Dim sVBPData As String, sVBPLines() As String
    Dim iCnt As Integer
    
    Debug.Print sFilePathName
    sVBPData = ReadFile(sFilePathName)
    sVBPLines = Split(sVBPData, vbCrLf)
    
    'This code will be executed when a command line param is executed. We need to first check if required folder is there or otherwise.
    'No call to frmMain as it will cause the Interface to show.
    
    If bCopyFiles = True Then
        Debug.Print sDestFolder
        If FolderExists(sDestFolder) = False Then
            INISetup App.Path & "\AppCfg.cfg", 5000
            If CBool(Read_Ini("options", "notify", "0")) = False Then
                If MsgBox("Specified Folder to copy Dependency Files does not exists. Would you like to create a New folder at specified Location?", vbYesNo + vbQuestion, "Specified Folder does not exists. Create New?") = vbNo Then
                    Exit Function
                Else
                    MkDir sDestFolder
                End If
            Else
                MkDir sDestFolder
            End If
        End If
        
        If Right(sDestFolder, 1) <> "\" Then
            sDestFolder = sDestFolder & "\"
        End If
    Else
        'Prepare Listview ...
        frmMain.lstDep.ListItems.Clear
        frmMain.lblSize.Caption = "0"
    End If
    
    For iCnt = LBound(sVBPLines) To UBound(sVBPLines)
        If IsStringEmpty(sVBPLines(iCnt)) = False Then
            Select Case ReferenceType(sVBPLines(iCnt))
    
                Case [DLL]
                        
                        ParseRefLine sVBPLines(iCnt), bCopyFiles, sDestFolder
                        
                Case [OCX]
                        
                        ParseObjectLine sVBPLines(iCnt), bCopyFiles, sDestFolder
                        
                Case [NONE]
    
                        'Do nothing
                        
            End Select
        End If
    Next iCnt
    
    If bCopyFiles Then
        INISetup App.Path & "\AppCfg.cfg", 5000
        If CInt(Read_Ini("options", "batch", "1")) = 1 Then
            GetFileFromRes "bat", "files", sDestFolder & "AutoRename.bat"
        End If
        
        If CInt(Read_Ini("options", "regsvr", "0")) Then
            GetFileFromRes "regsvr32", "files", sDestFolder & "regsvr32.exe"
        End If
    End If
    

End Function

Private Function ParseRefLine(sVBPLine As String, Optional bCopyFiles As Boolean, Optional sDestFolder As String = "")
    
    Dim sRefLine() As String
    Dim sPath As String, sTypeLib As String, sDescription As String
    Dim sRegKeyForPath As String
    Dim oListItem As ListItem
    
    'On Error Resume Next
    
    sRefLine = Split(sVBPLine, "#")
    
    'Get TypeLib ...
    sTypeLib = Mid$(sRefLine(0), InStr(sRefLine(0), "{"), Len(sRefLine(0)))
    
    'RegKey Path ...
    sRegKeyForPath = sTypeLib & "\" & sRefLine(1) & "\" & sRefLine(2) & "\win32"
    
    'Check if DLL is registered ...
    If KeyExists(HKEY_CLASSES_ROOT, "TypeLib\" & sTypeLib) Then
        sDescription = GetString(HKEY_CLASSES_ROOT, "TypeLib\" & sTypeLib & "\" & sRefLine(1), "")
        sPath = GetString(HKEY_CLASSES_ROOT, "TypeLib\" & sRegKeyForPath, "")
    Else
        sDescription = sRefLine(UBound(sRefLine))
        sPath = "- NA -"
    End If
    
    'There might be some special cases like the one I noticed with vbscript.dll file.
    'When this DLL is referenced, Visual Basic records its reference as vbscript.dll\3 and I really don't know why!! :)
    'So, this forced me to write some extra condition check to handle these sort of situations as well.
    'If still something manages to escape and generate error, then please let me know on mailme_friends@yahoo.com
    'Thanks.
    If sPath <> "- NA -" And InStr(Mid$(sPath, InStrRev(sPath, "\") + 1, Len(sPath)), ".") = 0 Then
        sPath = GetParentFolderPath(sPath, False)
    End If
    
    'See if directly to copy files ...
    If bCopyFiles Then
        If sPath <> "- NA -" Then
            INISetup App.Path & "\AppCfg.cfg", 5000
            If CInt(Read_Ini("options", "rename", "1")) = 1 Then
                FileCopy sPath, sDestFolder & FileNameFromPath(sPath, False) & "._" & Right(LCase(sPath), 2)
            Else
                FileCopy sPath, sDestFolder & FileNameFromPath(sPath)
            End If
        End If
        
'        If CInt(Read_Ini("options", "batch", "1")) = 1 Then
'            GetFileFromRes "bat", "files", sDestFolder & "AutoRename.bat"
'        End If
'
'        If CInt(Read_Ini("options", "regsvr", "0")) Then
'            GetFileFromRes "regsvr32", "files", sDestFolder & "regsvr32.exe"
'        End If
    
    Else
        'No Direct copy, then Display List on frmMain ...
        With frmMain.lstDep.ListItems
            Set oListItem = .Add(, , sRefLine(4), ImageIndexByType(sPath), ImageIndexByType(sPath))
                
                If sPath = "- NA -" Then
                    oListItem.SubItems(2) = "- NA -"
                    oListItem.SubItems(3) = "Unregistered!"
                    oListItem.Bold = True
                    oListItem.ForeColor = vbRed
                ElseIf FileExists(sPath) Then
                    oListItem.SubItems(2) = Round(FileLen(sPath) / 1024, 0) & " KB"
                    oListItem.SubItems(3) = "File Exists."
                Else
                    oListItem.SubItems(2) = "- NA -"
                    oListItem.SubItems(3) = "File Missing!"
                    oListItem.Bold = True
                    oListItem.ForeColor = vbRed
                End If
    
                If sPath <> "- NA -" Then
                    oListItem.SubItems(1) = Right(UCase(sPath), 3)
                Else
                    oListItem.SubItems(1) = Right(UCase(sRefLine(3)), 3)
                End If
                oListItem.SubItems(4) = sPath
                
            Set oListItem = Nothing
        End With
    End If
    
End Function
Private Function ParseObjectLine(sVBPLine As String, Optional bCopyFiles As Boolean = False, Optional sDestFolder As String = "")
    Dim sObjLine() As String, sObjElements() As String
    Dim sDescrption As String, sPath As String
    Dim sTypeLib As String, sRegKeyForPath As String
    Dim oListItem As ListItem
    
    sObjLine = Split(sVBPLine, "#")
    sObjElements = Split(sVBPLine, ";")
    
    sTypeLib = Mid$(sObjLine(0), InStr(sObjLine(0), "=") + 1, Len(sObjLine(0)))
    sRegKeyForPath = Replace(Mid$(sObjElements(0), InStr(sObjElements(0), "=") + 1, Len(sObjElements(0))), "#", "\") & "\win32"
    
    If KeyExists(HKEY_CLASSES_ROOT, "TypeLib\" & sTypeLib) Then
        sDescrption = GetString(HKEY_CLASSES_ROOT, "TypeLib\" & sTypeLib & "\" & sObjLine(1), "")
        sPath = GetString(HKEY_CLASSES_ROOT, "TypeLib\" & sRegKeyForPath, "")
    Else
        sDescrption = sObjElements(1)
        sPath = "- NA -"
    End If
    
    'See if directly to copy files ...
    If bCopyFiles Then
        If sPath <> "- NA -" Then
            INISetup App.Path & "\AppCfg.cfg", 5000
            If CInt(Read_Ini("options", "rename", "1")) = 1 Then
                FileCopy sPath, sDestFolder & FileNameFromPath(sPath, False) & "._" & Right(LCase(sPath), 2)
            Else
                FileCopy sPath, sDestFolder & FileNameFromPath(sPath)
            End If
        End If
    
    Else
    
        'No Direct copy, then Display List on frmMain ...
        With frmMain.lstDep.ListItems
            Set oListItem = .Add(, , sDescrption, 2, 2)
                
                If sPath = "- NA -" Then
                    oListItem.SubItems(2) = "- NA -"
                    oListItem.SubItems(3) = "Unregistered!"
                    oListItem.Bold = True
                    oListItem.ForeColor = vbRed
                ElseIf FileExists(sPath) Then
                    oListItem.SubItems(2) = Round(FileLen(sPath) / 1024, 0) & " KB"
                    oListItem.SubItems(3) = "File Exists."
                Else
                    oListItem.SubItems(2) = "- NA -"
                    oListItem.SubItems(3) = "File Missing!"
                    oListItem.Bold = True
                    oListItem.ForeColor = vbRed
                End If
            
                oListItem.SubItems(1) = "OCX"
                oListItem.SubItems(4) = sPath
                
            Set oListItem = Nothing
        End With
    End If
    
End Function
Public Function SaveSettingsToINI()
    
    INISetup App.Path & "\AppCfg.cfg", 5000
    
    With frmMain
    
        Write_Ini "options", "rename", .chkRenameExt.Value
        Write_Ini "options", "batch", .chkBatchFile.Value
        Write_Ini "options", "regsvr", .chkRegSvr.Value
        Write_Ini "options", "notify", .optNewFolder(0).Value
        Write_Ini "options", "shell", .chkShellMenu.Value
        Write_Ini "options", "ontop", .chkOnTop.Value
        Write_Ini "options", "openfld", .chkOpenFolder.Value
        
    End With
        
    
End Function

Public Function LoadSettingsFromINI()
    
    INISetup App.Path & "\AppCfg.cfg", 5000
    
    With frmMain
        
        .chkRenameExt.Value = CInt(Read_Ini("options", "rename", "1"))
        .chkBatchFile.Value = CInt(Read_Ini("options", "batch", "1"))
        .chkRegSvr.Value = CInt(Read_Ini("options", "regsvr", "0"))
        .optNewFolder(0).Value = CBool(Read_Ini("options", "notify", "0"))
        .optNewFolder(1).Value = Not (.optNewFolder(0).Value)
        .chkShellMenu.Value = CInt(Read_Ini("options", "shell", "0"))
        .chkOnTop.Value = CInt(Read_Ini("options", "ontop", "1"))
        .chkOpenFolder.Value = CInt(Read_Ini("options", "openfld", "0"))
        
    End With
    
End Function

Public Function IntegrateInShellMenu(Optional bRemove As Boolean = False)
    Dim sVBPFileKey As String
    
    'Get the VBP File Shell context menu handler Reg Key Name ...
    sVBPFileKey = GetString(HKEY_CLASSES_ROOT, ".vbp", "")
    
    If bRemove Then GoTo RemoveShellMenu
    
    'No such key? Then create it. very very very rare; but should consider it ...
    If IsStringEmpty(sVBPFileKey) Then
        sVBPFileKey = "VisualBasic.Project"
        SaveString HKEY_CLASSES_ROOT, ".vbp", "", sVBPFileKey
        If KeyExists(HKEY_CLASSES_ROOT, sVBPFileKey) = False Then
            SaveKey HKEY_CLASSES_ROOT, sVBPFileKey
        End If
    End If
    
    'If no Shell Menu handler added for VBP file then prepare reg key for it now ...
    If KeyExists(HKEY_CLASSES_ROOT, sVBPFileKey & "\shell") = False Then
        SaveKey HKEY_CLASSES_ROOT, sVBPFileKey & "\shell"
    End If
    
    'That's it. We are set to integrate this Application now!
    'First, add Scan option ...
    If KeyExists(HKEY_CLASSES_ROOT, sVBPFileKey & Scan_Key) = False Then
        SaveKey HKEY_CLASSES_ROOT, sVBPFileKey & Scan_Key
        SaveString HKEY_CLASSES_ROOT, sVBPFileKey & Scan_Key, "", Scan_MenuCaption
        SaveKey HKEY_CLASSES_ROOT, sVBPFileKey & ScanCommand_Key
        SaveString HKEY_CLASSES_ROOT, sVBPFileKey & ScanCommand_Key, "", Chr(34) & App.Path & "\" & App.EXEName & ".exe" & Chr(34) & " /s " & Chr(34) & "%1" & Chr(34)
    End If
    
    'Now, let's add Collect option ...
    If KeyExists(HKEY_CLASSES_ROOT, sVBPFileKey & Collect_Key) = False Then
        SaveKey HKEY_CLASSES_ROOT, sVBPFileKey & Collect_Key
        SaveString HKEY_CLASSES_ROOT, sVBPFileKey & Collect_Key, "", Collect_MenuCaption
        SaveKey HKEY_CLASSES_ROOT, sVBPFileKey & CollectCommand_Key
        SaveString HKEY_CLASSES_ROOT, sVBPFileKey & CollectCommand_Key, "", Chr(34) & App.Path & "\" & App.EXEName & ".exe" & Chr(34) & " /c " & Chr(34) & "%1" & Chr(34)
    End If
    
    Exit Function
    
RemoveShellMenu:
    'Remove all Shell Menu keys ...
    DeleteKey HKEY_CLASSES_ROOT, sVBPFileKey & ScanCommand_Key
    DeleteKey HKEY_CLASSES_ROOT, sVBPFileKey & Scan_Key
    DeleteKey HKEY_CLASSES_ROOT, sVBPFileKey & CollectCommand_Key
    DeleteKey HKEY_CLASSES_ROOT, sVBPFileKey & Collect_Key
    
End Function
