Attribute VB_Name = "mMain"
Option Explicit

Sub Main()
    Dim sProjPath As String
    Dim sProjFolder As String
    
'    sProjPath = "F:\My Projects\VB\O,P,Q\Programming Utility\Other\Rutu's PDCT\CollectDEP.vbp"
    sProjPath = Mid$(Command$, InStr(Command$, " ") + 1, Len(Command$))
    sProjPath = Replace(sProjPath, Chr(34), "")     '"File Name" to File Name. This happens when Shell Menu passes file name.
    sProjFolder = GetParentFolderPath(sProjPath, True)

    Select Case LCase(Left(Command$, 2))
        
        Case "/c"
            ProcessVBP sProjPath, True, sProjFolder & "Dependencies"
            
        Case "/s"
            frmMain.Show
            With frmMain
                .txtProjFile.Text = sProjPath
                .txtDepFolder.Text = sProjFolder & "Dependencies"
                .btnCopy.Enabled = True
            End With
            ProcessVBP sProjPath
    
        Case Else
            frmSplash.Show
            
            
    End Select
    
End Sub
