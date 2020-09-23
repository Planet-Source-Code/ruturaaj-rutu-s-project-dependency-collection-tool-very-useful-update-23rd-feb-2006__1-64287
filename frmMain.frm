VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   7860
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9330
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7860
   ScaleWidth      =   9330
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picHeader 
      Align           =   1  'Align Top
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   1560
      Left            =   0
      Picture         =   "frmMain.frx":27A2
      ScaleHeight     =   1500
      ScaleWidth      =   9270
      TabIndex        =   0
      Top             =   0
      Width           =   9330
   End
   Begin MSComctlLib.ImageList imgDep 
      Left            =   1200
      Top             =   10080
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":30964
            Key             =   "dll"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":30DB6
            Key             =   "ocx"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":31208
            Key             =   "tlb"
         EndProperty
      EndProperty
   End
   Begin CollectDEP.CDlg CDlg1 
      Height          =   420
      Left            =   600
      TabIndex        =   24
      Top             =   10080
      Width           =   420
      _ExtentX        =   741
      _ExtentY        =   741
   End
   Begin VB.Frame fmOptions 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Options && Settings ..."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H001135B3&
      Height          =   3135
      Left            =   5160
      TabIndex        =   7
      Top             =   1680
      Width           =   4095
      Begin VB.Frame fmAppOptions 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   120
         TabIndex        =   16
         Top             =   2040
         Width           =   3855
         Begin VB.CheckBox chkShellMenu 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Integrate with Shell Context Menu."
            ForeColor       =   &H009C6838&
            Height          =   255
            Left            =   120
            TabIndex        =   26
            Top             =   240
            Width           =   3615
         End
         Begin VB.CheckBox chkOnTop 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Set Application Window On Top."
            ForeColor       =   &H009C6838&
            Height          =   255
            Left            =   120
            TabIndex        =   17
            Top             =   480
            Width           =   3375
         End
      End
      Begin VB.Frame fmDepFileManipulation 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   3855
         Begin VB.CheckBox chkRegSvr 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Add RegSvr32.exe to Dependencies."
            ForeColor       =   &H009C6838&
            Height          =   255
            Left            =   120
            TabIndex        =   11
            Top             =   600
            Width           =   3615
         End
         Begin VB.CheckBox chkBatchFile 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Create Auto-Rename Batch File."
            Enabled         =   0   'False
            ForeColor       =   &H009C6838&
            Height          =   255
            Left            =   120
            TabIndex        =   10
            Top             =   360
            Width           =   3255
         End
         Begin VB.CheckBox chkRenameExt 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Change Dependency File Extensions."
            ForeColor       =   &H009C6838&
            Height          =   255
            Left            =   120
            TabIndex        =   9
            Top             =   120
            Width           =   3615
         End
      End
      Begin VB.Frame fmNewFolder 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   120
         TabIndex        =   12
         Top             =   1200
         Width           =   3855
         Begin VB.OptionButton optNewFolder 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Notify me."
            ForeColor       =   &H009C6838&
            Height          =   255
            Index           =   1
            Left            =   2280
            TabIndex        =   15
            Top             =   480
            Value           =   -1  'True
            Width           =   1215
         End
         Begin VB.OptionButton optNewFolder 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Create New Folder."
            ForeColor       =   &H009C6838&
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   14
            Top             =   480
            Width           =   2055
         End
         Begin VB.Label lblInfo 
            BackStyle       =   0  'Transparent
            Caption         =   "If Specified Folder to copy Dependency Files does not exist, then ..."
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   375
            Index           =   1
            Left            =   120
            TabIndex        =   13
            Top             =   120
            Width           =   3495
         End
      End
   End
   Begin VB.Frame fmCopyDep 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Copy all Selected Dependency File(s) ..."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H009C6838&
      Height          =   1215
      Left            =   120
      TabIndex        =   18
      Top             =   3600
      Width           =   4935
      Begin VB.CheckBox chkOpenFolder 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Open Dependency Folder, when done."
         ForeColor       =   &H009C6838&
         Height          =   255
         Left            =   720
         TabIndex        =   25
         Top             =   840
         Width           =   3615
      End
      Begin CollectDEP.XPButton btnCopy 
         Height          =   375
         Left            =   840
         TabIndex        =   19
         Top             =   360
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   661
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "&Copy Dependency File(s) Now!"
         ForeColor       =   4210752
         ForeHover       =   1127859
      End
   End
   Begin VB.Frame fmDepFolder 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Folder for Dependency File(s) ... "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H009C6838&
      Height          =   855
      Left            =   120
      TabIndex        =   4
      Top             =   2640
      Width           =   4935
      Begin VB.TextBox txtDepFolder 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         OLEDropMode     =   1  'Manual
         TabIndex        =   5
         Top             =   360
         Width           =   4215
      End
      Begin CollectDEP.XPButton btnBrowse 
         Height          =   285
         Index           =   1
         Left            =   4440
         TabIndex        =   6
         Top             =   360
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   503
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "..."
         ForeColor       =   -2147483642
         ForeHover       =   0
      End
   End
   Begin VB.Frame fmDep 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Select Dependency File(s) ... "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H009C6838&
      Height          =   2895
      Left            =   120
      TabIndex        =   20
      Top             =   4920
      Width           =   9135
      Begin MSComctlLib.ListView lstDep 
         Height          =   2175
         Left            =   120
         TabIndex        =   21
         Top             =   360
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   3836
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         OLEDropMode     =   1
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         Icons           =   "imgDep"
         SmallIcons      =   "imgDep"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         OLEDropMode     =   1
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Dependency Desciption"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Type"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "File Size"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "File Status"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "File Path"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label lblSize 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   6000
         TabIndex        =   23
         Top             =   2640
         Width           =   120
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Dependency Overheads (KB):"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H001135B3&
         Height          =   195
         Index           =   0
         Left            =   3000
         TabIndex        =   22
         Top             =   2640
         Width           =   2850
      End
   End
   Begin VB.Frame fmVBP 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Select VBP Project ... "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H009C6838&
      Height          =   855
      Left            =   120
      TabIndex        =   1
      Top             =   1680
      Width           =   4935
      Begin CollectDEP.XPButton btnBrowse 
         Height          =   285
         Index           =   0
         Left            =   4440
         TabIndex        =   3
         Top             =   360
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   503
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "..."
         ForeColor       =   -2147483642
         ForeHover       =   0
      End
      Begin VB.TextBox txtProjFile 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         OLEDropMode     =   1  'Manual
         TabIndex        =   2
         Top             =   360
         Width           =   4215
      End
   End
   Begin VB.Menu mnuPopDepList 
      Caption         =   "PopupDepList"
      Visible         =   0   'False
      Begin VB.Menu mnuPopDepListAdd 
         Caption         =   "Add to List"
      End
      Begin VB.Menu mnuPopDepListRemove 
         Caption         =   "Remove from List"
      End
      Begin VB.Menu mnuPopDepListLine1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPopDepListSelAll 
         Caption         =   "Select All"
      End
      Begin VB.Menu mnuPopDepListDeSelAll 
         Caption         =   "Deselect All"
      End
      Begin VB.Menu mnuPopDepListLine2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPopDepListLocate 
         Caption         =   "Locate on Disk"
      End
      Begin VB.Menu mnuPopDepListLine3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPopDepListPropDlg 
         Caption         =   "Properties"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnBrowse_Click(Index As Integer)
    Select Case Index
        
        Case 0
            'Browse for VBP File ...
            With CDlg1
                .DialogTitle = "Select Visual Basic 6.0 Project File ..."
                .DefaultExt = "vbp"
                .Filename = ""
                .Flags = cdlOFNFileMustExist
                .Filter = "Visual Basic 6.0 Project File    (*.vbp)|*.vbp"
                .ShowOpen
                
                If IsStringEmpty(.Filename) Then Exit Sub
                
                txtProjFile.Text = .Filename
                txtDepFolder.Text = GetParentFolderPath(.Filename, True) & "Dependencies"
                txtDepFolder.SelStart = Len(txtDepFolder.Text)
                
            End With
            
            'Process VBP file to find the dependancy files ...
            ProcessVBP txtProjFile.Text
            
            'Toggle Enabled Status of Copy button ...
            If lstDep.ListItems.Count > 0 Then
                btnCopy.Enabled = True
            Else
                btnCopy.Enabled = False
            End If
        
        
        Case 1
            
            'Browse for Dependancy Folder ...
            With CDlg1
                .DialogTitle = "Select Folder to copy all Dependency Files ..."
                .FolderPath = ""
                .Flags = BIF_EDITBOX + BIF_DONTGOBELOWDOMAIN
                .ShowBrowseForFolder
                
                If IsStringEmpty(.FolderPath) Then Exit Sub
                
                txtDepFolder.Text = .FolderPath
            End With
            
    End Select
    
End Sub

Private Sub btnCopy_Click()
    Dim iCnt As Integer
    Dim bAtleastOne As Boolean
    Dim sDepFolder As String
    
    bAtleastOne = False
    For iCnt = 1 To lstDep.ListItems.Count
        If lstDep.ListItems(iCnt).Checked = True Then
            bAtleastOne = True
            Exit For
        End If
    Next iCnt
    
    If bAtleastOne = False Then
        MsgBox "No Dependency Files selected from List of Project Dependency Files. Please make a Check mark for the Dependency File which you want to copy to specified Folder.", vbInformation, "No File Selected!"
        Exit Sub
    End If
    
    If FolderExists(txtDepFolder.Text) = False Then
        If optNewFolder(0).Value = False Then
            If MsgBox("Specified Folder to copy Dependency Files does not exists. Would you like to create a New folder at specified Location?", vbYesNo + vbQuestion, "Specified Folder does not exists. Create New?") = vbNo Then
                Exit Sub
            Else
                MkDir txtDepFolder.Text
            End If
        Else
            MkDir txtDepFolder.Text
        End If
    End If
    
    If Right(txtDepFolder.Text, 1) = "\" Then
        sDepFolder = txtDepFolder.Text
    Else
        sDepFolder = txtDepFolder.Text & "\"
    End If
    
    btnCopy.Enabled = False
    
    For iCnt = 1 To lstDep.ListItems.Count
        If lstDep.ListItems(iCnt).Checked And lstDep.ListItems(iCnt).Bold = False Then
            If chkRenameExt.Value = vbChecked Then
                FileCopy lstDep.ListItems(iCnt).SubItems(4), sDepFolder & FileNameFromPath(lstDep.ListItems(iCnt).SubItems(4), False) & "._" & Right(LCase(lstDep.ListItems(iCnt).SubItems(4)), 2)
            Else
                FileCopy lstDep.ListItems(iCnt).SubItems(4), sDepFolder & FileNameFromPath(lstDep.ListItems(iCnt).SubItems(4))
            End If
        End If
    Next iCnt
    
    If chkBatchFile.Enabled And chkBatchFile.Value = vbChecked Then
        GetFileFromRes "bat", "files", sDepFolder & "AutoRename.bat"
    End If
    
    If chkRegSvr.Value = vbChecked Then
        GetFileFromRes "regsvr32", "files", sDepFolder & "regsvr32.exe"
    End If
    
    MsgBox "Process of collecting Project Dependencies completed successfully!", vbInformation, "Done!"
    
    btnCopy.Enabled = True
    
    If chkOpenFolder.Value = vbChecked Then
        OpenURL txtDepFolder.Text, Me.hwnd
    End If

End Sub

Private Sub chkOnTop_Click()
    SetAlwaysOnTopMode Me.hwnd, CBool(chkOnTop.Value)
End Sub

Private Sub chkRenameExt_Click()
    chkBatchFile.Enabled = CBool(chkRenameExt.Value)
End Sub

Private Sub chkShellMenu_Click()
    If FileExists(App.Path & "\" & App.EXEName & ".exe") = False Then
        MsgBox "Application Executable File (" & App.EXEName & ".exe ) does not found in Application Folder. Shell Menu Integration requires it. In absense of " & App.EXEName & ".exe file, Shell Menu integration can not be performed.", vbInformation, "Application's EXE file Missing!"
        chkShellMenu.Value = vbUnchecked
        Exit Sub
    End If
    
    IntegrateInShellMenu Not (CBool(chkShellMenu.Value))
    
End Sub

Private Sub Form_Load()
    'Only one instance at a time ...
    If App.PrevInstance Then End
    
    'Set Application Title ...
    Me.Caption = "Rutu's Project Dependency Collection Tool (version: " & ApplicationBuild(False) & "; build: " & App.Revision & " )"
    
    'Lstview display adjustments ...
    With lstDep
        .ColumnHeaders(1).Width = .Width * 0.4
        .ColumnHeaders(2).Width = .Width * 0.1
        .ColumnHeaders(3).Width = .Width * 0.15
        .ColumnHeaders(4).Width = .Width * 0.2
        .ColumnHeaders(5).Width = .Width * 0.4
    End With
    
    
    btnCopy.Enabled = (Left(Command$, 1) = "/")
    
    'Load Settings ...
    LoadSettingsFromINI
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveSettingsToINI
End Sub

Private Sub lstDep_DblClick()
    
    If lstDep.SelectedItem.Bold = False Then
        OpenURL GetParentFolderPath(lstDep.SelectedItem.SubItems(4), False), Me.hwnd
    End If
    
End Sub

Private Sub lstDep_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    If Item.Bold Then
        Item.Checked = False
        Exit Sub
    End If
    
    If Item.Checked Then
        lblSize.Caption = Round(CLng(lblSize.Caption) + FileLen(Item.SubItems(4)) / 1024, 0)
    Else
        lblSize.Caption = Round(CLng(lblSize.Caption) - FileLen(Item.SubItems(4)) / 1024, 0)
    End If
    
End Sub

Private Sub lstDep_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If IsStringEmpty(txtProjFile.Text) Then Exit Sub
    
    If Button = vbRightButton Then
        If lstDep.ListItems.Count > 0 Then
            mnuPopDepListLocate.Enabled = Not (lstDep.SelectedItem.Bold)
            mnuPopDepListRemove.Enabled = True
            mnuPopDepListSelAll.Enabled = True
            mnuPopDepListDeSelAll.Enabled = True
        Else
            mnuPopDepListLocate.Enabled = False
            mnuPopDepListRemove.Enabled = False
            mnuPopDepListSelAll.Enabled = False
            mnuPopDepListDeSelAll.Enabled = False
        End If
        
        PopupMenu mnuPopDepList
        
    End If
    
End Sub

Private Sub lstDep_OLEDragDrop(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim sFileExt As String
    
    If IsStringEmpty(txtProjFile.Text) Then
        MsgBox "No Visual Basic 6.0 Project is selected. Please select a VBP Project and then try adding Dependency Files manually.", vbInformation, "Please select VBP Project File first ..."
        Exit Sub
    End If
    
    sFileExt = LCase(Right(Data.Files(1), 3))
    
    If sFileExt <> "dll" And sFileExt <> "ocx" And sFileExt <> "tlb" Then
        Exit Sub
    Else
        AddToList Data.Files(1)
    End If
    
End Sub

Private Sub mnuPopDepListAdd_Click()
    
    With CDlg1
        .DialogTitle = "Select DLL, OCX or TLB to add to Dependency List ..."
        .Filename = ""
        .Flags = cdlOFNFileMustExist
        .Filter = "Dependency Files|*.ocx;*.dll;*.tlb|Active-X Components|*.ocx|Dynamic Link Libraries|*.dll|Type Libraries|*.tlb|All Files|*.*"
        .ShowOpen
        
        If IsStringEmpty(.Filename) Then Exit Sub

        AddToList .Filename
    
    End With
    
End Sub

Private Sub mnuPopDepListDeSelAll_Click()
    Dim iCnt As Integer
    
    For iCnt = 1 To lstDep.ListItems.Count
        If lstDep.ListItems(iCnt).Bold = False Then
            lstDep.ListItems(iCnt).Checked = False
        End If
    Next iCnt
End Sub

Private Sub mnuPopDepListLocate_Click()
    If lstDep.SelectedItem.Bold Then Exit Sub
    
    OpenURL GetParentFolderPath(lstDep.SelectedItem.SubItems(4), True), Me.hwnd
    
End Sub

Private Sub mnuPopDepListPropDlg_Click()
    Dim sFilename As String
    
    If lstDep.SelectedItem Is Nothing Then Exit Sub
    If lstDep.SelectedItem.Bold = True Then Exit Sub
    
    sFilename = lstDep.SelectedItem.SubItems(4)
    
    Call ShowPropertiesDlg(sFilename, Me.hwnd)
    
End Sub

Private Sub mnuPopDepListRemove_Click()
    
    If lstDep.SelectedItem.Bold Then
        lstDep.ListItems.Remove lstDep.SelectedItem.Index
    Else
        If MsgBox("Are you sure you want to Remove " & FileNameFromPath(lstDep.SelectedItem.SubItems(4), True) & " ?" & vbCrLf & vbCrLf & "Please note that Application will just remove the name of " & FileNameFromPath(lstDep.SelectedItem.SubItems(4), True) & " from the list of Dependencies leaving its Project Reference from " & FileNameFromPath(txtProjFile.Text, True) & " unchanged.", vbQuestion + vbYesNo, "Remove " & FileNameFromPath(lstDep.SelectedItem.SubItems(4), True) & " from List?") = vbYes Then
            lstDep.ListItems.Remove lstDep.SelectedItem.Index
        End If
    End If
    
End Sub

Private Sub mnuPopDepListSelAll_Click()
    Dim iCnt As Integer
    
    For iCnt = 1 To lstDep.ListItems.Count
        If lstDep.ListItems(iCnt).Bold = False Then
            lstDep.ListItems(iCnt).Checked = True
        End If
    Next iCnt
    
End Sub

Private Sub txtDepFolder_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If FileExists(Data.Files(1)) = False Then
        txtDepFolder.Text = Data.Files(1)
    End If
    
End Sub

Private Sub txtProjFile_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    'Check if Dragged file is a Visual Basic Project File ...
    If Right(LCase(Data.Files(1)), 3) = "vbp" Then
        txtProjFile.Text = Data.Files(1)
    Else
        Effect = vbDropEffectScroll
        Exit Sub
    End If
    
    'Suggest Dependency Folder Location under Project Folder itself ...
    If IsStringEmpty(txtDepFolder.Text) Then
        txtDepFolder.Text = GetParentFolderPath(txtProjFile.Text, True) & "Dependencies"
        txtDepFolder.SelStart = Len(txtDepFolder.Text)
    End If
    
    'Process VBP file to find the dependancy files ...
    ProcessVBP txtProjFile.Text
    
    'Toggle Enabled Status of Copy button ...
    If lstDep.ListItems.Count > 0 Then
        btnCopy.Enabled = True
    Else
        btnCopy.Enabled = False
    End If
    
End Sub
