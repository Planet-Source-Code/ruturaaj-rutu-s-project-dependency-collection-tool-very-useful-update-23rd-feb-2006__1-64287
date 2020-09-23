VERSION 5.00
Begin VB.Form frmSplash 
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   5625
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   6795
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   5625
   ScaleWidth      =   6795
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrDelay 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   0
      Top             =   5280
   End
   Begin VB.Label lblVersion 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Version"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   2280
      TabIndex        =   2
      Top             =   4680
      Width           =   735
   End
   Begin VB.Label lblTime 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Press Esc or else this dialog will be closed automatically after x seconds."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   150
      Left            =   840
      TabIndex        =   1
      Top             =   5400
      Width           =   5070
   End
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "This Project is developed by Ruturaj. (mailme_friends@yahoo.com)"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H001135B3&
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6735
   End
   Begin VB.Image imgSplash 
      Height          =   5175
      Left            =   240
      Picture         =   "frmSplash.frx":0000
      Top             =   240
      Width           =   6540
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim iSecondCnt As Integer

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        tmrDelay.Enabled = False
        frmMain.Show
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    tmrDelay.Interval = 1000
    iSecondCnt = 5
    lblTime.Caption = Replace("Press Esc or else this dialog will be closed automatically after ##x## seconds.", "##x##", iSecondCnt)
    lblVersion = "Version: " & ApplicationBuild(False) & "; build: " & App.Revision
    tmrDelay.Enabled = True
End Sub

Private Sub lblInfo_Click()
    OpenURL "mailto:mailme_friends@yahoo.com", Me.hwnd
End Sub

Private Sub tmrDelay_Timer()
    If iSecondCnt > 0 Then
        iSecondCnt = iSecondCnt - 1
        lblTime.Caption = Replace("Press Esc or else this dialog will be closed automatically after ##x## seconds.", "##x##", iSecondCnt)
        tmrDelay.Enabled = True
    Else
        tmrDelay.Enabled = False
        frmMain.Show
        Unload Me
    End If
    
End Sub
