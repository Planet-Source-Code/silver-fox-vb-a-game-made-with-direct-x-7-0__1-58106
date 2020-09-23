VERSION 5.00
Begin VB.Form frmStart 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3210
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   3210
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkVSYNC 
      BackColor       =   &H00000000&
      Caption         =   "Disable VSYNC"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1620
      TabIndex        =   10
      Top             =   2880
      Value           =   1  'Checked
      Width           =   1485
   End
   Begin VB.CheckBox chkLogging 
      BackColor       =   &H00000000&
      Caption         =   "Enable logging"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1620
      TabIndex        =   9
      Top             =   2640
      Value           =   1  'Checked
      Width           =   1485
   End
   Begin VB.TextBox txtKeys 
      Height          =   1575
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   8
      Top             =   720
      Visible         =   0   'False
      Width           =   4455
   End
   Begin VB.CommandButton cmdKeys 
      Caption         =   "&View Keys"
      Height          =   255
      Left            =   3240
      TabIndex        =   7
      Top             =   2400
      Width           =   1335
   End
   Begin VB.CommandButton cmdRun 
      Caption         =   "&Run Galaxy"
      Default         =   -1  'True
      Height          =   255
      Left            =   3240
      TabIndex        =   6
      Top             =   2700
      Width           =   1335
   End
   Begin VB.CheckBox chkFPS 
      BackColor       =   &H00000000&
      Caption         =   "Display FPS"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1620
      TabIndex        =   5
      Top             =   2400
      Value           =   1  'Checked
      Width           =   1485
   End
   Begin VB.CheckBox chkMusic 
      BackColor       =   &H00000000&
      Caption         =   "Enable Music"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   2640
      Width           =   1815
   End
   Begin VB.CheckBox chkSound 
      BackColor       =   &H00000000&
      Caption         =   "Enable Sound"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   2400
      Value           =   1  'Checked
      Width           =   1815
   End
   Begin VB.Label lblVersion 
      BackStyle       =   0  'Transparent
      Caption         =   "version"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   255
      Left            =   3550
      TabIndex        =   2
      Top             =   360
      Width           =   975
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Early Alpha, "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   255
      Left            =   2400
      TabIndex        =   1
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "GALAXY"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2055
   End
End
Attribute VB_Name = "frmStart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**************************************************************
'
' THIS WORK, INCLUDING THE SOURCE CODE, DOCUMENTATION
' AND RELATED MEDIA AND DATA, IS PLACED INTO THE PUBLIC DOMAIN.
'
' THE ORIGINAL AUTHOR IS SILVER FOX.
'
' THIS SOFTWARE IS PROVIDED AS-IS WITHOUT WARRANTY
' OF ANY KIND, NOT EVEN THE IMPLIED WARRANTY OF
' MERCHANTABILITY. THE AUTHOR OF THIS SOFTWARE,
' ASSUMES _NO_ RESPONSIBILITY FOR ANY CONSEQUENCE
' RESULTING FROM THE USE, MODIFICATION, OR
' REDISTRIBUTION OF THIS SOFTWARE.
'
'**************************************************************
'
' This file was downloaded from The Game Programming Wiki.
' Come and visit us at http://gpwiki.org
'
'**************************************************************

Option Explicit

Private Sub chkSound_Click()

    'Disable/enable music checkbox
    If chkSound.Value = vbChecked Then
        chkMusic.Enabled = True
    Else
        chkMusic.Value = vbUnchecked
        chkMusic.Enabled = False
    End If

End Sub

Private Sub cmdKeys_Click()

Dim strTemp As String
Dim bytTemp As Byte
Dim i As Long

    'Is the file there?
    If Dir(App.Path & "\keys.txt") = "" Then
        MsgBox "keys.txt file not found."
        Exit Sub
    End If

    'Display the keys.txt
    strTemp = ""
    Open App.Path & "\keys.txt" For Binary Access Read Lock Write As #1
        'Get the key text
        For i = 0 To LOF(1)
            'Read..
            Get 1, , bytTemp
            strTemp = strTemp & Chr(bytTemp)
        Next i
    Close #1

    'Let it be seen!
    txtKeys.Text = strTemp
    txtKeys.Visible = True

End Sub

Private Sub cmdRun_Click()

    'Set the global vars
    gblnSound = False
    gblnMusic = False
    gblnDisplayFPS = False
    gblnLogging = False
    gblnVSYNC = False
    If chkSound.Value = vbChecked Then gblnSound = True
    If chkMusic.Value = vbChecked Then gblnMusic = True
    If chkFPS.Value = vbChecked Then gblnDisplayFPS = True
    If chkLogging.Value = vbChecked Then gblnLogging = True
    If chkVSYNC.Value = vbChecked Then gblnVSYNC = True

    'Unload the form
    Unload Me

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    'Exit on escape
    If KeyCode = vbKeyEscape Then End

End Sub

Private Sub Form_Load()

    'Display version
    lblVersion.Caption = "v" & App.Major & "." & App.Minor & "." & App.Revision

End Sub
