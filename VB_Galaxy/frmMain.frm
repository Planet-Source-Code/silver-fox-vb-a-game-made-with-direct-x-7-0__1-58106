VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Galaxy"
   ClientHeight    =   90
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   90
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   90
   ScaleWidth      =   90
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmMain"
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
'**************************************************************

Option Explicit

Private Sub Form_Load()

    'Display the startup form
    frmStart.Show vbModal
    
    'Logging
    Log "frmMain", "Form_Load", "****************** Galaxy starting up ******************"
    If gblnSound = False Then
        Log "frmMain", "Form_Load", "Sound is OFF"
    Else
        Log "frmMain", "Form_Load", "Sound is ON"
    End If
    If gblnMusic = False Then
        Log "frmMain", "Form_Load", "Music is OFF"
    Else
        Log "frmMain", "Form_Load", "Music is ON"
    End If
    If gblnDisplayFPS = False Then
        Log "frmMain", "Form_Load", "FPS display is OFF"
    Else
        Log "frmMain", "Form_Load", "FPS display is ON"
    End If
    If gblnVSYNC = True Then
        Log "frmMain", "Form_Load", "VSYNC is disabled"
    Else
        Log "frmMain", "Form_Load", "VSYNC is enabled"
    End If

    'TEMPORARY!: (Load game options)
    gstrUniverse = DEFAULT_UNIVERSE
    gbytDisplay = DISPLAY_TITLE
    'gbytDisplay = DISPLAY_TACTICAL

    'Ensure there's a Music directory
    Log "frmMain", "Form_Load", "Checking for music directory"
    If Dir(App.Path & "\Music", vbDirectory) = "" Then MkDir App.Path & "\Music"

    'Load everything up
    Me.Show
    On Local Error GoTo DIRECTX7ERROR
    Log "frmMain", "Form_Load", "Initializing DirectX7 object"
    Set gobjDX = New DirectX7
    On Error GoTo 0
    DInput.Initialize Me
    DDraw.Initialize Me
    DSound.Initialize Me
    Log "frmMain", "Form_Load", "Initializing tickcount timer"
    glngTimer = gobjDX.TickCount()

    'Loop until program ends
    Log "frmMain", "Form_Load", "Commencing main game loop"
    gblnRunning = True
    Do While gblnRunning
        'Display the appropriate screen
        Select Case gbytDisplay
            Case DISPLAY_TITLE
                Title.Main
            Case DISPLAY_TACTICAL
                Tactical.Main
        End Select
        'Let windows have a go
        DoEvents
        'Let ddraw, dsound, and dinput do their thing
        DDraw.Main
        DSound.Main
        DInput.Main
    Loop
    
    'Kill everything
    Log "frmMain", "Form_Load", "****************** Shutting down ******************"
    EndProgram Me
    
    'End the program
    End
    
    'Exit before error handler
    Exit Sub
    
'Error initializing DirectX7
DIRECTX7ERROR:
    Log "frmMain", "Form_Load", "Error initializing DirectX!"
    MsgBox "Error initializing DirectX.  A minimum of DirectX 7.0 is required to play Galaxy.  Visit www.microsoft.com/directx for the latest downloads."
    End
    
End Sub
