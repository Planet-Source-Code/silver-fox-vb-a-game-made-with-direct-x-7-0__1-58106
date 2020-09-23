Attribute VB_Name = "DInput"
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

'dX Variables
Dim mobjDI As DirectInput
Dim mobjDIKey As DirectInputDevice
Dim mobjDIKState As DIKEYBOARDSTATE
Dim mobjDIMouse As DirectInputDevice
Dim mobjDIMState As DIMOUSESTATE

Const MOUSE_SPEED = 2               'Speed of mouse cursor movement
Const CURSOR_RADIUS = 3             'Radius of mouse cursor circle

Global gintMouseX As Integer           'X Coordinate of the mouse cursor
Global gintMouseY As Integer           'Y Coordinate of the mouse cursor
Global gblnLMouseButton As Boolean     'Is the left mouse button being pressed?
Global gblnRMouseButton As Boolean     'Is the right mouse button being pressed?
Global gblnLMouseButtonUp As Boolean   'Was the left mouse button just released?
Global gblnRMouseButtonUp As Boolean   'Was the right mouse button just released?

'Loop counter
Dim i As Integer

'Public array showing which keys are active
Global gblnKey(211) As Boolean

'Keycode constants
Global Const DIK_ESCAPE = 1
Global Const DIK_1 = 2
Global Const DIK_2 = 3
Global Const DIK_3 = 4
Global Const DIK_4 = 5
Global Const DIK_5 = 6
Global Const DIK_6 = 7
Global Const DIK_7 = 8
Global Const DIK_8 = 9
Global Const DIK_9 = 10
Global Const DIK_0 = 11
Global Const DIK_MINUS = 12
Global Const DIK_EQUALS = 13
Global Const DIK_BACKSPACE = 14
Global Const DIK_TAB = 15
Global Const DIK_Q = 16
Global Const DIK_W = 17
Global Const DIK_E = 18
Global Const DIK_R = 19
Global Const DIK_T = 20
Global Const DIK_Y = 21
Global Const DIK_U = 22
Global Const DIK_I = 23
Global Const DIK_O = 24
Global Const DIK_P = 25
Global Const DIK_LBRACKET = 26
Global Const DIK_RBRACKET = 27
Global Const DIK_RETURN = 28
Global Const DIK_LCONTROL = 29
Global Const DIK_A = 30
Global Const DIK_S = 31
Global Const DIK_D = 32
Global Const DIK_F = 33
Global Const DIK_G = 34
Global Const DIK_H = 35
Global Const DIK_J = 36
Global Const DIK_K = 37
Global Const DIK_L = 38
Global Const DIK_SEMICOLON = 39
Global Const DIK_APOSTROPHE = 40
Global Const DIK_GRAVE = 41
Global Const DIK_LSHIFT = 42
Global Const DIK_BACKSLASH = 43
Global Const DIK_Z = 44
Global Const DIK_X = 45
Global Const DIK_C = 46
Global Const DIK_V = 47
Global Const DIK_B = 48
Global Const DIK_N = 49
Global Const DIK_M = 50
Global Const DIK_COMMA = 51
Global Const DIK_PERIOD = 52
Global Const DIK_SLASH = 53
Global Const DIK_RSHIFT = 54
Global Const DIK_MULTIPLY = 55
Global Const DIK_LALT = 56
Global Const DIK_SPACE = 57
Global Const DIK_CAPSLOCK = 58
Global Const DIK_F1 = 59
Global Const DIK_F2 = 60
Global Const DIK_F3 = 61
Global Const DIK_F4 = 62
Global Const DIK_F5 = 63
Global Const DIK_F6 = 64
Global Const DIK_F7 = 65
Global Const DIK_F8 = 66
Global Const DIK_F9 = 67
Global Const DIK_F10 = 68
Global Const DIK_NUMLOCK = 69
Global Const DIK_SCROLL = 70
Global Const DIK_NUMPAD7 = 71
Global Const DIK_NUMPAD8 = 72
Global Const DIK_NUMPAD9 = 73
Global Const DIK_SUBTRACT = 74
Global Const DIK_NUMPAD4 = 75
Global Const DIK_NUMPAD5 = 76
Global Const DIK_NUMPAD6 = 77
Global Const DIK_ADD = 78
Global Const DIK_NUMPAD1 = 79
Global Const DIK_NUMPAD2 = 80
Global Const DIK_NUMPAD3 = 81
Global Const DIK_NUMPAD0 = 82
Global Const DIK_DECIMAL = 83
Global Const DIK_F11 = 87
Global Const DIK_F12 = 88
Global Const DIK_NUMPADENTER = 156
Global Const DIK_RCONTROL = 157
Global Const DIK_DIVIDE = 181
Global Const DIK_RALT = 184
Global Const DIK_HOME = 199
Global Const DIK_UP = 200
Global Const DIK_PAGEUP = 201
Global Const DIK_LEFT = 203
Global Const DIK_RIGHT = 205
Global Const DIK_END = 207
Global Const DIK_DOWN = 208
Global Const DIK_PAGEDOWN = 209
Global Const DIK_INSERT = 210
Global Const DIK_DELETE = 211

Public Sub Initialize(frmInit As Form)

    'Create the direct input object
    On Local Error GoTo DIERROR
    Set mobjDI = gobjDX.DirectInputCreate()
        
    'Aquire the keyboard as the device
    On Local Error GoTo DIKEYERROR
    Set mobjDIKey = mobjDI.CreateDevice("GUID_SysKeyboard")
    
    'Get input nonexclusively, only when in foreground mode
    mobjDIKey.SetCommonDataFormat DIFORMAT_KEYBOARD
    mobjDIKey.SetCooperativeLevel frmInit.hWnd, DISCL_FOREGROUND Or DISCL_NONEXCLUSIVE
    mobjDIKey.Acquire
    
    'Initialize the key array
    For i = 1 To 211
        gblnKey(i) = False
    Next
    
    'Aquire the mouse as the diMouse device
    On Local Error GoTo DIMOUSEERROR
    Set mobjDIMouse = mobjDI.CreateDevice("GUID_SysMouse")
    
    'Get mouse input exclusively, but only when in foreground mode
    mobjDIMouse.SetCommonDataFormat DIFORMAT_MOUSE
    mobjDIMouse.SetCooperativeLevel frmInit.hWnd, DISCL_FOREGROUND Or DISCL_EXCLUSIVE
    mobjDIMouse.Acquire
    
    'Initialize the mouse variables
    gintMouseX = SCREEN_WIDTH \ 2
    gintMouseY = SCREEN_HEIGHT \ 2
    gblnLMouseButton = False
    gblnRMouseButton = False
    
    'Exit before error code
    On Error GoTo 0
    Exit Sub
    
'Error handlers
DIERROR:
    TerminateSpecific False, False, False, True
    Log "DInput", "Initialize", "Error initializing DInput!"
    MsgBox "Error initializing DirectInput.  Please report this to moml31@hotmail.com along with your system specifications, the log.txt file, and any information you think may be helpful."
    End
DIKEYERROR:
    TerminateSpecific False, False, False, True
    Log "DInput", "Initialize", "Error acquiring keyboard!"
    MsgBox "Can't acquire keyboard.  Please report this to moml31@hotmail.com along with your system specifications, the log.txt file, and any information you think may be helpful."
    End
DIMOUSEERROR:
    TerminateSpecific False, False, False, True
    Log "DInput", "Initialize", "Error acquiring mouse!"
    MsgBox "Can't acquire mouse.  Please report this to moml31@hotmail.com along with your system specifications, the log.txt file, and any information you think may be helpful."
    End
    
End Sub

Public Sub Main()

    'Refresh DInput info
    On Local Error GoTo DIMAINERROR
    RefreshKeyState
    RefreshMouseState
    
    'Exit before error code
    On Error GoTo 0
    Exit Sub
    
'Error handlers
DIMAINERROR:
    TerminateSpecific True, True, True, True
    Log "DInput", "Main", "Error in main DInput sub!"
    MsgBox "Error in main input subroutine.  Please report this to moml31@hotmail.com along with your system specifications, the log.txt file, and any information you think may be helpful."
    End

End Sub

Public Sub Terminate()
    
    'Unaquire and destroy
    On Local Error GoTo DITERMERROR
    mobjDIMouse.Unacquire
    Set mobjDIMouse = Nothing
    mobjDIKey.Unacquire
    Set mobjDIKey = Nothing
    Set mobjDI = Nothing
    
    'Exit before error code
    On Error GoTo 0
    Exit Sub
    
'Error handlers
DITERMERROR:
    TerminateSpecific False, False, False, True
    Log "DInput", "Terminate", "Error terminating DInput!"
    MsgBox "Error terminating DirectInput.  Please report this to moml31@hotmail.com along with your system specifications, the log.txt file, and any information you think may be helpful."
    End
    
End Sub

Public Sub RefreshKeyState()
    
    'Get the current state of the keyboard
    On Local Error Resume Next
    mobjDIKey.GetDeviceStateKeyboard mobjDIKState
    
    'If we've been forced to unaquire, try to reaquire
    If Err.Number <> 0 Then mobjDIKey.Acquire
    'If this fails, exit sub
    If Err.Number <> 0 Then Exit Sub
    On Error GoTo 0
    
    'Scan through all the keys to check which are depressed
    On Local Error GoTo DIKEYSCANERROR
    For i = 1 To 211
        If mobjDIKState.Key(i) <> 0 Then
            gblnKey(i) = True
        Else
            gblnKey(i) = False
        End If
    Next
    
    'Exit before error code
    On Error GoTo 0
    Exit Sub
    
'Error handlers
DIKEYSCANERROR:
    TerminateSpecific True, True, True, True
    Log "DInput", "RefreshKeyState", "Error polling keyboard for input!"
    MsgBox "Error polling keyboard for input.  Please report this to moml31@hotmail.com along with your system specifications, the log.txt file, and any information you think may be helpful."
    End
    
End Sub

Public Sub RefreshMouseState()

    'Get the current state of the mouse
    On Local Error Resume Next
    mobjDIMouse.GetDeviceStateMouse mobjDIMState
    
    'If we've been forced to unaquire, try to reaquire
    If Err.Number <> 0 Then mobjDIMouse.Acquire
    'If this fails, exit sub
    If Err.Number <> 0 Then Exit Sub
    On Error GoTo 0
    
    'Adjust the mouse cursor x coordinate
    On Local Error GoTo DIMOUSESCANERROR
    gintMouseX = gintMouseX + mobjDIMState.x * MOUSE_SPEED
    If gintMouseX < 0 Then gintMouseX = 0
    If gintMouseX > SCREEN_WIDTH - 1 Then gintMouseX = SCREEN_WIDTH - 1
    
    'Adjust the mouse cursor y coordinate
    gintMouseY = gintMouseY + mobjDIMState.y * MOUSE_SPEED
    If gintMouseY < 0 Then gintMouseY = 0
    If gintMouseY > SCREEN_HEIGHT - 1 Then gintMouseY = SCREEN_HEIGHT - 1
    
    'Check the left mouse button state
    gblnLMouseButtonUp = False
    If mobjDIMState.buttons(0) <> 0 Then gblnLMouseButton = True
    If mobjDIMState.buttons(0) = 0 Then
        'If it WAS down, but not anymore, set released
        If gblnLMouseButton = True Then gblnLMouseButtonUp = True
        gblnLMouseButton = False
    End If
    
    'Check the right mouse button state
    gblnRMouseButtonUp = False
    If mobjDIMState.buttons(1) <> 0 Then gblnRMouseButton = True
    If mobjDIMState.buttons(1) = 0 Then
        'If it WAS down, but not anymore, set released
        If gblnRMouseButton = True Then gblnRMouseButtonUp = True
        gblnRMouseButton = False
    End If
        
    'Exit before error code
    On Error GoTo 0
    Exit Sub
    
'Error handlers
DIMOUSESCANERROR:
    TerminateSpecific True, True, True, True
    Log "DInput", "RefreshMouseState", "Error refreshing mouse state!"
    MsgBox "Error polling mouse for input.  Please report this to moml31@hotmail.com along with your system specifications, the log.txt file, and any information you think may be helpful."
    End
                
End Sub
