Attribute VB_Name = "DDraw"
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

'DirectX variables
Dim mobjDD As DirectDraw7
Dim msurfFront As DirectDrawSurface7
Dim msurfBack As DirectDrawSurface7

'Gamma
Dim mobjGammaControler As DirectDrawGammaControl 'The object that gets/sets gamma ramps
Dim mudtGammaRamp As DDGAMMARAMP                 'The gamma ramp we'll use to alter the screen state
Dim mudtOriginalRamp As DDGAMMARAMP              'The gamma ramp we'll use to store the original screen state
Dim mintRedVal As Integer                        'Store the current red value w.r.t. original
Dim mintGreenVal As Integer                      'Store the current green value w.r.t. original
Dim mintBlueVal As Integer                       'Store the current blue value w.r.t. original
Dim mintTargetRedVal As Integer                  'Store the target red value
Dim mintTargetGreenVal As Integer                'Store the target green value
Dim mintTargetBlueVal As Integer                 'Store the target blue value
Public pblnGamma As Boolean                      'Do we have gamma support?
Const FADE_DELAY_MS = 12                         'How fast should the fade be?

'Our permanent rectangles
Dim mrectScreen As RECT                     'Rectangle the size of the screen
Dim mrectTactical As RECT                   'Rectangle the size of the tactical viewport

'Program flow variables
Dim mlngFrameTime As Long                   'How long since last frame?
Dim mlngTimer As Long                       'How long since last FPS count update?
Dim mintFPSCounter As Integer               'Our FPS counter
Dim mintFPS As Integer                      'Our FPS storage variable

'Surfaces array
Private Type SURFACETYPE
    surfSprite As DirectDrawSurface7
    intWidth As Integer
    intHeight As Integer
    strResName As String
    blnExists As Boolean
    blnReserved As Boolean
End Type
Dim mudtSurface() As SURFACETYPE

'Queue array
Dim mlngQueue() As Long

'Sprite Objects array
Private Type SPRITEOBJECTTYPE
    blnExists As Boolean
    strResName As String
    lngIndex() As Long
    intUses As Integer          'How many different instances of this object are there?
End Type
Dim mudtSpriteObject() As SPRITEOBJECTTYPE

'ROPs
Dim mblnSrcPaint As Boolean

Public Sub Initialize(frmInit As Form)

Dim hwCaps As DDCAPS
Dim helCaps As DDCAPS
Dim ddsdMain As DDSURFACEDESC2
Dim ddsdFlip As DDSURFACEDESC2
Dim i As Integer
Dim j As Integer
    
    'Initialize DirectDraw
    On Local Error GoTo DDCREATEERROR
    Log "DDraw", "Initialize", "Initializing DDraw object"
    Set mobjDD = gobjDX.DirectDrawCreate("")
    
    'Check for Gamma Ramp Support
    On Local Error GoTo DDGAMMATESTERROR
    Log "DDraw", "Initialize", "Checking for gamma capability"
    pblnGamma = True
    mobjDD.GetCaps hwCaps, helCaps
    If (hwCaps.lCaps2 And DDCAPS2_PRIMARYGAMMA) = 0 Then pblnGamma = False
    
    'Set the cooperative level (Fullscreen exclusive)
    On Local Error GoTo DDCOOPERROR
    Log "DDraw", "Initialize", "Setting cooperative level"
    mobjDD.SetCooperativeLevel frmInit.hWnd, DDSCL_FULLSCREEN Or DDSCL_EXCLUSIVE
    
    'Set the resolution
    On Local Error GoTo DDDISPLAYERROR
    Log "DDraw", "Initialize", "Changing resolution"
    mobjDD.SetDisplayMode SCREEN_WIDTH, SCREEN_HEIGHT, SCREEN_BITDEPTH, 0, DDSDM_DEFAULT

    'Describe the flipping chain architecture we'd like to use
    On Local Error GoTo DDCREATEFRONTERROR
    Log "DDraw", "Initialize", "Setting up the flipping chain"
    ddsdMain.lFlags = DDSD_CAPS Or DDSD_BACKBUFFERCOUNT
    ddsdMain.lBackBufferCount = 1
    ddsdMain.ddsCaps.lCaps = DDSCAPS_COMPLEX Or DDSCAPS_FLIP Or DDSCAPS_PRIMARYSURFACE
    
    'Create the primary surface
    Log "DDraw", "Initialize", "Creating primary surface"
    Set msurfFront = mobjDD.CreateSurface(ddsdMain)
    
    'Create the backbuffer
    On Local Error GoTo DDCREATEBACKERROR
    Log "DDraw", "Initialize", "Creating backbuffer"
    ddsdFlip.ddsCaps.lCaps = DDSCAPS_BACKBUFFER
    Set msurfBack = msurfFront.GetAttachedSurface(ddsdFlip.ddsCaps)
    
    'Set the text colour for the backbuffer
    Log "DDraw", "Initialize", "Setting backbuffer text color"
    msurfBack.SetForeColor vbGreen
    msurfBack.SetFontTransparency True

    'If we have the gamma cap
    On Local Error GoTo DDINITGAMMAERROR
    If (pblnGamma = True) Then
        'Make a new gamma controller
        Log "DDraw", "Initialize", "Creating gamma controller"
        Set mobjGammaControler = msurfFront.GetDirectDrawGammaControl
        'Fill out the original gamma ramps
        mobjGammaControler.GetGammaRamp DDSGR_DEFAULT, mudtOriginalRamp
    End If
    
    'Set our initial colour values to zero
    On Local Error GoTo DDINITVARSERROR
    gblnFadeComplete = True
    mintTargetRedVal = 0
    mintTargetGreenVal = 0
    mintTargetBlueVal = 0
    mintRedVal = 0
    mintGreenVal = 0
    mintBlueVal = 0
    
    'Create our screen-sized rectangle
    mrectScreen.Bottom = SCREEN_HEIGHT
    mrectScreen.Right = SCREEN_WIDTH
    
    'Create our tactical viewport rectangle
    mrectTactical.Bottom = TACTICAL_BOTTOM
    mrectTactical.Top = TACTICAL_TOP
    mrectTactical.Right = TACTICAL_RIGHT
    mrectTactical.Left = TACTICAL_LEFT
    
    'Load the footer
    LoadFooter App.Path & "\resource.bnk"
    
    'Init our surface array
    ReDim mudtSurface(0)
    
    'Init the queue array
    ReDim mlngQueue(0)
    mlngQueue(0) = -1
    
    'Init the sprite object array
    ReDim mudtSpriteObject(0)
    mudtSpriteObject(0).blnExists = False
    
    'Clear the backbuffer
    ClearBuffer
                
    'Test the SrcPaint ROP
    mblnSrcPaint = TestROP(msurfBack, vbSrcPaint)
    
    'Exit before error code
    On Error GoTo 0
    Exit Sub
    
'Error handlers
DDCREATEERROR:
    TerminateSpecific False, False, True, True
    Log "DDraw", "Initialize", "Error initializing DirectDraw!"
    MsgBox "Error initializing DirectDraw.  Please report this to moml31@hotmail.com along with your system specifications, the log.txt file,, the log.txt file, and any information you think may be helpful."
    End
DDGAMMATESTERROR:
    TerminateSpecific False, True, True, True
    Log "DDraw", "Initialize", "Error testing gamma capabilities!"
    MsgBox "Error testing for gamma capabilities.  Please report this to moml31@hotmail.com along with your system specifications, the log.txt file, and any information you think may be helpful."
    End
DDCOOPERROR:
    TerminateSpecific False, True, True, True
    Log "DDraw", "Initialize", "Error setting cooperative level!"
    MsgBox "Error setting cooperative level.  Please report this to moml31@hotmail.com along with your system specifications, the log.txt file, and any information you think may be helpful."
    End
DDDISPLAYERROR:
    TerminateSpecific False, True, True, True
    Log "DDraw", "Initialize", "Error setting display mode!"
    MsgBox "Error setting display mode.  Please report this to moml31@hotmail.com along with your system specifications, the log.txt file, and any information you think may be helpful."
    End
DDCREATEFRONTERROR:
    TerminateSpecific False, True, True, True
    Log "DDraw", "Initialize", "Error initializing front buffer!"
    MsgBox "Error initializing front buffer.  Please report this to moml31@hotmail.com along with your system specifications, the log.txt file, and any information you think may be helpful."
    End
DDCREATEBACKERROR:
    TerminateSpecific False, True, True, True
    Log "DDraw", "Initialize", "Error initializing backbuffer!"
    MsgBox "Error initializing back buffer.  Please report this to moml31@hotmail.com along with your system specifications, the log.txt file, and any information you think may be helpful."
    End
DDINITGAMMAERROR:
    TerminateSpecific False, True, True, True
    Log "DDraw", "Initialize", "Error initializing gamma functions!"
    MsgBox "Error initializing gamma functions.  Please report this to moml31@hotmail.com along with your system specifications, the log.txt file, and any information you think may be helpful."
    End
DDINITVARSERROR:
    TerminateSpecific False, True, True, True
    Log "DDraw", "Initialize", "Error initializing DDraw variables!"
    MsgBox "Error initializing DirectDraw variables.  Please report this to moml31@hotmail.com along with your system specifications, the log.txt file, and any information you think may be helpful."
    End

End Sub

Public Sub Main()

    'Handle lost surfaces
    HandleLostSurfaces
    
    'Process the queue
    DDraw.ProcessQueue

    'Call the timer
    Timer
    
    'Display the FPS
    If gblnDisplayFPS Then TextOut gintFPS & " FPS", 20, 20, vbGreen

    'Perform gamma fade
    Fade

    'Flip the backbuffer to frontbuffer
    Flip
            
End Sub

Public Sub Flip()

    'Flip the backbuffer to frontbuffer
    On Local Error GoTo DDFLIPERROR
    If gblnVSYNC = True Then
        msurfFront.Flip Nothing, DDFLIP_NOVSYNC
    Else
        msurfFront.Flip Nothing, DDFLIP_WAIT
    End If
    'msurfFront.BltFast 0, 0, msurfBack, mrectScreen, DDBLTFAST_WAIT
    
    'Clear the backbuffer
    ClearBuffer

    'Exit before error code
    On Error GoTo 0
    Exit Sub
    
'Error handlers
DDFLIPERROR:
    TerminateSpecific False, True, True, True
    Log "DDraw", "Flip", "Error flipping backbuffer to primary surface!"
    MsgBox "Error flipping front buffer.  Please report this to moml31@hotmail.com along with your system specifications, the log.txt file, and any information you think may be helpful."
    End

End Sub

Public Sub ClearBuffer()

    'Clear the backbuffer
    On Local Error GoTo DDCLEARERROR
    msurfBack.BltColorFill mrectScreen, 0

    'Exit before error code
    On Error GoTo 0
    Exit Sub
    
'Error handlers
DDCLEARERROR:
    TerminateSpecific False, True, True, True
    Log "DDraw", "ClearBuffer", "Error clearing backbuffer!"
    MsgBox "Error clearing backbuffer.  Please report this to moml31@hotmail.com along with your system specifications, the log.txt file, and any information you think may be helpful."
    End

End Sub

Public Sub TextOut(strText As String, intX As Integer, intY As Integer, lngColor As Long)

    'This function writes text to the backbuffer
    On Local Error GoTo DDTEXTOUTERROR
    msurfBack.SetForeColor lngColor
    msurfBack.DrawText intX, intY, strText, False

    'Exit before error code
    On Error GoTo 0
    Exit Sub
    
'Error handlers
DDTEXTOUTERROR:
    Log "DDraw", "TextOut", "Error using textout!  " & intX & ", " & intY & "  color: " & lngColor & " string: " & strText & " err: " & Err.Description
    TerminateSpecific False, True, True, True
    MsgBox "Error using DirectDraw Text Out.  Please report this to moml31@hotmail.com along with your system specifications, the log.txt file, and any information you think may be helpful."
    End

End Sub

Public Sub SetFade(intRed, intGreen, intBlue)

    'Fade is not complete
    On Local Error GoTo DDSETFADEERROR
    gblnFadeComplete = False
    
    'Set the target values
    mintTargetRedVal = intRed
    mintTargetGreenVal = intGreen
    mintTargetBlueVal = intBlue

    'Exit before error code
    On Error GoTo 0
    Exit Sub
    
'Error handlers
DDSETFADEERROR:
    TerminateSpecific False, True, True, True
    Log "DDraw", "SetFade", "Error setting gamma fade! (" & intRed & ", " & intGreen & ", " & intBlue & ")"
    MsgBox "Error setting gamma fade values.  Please report this to moml31@hotmail.com along with your system specifications, the log.txt file, and any information you think may be helpful."
    End

End Sub

Private Sub Fade()

    'Gamma fade
    On Local Error GoTo DDFADEERROR
    If pblnGamma And gblnFadeComplete = False Then
        'Set the new values
        If mintRedVal > mintTargetRedVal Then mintRedVal = mintRedVal - 1
        If mintRedVal < mintTargetRedVal Then mintRedVal = mintRedVal + 1
        If mintGreenVal > mintTargetGreenVal Then mintGreenVal = mintGreenVal - 1
        If mintGreenVal < mintTargetGreenVal Then mintGreenVal = mintGreenVal + 1
        If mintBlueVal > mintTargetBlueVal Then mintBlueVal = mintBlueVal - 1
        If mintBlueVal < mintTargetBlueVal Then mintBlueVal = mintBlueVal + 1
        'Display the new ramp
        SetGamma mintRedVal, mintGreenVal, mintBlueVal
        'Check for fade completion
        If (mintRedVal = mintTargetRedVal) And (mintGreenVal = mintTargetGreenVal) And (mintBlueVal = mintTargetBlueVal) Then gblnFadeComplete = True
    End If

    'Exit before error code
    On Error GoTo 0
    Exit Sub
    
'Error handlers
DDFADEERROR:
    TerminateSpecific False, True, True, True
    Log "DDraw", "Fade", "Error performing gamma fade! (" & mintRedVal & ", " & mintGreenVal & ", " & mintBlueVal & ")"
    MsgBox "Error performing gamma fade.  Please report this to moml31@hotmail.com along with your system specifications, the log.txt file, and any information you think may be helpful."
    End

End Sub

Public Sub FadeIn()

Dim i As Integer

    'Gamma fade in!
    On Local Error GoTo DDFADEINERROR
    If pblnGamma Then
        For i = -100 To 0 Step 1
            SlowDown
            mintRedVal = i
            mintGreenVal = i
            mintBlueVal = i
            SetGamma mintRedVal, mintGreenVal, mintBlueVal
        Next
    End If

    'Exit before error code
    On Error GoTo 0
    Exit Sub
    
'Error handlers
DDFADEINERROR:
    TerminateSpecific False, True, True, True
    Log "DDraw", "FadeIn", "Error calling gamma fade in! (" & mintRedVal & ", " & mintGreenVal & ", " & mintBlueVal & ")"
    MsgBox "Error performing gamma fade in.  Please report this to moml31@hotmail.com along with your system specifications, the log.txt file, and any information you think may be helpful."
    End

End Sub

Public Sub FadeOut()

Dim i As Integer

    'Gamma fade out!
    On Local Error GoTo DDFADEOUTERROR
    If pblnGamma Then
        For i = 0 To -100 Step -1
            SlowDown
            mintRedVal = i
            mintGreenVal = i
            mintBlueVal = i
            SetGamma mintRedVal, mintGreenVal, mintBlueVal
        Next
    End If

    'Exit before error code
    On Error GoTo 0
    Exit Sub
    
'Error handlers
DDFADEOUTERROR:
    TerminateSpecific False, True, True, True
    Log "DDraw", "FadeOut", "Error calling gamma fade out! (" & mintRedVal & ", " & mintGreenVal & ", " & mintBlueVal & ")"
    MsgBox "Error performing gamma fade out.  Please report this to moml31@hotmail.com along with your system specifications, the log.txt file, and any information you think may be helpful."
    End

End Sub

Public Sub SetGamma(intRed As Integer, intGreen As Integer, intBlue As Integer)

Dim i As Integer

    'Exit if user has no gamma!
    If pblnGamma = False Then Exit Sub

    'Alter the gamma ramp to the percent given by comparing to original state
    'A value of zero ("0") for intRed, intGreen, or intBlue will result in the
    'gamma level being set back to the original levels. Anything ABOVE zero will
    'fade towards FULL colour, anything below zero will fade towards NO colour
    On Local Error GoTo DDSETGAMMAERROR
    mintRedVal = intRed
    mintGreenVal = intGreen
    mintBlueVal = intBlue
    For i = 0 To 255
        If intRed < 0 Then mudtGammaRamp.red(i) = ConvToSignedValue(ConvToUnSignedValue(mudtOriginalRamp.red(i)) * (100 - Abs(intRed)) / 100)
        If intRed = 0 Then mudtGammaRamp.red(i) = mudtOriginalRamp.red(i)
        If intRed > 0 Then mudtGammaRamp.red(i) = ConvToSignedValue(65535 - ((65535 - ConvToUnSignedValue(mudtOriginalRamp.red(i))) * (100 - intRed) / 100))
        If intGreen < 0 Then mudtGammaRamp.green(i) = ConvToSignedValue(ConvToUnSignedValue(mudtOriginalRamp.green(i)) * (100 - Abs(intGreen)) / 100)
        If intGreen = 0 Then mudtGammaRamp.green(i) = mudtOriginalRamp.green(i)
        If intGreen > 0 Then mudtGammaRamp.green(i) = ConvToSignedValue(65535 - ((65535 - ConvToUnSignedValue(mudtOriginalRamp.green(i))) * (100 - intGreen) / 100))
        If intBlue < 0 Then mudtGammaRamp.blue(i) = ConvToSignedValue(ConvToUnSignedValue(mudtOriginalRamp.blue(i)) * (100 - Abs(intBlue)) / 100)
        If intBlue = 0 Then mudtGammaRamp.blue(i) = mudtOriginalRamp.blue(i)
        If intBlue > 0 Then mudtGammaRamp.blue(i) = ConvToSignedValue(65535 - ((65535 - ConvToUnSignedValue(mudtOriginalRamp.blue(i))) * (100 - intBlue) / 100))
    Next
    mobjGammaControler.SetGammaRamp DDSGR_DEFAULT, mudtGammaRamp

    'Exit before error code
    On Error GoTo 0
    Exit Sub
    
'Error handlers
DDSETGAMMAERROR:
    TerminateSpecific False, True, True, True
    Log "DDraw", "SetGamma", "Error setting gamma ramps! (" & mintRedVal & ", " & mintGreenVal & ", " & mintBlueVal & ")"
    MsgBox "Error setting gamma values.  Please report this to moml31@hotmail.com along with your system specifications, the log.txt file, and any information you think may be helpful."
    End

End Sub

Private Sub SlowDown()

Dim lngTickStore As Long

    'Delay the effect somewhat
    On Local Error GoTo DDSLOWDOWNERROR
    lngTickStore = gobjDX.TickCount()
    Do While lngTickStore + FADE_DELAY_MS > gobjDX.TickCount()
        DoEvents
    Loop

    'Exit before error code
    On Error GoTo 0
    Exit Sub
    
'Error handlers
DDSLOWDOWNERROR:
    TerminateSpecific False, True, True, True
    Log "DDraw", "SlowDown", "Error in slowdown function!"
    MsgBox "Error performing DirectDraw slowdown.  Please report this to moml31@hotmail.com along with your system specifications, the log.txt file, and any information you think may be helpful."
    End

End Sub

Private Function ExclusiveMode() As Boolean

Dim lngTestExMode As Long
    
    'This function tests if we're still in exclusive mode
    On Local Error GoTo DDEXCLUSIVEERROR
    lngTestExMode = mobjDD.TestCooperativeLevel
    
    If (lngTestExMode = DD_OK) Then
        ExclusiveMode = True
    Else
        ExclusiveMode = False
    End If
    
    'Exit before error code
    On Error GoTo 0
    Exit Function
    
'Error handlers
DDEXCLUSIVEERROR:
    TerminateSpecific False, True, True, True
    Log "DDraw", "ExclusiveMode", "Error testing for exclusivity!  TestCooperativeLevel: " & lngTestExMode
    MsgBox "Error testing exclusivity.  Please report this to moml31@hotmail.com along with your system specifications, the log.txt file, and any information you think may be helpful."
    End

End Function

Public Sub HandleLostSurfaces()

Dim blnLostSurfaces As Boolean
Dim i As Long

    'This sub will reload lost surfaces if necessary
    On Local Error GoTo DDHANDLELOSTERROR
    blnLostSurfaces = False
    Do Until ExclusiveMode
        DoEvents
        blnLostSurfaces = True
    Loop
    
    'If we did lose our bitmaps, restore the surfaces
    DoEvents
    If blnLostSurfaces Then
        mobjDD.RestoreAllSurfaces
        For i = 0 To UBound(mudtSurface)
            If mudtSurface(i).blnExists Then LoadSprite mudtSurface(i).strResName, mudtSurface(i).intWidth, mudtSurface(i).intHeight, 0, i
        Next i
        'Reset the general timer
        glngTimer = gobjDX.TickCount()
    End If

    'Exit before error code
    On Error GoTo 0
    Exit Sub
    
'Error handlers
DDHANDLELOSTERROR:
    TerminateSpecific False, True, True, True
    Log "DDraw", "HandleLostSurfaces", "Error handling lost surfaces!"
    MsgBox "Error handling lost surfaces.  Please report this to moml31@hotmail.com along with your system specifications, the log.txt file, and any information you think may be helpful."
    End

End Sub

Public Function LoadSprite(strResName As String, intWidth As Integer, intHeight As Integer, lngColorKey As Long, Optional lngSetNum As Long = -1) As Long

Dim ddckKey As DDCOLORKEY
Dim ddsdNewSprite As DDSURFACEDESC2
Dim lngSprite As Long
Dim lngTemp As Long
Dim blnFirstTry As Boolean
    
    'Find a new spot in the array
    On Local Error GoTo DDLOADSPRITEERROR
    If lngSetNum = -1 Then
        lngSprite = FreeSpot()
        LoadSprite = lngSprite
    Else
        lngSprite = lngSetNum
        LoadSprite = lngSetNum
    End If
    
    'Some people have problems loading large surfaces to vidmem
    blnFirstTry = True
TRYAGAIN:
    If blnFirstTry = True Then
        'Try loading in vidmem first
        On Local Error GoTo FIRSTTRY
        ddsdNewSprite.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
    Else
        'If vidmem fails, go to system
        On Local Error GoTo DDLOADSPRITEERROR
        ddsdNewSprite.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
    End If
        
    'Fill out the surface description
    ddsdNewSprite.lFlags = DDSD_CAPS Or DDSD_WIDTH Or DDSD_HEIGHT
    ddsdNewSprite.lWidth = intWidth
    ddsdNewSprite.lHeight = intHeight
    
    'Create the surface
    Set mudtSurface(lngSprite).surfSprite = mobjDD.CreateSurface(ddsdNewSprite)
    
    'Sets the Sprite's colour key
    ddckKey.low = lngColorKey
    ddckKey.high = lngColorKey
    mudtSurface(lngSprite).surfSprite.SetColorKey DDCKEY_SRCBLT, ddckKey
    
    'Find the sprite in the resource and extract the data
    ExtractData App.Path & "\resource.bnk", strResName
    
    'Blit to surface
    lngTemp = mudtSurface(lngSprite).surfSprite.GetDC
    StretchDIBits lngTemp, 0, 0, gudtBMPInfo.bmiHeader.biWidth, gudtBMPInfo.bmiHeader.biHeight, 0, 0, gudtBMPInfo.bmiHeader.biWidth, gudtBMPInfo.bmiHeader.biHeight, gudtBMPData(0), gudtBMPInfo, DIB_RGB_COLORS, vbSrcCopy
    mudtSurface(lngSprite).surfSprite.ReleaseDC lngTemp
    
    'Set the surface info
    mudtSurface(lngSprite).blnExists = True
    mudtSurface(lngSprite).blnReserved = False
    mudtSurface(lngSprite).intHeight = intHeight
    mudtSurface(lngSprite).intWidth = intWidth
    mudtSurface(lngSprite).strResName = strResName

    'Exit before error code
    On Error GoTo 0
    Exit Function
    
'Error handlers
FIRSTTRY:
    blnFirstTry = False
    GoTo TRYAGAIN
DDLOADSPRITEERROR:
    TerminateSpecific False, True, True, True
    Log "DDraw", "LoadSprite", "Error loading sprite! (" & strResName & ", " & intWidth & ", " & intHeight & ", " & lngColorKey & ", " & lngSetNum & ")"
    MsgBox "Error loading sprite surfaces.  Please report this to moml31@hotmail.com along with your system specifications, the log.txt file, and any information you think may be helpful."
    End

End Function

Private Function FreeSpot() As Long

Dim i As Long

    'Check for empty slots
    For i = 0 To UBound(mudtSurface)
        If mudtSurface(i).blnExists = False And Not (mudtSurface(i).blnReserved) Then
            FreeSpot = i
            Exit Function
        End If
    Next i
    
    'Make a new slot
    ReDim Preserve mudtSurface(UBound(mudtSurface) + 1)
    FreeSpot = UBound(mudtSurface)

End Function

Public Sub DeleteSprite(lngSprite As Long)

    'Kill the sprite
    On Local Error GoTo DDDELETESPRITEERROR
    mudtSurface(lngSprite).blnExists = False
    mudtSurface(lngSprite).blnReserved = False
    Set mudtSurface(lngSprite).surfSprite = Nothing

    'Exit before error code
    On Error GoTo 0
    Exit Sub
    
'Error handlers
DDDELETESPRITEERROR:
    TerminateSpecific False, True, True, True
    Log "DDraw", "DeleteSprite", "Error deleting sprite!  SpriteID: " & lngSprite
    MsgBox "Error deleting sprite.  Please report this to moml31@hotmail.com along with your system specifications, the log.txt file, and any information you think may be helpful."
    End

End Sub

Public Sub DisplaySprite(lngSprite As Long, intX As Integer, intY As Integer)

Dim rectDest As RECT
Dim rectSource As RECT

    'Ensure the sprite exists before displaying it!
    On Local Error GoTo DDDISPLAYSPRITEERROR
    If Not (mudtSurface(lngSprite).blnExists) Then Exit Sub

    'Calc the rects
    With rectDest
        .Left = intX
        .Right = intX + mudtSurface(lngSprite).intWidth
        .Top = intY
        .Bottom = intY + mudtSurface(lngSprite).intHeight
    End With
    IntersectRect rectSource, rectDest, mrectScreen
    IntersectRect rectDest, rectDest, mrectScreen
    With rectSource
        If .Left > 0 Then
            .Right = .Right - .Left
            .Left = 0
        End If
        If .Top > 0 Then
            .Bottom = .Bottom - .Top
            .Top = 0
        End If
    End With
    
    'Display
    msurfBack.BltFast rectDest.Left, rectDest.Top, mudtSurface(lngSprite).surfSprite, rectSource, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT

    'Exit before error code
    On Error GoTo 0
    Exit Sub
    
'Error handlers
DDDISPLAYSPRITEERROR:
    TerminateSpecific False, True, True, True
    Log "DDraw", "DisplaySprite", "Error displaying sprite! (" & lngSprite & ", " & intX & ", " & intY & ")"
    MsgBox "Error displaying sprite.  Please report this to moml31@hotmail.com along with your system specifications, the log.txt file, and any information you think may be helpful."
    End

End Sub

Public Sub DisplayBar(lngSprite As Long, intX As Integer, intY As Integer, intWidth As Integer)

Dim rectDest As RECT
Dim rectSource As RECT

    'If there's no width, don't bother!
    On Local Error GoTo DDDISPLAYBARERROR
    If intWidth < 1 Then Exit Sub

    'Ensure the sprite exists before displaying it!
    If Not (mudtSurface(lngSprite).blnExists) Then Exit Sub

    'Calc the rects
    With rectDest
        .Left = intX
        .Right = intX + intWidth
        .Top = intY
        .Bottom = intY + mudtSurface(lngSprite).intHeight
    End With
    IntersectRect rectSource, rectDest, mrectScreen
    IntersectRect rectDest, rectDest, mrectScreen
    With rectSource
        If .Left > 0 Then
            .Right = .Right - .Left
            .Left = 0
        End If
        If .Top > 0 Then
            .Bottom = .Bottom - .Top
            .Top = 0
        End If
    End With
    
    'Display
    msurfBack.BltFast rectDest.Left, rectDest.Top, mudtSurface(lngSprite).surfSprite, rectSource, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT

    'Exit before error code
    On Error GoTo 0
    Exit Sub
    
'Error handlers
DDDISPLAYBARERROR:
    TerminateSpecific False, True, True, True
    Log "DDraw", "DisplayBar", "Error displaying bar sprite! (" & lngSprite & ", " & intX & ", " & intY & ", " & intWidth & ")"
    MsgBox "Error displaying bar surface.  Please report this to moml31@hotmail.com along with your system specifications, the log.txt file, and any information you think may be helpful."
    End

End Sub

Public Sub DisplaySpriteClip(lngSpriteObject As Long, lngIndex As Long, intX As Integer, intY As Integer, Optional blnClip As Boolean = True, Optional blnSrcPaint As Boolean = False)

Dim rectDest As RECT
Dim rectSource As RECT
Dim objBltFx As DDBLTFX

    'Ensure the sprite exists before displaying it!
    On Local Error GoTo DDDISPLAYSPRITECLIPERROR
    If Not (mudtSurface(mudtSpriteObject(lngSpriteObject).lngIndex(lngIndex)).blnExists) Then Exit Sub

    'Calc the rects
    With rectDest
        .Left = intX
        .Right = intX + mudtSurface(mudtSpriteObject(lngSpriteObject).lngIndex(lngIndex)).intWidth
        .Top = intY
        .Bottom = intY + mudtSurface(mudtSpriteObject(lngSpriteObject).lngIndex(lngIndex)).intHeight
    End With
    'Clip?
    If blnClip = True Then
        IntersectRect rectSource, rectDest, mrectTactical
        IntersectRect rectDest, rectDest, mrectTactical
        With rectSource
            If .Left > 0 Then
                .Right = .Right - .Left
                .Left = 0
            End If
            If .Top > 0 Then
                .Bottom = .Bottom - .Top
                .Top = 0
            End If
            If intX < mrectTactical.Left Then
                .Left = .Left + (mrectTactical.Left - intX)
                .Right = .Right + (mrectTactical.Left - intX)
            End If
            If intY < mrectTactical.Top Then
                .Top = .Top + (mrectTactical.Top - intY)
                .Bottom = .Bottom + (mrectTactical.Top - intY)
            End If
        End With
    Else
        'No clip!
        rectSource.Left = 0
        rectSource.Right = mudtSurface(mudtSpriteObject(lngSpriteObject).lngIndex(lngIndex)).intWidth
        rectSource.Top = 0
        rectSource.Bottom = mudtSurface(mudtSpriteObject(lngSpriteObject).lngIndex(lngIndex)).intHeight
    End If
    
    'Blit! (Check for ROP caps!)
    If (blnSrcPaint = True) And (mblnSrcPaint = True) Then
        'Display SrcPaint
        objBltFx.lROP = vbSrcPaint
        msurfBack.BltFx rectDest, mudtSurface(mudtSpriteObject(lngSpriteObject).lngIndex(lngIndex)).surfSprite, rectSource, DDBLT_ROP Or DDBLT_WAIT, objBltFx
    Else
        'Display normal
        msurfBack.BltFast rectDest.Left, rectDest.Top, mudtSurface(mudtSpriteObject(lngSpriteObject).lngIndex(lngIndex)).surfSprite, rectSource, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
    End If

    'Exit before error code
    On Error GoTo 0
    Exit Sub
    
'Error handlers
DDDISPLAYSPRITECLIPERROR:
    TerminateSpecific False, True, True, True
    Log "DDraw", "DisplaySpriteClip", "Error displaying clipped sprite! (" & lngSpriteObject & ", " & lngIndex & ", " & intX & ", " & intY & ")"
    MsgBox "Error displaying clipped sprite.  Please report this to moml31@hotmail.com along with your system specifications, the log.txt file, and any information you think may be helpful."
    End

End Sub

Public Sub DisplayClip(lngIndex As Long, intX As Integer, intY As Integer, Optional blnClip As Boolean = True, Optional blnSrcPaint As Boolean = False)

Dim rectDest As RECT
Dim rectSource As RECT
Dim objBltFx As DDBLTFX

    'Ensure the sprite exists before displaying it!
    On Local Error GoTo DDDISPLAYCLIPERROR
    If Not (mudtSurface(lngIndex).blnExists) Then Exit Sub

    'Calc the rects
    With rectDest
        .Left = intX
        .Right = intX + mudtSurface(lngIndex).intWidth
        .Top = intY
        .Bottom = intY + mudtSurface(lngIndex).intHeight
    End With
    'Clip?
    If blnClip = True Then
        IntersectRect rectSource, rectDest, mrectTactical
        IntersectRect rectDest, rectDest, mrectTactical
        With rectSource
            If .Left > 0 Then
                .Right = .Right - .Left
                .Left = 0
            End If
            If .Top > 0 Then
                .Bottom = .Bottom - .Top
                .Top = 0
            End If
            If intX < mrectTactical.Left Then
                .Left = .Left + (mrectTactical.Left - intX)
                .Right = .Right + (mrectTactical.Left - intX)
            End If
            If intY < mrectTactical.Top Then
                .Top = .Top + (mrectTactical.Top - intY)
                .Bottom = .Bottom + (mrectTactical.Top - intY)
            End If
        End With
    Else
        'No clip!
        rectSource.Left = 0
        rectSource.Right = mudtSurface(lngIndex).intWidth
        rectSource.Top = 0
        rectSource.Bottom = mudtSurface(lngIndex).intHeight
    End If
    
    'Blit! (Check for ROP caps!)
    If (blnSrcPaint = True) And (mblnSrcPaint = True) Then
        'Display SrcPaint
        objBltFx.lROP = vbSrcPaint
        msurfBack.BltFx rectDest, mudtSurface(lngIndex).surfSprite, rectSource, DDBLT_ROP Or DDBLT_WAIT, objBltFx
    Else
        'Display normal
        msurfBack.BltFast rectDest.Left, rectDest.Top, mudtSurface(lngIndex).surfSprite, rectSource, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
    End If

    'Exit before error code
    On Error GoTo 0
    Exit Sub
    
'Error handlers
DDDISPLAYCLIPERROR:
    TerminateSpecific False, True, True, True
    Log "DDraw", "DisplayClip", "Error displaying clipped surface! (" & lngIndex & ", " & intX & ", " & intY & ")"
    MsgBox "Error displaying clipped surface.  Please report this to moml31@hotmail.com along with your system specifications, the log.txt file, and any information you think may be helpful."
    End

End Sub

Private Function TestROP(ByRef surfBack As DirectDrawSurface7, lngROP As Long) As Boolean

Dim objBltFx As DDBLTFX
Dim rectTemp As RECT
Dim surfTemp As DirectDrawSurface7
Dim udtDDSD As DDSURFACEDESC2

    'Create a small temporary surface
    On Local Error GoTo DDTESTROPERROR
    udtDDSD.lFlags = DDSD_HEIGHT Or DDSD_WIDTH
    udtDDSD.lHeight = 1
    udtDDSD.lWidth = 1
    Set surfTemp = mobjDD.CreateSurface(udtDDSD)
    
    'Set the BltFx ROP code
    objBltFx.lROP = lngROP
    
    'Our source and dest rectangle
    rectTemp.Right = 1
    rectTemp.Bottom = 1
    
    'Test the BltFx capability
    If surfBack.BltFx(rectTemp, surfTemp, rectTemp, DDBLT_ROP Or DDBLT_WAIT, objBltFx) <> 0 Then
        TestROP = False
    Else
        TestROP = True
    End If

    'Exit before error code
    On Error GoTo 0
    Exit Function
    
'Error handlers
DDTESTROPERROR:
    TerminateSpecific False, True, True, True
    Log "DDraw", "TestROP", "Error testing ROP capabilities! (" & lngROP & ")"
    MsgBox "Error testing ROP " & CStr(lngROP) & ".  Please report this to moml31@hotmail.com along with your system specifications, the log.txt file, and any information you think may be helpful."
    End

End Function

Public Function Queue(strResName As String, intWidth As Integer, intHeight As Integer) As Long

Dim lngSprite As Long

    'Find a spot
    On Local Error GoTo DDQUEUEERROR
    lngSprite = FreeSpot()
    Queue = lngSprite
    
    'Store it's data
    mudtSurface(lngSprite).blnReserved = True
    mudtSurface(lngSprite).intHeight = intHeight
    mudtSurface(lngSprite).intWidth = intWidth
    mudtSurface(lngSprite).strResName = strResName
    
    'Add it to the queue
    If mlngQueue(0) = -1 Then
        mlngQueue(0) = lngSprite
    Else
        ReDim Preserve mlngQueue(UBound(mlngQueue) + 1)
        mlngQueue(UBound(mlngQueue)) = lngSprite
    End If
   
    'Exit before error code
    On Error GoTo 0
    Exit Function
    
'Error handlers
DDQUEUEERROR:
    TerminateSpecific False, True, True, True
    Log "DDraw", "Queue", "Error queuing sprite! (" & strResName & ", " & intWidth & ", " & intHeight & ")"
    MsgBox "Error queuing sprite.  Please report this to moml31@hotmail.com along with your system specifications, the log.txt file, and any information you think may be helpful."
    End
   
End Function

Public Sub ProcessQueue()

Dim i As Long

    'Exit if there's nothing in the queue
    On Local Error GoTo DDPROCESSQUEUEERROR
    If mlngQueue(0) = -1 Then Exit Sub
    
    'Process the next sprite in the queue
    With mudtSurface(mlngQueue(0))
        LoadSprite .strResName, .intWidth, .intHeight, 0, mlngQueue(0)
    End With
    
    'Shorten the queue
    If UBound(mlngQueue) = 0 Then
        mlngQueue(0) = -1
    Else
        For i = 0 To UBound(mlngQueue) - 1
            mlngQueue(i) = mlngQueue(i + 1)
        Next i
        ReDim Preserve mlngQueue(UBound(mlngQueue) - 1)
    End If

    'Exit before error code
    On Error GoTo 0
    Exit Sub
    
'Error handlers
DDPROCESSQUEUEERROR:
    TerminateSpecific False, True, True, True
    Log "DDraw", "ProcessQueue", "Error processing queue! (" & i & ")"
    MsgBox "Error processing queue.  Please report this to moml31@hotmail.com along with your system specifications, the log.txt file, and any information you think may be helpful."
    End

End Sub

Public Function LoadSpriteObject(strResName As String, intWidth As Integer, intHeight As Integer, bytFrameAmt As Byte, bytAnimAmt As Byte, blnLoadImmediate As Boolean, Optional blnLoadShipPic As Boolean = True) As Long

Dim i As Long
Dim j As Long
Dim lngFreeSpot As Long

    'Check if this resource name is already in the array
    On Local Error GoTo DDLOADSPRITEOBJECTERROR
    lngFreeSpot = -1
    For i = 0 To UBound(mudtSpriteObject)
        If mudtSpriteObject(i).blnExists = True And mudtSpriteObject(i).strResName = strResName Then
            lngFreeSpot = i
            Exit For
        End If
    Next i
    'Find a spot in the Sprite Object array (if a pre-existing spot hasn't been found)
    If lngFreeSpot = -1 Then
        For i = 0 To UBound(mudtSpriteObject)
            If mudtSpriteObject(i).blnExists = False Then
                lngFreeSpot = i
                Exit For
            End If
        Next i
    End If
    'Make a new spot if there aren't any free
    If lngFreeSpot = -1 Then
        ReDim Preserve mudtSpriteObject(UBound(mudtSpriteObject) + 1)
        lngFreeSpot = UBound(mudtSpriteObject)
    End If
    
    'Load the index array if we haven't already..
    If mudtSpriteObject(lngFreeSpot).intUses < 1 Then
        'Ship pic?
        If blnLoadShipPic = True Then
            ReDim mudtSpriteObject(lngFreeSpot).lngIndex((bytAnimAmt + 1) * (bytFrameAmt + 1))
        Else
            ReDim mudtSpriteObject(lngFreeSpot).lngIndex((bytAnimAmt + 1) * (bytFrameAmt + 1) - 1)
        End If
        For i = 0 To bytAnimAmt
            For j = 0 To bytFrameAmt
                'If it's immediate..
                If blnLoadImmediate Then
                    mudtSpriteObject(lngFreeSpot).lngIndex((i * (bytFrameAmt + 1)) + j) = DDraw.LoadSprite(strResName & i & j, intWidth, intHeight, 0)
                Else
                    mudtSpriteObject(lngFreeSpot).lngIndex((i * (bytFrameAmt + 1)) + j) = DDraw.Queue(strResName & i & j, intWidth, intHeight)
                End If
            Next j
        Next i
        'Load the shippic
        If blnLoadShipPic = True Then
            If blnLoadImmediate Then
                mudtSpriteObject(lngFreeSpot).lngIndex(UBound(mudtSpriteObject(lngFreeSpot).lngIndex)) = DDraw.LoadSprite(strResName, SHIPPIC_WIDTH, SHIPPIC_HEIGHT, 0)
            Else
                mudtSpriteObject(lngFreeSpot).lngIndex(UBound(mudtSpriteObject(lngFreeSpot).lngIndex)) = DDraw.Queue(strResName, SHIPPIC_WIDTH, SHIPPIC_HEIGHT)
            End If
        End If
    End If
    
    'Load remaining fields
    mudtSpriteObject(lngFreeSpot).blnExists = True
    mudtSpriteObject(lngFreeSpot).intUses = mudtSpriteObject(lngFreeSpot).intUses + 1
    mudtSpriteObject(lngFreeSpot).strResName = strResName
    
    'Return index
    LoadSpriteObject = lngFreeSpot

    'Exit before error code
    On Error GoTo 0
    Exit Function
    
'Error handlers
DDLOADSPRITEOBJECTERROR:
    TerminateSpecific False, True, True, True
    Log "DDraw", "LoadSpriteObject", "Error loading sprite object! (" & strResName & ", " & intWidth & ", " & intHeight & ", " & bytFrameAmt & ", " & bytAnimAmt & ", " & blnLoadImmediate & ", " & blnLoadShipPic & ")"
    MsgBox "Error loading sprite object.  Please report this to moml31@hotmail.com along with your system specifications, the log.txt file, and any information you think may be helpful."
    End

End Function

Public Sub DeleteSpriteObject(lngIndex As Long)

Dim i As Integer

    'Decrement the intUses
    On Local Error GoTo DDDELETESPRITEOBJECTERROR
    mudtSpriteObject(lngIndex).intUses = mudtSpriteObject(lngIndex).intUses - 1
    
    'Are we down to zero uses?
    If mudtSpriteObject(lngIndex).intUses < 1 Then
        'Kill this entry
        mudtSpriteObject(lngIndex).blnExists = False
        mudtSpriteObject(lngIndex).strResName = ""
        mudtSpriteObject(lngIndex).intUses = 0
        For i = 0 To UBound(mudtSpriteObject(lngIndex).lngIndex)
            DeleteSprite mudtSpriteObject(lngIndex).lngIndex(i)
        Next i
    End If

    'Exit before error code
    On Error GoTo 0
    Exit Sub
    
'Error handlers
DDDELETESPRITEOBJECTERROR:
    TerminateSpecific False, True, True, True
    Log "DDraw", "DeleteSpriteObject", "Error deleting sprite object! (" & lngIndex & ")"
    MsgBox "Error deleting sprite object.  Please report this to moml31@hotmail.com along with your system specifications, the log.txt file, and any information you think may be helpful."
    End

End Sub

Public Function GetBufferDC() As Long

    'Locks the backbuffer and gets the device context handle
    GetBufferDC = msurfBack.GetDC

End Function

Public Sub ReleaseBufferDC(lngDC As Long)

    'Releases the backbuffer lock and device context handle
    msurfBack.ReleaseDC lngDC
    
End Sub

Public Sub DeleteAllSurfaces()

Dim i As Long

    'Kill 'em all!
    On Local Error GoTo DDDELETEALLSURFACESERROR
    For i = 0 To UBound(mudtSurface)
        Set mudtSurface(i).surfSprite = Nothing
    Next i

    'Exit before error code
    On Error GoTo 0
    Exit Sub
    
'Error handlers
DDDELETEALLSURFACESERROR:
    DDraw.RestoreDisplay
    Log "DDraw", "DeleteAllSurfaces", "Error destroying all surfaces! (" & i & ")"
    MsgBox "Error destroying all surfaces.  Please report this to moml31@hotmail.com along with your system specifications, the log.txt file, and any information you think may be helpful."
    End

End Sub

Public Sub RestoreDisplay()

    'Restore the display mode..
    Call mobjDD.RestoreDisplayMode
    Call mobjDD.SetCooperativeLevel(frmMain.hWnd, DDSCL_NORMAL)
    Set mobjDD = Nothing

End Sub

Public Sub Terminate(frmTerm As Form)

    'Terminate DirectDraw
    On Local Error GoTo DDTERMERROR
    DeleteAllSurfaces
    Set mobjGammaControler = Nothing
    Set msurfBack = Nothing
    Set msurfFront = Nothing
    Call mobjDD.RestoreDisplayMode
    Call mobjDD.SetCooperativeLevel(frmTerm.hWnd, DDSCL_NORMAL)
    Set mobjDD = Nothing

    'Exit before error code
    On Error GoTo 0
    Exit Sub
    
'Error handlers
DDTERMERROR:
    TerminateSpecific False, False, True, True
    Log "DDraw", "Terminate", "Error terminating DDraw!"
    MsgBox "Error terminating DirectDraw.  Please report this to moml31@hotmail.com along with your system specifications, the log.txt file, and any information you think may be helpful."
    End

End Sub
