Attribute VB_Name = "Title"
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

Dim mblnLoaded As Boolean       'Has this screen been loaded yet?
Dim mblnTerminating As Boolean  'Are we terminating this screen?
Dim mbytNextScreen As Byte      'What screen should be loaded upon this screen's termination?

Dim mlngSpriteTitle As Long     'Reference for title screen sprite
Dim mlngSpriteHighlight As Long 'Reference for menu item highlight sprite
Dim mlngSpriteCursor As Long    'Reference for mouse cursor
Dim mlngSpriteLoading As Long   'Reference for loading bitmap

'Mouse cursor constants
Const CURSOR_WIDTH = 10
Const CURSOR_HEIGHT = 10

'Loading screen constants
Const LOAD_WIDTH = 800
Const LOAD_HEIGHT = 100

'Menu item location constants
Const MENU_X = 606
Const MENU_XDELTA = 160
Const MENU_Y = 30
Const MENU_YDELTA = 30
Const MENU_YPAD = 5
Const MENU_NUM = 4

Public Sub Main()

Dim i As Byte
Static intSound As Integer  'Is a onmouseover sound currently being heard?

    'Check if we're terminating
    If mblnTerminating Then Terminate

    'Check if we're the screen that's supposed to be currently displayed
    If gbytDisplay <> DISPLAY_TITLE Then Exit Sub

    'If not yet loaded, load!
    If Not (mblnLoaded) Then Initialize
    
    'Display our title bitmap
    DDraw.DisplaySprite mlngSpriteTitle, 0, 0
    
    'Display our mouse cursor
    DDraw.DisplaySprite mlngSpriteCursor, gintMouseX, gintMouseY
    
    'Display menu highlights, if neccessary
    For i = 0 To MENU_NUM
        If MenuItem(i) Then
            'Display the highlight
            DisplayHighlight i
            'Make sound (ensure no repeating!)
            If intSound <> i Then
                DSound.Beep1
                intSound = i
            End If
        'If menu item is no longer highlighted, reset sound variable
        ElseIf intSound = i Then
            intSound = -1
        End If
    Next i
        
    'Handle clicks
    If MenuItem(2) And gblnLMouseButtonUp Then
        DSound.Beep2                                                    'Make the beep sound
        DDraw.SetFade FADE_OUT_GAMMA, FADE_OUT_GAMMA, FADE_OUT_GAMMA    'Fade out ddraw
        gstrUniverse = DEFAULT_UNIVERSE                                 'Load the default universe
        mbytNextScreen = DISPLAY_TACTICAL                               'Display the tactical screen
        mblnTerminating = True                                          'Commence termination of this screen
    End If
    If MenuItem(4) And gblnLMouseButtonUp Then
        DSound.Beep2                                                    'Make the beep sound
        DDraw.SetFade FADE_OUT_GAMMA, FADE_OUT_GAMMA, FADE_OUT_GAMMA    'Fade out ddraw
        DSound.SetFade FADE_OUT_MUSIC                                   'Fade out music
        mbytNextScreen = DISPLAY_NONE                                   'Display nothing next!
        mblnTerminating = True                                          'Commence termination of this screen
    End If

End Sub

Private Function MenuItem(bytIndex As Byte) As Boolean

    'Determine if the mouse is within the menu item in question
    MenuItem = False
    If (gintMouseX >= MENU_X) And (gintMouseX <= MENU_X + MENU_XDELTA) And (gintMouseY >= MENU_Y + (MENU_YDELTA + MENU_YPAD) * bytIndex) And (gintMouseY <= MENU_Y + (MENU_YDELTA + MENU_YPAD) * (bytIndex + 1) - MENU_YPAD) Then MenuItem = True

End Function

Private Sub DisplayHighlight(bytIndex As Byte)

    'Display the highlight for the given menu item
    DDraw.DisplaySprite mlngSpriteHighlight, MENU_X, MENU_Y + (MENU_YDELTA + MENU_YPAD) * bytIndex

End Sub

Private Sub Initialize()

    'log
    Log "Title", "Initialize", "Initializing title screen display"

    'Display loading bitmap
    DDraw.SetGamma -100, -100, -100
    mlngSpriteLoading = DDraw.LoadSprite("Loading", LOAD_WIDTH, LOAD_HEIGHT, 0)
    DDraw.DisplaySprite mlngSpriteLoading, 0, (SCREEN_HEIGHT \ 2) - (LOAD_HEIGHT \ 2)
    DDraw.Flip
    DDraw.FadeIn

    'Load the music and start it playing
    DSound.LoadMusic "intro"
    DSound.plngFadeValue = 0
    DSound.PlayMusic
    
    'Load our sprites
    mlngSpriteTitle = DDraw.LoadSprite("Title", SCREEN_WIDTH, SCREEN_HEIGHT, 0)
    mlngSpriteHighlight = DDraw.LoadSprite("Highlight", MENU_XDELTA, MENU_YDELTA, 0)
    mlngSpriteCursor = DDraw.LoadSprite("Cursor", CURSOR_WIDTH, CURSOR_HEIGHT, 0)
    
    'Fade out and slowly in
    DDraw.FadeOut
    DDraw.SetFade 0, 0, 0
    
    'Set as loaded
    mblnLoaded = True
    mblnTerminating = False

End Sub

Private Sub Terminate()

    'Ensure termination go-ahead
    If mblnTerminating = False Or (gblnFadeComplete = False And DDraw.pblnGamma = True) Or (gblnMusicFadeComplete = False And gblnMusic = True And gblnSound = True) Then Exit Sub

    'log
    Log "Title", "Terminate", "Terminating title screen display"
    
    'Delete our sprites
    DDraw.DeleteSprite mlngSpriteTitle
    DDraw.DeleteSprite mlngSpriteHighlight
    DDraw.DeleteSprite mlngSpriteCursor
    DDraw.DeleteSprite mlngSpriteLoading

    'Set as unloaded
    mblnLoaded = False
    mblnTerminating = False
    
    'Load next screen
    gbytDisplay = mbytNextScreen
    
    'If there's no screen to load next, then exit program
    If mbytNextScreen = DISPLAY_NONE Then gblnRunning = False
    
End Sub
