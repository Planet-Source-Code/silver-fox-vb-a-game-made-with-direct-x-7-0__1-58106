Attribute VB_Name = "Tactical"
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

'Loading screen constants
Const LOAD_WIDTH = 800
Const LOAD_HEIGHT = 100

'Our sprite references
Dim mlngSpriteTacticalLeft As Long
Dim mlngSpriteTacticalRight As Long
Dim mlngSpriteTacticalBottom As Long
Dim mlngSpriteTacticalTop As Long
Dim mlngSpriteStar As Long
Dim mlngSpriteGradientBar As Long
Dim mlngSpriteSolidBar As Long
Dim mlngSpriteSlider As Long
Dim mlngSpriteMarker As Long

'Smoke sprites
Private Const SMOKE_SPRITE_NUM = 5
Private Const SMOKE_HEIGHT = 10
Private Const SMOKE_WIDTH = 10
Dim mlngSpriteSmoke(SMOKE_SPRITE_NUM - 1) As Long

'Shield sprites
Private Const SHIELD_SPRITE_NUM = 7
Dim mlngSpriteShield(SHIELD_SPRITE_NUM - 1) As Long

'FTL effect sprites
Private Const FTL_EFFECT_SPRITE_NUM = 6
Private Const FTL_EFFECT_DISTANCE = 10
Private Const FTL_EFFECT_RAND = 1

'Radar dot sprites
Private Const RADAR_DOT_NUM = 7
Private Type RADARSPRITETYPE
    lngUnknown(RADAR_DOT_NUM) As Long
    lngEnemy(RADAR_DOT_NUM) As Long
    lngNeutral(RADAR_DOT_NUM) As Long
    lngFriendly(RADAR_DOT_NUM) As Long
    lngPlanet(RADAR_DOT_NUM) As Long
    lngWeapon As Long
End Type
Dim mudtSpriteRadar As RADARSPRITETYPE

'Star array
Private Type STARTYPE
    dblX As Double
    dblY As Double
    sngRelSpeed As Single
End Type
Dim mudtStar() As STARTYPE
Const STAR_NUM = 30             'How many stars?
Const STAR_SPEEDS = 2           'How many different speeds? (zero based)

'Comm messages
Private Type COMMTYPE
    strMessage As String        'What are the messages?
    lngDecay As Long            'At what clock-tick should the message decay?
End Type
Dim mudtComm() As COMMTYPE      'Our comm array
Dim mintNumComm As Integer      'How many messages are there?
Const COMM_DECAY = 10000        'Number of MS each message lives
Const COMM_MAX = 28             'Maximum number of messages
Const COMM_X = 20               'At what X value do we display these messages?
Const COMM_Y = 470              'At what Y value?
Const COMM_Y_HEIGHT = 17        'What's the Y delta?

'Textout and linedraw variables
Const TEXT_PLAYER_HEIGHT = 17
Const TEXT_PLAYER_ROW1_WIDTH = 73
Const TEXT_PLAYER_ROW2_WIDTH = 82
Const TEXT_PLAYER_ROW3_WIDTH = 62
Const TEXT_PLAYER_Y = 503
Const TEXT_PLAYER_Y_BARHEIGHT = 6
Const TEXT_PLAYER_Y_SLIDERHEIGHT = 2
Const TEXT_PLAYER_Y_MARKERHEIGHT = 3
Const TEXT_PLAYER_ROW1X = 20
Const TEXT_PLAYER_ROW2X = 237
Const TEXT_PLAYER_ROW3X = 445
Const TEXT_PLAYER_DELTA_MIS = 28
Const TEXT_PLAYER_DELTA_MINE = 85
Const TEXT_PLAYER_DELTA_AP = 111
Const TEXT_PLAYER_DELTA_LD = 136
Const TEXT_PLAYER_DELTA_JAM = 162
Const TEXT_RADAR_Y_RANGE = 161
Const TEXT_RADAR_Y_NAME = 211
Const TEXT_RADAR_Y_INFO = 250
Const TEXT_RADAR_Y_CONDITION = 323
Const TEXT_RADAR_X = 667
Const TEXT_RADAR_X_RANGE = 739
Const TEXT_RADAR_X_INFO = 710
Const TEXT_RADAR_X_CONDITION = 696
Const SHIPPIC_X = 673
Const SHIPPIC_Y = 372
Dim mlngDC As Long

'Tactical viewport size
Const TACTICAL_VIEW_WIDTH = 644
Const TACTICAL_VIEW_HEIGHT = 484

Public Sub Main()

    'Check if we're terminating
    If mblnTerminating Then Terminate

    'Check if we're the screen that's supposed to be currently displayed
    If gbytDisplay <> DISPLAY_TACTICAL Then Exit Sub

    'If not yet loaded, load!
    If Not (mblnLoaded) Then Initialize
    
    'Get input
    GetInput
    
    'Physics
    Universe.Physics
    
    'Radar
    UpdateRadar
    
    'Display
    If Not (gudtPlayer.udtSystems.blnARCDActive = True Or gudtPlayer.udtSystems.blnFTLDActive = True) Then
        'If we're going FTL, don't display stuff :)
        DisplayStars
        DisplayLasers
        DisplayBullets
        DisplayMissiles
        DisplayObjects
    End If
    DisplayPlayer
    DisplayExplosions
    DisplayRadar
            
    'Display lines
    mlngDC = DDraw.GetBufferDC
    DisplayRadarLines
    DDraw.ReleaseBufferDC mlngDC
        
    'Display the tactical frame
    DisplayFrame
    DisplayBars
    
    'Display text
    mlngDC = DDraw.GetBufferDC
    DisplayText
    DDraw.ReleaseBufferDC mlngDC
    
    'Display shippic
    If gudtPlayer.lngRadarObject >= 0 Then DDraw.DisplaySpriteClip gudtObject(gudtRadar(gudtPlayer.lngRadarObject).lngObject).udtSprite.lngSpriteObject, (gudtObject(gudtRadar(gudtPlayer.lngRadarObject).lngObject).udtSprite.bytAnimAmt + 1) * (gudtObject(gudtRadar(gudtPlayer.lngRadarObject).lngObject).udtSprite.bytFrameAmt + 1), SHIPPIC_X, SHIPPIC_Y, False
                
End Sub

Private Sub DisplayBullets()

Dim i As Long

    'Check for bullets
    If glngNumBullets <= 0 Then Exit Sub
    
    'Display 'em if they're close enough
    For i = 0 To glngNumBullets - 1
        'Check distance
        If GetDist(gudtPlayer.udtPhysics.dblX, gudtPlayer.udtPhysics.dblY, gudtBullet(i).dblX, gudtBullet(i).dblY) <= 400 Then
            'Display!
            DDraw.DisplayClip gudtCannon(gudtBullet(i).bytCannon).lngSprite, 330 - (gudtPlayer.udtPhysics.dblX - gudtBullet(i).dblX) - BULLET_WIDTH \ 2, 250 - (gudtPlayer.udtPhysics.dblY - gudtBullet(i).dblY) - BULLET_HEIGHT \ 2
        End If
    Next i

End Sub

Private Sub DisplayMissiles()

Dim i As Long
Dim j As Long
Dim sngRelX As Single
Dim sngRelY As Single
Dim sngSmokeRelX As Single
Dim sngSmokeRelY As Single
Dim dblDist As Double
    
    'Check for missiles
    If glngNumLiveMissiles <= 0 Then Exit Sub
    
    'Display 'em if they're close enough
    For i = 0 To glngNumLiveMissiles - 1
        'Store dist
        dblDist = GetDist(gudtPlayer.udtPhysics.dblX, gudtPlayer.udtPhysics.dblY, gudtLiveMissile(i).dblX, gudtLiveMissile(i).dblY)
        'Check distance
        If dblDist <= 400 Then
            'If it's within visible range..
            sngRelX = gudtLiveMissile(i).dblX - gudtPlayer.udtPhysics.dblX
            sngRelY = gudtLiveMissile(i).dblY - gudtPlayer.udtPhysics.dblY
            If sngRelX > -321 - gudtMissile(gudtLiveMissile(i).bytMissile).udtSprite.intWidth \ 2 And sngRelX < 321 + gudtMissile(gudtLiveMissile(i).bytMissile).udtSprite.intWidth \ 2 And sngRelY > -241 - gudtMissile(gudtLiveMissile(i).bytMissile).udtSprite.intHeight \ 2 And sngRelY < 241 + gudtMissile(gudtLiveMissile(i).bytMissile).udtSprite.intHeight \ 2 Then
                'Display!
                DDraw.DisplaySpriteClip gudtLiveMissile(i).udtSprite.lngSpriteObject, (gudtLiveMissile(i).udtSprite.bytAnimNum * (gudtMissile(gudtLiveMissile(i).bytMissile).udtSprite.bytFrameAmt + 1)) + gudtLiveMissile(i).udtSprite.bytFrameNum, sngRelX + 330 - gudtMissile(gudtLiveMissile(i).bytMissile).udtSprite.intWidth \ 2, sngRelY + 250 - gudtMissile(gudtLiveMissile(i).bytMissile).udtSprite.intHeight \ 2
            End If
        End If
        'Check smoke distance
        If dblDist <= 800 Then
            'Display smoke
            For j = 0 To SMOKE_SPRITE_NUM - 1
                sngSmokeRelX = gudtLiveMissile(i).dblXPrev(j) - gudtPlayer.udtPhysics.dblX
                sngSmokeRelY = gudtLiveMissile(i).dblYPrev(j) - gudtPlayer.udtPhysics.dblY
                If gudtLiveMissile(i).dblXPrev(j) <> 0 Then DDraw.DisplayClip mlngSpriteSmoke(j), sngSmokeRelX + 330 - SMOKE_WIDTH \ 2, sngSmokeRelY + 250 - SMOKE_HEIGHT \ 2
            Next j
        End If
    Next i

End Sub

Private Sub DisplayLasers()

Dim lngDC As Long
Dim i As Long
Dim blnDisplay As Boolean
Dim intX1 As Integer
Dim intY1 As Integer
Dim intX2 As Integer
Dim intY2 As Integer

    'Check for lasers
    If glngNumLaserDisplay = 0 Then Exit Sub
    
    'Get the DC
    lngDC = DDraw.GetBufferDC
    
    'Display 'em if they're close enough
    For i = 0 To glngNumLaserDisplay - 1
        'If either are the player, then do it
        blnDisplay = False
        If (gudtLaserDisplay(i).lngOwner = -1) Or (gudtLaserDisplay(i).lngTarget = -1) Then
            blnDisplay = True
        'Otherwise, check distances
        ElseIf GetDist(gudtObject(gudtLaserDisplay(i).lngOwner).udtPhysics.dblX, gudtObject(gudtLaserDisplay(i).lngOwner).udtPhysics.dblY, gudtPlayer.udtPhysics.dblX, gudtPlayer.udtPhysics.dblY) <= LASER_DISPLAY_DIST Then
            blnDisplay = True
        ElseIf GetDist(gudtObject(gudtLaserDisplay(i).lngTarget).udtPhysics.dblX, gudtObject(gudtLaserDisplay(i).lngTarget).udtPhysics.dblY, gudtPlayer.udtPhysics.dblX, gudtPlayer.udtPhysics.dblY) <= LASER_DISPLAY_DIST Then
            blnDisplay = True
        End If
        'If we are to display this one..
        If blnDisplay = True Then
            'Find the screen coords of the two objects
            If gudtLaserDisplay(i).lngOwner = -1 Then
                intX1 = 330
                intY1 = 250
            Else
                intX1 = gudtObject(gudtLaserDisplay(i).lngOwner).udtPhysics.dblX - gudtPlayer.udtPhysics.dblX + 330
                intY1 = gudtObject(gudtLaserDisplay(i).lngOwner).udtPhysics.dblY - gudtPlayer.udtPhysics.dblY + 250
            End If
            If gudtLaserDisplay(i).lngTarget = -1 Then
                intX2 = 330
                intY2 = 250
            Else
                intX2 = gudtObject(gudtLaserDisplay(i).lngTarget).udtPhysics.dblX - gudtPlayer.udtPhysics.dblX + 330
                intY2 = gudtObject(gudtLaserDisplay(i).lngTarget).udtPhysics.dblY - gudtPlayer.udtPhysics.dblY + 250
            End If
            'Clip
            If intX1 < TACTICAL_LEFT - LASER_WIDTH Then intX1 = TACTICAL_LEFT - LASER_WIDTH
            If intX1 > TACTICAL_RIGHT + LASER_WIDTH Then intX1 = TACTICAL_RIGHT + LASER_WIDTH
            If intY1 < TACTICAL_TOP - LASER_WIDTH Then intY1 = TACTICAL_TOP - LASER_WIDTH
            If intY1 > TACTICAL_BOTTOM + LASER_WIDTH Then intY1 = TACTICAL_BOTTOM + LASER_WIDTH
            If intX2 < TACTICAL_LEFT - LASER_WIDTH Then intX2 = TACTICAL_LEFT - LASER_WIDTH
            If intX2 > TACTICAL_RIGHT + LASER_WIDTH Then intX2 = TACTICAL_RIGHT + LASER_WIDTH
            If intY2 < TACTICAL_TOP - LASER_WIDTH Then intY2 = TACTICAL_TOP - LASER_WIDTH
            If intY2 > TACTICAL_BOTTOM + LASER_WIDTH Then intY2 = TACTICAL_BOTTOM + LASER_WIDTH
            'Display it!
            SetPen lngDC, LASER_WIDTH, gudtLaser(gudtLaserDisplay(i).bytType).lngColour
            LineDraw CLng(intX1), CLng(intY1), CLng(intX2), CLng(intY2), lngDC
            RemovePen lngDC
        End If
    Next i

    'Release the DC
    DDraw.ReleaseBufferDC lngDC

    'Erase the laser array
    Erase gudtLaserDisplay
    glngNumLaserDisplay = 0

End Sub

Private Sub DisplayRadarText()

Dim sngTemp As Single

    'Scanner range
    'ShowText "Scan Range:", TEXT_RADAR_X, TEXT_RADAR_Y_RANGE, vbGreen, mlngDC
    ShowText NormalizeDistance(CSng(gudtPlayer.dblCurrentRange)), TEXT_RADAR_X_RANGE, TEXT_RADAR_Y_RANGE, vbGreen, mlngDC

    'Exit if no object selected
    If gudtPlayer.lngRadarObject = -1 Then Exit Sub
    
    'Static radar text
    'ShowText "Name:", TEXT_RADAR_X, TEXT_RADAR_Y_NAME, vbGreen, mlngDC
    'ShowText "Race:", TEXT_RADAR_X, TEXT_RADAR_Y_NAME + TEXT_PLAYER_HEIGHT, vbGreen, mlngDC
    'ShowText "Mass:", TEXT_RADAR_X, TEXT_RADAR_Y_INFO, vbGreen, mlngDC
    'ShowText "Speed:", TEXT_RADAR_X, TEXT_RADAR_Y_INFO + TEXT_PLAYER_HEIGHT, vbGreen, mlngDC
    'ShowText "Dir:", TEXT_RADAR_X, TEXT_RADAR_Y_INFO + TEXT_PLAYER_HEIGHT * 2, vbGreen, mlngDC
    'ShowText "Dist:", TEXT_RADAR_X, TEXT_RADAR_Y_INFO + TEXT_PLAYER_HEIGHT * 3, vbGreen, mlngDC
    'ShowText "Condition", TEXT_RADAR_X_CONDITION, TEXT_RADAR_Y_CONDITION, vbGreen, mlngDC
    
    'Display radar textout
    ShowText gudtObject(gudtRadar(gudtPlayer.lngRadarObject).lngObject).udtInfo.strName, TEXT_RADAR_X_INFO, TEXT_RADAR_Y_NAME, vbGreen, mlngDC
    ShowText RaceName(gudtObject(gudtRadar(gudtPlayer.lngRadarObject).lngObject).udtInfo.bytRace), TEXT_RADAR_X_INFO, TEXT_RADAR_Y_NAME + TEXT_PLAYER_HEIGHT, vbGreen, mlngDC
    ShowText CStr(gudtObject(gudtRadar(gudtPlayer.lngRadarObject).lngObject).udtPhysics.lngMass) & " tons", TEXT_RADAR_X_INFO, TEXT_RADAR_Y_INFO, vbGreen, mlngDC
    ShowText NormalizeSpeed(gudtObject(gudtRadar(gudtPlayer.lngRadarObject).lngObject).udtPhysics.sngSpeed), TEXT_RADAR_X_INFO, TEXT_RADAR_Y_INFO + TEXT_PLAYER_HEIGHT, vbGreen, mlngDC
    ShowText CStr(Fix((180 / Pi) * FixAngle(FindAngle(gudtPlayer.udtPhysics.dblX, gudtPlayer.udtPhysics.dblY, gudtObject(gudtRadar(gudtPlayer.lngRadarObject).lngObject).udtPhysics.dblX, gudtObject(gudtRadar(gudtPlayer.lngRadarObject).lngObject).udtPhysics.dblY)))) & " degrees", TEXT_RADAR_X_INFO, TEXT_RADAR_Y_INFO + TEXT_PLAYER_HEIGHT * 2, vbGreen, mlngDC
    ShowText NormalizeDistance(GetDist(gudtPlayer.udtPhysics.dblX, gudtPlayer.udtPhysics.dblY, gudtObject(gudtRadar(gudtPlayer.lngRadarObject).lngObject).udtPhysics.dblX, gudtObject(gudtRadar(gudtPlayer.lngRadarObject).lngObject).udtPhysics.dblY)), TEXT_RADAR_X_INFO, TEXT_RADAR_Y_INFO + TEXT_PLAYER_HEIGHT * 3, vbGreen, mlngDC
    'Condition
    If gudtArmour(gudtObject(gudtRadar(gudtPlayer.lngRadarObject).lngObject).udtSystems.bytArmour).lngMaxArmour > 0 Then
        sngTemp = gudtObject(gudtRadar(gudtPlayer.lngRadarObject).lngObject).udtSystems.lngArmour / gudtArmour(gudtObject(gudtRadar(gudtPlayer.lngRadarObject).lngObject).udtSystems.bytArmour).lngMaxArmour
        If sngTemp <= 0.33 Then
            ShowText "Condition", TEXT_RADAR_X_CONDITION, TEXT_RADAR_Y_CONDITION, vbRed, mlngDC
        ElseIf sngTemp <= 0.67 Then
            ShowText "Condition", TEXT_RADAR_X_CONDITION, TEXT_RADAR_Y_CONDITION, vbYellow, mlngDC
        End If
    End If
    
End Sub

Public Sub RadarTab()

Dim lngLowestOrder As Long
Dim lngCurrentOrder As Long
Dim i As Long
Dim j As Long

    'If there are no entries in radar array, exit
    If glngNumRadar = 0 Then
        gudtPlayer.lngRadarObject = -1
        Exit Sub
    End If
    
    'If there is only one entry, select it
    If glngNumRadar = 1 Then
        gudtPlayer.lngRadarObject = 0
        Exit Sub
    End If

    'Determine the current order
    If gudtPlayer.lngRadarObject >= 0 Then
        lngCurrentOrder = gudtRadar(gudtPlayer.lngRadarObject).lngOrder
    Else
        lngCurrentOrder = glngRadarOrder
    End If

    'Determine the lowest order in the list
    For i = 0 To glngNumRadar - 1
        If gudtRadar(i).lngOrder < lngLowestOrder Or lngLowestOrder = 0 Then
            lngLowestOrder = gudtRadar(i).lngOrder
        End If
    Next i
    
    'Loop until we find the next in the array
    i = lngCurrentOrder + 1
    Do
        'Loop at the top
        If i > glngRadarOrder Then i = lngLowestOrder
        'Check if there's a gudtRadar entry with matching lngOrder
        For j = 0 To glngNumRadar - 1
            'Is this the one?
            If (gudtRadar(j).lngOrder = i) And (gudtObject(gudtRadar(j).lngObject).blnExists = True) Then
                'This is it!
                gudtPlayer.lngRadarObject = j
                Exit Sub
            End If
        Next j
        'Increment
        i = i + 1
    Loop Until i = lngCurrentOrder

    'Error?  Next in order wasn't found..
    gudtPlayer.lngRadarObject = -1

End Sub

Private Sub ShiftRadarTab()

Dim lngLowestOrder As Long
Dim lngCurrentOrder As Long
Dim i As Long
Dim j As Long

    'If there are no entries in radar array, exit
    If glngNumRadar = 0 Then
        gudtPlayer.lngRadarObject = -1
        Exit Sub
    End If
    
    'If there is only one entry, select it
    If glngNumRadar = 1 Then
        gudtPlayer.lngRadarObject = 0
        Exit Sub
    End If

    'Determine the current order
    If gudtPlayer.lngRadarObject >= 0 Then
        lngCurrentOrder = gudtRadar(gudtPlayer.lngRadarObject).lngOrder
    Else
        lngCurrentOrder = glngRadarOrder
    End If

    'Determine the lowest order in the list
    For i = 0 To glngNumRadar - 1
        If gudtRadar(i).lngOrder < lngLowestOrder Or lngLowestOrder = 0 Then
            lngLowestOrder = gudtRadar(i).lngOrder
        End If
    Next i
    
    'Loop until we find the next in the array
    i = lngCurrentOrder - 1
    Do
        'Loop at the top
        If i < lngLowestOrder Then i = glngRadarOrder
        'Check if there's a gudtRadar entry with matching lngOrder
        For j = 0 To glngNumRadar - 1
            'Is this the one?
            If (gudtRadar(j).lngOrder = i) And (gudtObject(gudtRadar(j).lngObject).blnExists = True) Then
                'This is it!
                gudtPlayer.lngRadarObject = j
                Exit Sub
            End If
        Next j
        'Increment
        i = i - 1
    Loop Until i = lngCurrentOrder

    'Error?  Next in order wasn't found..
    gudtPlayer.lngRadarObject = -1

End Sub

Public Sub RadarEnemyTab()

Dim lngLowestOrder As Long
Dim lngCurrentOrder As Long
Dim i As Long
Dim j As Long

    'If there are no entries in radar array, exit
    If glngNumRadar = 0 Then
        gudtPlayer.lngRadarObject = -1
        Exit Sub
    End If
    
    'If there is only one entry, select it
    If glngNumRadar = 1 Then
        gudtPlayer.lngRadarObject = 0
        Exit Sub
    End If

    'Determine the current order
    If gudtPlayer.lngRadarObject >= 0 Then
        lngCurrentOrder = gudtRadar(gudtPlayer.lngRadarObject).lngOrder
    Else
        lngCurrentOrder = glngRadarOrder
    End If

    'Determine the lowest order in the list
    For i = 0 To glngNumRadar - 1
        If gudtRadar(i).lngOrder < lngLowestOrder Or lngLowestOrder = 0 Then
            lngLowestOrder = gudtRadar(i).lngOrder
        End If
    Next i
    
    'Loop until we find the next in the array
    i = lngCurrentOrder + 1
    Do
        'Loop at the top
        If i > glngRadarOrder Then i = lngLowestOrder
        'Check if there's a gudtRadar entry with matching lngOrder and blnEnemy
        For j = 0 To glngNumRadar - 1
            'Is this the one?
            If (gudtRadar(j).lngOrder = i) And (gudtRadar(j).blnEnemy = True) And (gudtObject(gudtRadar(j).lngObject).blnExists = True) Then
                'This is it!
                gudtPlayer.lngRadarObject = j
                Exit Sub
            End If
        Next j
        'Increment
        i = i + 1
    Loop Until (i = lngCurrentOrder) Or ((i = glngRadarOrder + 1) And (lngLowestOrder = lngCurrentOrder))

End Sub

Private Sub RadarClosestEnemyTab()

Dim i As Long
Dim lngObject As Long
Dim dblDistance As Double

    'Find the closest enemy on radar
    lngObject = -1
    For i = 0 To glngNumRadar - 1
        If (gudtRadar(i).blnEnemy = True) And (gudtObject(gudtRadar(i).lngObject).udtInfo.dblDistance < dblDistance Or dblDistance = 0) And (gudtObject(gudtRadar(i).lngObject).blnExists = True) Then
            'Set this as the current closest
            dblDistance = gudtObject(gudtRadar(i).lngObject).udtInfo.dblDistance
            lngObject = i
        End If
    Next i
    
    'Set new radar object
    If lngObject >= 0 Then gudtPlayer.lngRadarObject = lngObject

End Sub

Private Sub UpdateRadar()

Dim i As Long
Dim j As Long
Dim blnFound As Boolean
    
    'Check for leavers
    i = 0
    Do While i <= glngNumRadar - 1
        If (gudtObject(gudtRadar(i).lngObject).udtInfo.dblDistance - gudtObject(gudtRadar(i).lngObject).udtSprite.intWidth \ 2 > gudtPlayer.dblCurrentRange) Or (gudtObject(gudtRadar(i).lngObject).blnExists = False) Then
            'If current radar object is higher in the array, decrement its pointer
            If gudtPlayer.lngRadarObject > i Then
                gudtPlayer.lngRadarObject = gudtPlayer.lngRadarObject - 1
            'If this is the current radar object (and not the only one), then tab!
            ElseIf gudtPlayer.lngRadarObject = i And glngNumRadar > 1 Then
                'Tab first..
                RadarTab
                'See if there's an enemy first
                RadarClosestEnemyTab
                'Decrement..
                If gudtPlayer.lngRadarObject >= i Then gudtPlayer.lngRadarObject = gudtPlayer.lngRadarObject - 1
            'If this is the only remaining object then set to -1
            ElseIf glngNumRadar = 1 Then
                gudtPlayer.lngRadarObject = -1
            End If
            'Remove object from array
            For j = i To glngNumRadar - 2
                gudtRadar(j).lngObject = gudtRadar(j + 1).lngObject
                gudtRadar(j).lngOrder = gudtRadar(j + 1).lngOrder
                gudtRadar(j).intSize = gudtRadar(j + 1).intSize
                gudtRadar(j).blnEnemy = gudtRadar(j + 1).blnEnemy
            Next j
            'Resize the array
            glngNumRadar = glngNumRadar - 1
            If glngNumRadar > 0 Then
                ReDim Preserve gudtRadar(glngNumRadar - 1)
            Else
                Erase gudtRadar
            End If
            i = i - 1
        End If
        i = i + 1
    Loop
    
    'Check for newbies
    For i = 0 To UBound(gudtObject)
        If (gudtObject(i).udtInfo.dblDistance - gudtObject(i).udtSprite.intWidth \ 2 <= gudtPlayer.dblCurrentRange) And (gudtObject(i).blnExists = True) Then
            'See if this one is already in the array
            blnFound = False
            For j = 0 To glngNumRadar - 1
                If gudtRadar(j).lngObject = i Then
                    blnFound = True
                    Exit For
                End If
            Next j
            'If this is a new one, add to the array
            If blnFound = False Then
                'Add to array
                glngNumRadar = glngNumRadar + 1
                glngRadarOrder = glngRadarOrder + 1
                ReDim Preserve gudtRadar(glngNumRadar - 1)
                gudtRadar(glngNumRadar - 1).lngObject = i
                gudtRadar(glngNumRadar - 1).lngOrder = glngRadarOrder
                'If we don't have an object selected, select one
                If gudtPlayer.lngRadarObject = -1 Then RadarTab
            End If
        End If
    Next i

End Sub

Private Sub DisplayRadar()

Dim i As Long
Dim strRelation As String
Dim intSize As Integer

    'Exit sub if nothing to display
    If glngNumRadar = 0 Then Exit Sub

    'Check for objects in range
    For i = glngNumRadar - 1 To 0 Step -1
        'Determine size
        intSize = ((gudtObject(gudtRadar(i).lngObject).udtSprite.intWidth) / gudtPlayer.dblCurrentRange / 2 * RADAR_WIDTH)
        If intSize < 2 Then intSize = 2
        If intSize > 2 And intSize < 4 Then intSize = 4
        If intSize > 4 And intSize < 6 Then intSize = 6
        If intSize > 6 And intSize < 8 Then intSize = 8
        If intSize > 8 And intSize < 10 Then intSize = 10
        If intSize > 10 And intSize < 12 Then intSize = 12
        If intSize > 12 And intSize < 14 Then intSize = 14
        If intSize > 14 Then intSize = 16
        gudtRadar(i).intSize = intSize
        'Determine relations and display
        gudtRadar(i).blnEnemy = False
        If gudtObject(gudtRadar(i).lngObject).udtInfo.bytRace = RACE_PLANET Then
            DDraw.DisplaySprite mudtSpriteRadar.lngPlanet(intSize \ 2 - 1), ((gudtObject(gudtRadar(i).lngObject).udtPhysics.dblX - gudtPlayer.udtPhysics.dblX) / gudtPlayer.dblCurrentRange / 2 * RADAR_WIDTH) + 725 - (intSize \ 2), ((gudtObject(gudtRadar(i).lngObject).udtPhysics.dblY - gudtPlayer.udtPhysics.dblY) / gudtPlayer.dblCurrentRange / 2 * RADAR_WIDTH) + 73 - (intSize \ 2)
        ElseIf gudtRace(gudtObject(gudtRadar(i).lngObject).udtInfo.bytRace).blnEncountered = False Then
            DDraw.DisplaySprite mudtSpriteRadar.lngUnknown(intSize \ 2 - 1), ((gudtObject(gudtRadar(i).lngObject).udtPhysics.dblX - gudtPlayer.udtPhysics.dblX) / gudtPlayer.dblCurrentRange / 2 * RADAR_WIDTH) + 725 - (intSize \ 2), ((gudtObject(gudtRadar(i).lngObject).udtPhysics.dblY - gudtPlayer.udtPhysics.dblY) / gudtPlayer.dblCurrentRange / 2 * RADAR_WIDTH) + 73 - (intSize \ 2)
        ElseIf gudtRace(gudtObject(gudtRadar(i).lngObject).udtInfo.bytRace).intRelations(RACE_PLAYER) > RELATIONS_BAD And gudtRace(gudtObject(gudtRadar(i).lngObject).udtInfo.bytRace).intRelations(RACE_PLAYER) < RELATIONS_GOOD Then
            DDraw.DisplaySprite mudtSpriteRadar.lngNeutral(intSize \ 2 - 1), ((gudtObject(gudtRadar(i).lngObject).udtPhysics.dblX - gudtPlayer.udtPhysics.dblX) / gudtPlayer.dblCurrentRange / 2 * RADAR_WIDTH) + 725 - (intSize \ 2), ((gudtObject(gudtRadar(i).lngObject).udtPhysics.dblY - gudtPlayer.udtPhysics.dblY) / gudtPlayer.dblCurrentRange / 2 * RADAR_WIDTH) + 73 - (intSize \ 2)
        ElseIf gudtRace(gudtObject(gudtRadar(i).lngObject).udtInfo.bytRace).intRelations(RACE_PLAYER) <= RELATIONS_BAD Then
            DDraw.DisplaySprite mudtSpriteRadar.lngEnemy(intSize \ 2 - 1), ((gudtObject(gudtRadar(i).lngObject).udtPhysics.dblX - gudtPlayer.udtPhysics.dblX) / gudtPlayer.dblCurrentRange / 2 * RADAR_WIDTH) + 725 - (intSize \ 2), ((gudtObject(gudtRadar(i).lngObject).udtPhysics.dblY - gudtPlayer.udtPhysics.dblY) / gudtPlayer.dblCurrentRange / 2 * RADAR_WIDTH) + 73 - (intSize \ 2)
            gudtRadar(i).blnEnemy = True
        Else
            DDraw.DisplaySprite mudtSpriteRadar.lngFriendly(intSize \ 2 - 1), ((gudtObject(gudtRadar(i).lngObject).udtPhysics.dblX - gudtPlayer.udtPhysics.dblX) / gudtPlayer.dblCurrentRange / 2 * RADAR_WIDTH) + 725 - (intSize \ 2), ((gudtObject(gudtRadar(i).lngObject).udtPhysics.dblY - gudtPlayer.udtPhysics.dblY) / gudtPlayer.dblCurrentRange / 2 * RADAR_WIDTH) + 73 - (intSize \ 2)
        End If
    Next i

    'Check for missiles in range
    For i = 0 To glngNumLiveMissiles - 1
        'Check range
        If GetDist(gudtLiveMissile(i).dblX, gudtLiveMissile(i).dblY, gudtPlayer.udtPhysics.dblX, gudtPlayer.udtPhysics.dblY) <= gudtPlayer.dblCurrentRange Then
            'Display
            DDraw.DisplaySprite mudtSpriteRadar.lngWeapon, ((gudtLiveMissile(i).dblX - gudtPlayer.udtPhysics.dblX) / gudtPlayer.dblCurrentRange / 2 * RADAR_WIDTH) + 725, ((gudtLiveMissile(i).dblY - gudtPlayer.udtPhysics.dblY) / gudtPlayer.dblCurrentRange / 2 * RADAR_WIDTH) + 73
        End If
    Next i

End Sub

Private Sub DisplayRadarLines()

Dim lngX1 As Long
Dim lngX2 As Long
Dim lngY1 As Long
Dim lngY2 As Long
Dim lngXCenter As Long
Dim lngYCenter As Long
Dim lngLineDist As Long
Dim lngLineLength As Long
Dim lngX1LineLength As Long
Dim lngX2LineLength As Long
Dim lngY1LineLength As Long
Dim lngY2LineLength As Long

    'Load the pen
    SetPen mlngDC, 1, vbRed

    'Display the central radar "viewport" box
    lngX1 = 725 - (TACTICAL_VIEW_WIDTH / gudtPlayer.dblCurrentRange / 2 * RADAR_WIDTH) / 2
    lngX2 = 725 + (TACTICAL_VIEW_WIDTH / gudtPlayer.dblCurrentRange / 2 * RADAR_WIDTH) / 2
    lngY1 = 73 - (TACTICAL_VIEW_HEIGHT / gudtPlayer.dblCurrentRange / 2 * RADAR_WIDTH) / 2
    lngY2 = 73 + (TACTICAL_VIEW_HEIGHT / gudtPlayer.dblCurrentRange / 2 * RADAR_WIDTH) / 2
    BoxDraw lngX1, lngY1, lngX2, lngY2, mlngDC
    
    'Only do this stuff if there's something actually on the radar
    If gudtPlayer.lngRadarObject > -1 Then
        
        'Display the radar object highlight box
        lngXCenter = ((gudtObject(gudtRadar(gudtPlayer.lngRadarObject).lngObject).udtPhysics.dblX - gudtPlayer.udtPhysics.dblX) / gudtPlayer.dblCurrentRange / 2 * RADAR_WIDTH) + 725
        lngYCenter = ((gudtObject(gudtRadar(gudtPlayer.lngRadarObject).lngObject).udtPhysics.dblY - gudtPlayer.udtPhysics.dblY) / gudtPlayer.dblCurrentRange / 2 * RADAR_WIDTH) + 73
        Select Case gudtRadar(gudtPlayer.lngRadarObject).intSize
            Case 2
                lngLineDist = 2
            Case 4
                lngLineDist = 3
            Case 6
                lngLineDist = 4
            Case 8
                lngLineDist = 5
            Case 10
                lngLineDist = 7
            Case 12
                lngLineDist = 8
            Case 14
                lngLineDist = 10
            Case 16
                lngLineDist = 12
        End Select
        lngX1 = lngXCenter - lngLineDist - 1
        lngX2 = lngXCenter + lngLineDist
        lngY1 = lngYCenter - lngLineDist - 1
        lngY2 = lngYCenter + lngLineDist
        lngLineLength = gudtRadar(gudtPlayer.lngRadarObject).intSize / 2.5
        If lngLineLength < 2 Then lngLineLength = 2
        'Draw it
        If (lngX1 > 660) Then LineDraw lngX1, lngY1, lngX1 + lngLineLength, lngY1, mlngDC
        If (lngX1 > 660) Then LineDraw lngX1, lngY1, lngX1, lngY1 + lngLineLength, mlngDC
        If (lngX1 > 660) And (lngY2 < 138) Then LineDraw lngX1, lngY2, lngX1 + lngLineLength, lngY2, mlngDC
        If (lngX1 > 660) And (lngY2 < 138) Then LineDraw lngX1, lngY2, lngX1, lngY2 - lngLineLength, mlngDC
        If (lngY2 < 138) Then LineDraw lngX2, lngY2, lngX2 - lngLineLength, lngY2, mlngDC
        If (lngY2 < 138) Then LineDraw lngX2, lngY2, lngX2, lngY2 - lngLineLength, mlngDC
        LineDraw lngX2, lngY1, lngX2 - lngLineLength, lngY1, mlngDC
        LineDraw lngX2, lngY1, lngX2, lngY1 + lngLineLength, mlngDC
        
        'Display the viewport object highlight box
        lngXCenter = gudtObject(gudtRadar(gudtPlayer.lngRadarObject).lngObject).udtPhysics.dblX - gudtPlayer.udtPhysics.dblX
        lngYCenter = gudtObject(gudtRadar(gudtPlayer.lngRadarObject).lngObject).udtPhysics.dblY - gudtPlayer.udtPhysics.dblY
        lngLineDist = gudtObject(gudtRadar(gudtPlayer.lngRadarObject).lngObject).udtSprite.intWidth / 1.75
        If lngLineDist < 10 Then lngLineDist = 10
        lngLineLength = gudtObject(gudtRadar(gudtPlayer.lngRadarObject).lngObject).udtSprite.intWidth / 5
        If lngLineLength < 5 Then lngLineLength = 5
        lngX1 = lngXCenter - lngLineDist - 1 + 330
        lngX2 = lngXCenter + lngLineDist + 330
        lngY1 = lngYCenter - lngLineDist - 1 + 250
        lngY2 = lngYCenter + lngLineDist + 250
        'Clip
        lngX1LineLength = lngLineLength
        lngX2LineLength = lngLineLength
        lngY1LineLength = lngLineLength
        lngY2LineLength = lngLineLength
        If lngX1 < 10 Then
            lngX1LineLength = lngLineLength - (10 - lngX1)
            lngX1 = 9
        End If
        If lngX1 > TACTICAL_VIEW_WIDTH + 6 - lngX1LineLength Then
            lngX1LineLength = lngLineLength - (lngX1 - (TACTICAL_VIEW_WIDTH + 7 - lngX1LineLength))
            If lngX1LineLength < 0 Then lngX1LineLength = 0
        End If
        If lngX1 > TACTICAL_VIEW_WIDTH + 6 Then lngX1 = TACTICAL_VIEW_WIDTH + 7
        If lngX2 > TACTICAL_VIEW_WIDTH + 6 Then
            lngX2LineLength = lngLineLength - (lngX2 - (TACTICAL_VIEW_WIDTH + 7))
            If lngX2LineLength < 0 Then lngX2LineLength = 0
            lngX2 = TACTICAL_VIEW_WIDTH + 7
        End If
        If lngY1 < 9 Then
            lngY1LineLength = lngLineLength - (9 - lngY1)
            lngY1 = 8
        End If
        If lngY1 > TACTICAL_VIEW_HEIGHT + 6 - lngY1LineLength Then
            lngY1LineLength = lngLineLength - (lngY1 - (TACTICAL_VIEW_HEIGHT + 7 - lngY1LineLength))
            If lngY1LineLength < 0 Then lngY1LineLength = 0
        End If
        If lngY1 > TACTICAL_VIEW_HEIGHT + 6 Then lngY1 = TACTICAL_VIEW_HEIGHT + 7
        If lngY2 > TACTICAL_VIEW_HEIGHT + 6 Then
            lngY2LineLength = lngLineLength - (lngY2 - (TACTICAL_VIEW_HEIGHT + 7))
            If lngY2LineLength < 0 Then lngY2LineLength = 0
            lngY2 = TACTICAL_VIEW_HEIGHT + 7
        End If
        'Draw it
        LineDraw lngX1, lngY1, lngX1 + lngX1LineLength, lngY1, mlngDC
        LineDraw lngX1, lngY2, lngX1 + lngX1LineLength, lngY2, mlngDC
        LineDraw lngX2, lngY2, lngX2 - lngX2LineLength, lngY2, mlngDC
        LineDraw lngX2, lngY1, lngX2 - lngX2LineLength, lngY1, mlngDC
        LineDraw lngX1, lngY1, lngX1, lngY1 + lngY1LineLength, mlngDC
        LineDraw lngX2, lngY1, lngX2, lngY1 + lngY1LineLength, mlngDC
        LineDraw lngX1, lngY2, lngX1, lngY2 - lngY2LineLength, mlngDC
        LineDraw lngX2, lngY2, lngX2, lngY2 - lngY2LineLength, mlngDC
    
    End If
    
    'Remove pen
    RemovePen mlngDC

End Sub

Private Sub DisplayStars()

Dim i As Long

    'Display the stars!
    For i = 0 To STAR_NUM
        'Move the stars
        Motion mudtStar(i).dblX, mudtStar(i).dblY, mudtStar(i).sngRelSpeed * gudtPlayer.udtPhysics.sngSpeed, gudtPlayer.udtPhysics.sngHeading + Pi
        'Ensure they're still on screen
        If mudtStar(i).dblX > TACTICAL_VIEW_WIDTH \ 2 Then
            mudtStar(i).dblX = -TACTICAL_VIEW_WIDTH \ 2
            mudtStar(i).dblY = Fix(Rnd * TACTICAL_VIEW_HEIGHT) - TACTICAL_VIEW_HEIGHT \ 2
        End If
        If mudtStar(i).dblX < -TACTICAL_VIEW_WIDTH \ 2 Then
            mudtStar(i).dblX = TACTICAL_VIEW_WIDTH \ 2
            mudtStar(i).dblY = Fix(Rnd * TACTICAL_VIEW_HEIGHT) - TACTICAL_VIEW_HEIGHT \ 2
        End If
        If mudtStar(i).dblY > TACTICAL_VIEW_HEIGHT \ 2 Then
            mudtStar(i).dblY = -TACTICAL_VIEW_HEIGHT \ 2
            mudtStar(i).dblX = Fix(Rnd * TACTICAL_VIEW_WIDTH) - TACTICAL_VIEW_WIDTH \ 2
        End If
        If mudtStar(i).dblY < -TACTICAL_VIEW_HEIGHT \ 2 Then
            mudtStar(i).dblY = TACTICAL_VIEW_HEIGHT \ 2
            mudtStar(i).dblX = Fix(Rnd * TACTICAL_VIEW_WIDTH) - TACTICAL_VIEW_WIDTH \ 2
        End If
        'Display them
        DDraw.DisplaySprite mlngSpriteStar, CInt(mudtStar(i).dblX) + 330, CInt(mudtStar(i).dblY) + 250
    Next i

End Sub

Private Sub DisplayObjects()

Dim i As Long
Dim sngRelX As Single
Dim sngRelY As Single
Dim bytShield As Byte

    'Display all visible objects
    For i = 0 To UBound(gudtObject)
        'If the object exists and is loaded..
        If gudtObject(i).blnExists And gudtObject(i).udtSprite.blnLoaded Then
            'If it's within visible range..
            sngRelX = gudtObject(i).udtPhysics.dblX - gudtPlayer.udtPhysics.dblX
            sngRelY = gudtObject(i).udtPhysics.dblY - gudtPlayer.udtPhysics.dblY
            If sngRelX > -321 - gudtObject(i).udtSprite.intWidth \ 2 And sngRelX < 321 + gudtObject(i).udtSprite.intWidth \ 2 And sngRelY > -241 - gudtObject(i).udtSprite.intHeight \ 2 And sngRelY < 241 + gudtObject(i).udtSprite.intHeight \ 2 Then
                'Display!
                DDraw.DisplaySpriteClip gudtObject(i).udtSprite.lngSpriteObject, (gudtObject(i).udtSprite.bytAnimNum * (gudtObject(i).udtSprite.bytFrameAmt + 1)) + gudtObject(i).udtSprite.bytFrameNum, sngRelX + 330 - gudtObject(i).udtSprite.intWidth \ 2, sngRelY + 250 - gudtObject(i).udtSprite.intHeight \ 2
                'Display shields?
                If gudtObject(i).udtInfo.blnShieldUp = True Then
                    'Find the correct shield sprite
                    bytShield = FindShield(gudtObject(i).udtSprite.intWidth)
                    'Display!
                    If bytShield = SHIELD_9 Then
                        'Center on spacestation
                        DDraw.DisplayClip mlngSpriteShield(bytShield), sngRelX + 330 - gudtObject(i).udtSprite.intWidth \ 2, sngRelY + 200 - gudtObject(i).udtSprite.intHeight \ 2, , True
                    Else
                        'Ships are already symmetrical
                        DDraw.DisplayClip mlngSpriteShield(bytShield), sngRelX + 330 - gudtObject(i).udtSprite.intWidth \ 2, sngRelY + 250 - gudtObject(i).udtSprite.intHeight \ 2, , True
                    End If
                    'Remove shields?
                    If gudtObject(i).udtInfo.lngShieldDown <= glngGameTime Then gudtObject(i).udtInfo.blnShieldUp = False
                End If
            End If
        End If
    Next i

End Sub

Private Sub DisplayExplosions()

Dim i As Long
Dim sngRelX As Single
Dim sngRelY As Single

    'Are there any?
    If glngNumExplosions = 0 Then Exit Sub
    
    'Display those that are in range
    For i = 0 To glngNumExplosions - 1
        'Range?
        sngRelX = gudtExplosion(i).dblX - gudtPlayer.udtPhysics.dblX
        sngRelY = gudtExplosion(i).dblY - gudtPlayer.udtPhysics.dblY
        'Type?
        Select Case gudtExplosion(i).bytExplosionType
            Case 0
                If sngRelX > -321 - EXPLOSION0_WIDTH \ 2 And sngRelX < 321 + EXPLOSION0_WIDTH \ 2 And sngRelY > -241 - EXPLOSION0_HEIGHT \ 2 And sngRelY < 241 + EXPLOSION0_HEIGHT \ 2 Then
                    'Display!
                    DDraw.DisplayClip gudtExplosionSprite(0).lngSprite(gudtExplosion(i).bytAnimFrame), sngRelX + 330 - EXPLOSION0_WIDTH \ 2, sngRelY + 250 - EXPLOSION0_HEIGHT \ 2, , True
                End If
            Case 1
                If sngRelX > -321 - EXPLOSION1_WIDTH \ 2 And sngRelX < 321 + EXPLOSION1_WIDTH \ 2 And sngRelY > -241 - EXPLOSION1_HEIGHT \ 2 And sngRelY < 241 + EXPLOSION1_HEIGHT \ 2 Then
                    'Display!
                    DDraw.DisplayClip gudtExplosionSprite(1).lngSprite(gudtExplosion(i).bytAnimFrame), sngRelX + 330 - EXPLOSION1_WIDTH \ 2, sngRelY + 250 - EXPLOSION1_HEIGHT \ 2, , True
                End If
        End Select
    Next i

End Sub

Private Sub DisplayDeathExplosions()

Dim lngNumExplosions As Long
Dim lngInterval As Long
Dim dblX As Double
Dim dblY As Double

    'Is it time for the big explosion?
    If glngGameTime > glngPlayerExplodingStart + PLAYER_EXPLODE_DURATION Then
        'Big explosion?
        If gudtHull(gudtPlayer.udtSystems.bytHull).udtSprite.intWidth >= PIXELS_WIDTH_FOR_LARGE_EXPLOSION Then
            CreateExplosion gudtPlayer.udtPhysics.dblX, gudtPlayer.udtPhysics.dblY, gudtPlayer.udtPhysics.sngSpeed, gudtPlayer.udtPhysics.sngHeading, 1
        'Small explosion
        Else
            CreateExplosion gudtPlayer.udtPhysics.dblX, gudtPlayer.udtPhysics.dblY, gudtPlayer.udtPhysics.sngSpeed, gudtPlayer.udtPhysics.sngHeading, 0
        End If
        'Really dead!
        gblnPlayerExploded = True
        Exit Sub
    End If
    
    'Nope, but maybe a little explosion?  First, calc how many
    lngNumExplosions = gudtHull(gudtPlayer.udtSystems.bytHull).udtSprite.intWidth \ PIXELS_WIDTH_PER_EXPLOSION
    
    'Calc explosion interval
    lngInterval = PLAYER_EXPLODE_DURATION \ lngNumExplosions
    
    'Given the number of explosions that have already occurred, is it time for a new one?
    If ((glngPlayerExplosionNum + 1) * lngInterval) + glngPlayerExplodingStart < glngGameTime Then
        'Get the coords
        Randomize
        PointOnLine gudtPlayer.udtPhysics.dblX, gudtPlayer.udtPhysics.dblY, Rnd() * 2 * Pi, gudtHull(gudtPlayer.udtSystems.bytHull).udtSprite.intWidth / 4 + Rnd() * gudtHull(gudtPlayer.udtSystems.bytHull).udtSprite.intWidth / 4, dblX, dblY
        'PointOnLine gudtPlayer.udtPhysics.dblX, gudtPlayer.udtPhysics.dblY, Rnd() * 2 * Pi, Rnd() * gudtHull(gudtPlayer.udtSystems.bytHull).udtSprite.intWidth / 2, dblX, dblY
        'Make a small explosion
        CreateExplosion dblX, dblY, gudtPlayer.udtPhysics.sngSpeed, gudtPlayer.udtPhysics.sngHeading, 0
        'Increment the counter
        glngPlayerExplosionNum = glngPlayerExplosionNum + 1
    End If

End Sub

Private Sub DisplayPlayer()

Dim bytShield As Byte
Dim blnFTL As Boolean
Dim i As Long
Dim dblXDisp As Double
Dim dblYDisp As Double
Static dblXPrev(FTL_EFFECT_SPRITE_NUM - 1) As Double
Static dblYPrev(FTL_EFFECT_SPRITE_NUM - 1) As Double

    'Is it time to display the "you're dead" message?
    If (gblnPlayerDead = True) And (glngGameTime > glngPlayerExplodingStart + PLAYER_EXPLODE_CONTINUE_DELAY) And (gblnPlayerDeadMessage = False) Then
        AddComm "Your ship has been destroyed!  Press any key..."
        gblnPlayerDeadMessage = True
    End If

    'Is the player dead and exploding?
    If (gblnPlayerDead = True) And (gblnPlayerExploded = False) Then DisplayDeathExplosions

    'Is the player really dead?
    If (gblnPlayerDead = True) And (gblnPlayerExploded = True) Then Exit Sub

    'Are we going FTL?
    blnFTL = (gudtPlayer.udtSystems.blnARCDActive = True Or gudtPlayer.udtSystems.blnFTLDActive = True)

    'Display the player's ship
    DDraw.DisplaySpriteClip gudtPlayer.udtSprite.lngSpriteObject, (gudtPlayer.udtSprite.bytAnimNum * (gudtPlayer.udtSprite.bytFrameAmt + 1)) + gudtPlayer.udtSprite.bytFrameNum, 330 - gudtPlayer.udtSprite.intWidth / 2, 250 - gudtPlayer.udtSprite.intHeight / 2, , blnFTL
    
    'If we're going FTL, then show the neat effects
    If blnFTL Then
        'Make sure things look nice and random
        Randomize
        'Loop through the previous coords
        For i = FTL_EFFECT_SPRITE_NUM - 1 To 0 Step -1
            'Calc the new displacement
            dblXDisp = Sin(-gudtPlayer.udtPhysics.sngHeading + Rnd(FTL_EFFECT_RAND) - FTL_EFFECT_RAND / 2) * FTL_EFFECT_DISTANCE * (i + 1) * Rnd(FTL_EFFECT_RAND)
            dblYDisp = Cos(-gudtPlayer.udtPhysics.sngHeading + Rnd(FTL_EFFECT_RAND) - FTL_EFFECT_RAND / 2) * FTL_EFFECT_DISTANCE * (i + 1) * Rnd(FTL_EFFECT_RAND)
            'Display the image
            DDraw.DisplaySpriteClip gudtPlayer.udtSprite.lngSpriteObject, (gudtPlayer.udtSprite.bytAnimNum * (gudtPlayer.udtSprite.bytFrameAmt + 1)) + gudtPlayer.udtSprite.bytFrameNum, (330 - gudtPlayer.udtSprite.intWidth / 2) + dblXPrev(i), (250 - gudtPlayer.udtSprite.intHeight / 2) + dblYPrev(i), , True
            'Set the new coord
            If i <> 0 Then
                'Take the previous coord
                dblXPrev(i) = dblXPrev(i - 1) + dblXDisp
                dblYPrev(i) = dblYPrev(i - 1) + dblYDisp
            Else
                'Set the new coord
                dblXPrev(i) = dblXDisp
                dblYPrev(i) = dblYDisp
            End If
        Next i
    Else
        'Clear the previous coord arrays
        Erase dblXPrev
        Erase dblYPrev
    End If
    
    'Display shields
    If gudtPlayer.udtInfo.blnShieldUp = True Then
        'Find correct shield sprite
        bytShield = FindShield(gudtPlayer.udtSprite.intWidth)
        'Display!
        DDraw.DisplayClip mlngSpriteShield(bytShield), 330 - gudtPlayer.udtSprite.intWidth / 2, 250 - gudtPlayer.udtSprite.intHeight / 2, , True
        'Remove shields?
        If gudtPlayer.udtInfo.lngShieldDown <= glngGameTime Then gudtPlayer.udtInfo.blnShieldUp = False
    End If

End Sub

Private Sub DisplayFrame()

    'Display the tactical screen's frame
    DDraw.DisplaySprite mlngSpriteTacticalTop, 0, 0
    DDraw.DisplaySprite mlngSpriteTacticalRight, 642, 0
    DDraw.DisplaySprite mlngSpriteTacticalLeft, 0, 13
    DDraw.DisplaySprite mlngSpriteTacticalBottom, 0, 487

End Sub

Private Sub DisplayText()

    'Set the font
    SetFont mlngDC, "MS Sans Serif", 16
    
    'Display constant text
    'ShowText "Energy:", TEXT_PLAYER_ROW1X, TEXT_PLAYER_Y, vbGreen, mlngDC
    'ShowText "Generator:", TEXT_PLAYER_ROW1X, TEXT_PLAYER_Y + TEXT_PLAYER_HEIGHT * 1, vbGreen, mlngDC
    'ShowText "Engines:", TEXT_PLAYER_ROW1X, TEXT_PLAYER_Y + TEXT_PLAYER_HEIGHT * 2, vbGreen, mlngDC
    'ShowText "Shields:", TEXT_PLAYER_ROW1X, TEXT_PLAYER_Y + TEXT_PLAYER_HEIGHT * 3, vbGreen, mlngDC
    'ShowText "Weapons:", TEXT_PLAYER_ROW1X, TEXT_PLAYER_Y + TEXT_PLAYER_HEIGHT * 4, vbGreen, mlngDC
    'ShowText "Speed:", TEXT_PLAYER_ROW2X, TEXT_PLAYER_Y, vbGreen, mlngDC
    'ShowText "Mass:", TEXT_PLAYER_ROW2X, TEXT_PLAYER_Y + TEXT_PLAYER_HEIGHT * 1, vbGreen, mlngDC
    'ShowText "Destination:", TEXT_PLAYER_ROW2X, TEXT_PLAYER_Y + TEXT_PLAYER_HEIGHT * 2, vbGreen, mlngDC
    'ShowText " Direction:", TEXT_PLAYER_ROW2X, TEXT_PLAYER_Y + TEXT_PLAYER_HEIGHT * 3, vbGreen, mlngDC
    'ShowText " Distance:", TEXT_PLAYER_ROW2X, TEXT_PLAYER_Y + TEXT_PLAYER_HEIGHT * 4, vbGreen, mlngDC
    'ShowText "Armour:", TEXT_PLAYER_ROW3X, TEXT_PLAYER_Y, vbGreen, mlngDC
    'ShowText "Crew:", TEXT_PLAYER_ROW3X, TEXT_PLAYER_Y + TEXT_PLAYER_HEIGHT * 1, vbGreen, mlngDC
    'ShowText "Fuel:", TEXT_PLAYER_ROW3X, TEXT_PLAYER_Y + TEXT_PLAYER_HEIGHT * 2, vbGreen, mlngDC
    'ShowText "Cargo:", TEXT_PLAYER_ROW3X, TEXT_PLAYER_Y + TEXT_PLAYER_HEIGHT * 3, vbGreen, mlngDC
    'ShowText " Autopilot   Light Drive   Jammer", TEXT_PLAYER_ROW3X, TEXT_PLAYER_Y + TEXT_PLAYER_HEIGHT * 4, RGB(40, 40, 40), mlngDC
    
    'Display autopilot/FTLD/Jammer
    'ShowText "Mis: 99  Mine: 99   AP  LD  Jam", TEXT_PLAYER_ROW3X, TEXT_PLAYER_Y + TEXT_PLAYER_HEIGHT * 4, vbWhite, mlngDC
    If gudtPlayer.udtAI.bytAction = AI_AUTOPILOT Then
        'Autopilot
        ShowText "AP", TEXT_PLAYER_ROW3X + TEXT_PLAYER_DELTA_AP, TEXT_PLAYER_Y + TEXT_PLAYER_HEIGHT * 4, vbWhite, mlngDC
    End If
    If (gudtPlayer.udtSystems.blnFTLDActive = True Or gudtPlayer.udtSystems.blnARCDActive = True) Then
        'FTLD
        ShowText "LD", TEXT_PLAYER_ROW3X + TEXT_PLAYER_DELTA_LD, TEXT_PLAYER_Y + TEXT_PLAYER_HEIGHT * 4, vbWhite, mlngDC
    End If
    If gudtPlayer.udtSystems.blnJammerActive = True Then
        'Jammer
        ShowText "Jam", TEXT_PLAYER_ROW3X + TEXT_PLAYER_DELTA_JAM, TEXT_PLAYER_Y + TEXT_PLAYER_HEIGHT * 4, vbWhite, mlngDC
    End If
    
    'Display variable text
    ShowText NormalizeSpeed(gudtPlayer.udtPhysics.sngSpeed), TEXT_PLAYER_ROW2X + TEXT_PLAYER_ROW2_WIDTH, TEXT_PLAYER_Y, vbGreen, mlngDC, 1
    ShowText CStr(gudtPlayer.udtPhysics.lngMass) & " tons", TEXT_PLAYER_ROW2X + TEXT_PLAYER_ROW2_WIDTH, TEXT_PLAYER_Y + TEXT_PLAYER_HEIGHT * 1, vbGreen, mlngDC
    ShowText CStr(gudtPlayer.udtSystems.intMissileNum), TEXT_PLAYER_ROW3X + TEXT_PLAYER_DELTA_MIS, TEXT_PLAYER_Y + TEXT_PLAYER_HEIGHT * 4, vbGreen, mlngDC
    ShowText CStr(gudtPlayer.udtSystems.intMineNum), TEXT_PLAYER_ROW3X + TEXT_PLAYER_DELTA_MINE, TEXT_PLAYER_Y + TEXT_PLAYER_HEIGHT * 4, vbGreen, mlngDC
    
    'Display the destination and its direction/distance
    'Handle coords
    If gudtPlayer.udtAI.lngTarget = TARGET_COORDS Then
        ShowText "Coordinates", TEXT_PLAYER_ROW2X + TEXT_PLAYER_ROW2_WIDTH, TEXT_PLAYER_Y + TEXT_PLAYER_HEIGHT * 2, vbGreen, mlngDC
        ShowText CStr(Fix((180 / Pi) * FixAngle(FindAngle(gudtPlayer.udtPhysics.dblX, gudtPlayer.udtPhysics.dblY, gudtPlayer.udtAI.dblX, gudtPlayer.udtAI.dblY)))) & " degrees", TEXT_PLAYER_ROW2X + TEXT_PLAYER_ROW2_WIDTH, TEXT_PLAYER_Y + TEXT_PLAYER_HEIGHT * 3, vbGreen, mlngDC
        ShowText NormalizeDistance(GetDist(gudtPlayer.udtPhysics.dblX, gudtPlayer.udtPhysics.dblY, gudtPlayer.udtAI.dblX, gudtPlayer.udtAI.dblY)), TEXT_PLAYER_ROW2X + TEXT_PLAYER_ROW2_WIDTH, TEXT_PLAYER_Y + TEXT_PLAYER_HEIGHT * 4, vbGreen, mlngDC
    End If
    'Handle target
    If gudtPlayer.udtAI.lngTarget >= 0 Then
        ShowText gudtObject(gudtPlayer.udtAI.lngTarget).udtInfo.strName, TEXT_PLAYER_ROW2X + TEXT_PLAYER_ROW2_WIDTH, TEXT_PLAYER_Y + TEXT_PLAYER_HEIGHT * 2, vbGreen, mlngDC
        ShowText CStr(Fix((180 / Pi) * FixAngle(FindAngle(gudtPlayer.udtPhysics.dblX, gudtPlayer.udtPhysics.dblY, gudtObject(gudtPlayer.udtAI.lngTarget).udtPhysics.dblX, gudtObject(gudtPlayer.udtAI.lngTarget).udtPhysics.dblY)))) & " degrees", TEXT_PLAYER_ROW2X + TEXT_PLAYER_ROW2_WIDTH, TEXT_PLAYER_Y + TEXT_PLAYER_HEIGHT * 3, vbGreen, mlngDC
        ShowText NormalizeDistance(GetDist(gudtPlayer.udtPhysics.dblX, gudtPlayer.udtPhysics.dblY, gudtObject(gudtPlayer.udtAI.lngTarget).udtPhysics.dblX, gudtObject(gudtPlayer.udtAI.lngTarget).udtPhysics.dblY)), TEXT_PLAYER_ROW2X + TEXT_PLAYER_ROW2_WIDTH, TEXT_PLAYER_Y + TEXT_PLAYER_HEIGHT * 4, vbGreen, mlngDC, 1
    End If
    
    'Display communications messages
    DisplayComm
    
    'Display radar textout
    DisplayRadarText
    
    'Remove the font
    RemoveFont mlngDC

End Sub

Private Sub DisplayComm()

Dim i As Integer
Dim j As Integer

    'Exit if no messages
    If mintNumComm = 0 Then Exit Sub

    'Check if any have decayed
    i = 0
    Do While i < mintNumComm
        'Check time
        If mudtComm(i).lngDecay < gobjDX.TickCount Then
            'Remove
            For j = i To mintNumComm - 2
                mudtComm(j).lngDecay = mudtComm(j + 1).lngDecay
                mudtComm(j).strMessage = mudtComm(j + 1).strMessage
            Next j
            If mintNumComm > 1 Then ReDim Preserve mudtComm(UBound(mudtComm) - 1)
            mintNumComm = mintNumComm - 1
        End If
        i = i + 1
    Loop

    'Display all communications messages
    For i = 0 To mintNumComm - 1
        ShowText mudtComm(i).strMessage, COMM_X, COMM_Y - COMM_Y_HEIGHT * i, RGB(127, 127, 127), mlngDC, 1
    Next i

End Sub

Public Sub AddComm(strMessage As String)

Dim i As Integer

    'Add a message to the comm array
    If mintNumComm = 0 Then
        ReDim mudtComm(0)
        mintNumComm = 1
        mudtComm(0).strMessage = strMessage
        mudtComm(0).lngDecay = gobjDX.TickCount + COMM_DECAY
    Else
        'Make a new entry
        If mintNumComm < COMM_MAX Then
            ReDim Preserve mudtComm(mintNumComm)
            mintNumComm = mintNumComm + 1
        End If
        'Bump them up
        For i = mintNumComm - 1 To 1 Step -1
            mudtComm(i).lngDecay = mudtComm(i - 1).lngDecay
            mudtComm(i).strMessage = mudtComm(i - 1).strMessage
        Next i
        'Add new one
        mudtComm(0).strMessage = strMessage
        mudtComm(0).lngDecay = gobjDX.TickCount + COMM_DECAY
    End If

End Sub

Private Sub DisplayBars()

Dim i As Long
Dim lngTemp As Long
Dim sngTemp As Single
Dim bytMarker As Byte
 
    'Display the solid bars
    DDraw.DisplayBar mlngSpriteSolidBar, TEXT_PLAYER_ROW1X + TEXT_PLAYER_ROW1_WIDTH, TEXT_PLAYER_Y + TEXT_PLAYER_Y_BARHEIGHT, CInt(BAR_WIDTH * gudtPlayer.udtSystems.sngEnergy / gudtGenerator(gudtPlayer.udtSystems.bytGenerator).lngMaxBattery)
    If gudtPlayer.udtSystems.sngFuel > 0 Then DDraw.DisplayBar mlngSpriteSolidBar, TEXT_PLAYER_ROW1X + TEXT_PLAYER_ROW1_WIDTH, TEXT_PLAYER_Y + TEXT_PLAYER_Y_BARHEIGHT + TEXT_PLAYER_HEIGHT * 1, CInt(BAR_WIDTH * gudtPlayer.udtSystems.sngGeneratorEnergy / gudtGenerator(gudtPlayer.udtSystems.bytGenerator).lngMaxEnergy)
    DDraw.DisplayBar mlngSpriteSolidBar, TEXT_PLAYER_ROW1X + TEXT_PLAYER_ROW1_WIDTH, TEXT_PLAYER_Y + TEXT_PLAYER_Y_BARHEIGHT + TEXT_PLAYER_HEIGHT * 2, CInt(BAR_WIDTH * gudtPlayer.udtSystems.sngEngineEnergy / gudtEngine(gudtPlayer.udtSystems.bytEngine).lngMaxEnergy)
    DDraw.DisplayBar mlngSpriteSolidBar, TEXT_PLAYER_ROW1X + TEXT_PLAYER_ROW1_WIDTH, TEXT_PLAYER_Y + TEXT_PLAYER_Y_BARHEIGHT + TEXT_PLAYER_HEIGHT * 3, CInt(BAR_WIDTH * gudtPlayer.udtSystems.sngShieldEnergy / gudtShield(gudtPlayer.udtSystems.bytShield).lngMaxEnergy)
    DDraw.DisplayBar mlngSpriteSolidBar, TEXT_PLAYER_ROW1X + TEXT_PLAYER_ROW1_WIDTH, TEXT_PLAYER_Y + TEXT_PLAYER_Y_BARHEIGHT + TEXT_PLAYER_HEIGHT * 4, CInt(BAR_WIDTH * (gudtPlayer.udtSystems.sngWeaponEnergy) / (gudtCannon(gudtPlayer.udtSystems.bytCannon).lngMaxEnergy + gudtLaser(gudtPlayer.udtSystems.bytLaser).lngMaxEnergy))
    'Sum up all cargo + salvage
    lngTemp = 0
    If gudtPlayer.udtCargo.lngNumCargo > 0 Then
        For i = 0 To gudtPlayer.udtCargo.lngNumCargo - 1
            lngTemp = lngTemp + gudtPlayer.udtCargo.lngAmount(i)
        Next i
    End If
    lngTemp = lngTemp + gudtPlayer.udtCargo.lngSalvage
    DDraw.DisplayBar mlngSpriteSolidBar, TEXT_PLAYER_ROW3X + TEXT_PLAYER_ROW3_WIDTH, TEXT_PLAYER_Y + TEXT_PLAYER_Y_BARHEIGHT + TEXT_PLAYER_HEIGHT * 3, CInt(BAR_WIDTH * (lngTemp) / gudtHull(gudtPlayer.udtSystems.bytHull).lngMaxCargo)
    
    'And their sliders..
    DDraw.DisplaySprite mlngSpriteSlider, TEXT_PLAYER_ROW1X + TEXT_PLAYER_ROW1_WIDTH - 3 + gudtPlayer.udtControl.bytGenerator, TEXT_PLAYER_Y + TEXT_PLAYER_Y_SLIDERHEIGHT + TEXT_PLAYER_HEIGHT * 1
    DDraw.DisplaySprite mlngSpriteSlider, TEXT_PLAYER_ROW1X + TEXT_PLAYER_ROW1_WIDTH - 3 + gudtPlayer.udtControl.bytEngine, TEXT_PLAYER_Y + TEXT_PLAYER_Y_SLIDERHEIGHT + TEXT_PLAYER_HEIGHT * 2
    DDraw.DisplaySprite mlngSpriteSlider, TEXT_PLAYER_ROW1X + TEXT_PLAYER_ROW1_WIDTH - 3 + gudtPlayer.udtControl.bytShield, TEXT_PLAYER_Y + TEXT_PLAYER_Y_SLIDERHEIGHT + TEXT_PLAYER_HEIGHT * 3
    DDraw.DisplaySprite mlngSpriteSlider, TEXT_PLAYER_ROW1X + TEXT_PLAYER_ROW1_WIDTH - 3 + gudtPlayer.udtControl.bytWeapons, TEXT_PLAYER_Y + TEXT_PLAYER_Y_SLIDERHEIGHT + TEXT_PLAYER_HEIGHT * 4
    
    'Display "Maximal Efficiency" marker on generator bar
    If glngElapsed > 0 Then
        sngTemp = gudtGenerator(gudtPlayer.udtSystems.bytGenerator).sngOutPut * glngElapsed * (gudtPlayer.udtSystems.lngCrew / gudtHull(gudtPlayer.udtSystems.bytHull).lngMaxCrew)
        If sngTemp > 0 Then
            If gudtPlayer.udtSystems.blnJammerActive = True Then
                bytMarker = ConvByte((((gudtEngine(gudtPlayer.udtSystems.bytEngine).sngConsumption * gudtPlayer.udtSystems.sngEngineEnergy / gudtEngine(gudtPlayer.udtSystems.bytEngine).lngMaxEnergy * glngElapsed) + (gudtShield(gudtPlayer.udtSystems.bytShield).sngConsumption * gudtPlayer.udtSystems.sngShieldEnergy / gudtShield(gudtPlayer.udtSystems.bytShield).lngMaxEnergy * glngElapsed) + ((gudtCannon(gudtPlayer.udtSystems.bytCannon).sngConsumption + gudtLaser(gudtPlayer.udtSystems.bytLaser).sngConsumption) * gudtPlayer.udtSystems.sngWeaponEnergy / (gudtCannon(gudtPlayer.udtSystems.bytCannon).lngMaxEnergy + gudtLaser(gudtPlayer.udtSystems.bytLaser).lngMaxEnergy) * glngElapsed) + (gudtJammer.sngConsumption * glngElapsed)) / sngTemp) * BAR_WIDTH)
            Else
                bytMarker = ConvByte((((gudtEngine(gudtPlayer.udtSystems.bytEngine).sngConsumption * gudtPlayer.udtSystems.sngEngineEnergy / gudtEngine(gudtPlayer.udtSystems.bytEngine).lngMaxEnergy * glngElapsed) + (gudtShield(gudtPlayer.udtSystems.bytShield).sngConsumption * gudtPlayer.udtSystems.sngShieldEnergy / gudtShield(gudtPlayer.udtSystems.bytShield).lngMaxEnergy * glngElapsed) + ((gudtCannon(gudtPlayer.udtSystems.bytCannon).sngConsumption + gudtLaser(gudtPlayer.udtSystems.bytLaser).sngConsumption) * gudtPlayer.udtSystems.sngWeaponEnergy / (gudtCannon(gudtPlayer.udtSystems.bytCannon).lngMaxEnergy + gudtLaser(gudtPlayer.udtSystems.bytLaser).lngMaxEnergy) * glngElapsed)) / sngTemp) * BAR_WIDTH)
            End If
        End If
        If bytMarker <= BAR_WIDTH Then DDraw.DisplaySprite mlngSpriteMarker, TEXT_PLAYER_ROW1X + TEXT_PLAYER_ROW1_WIDTH - 1 + bytMarker, TEXT_PLAYER_Y + TEXT_PLAYER_Y_MARKERHEIGHT + TEXT_PLAYER_HEIGHT * 1
    End If
    
    'Display "Energy Needed for FTL Travel" marker on energy bar
    If gudtPlayer.udtSystems.blnARCD = True Or gudtPlayer.udtSystems.blnFTLD = True Then
        'Is this ARCD?
        If gudtPlayer.udtSystems.blnARCD = True And gudtGenerator(gudtPlayer.udtSystems.bytGenerator).lngMaxBattery > ARCD_CONSUMPTION Then
            bytMarker = CByte((ARCD_CONSUMPTION / gudtGenerator(gudtPlayer.udtSystems.bytGenerator).lngMaxBattery) * BAR_WIDTH)
            DDraw.DisplaySprite mlngSpriteMarker, TEXT_PLAYER_ROW1X + TEXT_PLAYER_ROW1_WIDTH + bytMarker, TEXT_PLAYER_Y + TEXT_PLAYER_Y_MARKERHEIGHT
        End If
        '...or FTLD?
        If gudtPlayer.udtSystems.blnFTLD = True And gudtGenerator(gudtPlayer.udtSystems.bytGenerator).lngMaxBattery > FTLD_CONSUMPTION Then
            bytMarker = CByte((FTLD_CONSUMPTION / gudtGenerator(gudtPlayer.udtSystems.bytGenerator).lngMaxBattery) * BAR_WIDTH)
            DDraw.DisplaySprite mlngSpriteMarker, TEXT_PLAYER_ROW1X + TEXT_PLAYER_ROW1_WIDTH + bytMarker, TEXT_PLAYER_Y + TEXT_PLAYER_Y_MARKERHEIGHT
        End If
    End If
    
    'Display the gradient bars
    DDraw.DisplayBar mlngSpriteGradientBar, TEXT_PLAYER_ROW3X + TEXT_PLAYER_ROW3_WIDTH, TEXT_PLAYER_Y + TEXT_PLAYER_Y_BARHEIGHT, CInt(gudtPlayer.udtSystems.lngArmour / gudtArmour(gudtPlayer.udtSystems.bytArmour).lngMaxArmour * BAR_WIDTH)
    DDraw.DisplayBar mlngSpriteGradientBar, TEXT_PLAYER_ROW3X + TEXT_PLAYER_ROW3_WIDTH, TEXT_PLAYER_Y + TEXT_PLAYER_Y_BARHEIGHT + TEXT_PLAYER_HEIGHT * 1, CInt(gudtPlayer.udtSystems.lngCrew / gudtHull(gudtPlayer.udtSystems.bytHull).lngMaxCrew * BAR_WIDTH)
    DDraw.DisplayBar mlngSpriteGradientBar, TEXT_PLAYER_ROW3X + TEXT_PLAYER_ROW3_WIDTH, TEXT_PLAYER_Y + TEXT_PLAYER_Y_BARHEIGHT + TEXT_PLAYER_HEIGHT * 2, CInt(gudtPlayer.udtSystems.sngFuel / gudtHull(gudtPlayer.udtSystems.bytHull).lngMaxFuel * BAR_WIDTH)

End Sub

Private Sub GetInput()

Dim i As Long
Dim lngTemp As Long
Dim bytMarker As Byte
Dim sngTemp As Single
Dim blnThrusting As Boolean
Dim blnFiringLasers As Boolean
Static blnOKey As Boolean
Static blnAKey As Boolean
Static blnDKey As Boolean
Static blnSKey As Boolean
Static blnFKey As Boolean
Static blnJKey As Boolean
Static blnTabKey As Boolean
Static blnGraveKey As Boolean
Static blnNKey As Boolean
Static blnNoMissiles As Boolean

    'Is the player long dead?
    If (gblnPlayerDead = True) And (glngGameTime > glngPlayerExplodingStart + PLAYER_EXPLODE_CONTINUE_DELAY) Then
        'Check for keypress
        For i = 0 To UBound(gblnKey)
            'If keypress, exit to main menu
            If gblnKey(i) = True Then
                mblnTerminating = True
                mbytNextScreen = DISPLAY_TITLE
                Exit For
            End If
        Next i
        Exit Sub
    End If

    'Is the player at all dead?
    If gblnPlayerDead = True Then Exit Sub

    'Exit on X key or ESC
    If (gblnKey(DIK_X)) Or (gblnKey(DIK_ESCAPE)) Then gblnRunning = False
    
    'Fire missiles
    If gblnKey(DIK_M) Then
        'Check if the player has anything selected on radar
        If gudtPlayer.lngRadarObject >= 0 Then
            'Check if it's time
            If gudtPlayer.udtSystems.lngMissileLastFire + gudtMissile(gudtPlayer.udtSystems.bytMissile).lngFireRate <= glngGameTime Then
                'Do we have any missiles?
                If (gudtPlayer.udtSystems.bytMissile > 0) And (gudtPlayer.udtSystems.intMissileNum > 0) Then
                    'Fire!
                    If CreateMissile(gudtPlayer.udtSystems.bytMissile, gudtPlayer.udtPhysics.dblX, gudtPlayer.udtPhysics.dblY, -1, gudtRadar(gudtPlayer.lngRadarObject).lngObject, gudtPlayer.udtPhysics.sngSpeed, gudtPlayer.udtPhysics.sngHeading, gudtPlayer.udtPhysics.sngFacing) = True Then
                        'Consume a missile
                        gudtPlayer.udtSystems.intMissileNum = gudtPlayer.udtSystems.intMissileNum - 1
                        'Set timer
                        gudtPlayer.udtSystems.lngMissileLastFire = glngGameTime
                    End If
                    'We still have missiles
                    blnNoMissiles = False
                Else
                    'No missiles!
                    If blnNoMissiles = False Then
                        AddComm "Zero missiles remaining."
                        blnNoMissiles = True
                    End If
                End If
            End If
        End If
    Else
        'We may have missiles..
        blnNoMissiles = False
    End If
    
    'Fire lasers
    blnFiringLasers = gudtPlayer.udtAI.blnLaserFire
    If gblnKey(DIK_V) Then
        'Try to fire lasers
        Universe.PlayerFireLaser
    Else
        'End sound
        gudtPlayer.udtAI.blnLaserFire = False
        If blnFiringLasers = True Then DSound.StopSound gudtPlayer.udtAI.lngLaserSound
    End If
    
    'If we WEREN'T firing, but are now, start sound
    If (blnFiringLasers = False) And (gudtPlayer.udtAI.blnLaserFire = True) Then gudtPlayer.udtAI.lngLaserSound = DSound.PlaySound(gudtLaser(gudtPlayer.udtSystems.bytLaser).lngSound, False, True, True, 0, 0)
    'If we WERE firing, but aren't now, stop sound
    If (blnFiringLasers = True) And (gudtPlayer.udtAI.blnLaserFire = False) Then DSound.StopSound gudtPlayer.udtAI.lngLaserSound
    
    'Fire cannons
    If gblnKey(DIK_C) Then
        'Try to fire cannons
        Universe.PlayerFireCannon
    End If
    
    'Set destination to currently highlighted radar object
    If gblnKey(DIK_N) Then
        'Ensure holding key down doesn't repeat
        If blnNKey = False And gudtPlayer.lngRadarObject <> -1 Then
            gudtPlayer.udtAI.lngTarget = gudtRadar(gudtPlayer.lngRadarObject).lngObject
        ElseIf blnNKey = False And gudtPlayer.lngRadarObject = -1 Then
            AddComm "No object selected on scanner."
        End If
        blnNKey = True
    Else
        blnNKey = False
    End If
    
    'Radar tab
    If gblnKey(DIK_TAB) And Not (gblnKey(DIK_LSHIFT) Or gblnKey(DIK_RSHIFT)) Then
        'Ensure holding key down doesn't repeat
        If blnTabKey = False Then
            RadarTab
        End If
        blnTabKey = True
    ElseIf gblnKey(DIK_TAB) And (gblnKey(DIK_LSHIFT) Or gblnKey(DIK_RSHIFT)) Then
        'Ensure holding key down doesn't repeat
        If blnTabKey = False Then
            ShiftRadarTab
        End If
        blnTabKey = True
    Else
        blnTabKey = False
    End If
    'Closest enemy radar tab
    If gblnKey(DIK_GRAVE) And (gblnKey(DIK_LSHIFT) Or gblnKey(DIK_RSHIFT)) Then
        'Ensure holding key down doesn't repeat
        If blnGraveKey = False Then
            RadarClosestEnemyTab
        End If
        blnGraveKey = True
    'Enemy radar tab
    ElseIf gblnKey(DIK_GRAVE) And Not (gblnKey(DIK_LSHIFT)) And Not (gblnKey(DIK_RSHIFT)) Then
        'Ensure holding key down doesn't repeat
        If blnGraveKey = False Then
            RadarEnemyTab
        End If
        blnGraveKey = True
    Else
        blnGraveKey = False
    End If
    
    'Radar range
    If gblnKey(DIK_SUBTRACT) Then
        'Decrease radar range
        gudtPlayer.dblCurrentRange = gudtPlayer.dblCurrentRange - (gudtPlayer.dblCurrentRange * RADAR_STEP * glngElapsed)
        'Ensure it's above min
        If gudtPlayer.dblCurrentRange < MIN_SCANNER_RANGE Then gudtPlayer.dblCurrentRange = MIN_SCANNER_RANGE
    End If
    If gblnKey(DIK_ADD) Then
        'Increase radar range
        gudtPlayer.dblCurrentRange = gudtPlayer.dblCurrentRange + (gudtPlayer.dblCurrentRange * RADAR_STEP * glngElapsed)
        'Ensure it's above min
        If gudtPlayer.dblCurrentRange > gudtScanner(gudtPlayer.udtSystems.bytScanner).dblMaxRange Then gudtPlayer.dblCurrentRange = gudtScanner(gudtPlayer.udtSystems.bytScanner).dblMaxRange
    End If
    
    'Distress
    If gblnKey(DIK_D) Then
        'Ensure holding key down doesn't repeat
        If blnDKey = False Then
            DistressCall gudtPlayer.udtPhysics.dblX, gudtPlayer.udtPhysics.dblY, gudtPlayer.udtInfo.strName, gudtPlayer.udtInfo.bytRace
        End If
        blnDKey = True
    Else
        blnDKey = False
    End If
    
    'Jammers
    If gblnKey(DIK_J) Then
        'Ensure holding key down doesn't repeat
        If blnJKey = False Then
            'Toggle jammer
            If gudtPlayer.udtSystems.blnJammerActive = True Then
                gudtPlayer.udtSystems.blnJammerActive = False
            ElseIf gudtPlayer.udtSystems.blnJammer = True Then
                gudtPlayer.udtSystems.blnJammerActive = True
            End If
        End If
        blnJKey = True
    Else
        blnJKey = False
    End If
    
    'FTL
    If gblnKey(DIK_F) Then
        'Ensure holding key down doesn't repeat
        If blnFKey = False Then
            'Check for requisite speed
            If (gudtPlayer.udtSystems.blnARCD = True Or gudtPlayer.udtSystems.blnFTLD = True) And (gudtPlayer.udtPhysics.sngSpeed < MIN_LIGHT_SPEED And NormalizeSpeed(gudtPlayer.udtPhysics.sngSpeed) <> "0.50c") Then
                'Insufficient speed
                AddComm "Minimum speed of " & NormalizeSpeed(MIN_LIGHT_SPEED) & " required for jump to Faster-Than-Light travel."
            Else
                'ARCD
                If gudtPlayer.udtSystems.blnARCD = True Then
                    If gudtPlayer.udtSystems.blnARCDActive = True Then
                        gudtPlayer.udtSystems.blnARCDActive = False
                    Else
                        'Do we have the energy?
                        If gudtPlayer.udtSystems.sngEnergy >= ARCD_CONSUMPTION Then
                            'Activate ARCD and consume energy
                            gudtPlayer.udtSystems.blnARCDActive = True
                            gudtPlayer.udtSystems.sngEnergy = gudtPlayer.udtSystems.sngEnergy - ARCD_CONSUMPTION
                        Else
                            'We do NOT have the energy!
                            AddComm "Insufficient energy for Faster-Than-Light travel."
                        End If
                    End If
                'FTLD
                ElseIf gudtPlayer.udtSystems.blnFTLD = True Then
                    If gudtPlayer.udtSystems.blnFTLDActive = True Then
                        gudtPlayer.udtSystems.blnFTLDActive = False
                    Else
                        'Do we have the energy?
                        If gudtPlayer.udtSystems.sngEnergy >= FTLD_CONSUMPTION Then
                            'Activate FTLD and consume energy
                            gudtPlayer.udtSystems.blnFTLDActive = True
                            gudtPlayer.udtSystems.sngEnergy = gudtPlayer.udtSystems.sngEnergy - FTLD_CONSUMPTION
                        Else
                            'We do NOT have the energy!
                            AddComm "Insufficient energy for Faster-Than-Light travel."
                        End If
                    End If
                End If
            End If
        End If
        blnFKey = True
    Else
        blnFKey = False
    End If
    
    'Set coords to distress call
    If gblnKey(DIK_S) Then
        'Ensure holding key down doesn't repeat
        If blnSKey = False Then
            'If there HAS been a distress call..
            If gblnDistress Then
                'Set coords
                gudtPlayer.udtAI.lngTarget = TARGET_COORDS
                gudtPlayer.udtAI.dblX = gdblDistressX
                gudtPlayer.udtAI.dblY = gdblDistressY
                'Notify player
                AddComm "Responding to distress call; course laid in."
            Else
                'No distress calls recorded!
                AddComm "No distress calls to respond to."
            End If
        End If
        blnSKey = True
    Else
        blnSKey = False
    End If
    
    'AI
    If gblnKey(DIK_O) Then
        'Ensure holding key down doesn't repeat
        If blnOKey = False Then
            gudtPlayer.udtAI.bytAction = AI_ALLSTOP
            AddComm "All Stop Engaged"
        End If
        blnOKey = True
    Else
        blnOKey = False
    End If
    If gblnKey(DIK_A) Then
        'Ensure that holding the key down doesn't repeatedly toggle the autopilot
        If gudtPlayer.udtAI.bytAction = AI_AUTOPILOT And blnAKey = False Then
            gudtPlayer.udtAI.bytAction = AI_NONE
        ElseIf blnAKey = False And gudtPlayer.udtAI.lngTarget <> TARGET_NONE Then
            gudtPlayer.udtAI.bytAction = AI_AUTOPILOT
        ElseIf blnAKey = False And gudtPlayer.udtAI.lngTarget = TARGET_NONE Then
            'No destination!
            AddComm "Select a destination for autopilot."
        End If
        blnAKey = True
    Else
        blnAKey = False
    End If
    
    'Rotate ship
    If gblnKey(DIK_NUMPAD6) Then
        gudtPlayer.udtPhysics.blnTurningRight = True
        gudtPlayer.udtAI.bytAction = AI_NONE
    End If
    If Not (gblnKey(DIK_NUMPAD6)) Then gudtPlayer.udtPhysics.blnTurningRight = False
    If gblnKey(DIK_NUMPAD4) Then
        gudtPlayer.udtPhysics.blnTurningLeft = True
        gudtPlayer.udtAI.bytAction = AI_NONE
    End If
    If Not (gblnKey(DIK_NUMPAD4)) Then gudtPlayer.udtPhysics.blnTurningLeft = False
    
    'Thrusting
    blnThrusting = gudtPlayer.udtAI.blnThrusting
    gudtPlayer.udtAI.blnThrusting = False
    If gblnKey(DIK_NUMPAD8) Then
        gudtPlayer.udtPhysics.blnThrusting = True
        gudtPlayer.udtAI.blnThrusting = True
        gudtPlayer.udtAI.bytAction = AI_NONE
    End If
    If Not (gblnKey(DIK_NUMPAD8)) Then gudtPlayer.udtPhysics.blnThrusting = False
    If gblnKey(DIK_NUMPAD2) Then
        gudtPlayer.udtPhysics.blnReverseThrusting = True
        gudtPlayer.udtAI.blnThrusting = True
        gudtPlayer.udtAI.bytAction = AI_NONE
    End If
    If Not (gblnKey(DIK_NUMPAD2)) Then gudtPlayer.udtPhysics.blnReverseThrusting = False
    
    'Skip if autopilot is on!
    If (gudtPlayer.udtAI.bytAction <> AI_AUTOPILOT) And (gudtPlayer.udtAI.bytAction <> AI_ALLSTOP) Then
        'If we WEREN'T thrusting, but are now, start sound
        If (blnThrusting = False) And (gudtPlayer.udtAI.blnThrusting = True) Then
            gudtPlayer.udtAI.lngThrustSound = DSound.PlaySound(gudtEngine(gudtPlayer.udtSystems.bytEngine).lngSound, False, True, True)
        'If we WERE thrusting, but aren't now, end sound
        ElseIf (blnThrusting = True) And (gudtPlayer.udtAI.blnThrusting = False) Then
            DSound.StopSound gudtPlayer.udtAI.lngThrustSound
        End If
    End If
    
    'Sliders
    If (gblnKey(DIK_LSHIFT)) Or (gblnKey(DIK_RSHIFT)) Then
        If gblnKey(DIK_1) And (gudtPlayer.udtControl.bytGenerator > 0) Then gudtPlayer.udtControl.bytGenerator = gudtPlayer.udtControl.bytGenerator - 1
        If gblnKey(DIK_2) And (gudtPlayer.udtControl.bytEngine > 0) Then gudtPlayer.udtControl.bytEngine = gudtPlayer.udtControl.bytEngine - 1
        If gblnKey(DIK_3) And (gudtPlayer.udtControl.bytShield > 0) Then gudtPlayer.udtControl.bytShield = gudtPlayer.udtControl.bytShield - 1
        If gblnKey(DIK_4) And (gudtPlayer.udtControl.bytWeapons > 0) Then gudtPlayer.udtControl.bytWeapons = gudtPlayer.udtControl.bytWeapons - 1
        If gblnKey(DIK_5) Then
            If gudtPlayer.udtControl.bytGenerator > 0 Then gudtPlayer.udtControl.bytGenerator = gudtPlayer.udtControl.bytGenerator - 1
            If gudtPlayer.udtControl.bytEngine > 0 Then gudtPlayer.udtControl.bytEngine = gudtPlayer.udtControl.bytEngine - 1
            If gudtPlayer.udtControl.bytShield > 0 Then gudtPlayer.udtControl.bytShield = gudtPlayer.udtControl.bytShield - 1
            If gudtPlayer.udtControl.bytWeapons > 0 Then gudtPlayer.udtControl.bytWeapons = gudtPlayer.udtControl.bytWeapons - 1
        End If
    ElseIf (gblnKey(DIK_RCONTROL)) Or (gblnKey(DIK_LCONTROL)) Then
        If gblnKey(DIK_1) Then gudtPlayer.udtControl.bytGenerator = BAR_WIDTH
        If gblnKey(DIK_2) Then gudtPlayer.udtControl.bytEngine = BAR_WIDTH
        If gblnKey(DIK_3) Then gudtPlayer.udtControl.bytShield = BAR_WIDTH
        If gblnKey(DIK_4) Then gudtPlayer.udtControl.bytWeapons = BAR_WIDTH
        If gblnKey(DIK_5) Then
            gudtPlayer.udtControl.bytGenerator = BAR_WIDTH
            gudtPlayer.udtControl.bytEngine = BAR_WIDTH
            gudtPlayer.udtControl.bytShield = BAR_WIDTH
            gudtPlayer.udtControl.bytWeapons = BAR_WIDTH
        End If
    ElseIf (gblnKey(DIK_LALT)) Or (gblnKey(DIK_RALT)) Then
        If gblnKey(DIK_1) Then gudtPlayer.udtControl.bytGenerator = 0
        If gblnKey(DIK_2) Then gudtPlayer.udtControl.bytEngine = 0
        If gblnKey(DIK_3) Then gudtPlayer.udtControl.bytShield = 0
        If gblnKey(DIK_4) Then gudtPlayer.udtControl.bytWeapons = 0
        If gblnKey(DIK_5) Then
            gudtPlayer.udtControl.bytGenerator = 0
            gudtPlayer.udtControl.bytEngine = 0
            gudtPlayer.udtControl.bytShield = 0
            gudtPlayer.udtControl.bytWeapons = 0
        End If
    Else
        If gblnKey(DIK_1) And (gudtPlayer.udtControl.bytGenerator < BAR_WIDTH) Then gudtPlayer.udtControl.bytGenerator = gudtPlayer.udtControl.bytGenerator + 1
        If gblnKey(DIK_2) And (gudtPlayer.udtControl.bytEngine < BAR_WIDTH) Then gudtPlayer.udtControl.bytEngine = gudtPlayer.udtControl.bytEngine + 1
        If gblnKey(DIK_3) And (gudtPlayer.udtControl.bytShield < BAR_WIDTH) Then gudtPlayer.udtControl.bytShield = gudtPlayer.udtControl.bytShield + 1
        If gblnKey(DIK_4) And (gudtPlayer.udtControl.bytWeapons < BAR_WIDTH) Then gudtPlayer.udtControl.bytWeapons = gudtPlayer.udtControl.bytWeapons + 1
        If gblnKey(DIK_5) Then
            If gudtPlayer.udtControl.bytGenerator < BAR_WIDTH Then gudtPlayer.udtControl.bytGenerator = gudtPlayer.udtControl.bytGenerator + 1
            If gudtPlayer.udtControl.bytEngine < BAR_WIDTH Then gudtPlayer.udtControl.bytEngine = gudtPlayer.udtControl.bytEngine + 1
            If gudtPlayer.udtControl.bytShield < BAR_WIDTH Then gudtPlayer.udtControl.bytShield = gudtPlayer.udtControl.bytShield + 1
            If gudtPlayer.udtControl.bytWeapons < BAR_WIDTH Then gudtPlayer.udtControl.bytWeapons = gudtPlayer.udtControl.bytWeapons + 1
        End If
    End If
    
    'Set to efficiency marker
    If gblnKey(DIK_6) Then
        'Ensure we have some elapsed time!
        If glngElapsed > 0 Then
            sngTemp = gudtGenerator(gudtPlayer.udtSystems.bytGenerator).sngOutPut * glngElapsed * (gudtPlayer.udtSystems.lngCrew / gudtHull(gudtPlayer.udtSystems.bytHull).lngMaxCrew)
            If gudtPlayer.udtSystems.blnJammerActive = True Then
                bytMarker = CByte((((gudtEngine(gudtPlayer.udtSystems.bytEngine).sngConsumption * gudtPlayer.udtSystems.sngEngineEnergy / gudtEngine(gudtPlayer.udtSystems.bytEngine).lngMaxEnergy * glngElapsed) + (gudtShield(gudtPlayer.udtSystems.bytShield).sngConsumption * gudtPlayer.udtSystems.sngShieldEnergy / gudtShield(gudtPlayer.udtSystems.bytShield).lngMaxEnergy * glngElapsed) + ((gudtCannon(gudtPlayer.udtSystems.bytCannon).sngConsumption + gudtLaser(gudtPlayer.udtSystems.bytLaser).sngConsumption) * gudtPlayer.udtSystems.sngWeaponEnergy / (gudtCannon(gudtPlayer.udtSystems.bytCannon).lngMaxEnergy + gudtLaser(gudtPlayer.udtSystems.bytLaser).lngMaxEnergy) * glngElapsed) + (gudtJammer.sngConsumption * glngElapsed)) / sngTemp) * BAR_WIDTH)
            Else
                bytMarker = CByte((((gudtEngine(gudtPlayer.udtSystems.bytEngine).sngConsumption * gudtPlayer.udtSystems.sngEngineEnergy / gudtEngine(gudtPlayer.udtSystems.bytEngine).lngMaxEnergy * glngElapsed) + (gudtShield(gudtPlayer.udtSystems.bytShield).sngConsumption * gudtPlayer.udtSystems.sngShieldEnergy / gudtShield(gudtPlayer.udtSystems.bytShield).lngMaxEnergy * glngElapsed) + ((gudtCannon(gudtPlayer.udtSystems.bytCannon).sngConsumption + gudtLaser(gudtPlayer.udtSystems.bytLaser).sngConsumption) * gudtPlayer.udtSystems.sngWeaponEnergy / (gudtCannon(gudtPlayer.udtSystems.bytCannon).lngMaxEnergy + gudtLaser(gudtPlayer.udtSystems.bytLaser).lngMaxEnergy) * glngElapsed)) / sngTemp) * BAR_WIDTH)
            End If
            If bytMarker < BAR_WIDTH Then
                gudtPlayer.udtControl.bytGenerator = bytMarker
            Else
                gudtPlayer.udtControl.bytGenerator = BAR_WIDTH
            End If
        End If
    End If

    'EASTER EGGS + CHEATS!
    'Penny-Arcade
    Static blnPAloaded As Boolean
    Static lngSpritePA As Long
    If (gblnKey(DIK_A) = True) And (gblnKey(DIK_P) = True) Then
        'Have we loaded it yet?
        If blnPAloaded = False Then
            blnPAloaded = True
            lngSpritePA = DDraw.LoadSprite("pa", 134, 95, 0)
        End If
        'Display!
        DDraw.DisplaySprite lngSpritePA, 50, 50
    End If
    'LUCKY
    If (gblnKey(DIK_L) = True) And (gblnKey(DIK_C) = True) And (gblnKey(DIK_K) = True) Then
        'MAX!
        gudtPlayer.udtSystems.lngArmour = gudtArmour(gudtPlayer.udtSystems.bytArmour).lngMaxArmour
        gudtPlayer.udtSystems.lngCrew = gudtHull(gudtPlayer.udtSystems.bytHull).lngMaxCrew
        gudtPlayer.udtSystems.sngFuel = gudtHull(gudtPlayer.udtSystems.bytHull).lngMaxFuel
        gudtPlayer.udtSystems.sngEnergy = gudtGenerator(gudtPlayer.udtSystems.bytGenerator).lngMaxBattery
    End If

End Sub

Private Sub Initialize()

Dim i As Integer
Dim lngSpriteLoading As Long

    'log
    Log "Tactical", "Initialize", "Initializing tactical display"

    'Need to display a "Boarding Ship..." screen while universe file and sprites are loaded
    lngSpriteLoading = DDraw.LoadSprite("Loading", LOAD_WIDTH, LOAD_HEIGHT, 0)
    DDraw.DisplaySprite lngSpriteLoading, 0, (SCREEN_HEIGHT \ 2) - (LOAD_HEIGHT \ 2)
    DDraw.Flip
    DDraw.FadeIn
    
    'Player is alive and not exploding
    gudtPlayer.udtPhysics.sngFacing = 0
    gudtPlayer.udtPhysics.sngSpeed = 0
    gblnPlayerDead = False
    gblnPlayerExploded = False
    glngPlayerExplodingStart = 0
    glngPlayerExplosionNum = 0
    
    'No explosions or bullets
    glngNumBullets = 0
    Erase gudtBullet
    glngNumExplosions = 0
    Erase gudtExplosion
    
    'Load the universe file
    Universe.LoadUniverse
    
    'Destroy the loading surface
    DDraw.DeleteSprite lngSpriteLoading
    
    'Load our sprites
    mlngSpriteTacticalLeft = DDraw.LoadSprite("TacticalLeft", 19, 474, 0)
    mlngSpriteTacticalRight = DDraw.LoadSprite("TacticalRight", 158, 487, 0)
    mlngSpriteTacticalTop = DDraw.LoadSprite("TacticalTop", 660, 13, 0)
    mlngSpriteTacticalBottom = DDraw.LoadSprite("TacticalBottom", 800, 113, 0)
    mlngSpriteStar = DDraw.LoadSprite("BackStar", 2, 2, 0)
    mlngSpriteGradientBar = DDraw.LoadSprite("GradientBar", 125, 5, 0)
    mlngSpriteSolidBar = DDraw.LoadSprite("SolidBar", 125, 5, 0)
    mlngSpriteSlider = DDraw.LoadSprite("Slider", 5, 13, 0)
    mlngSpriteMarker = DDraw.LoadSprite("Marker", 1, 11, 0)
    
    'Load shield sprites
    mlngSpriteShield(SHIELD_0) = DDraw.LoadSprite("Shield0", 10, 10, 0)
    mlngSpriteShield(SHIELD_1) = DDraw.LoadSprite("Shield1", 20, 20, 0)
    mlngSpriteShield(SHIELD_2) = DDraw.LoadSprite("Shield2", 40, 40, 0)
    mlngSpriteShield(SHIELD_3) = DDraw.LoadSprite("Shield3", 60, 60, 0)
    mlngSpriteShield(SHIELD_4) = DDraw.LoadSprite("Shield4", 80, 80, 0)
    mlngSpriteShield(SHIELD_5) = DDraw.LoadSprite("Shield5", 100, 100, 0)
    mlngSpriteShield(SHIELD_9) = DDraw.LoadSprite("Shield9", 400, 400, 0)
    
    'Load the radar dots
    For i = 0 To RADAR_DOT_NUM
        mudtSpriteRadar.lngEnemy(i) = DDraw.LoadSprite("dote" & (i + 1), (i + 1) * 2, (i + 1) * 2, 0)
        mudtSpriteRadar.lngFriendly(i) = DDraw.LoadSprite("dotf" & (i + 1), (i + 1) * 2, (i + 1) * 2, 0)
        mudtSpriteRadar.lngNeutral(i) = DDraw.LoadSprite("dotn" & (i + 1), (i + 1) * 2, (i + 1) * 2, 0)
        mudtSpriteRadar.lngPlanet(i) = DDraw.LoadSprite("dotp" & (i + 1), (i + 1) * 2, (i + 1) * 2, 0)
        mudtSpriteRadar.lngUnknown(i) = DDraw.LoadSprite("dotu" & (i + 1), (i + 1) * 2, (i + 1) * 2, 0)
    Next i
    mudtSpriteRadar.lngWeapon = DDraw.LoadSprite("dotw", 1, 1, 0)
    
    'Load smoke
    mlngSpriteSmoke(0) = DDraw.LoadSprite("0smoke00", SMOKE_WIDTH, SMOKE_HEIGHT, 0)
    mlngSpriteSmoke(1) = DDraw.LoadSprite("0smoke01", SMOKE_WIDTH, SMOKE_HEIGHT, 0)
    mlngSpriteSmoke(2) = DDraw.LoadSprite("0smoke02", SMOKE_WIDTH, SMOKE_HEIGHT, 0)
    mlngSpriteSmoke(3) = DDraw.LoadSprite("0smoke03", SMOKE_WIDTH, SMOKE_HEIGHT, 0)
    mlngSpriteSmoke(4) = DDraw.LoadSprite("0smoke04", SMOKE_WIDTH, SMOKE_HEIGHT, 0)
    
    'Init the stars array
    InitStars
    
    'Fade out
    DDraw.FadeOut
    DSound.SetFade FADE_OUT_MUSIC
    DDraw.SetFade 0, 0, 0
    
    'We're loaded!
    mblnLoaded = True
    mblnTerminating = False

End Sub

Private Sub InitStars()

Dim i As Long
Dim j As Long

    'Init our star array
    Randomize
    ReDim mudtStar(STAR_NUM)
    For i = 0 To STAR_NUM
        mudtStar(i).dblX = Fix(Rnd * 644) - 322 'Randomize star X coord
        mudtStar(i).dblY = Fix(Rnd * 484) - 242 'Randomize star Y coord
        'Determine star's relspeed
        For j = 0 To STAR_SPEEDS
            If i <= (STAR_NUM \ (STAR_SPEEDS + 1)) * (j + 1) Then
                mudtStar(i).sngRelSpeed = (1 / (STAR_SPEEDS + 2)) * (j + 1)
                Exit For
            End If
        Next j
    Next i

End Sub

Private Sub Terminate()

Dim i As Long

    'Ensure termination go-ahead
    If mblnTerminating = False Or (gblnFadeComplete = False And DDraw.pblnGamma = True) Or (gblnMusicFadeComplete = False And gblnMusic = True) Then Exit Sub
    
    'log
    Log "Tactical", "Terminate", "Terminating tactical display"
    
    'Delete our sprites
    For i = 0 To RADAR_DOT_NUM
        DDraw.DeleteSprite mudtSpriteRadar.lngEnemy(i)
        DDraw.DeleteSprite mudtSpriteRadar.lngFriendly(i)
        DDraw.DeleteSprite mudtSpriteRadar.lngNeutral(i)
        DDraw.DeleteSprite mudtSpriteRadar.lngPlanet(i)
        DDraw.DeleteSprite mudtSpriteRadar.lngUnknown(i)
    Next i
    DDraw.DeleteSprite mudtSpriteRadar.lngWeapon
    DDraw.DeleteSprite mlngSpriteMarker
    DDraw.DeleteSprite mlngSpriteSlider
    DDraw.DeleteSprite mlngSpriteSolidBar
    DDraw.DeleteSprite mlngSpriteGradientBar
    DDraw.DeleteSprite mlngSpriteStar
    DDraw.DeleteSprite mlngSpriteTacticalLeft
    DDraw.DeleteSprite mlngSpriteTacticalRight
    DDraw.DeleteSprite mlngSpriteTacticalBottom
    DDraw.DeleteSprite mlngSpriteTacticalTop
    DDraw.DeleteSprite mlngSpriteStar
    
    'Set as unloaded
    mblnLoaded = False
    mblnTerminating = False
    
    'Load next screen
    gbytDisplay = mbytNextScreen
    
    'If there's no screen to load next, then exit program
    If mbytNextScreen = DISPLAY_NONE Then gblnRunning = False

End Sub
