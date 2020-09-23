Attribute VB_Name = "DSound"
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

'DirectX Variables
Dim mobjDS As DirectSound
Dim mobjMusicBuffer As DirectSoundBuffer

'Fade variables
Public plngFadeValue As Long    'Current fade value
Dim mlngFadeTarget As Long      'Fade target

'Wave array
Private Type BUFFERTYPE
    objBuffer As DirectSoundBuffer
    strResName As String
    blnExists As Boolean
    blnFile As Boolean              'Is it from a file, or from the resource?
End Type
Dim mudtBuffer() As BUFFERTYPE

'Beep sound references
Dim mlngBeep1 As Long
Dim mlngBeep2 As Long

Public Sub Initialize(frmInit As Form)

    'Exit if no sound
    If gblnSound = False Then Exit Sub

    'Initialize DirectSound
    On Local Error GoTo DSINITERROR
    Log "DSound", "Initialize", "Initializing DSound object"
    Set mobjDS = gobjDX.DirectSoundCreate("")
    
    'Set the DirectSound object's cooperative level (Priority gives us sole control)
    Log "DSound", "Initialize", "Setting cooperative level (DSSCL_PRIORITY)"
    mobjDS.SetCooperativeLevel frmInit.hWnd, DSSCL_PRIORITY
    
    'Load the footer
    On Local Error GoTo DSLOADERROR
    LoadWavFooter App.Path & "\resource.snk"
    
    'Init our buffer array
    ReDim mudtBuffer(0)
    
    'Load the beep sounds!
    mlngBeep1 = LoadSound("ComputerBeep1")
    mlngBeep2 = LoadSound("ComputerBeep2")
    SetPan mlngBeep1, DSBPAN_LEFT
    SetPan mlngBeep2, DSBPAN_RIGHT
    
    'Exit before error code
    On Error GoTo 0
    Exit Sub
    
'Error handlers
DSINITERROR:
    TerminateSpecific False, True, True, True
    Log "DSound", "Initialize", "Error initializing DSound!"
    MsgBox "Error initializing DirectSound.  Ensure all other programs are closed and try again.  Failing that, please report to moml31@hotmail.com along with your system specifications, the log.txt file, and any information you think may be helpful."
    End
DSLOADERROR:
    TerminateSpecific True, True, True, True
    Log "DSound", "Initialize", "Error loading initial sound files!"
    MsgBox "Error loading initial sound files.  Please report this to moml31@hotmail.com along with your system specifications, the log.txt file, and any information you think may be helpful."
    End
    
End Sub

Public Function LoadSound(strResName As String) As Long

Dim lngTemp As Long
Dim udtBufferDesc As DSBUFFERDESC
Dim udtFormat As WAVEFORMATEX

    'Exit if no sound
    If gblnSound = False Then Exit Function

    'Find a new spot in the array
    On Local Error GoTo DSLOADSOUNDERROR
    lngTemp = FreeSpot
    
    'Set up a buffer description
    udtBufferDesc.lFlags = DSBCAPS_CTRLFREQUENCY Or DSBCAPS_CTRLPAN Or DSBCAPS_CTRLVOLUME
    
    'Extract the wave info
    ExtractWaveData App.Path & "\resource.snk", strResName
    
    'Set the Wave Format
    With udtFormat
        .nFormatTag = gudtHeader.intFormat
        .nChannels = gudtHeader.intChannels
        .lSamplesPerSec = gudtHeader.lngSamplesPerSec
        .nBitsPerSample = gudtHeader.intBitsPerSample
        .nBlockAlign = gudtHeader.intBlockAlign
        .lAvgBytesPerSec = gudtHeader.lngAvgBytesPerSec
    End With
    
    'Create the buffer
    udtBufferDesc.lBufferBytes = glngChunkSize
    Set mudtBuffer(lngTemp).objBuffer = mobjDS.CreateSoundBuffer(udtBufferDesc, udtFormat)
    
    'Load the buffer with data
    mudtBuffer(lngTemp).objBuffer.WriteBuffer 0, glngChunkSize, gbytData(0), DSBLOCK_ENTIREBUFFER
    
    'Set the udt data
    mudtBuffer(lngTemp).strResName = strResName
    mudtBuffer(lngTemp).blnExists = True
    mudtBuffer(lngTemp).blnFile = False
    
    'Return the value
    LoadSound = lngTemp
    
    'Exit before error code
    On Error GoTo 0
    Exit Function
    
'Error handlers
DSLOADSOUNDERROR:
    TerminateSpecific True, True, True, True
    Log "DSound", "LoadSound", "Error loading sound file! (" & strResName & ")"
    MsgBox "Error loading a sound.  Please report this to moml31@hotmail.com along with your system specifications, the log.txt file, and any information you think may be helpful."
    End

End Function

Public Sub LoadMusic(strFileName As String)

Dim udtBufferDesc As DSBUFFERDESC
Dim udtFormat As WAVEFORMATEX

    'Exit if no sound
    If gblnSound = False Then Exit Sub

    'Ensure we're supposed to be playing music
    If Not (gblnMusic) Then Exit Sub
        
    'log
    Log "DSound", "LoadMusic", "Loading music: " & strFileName & ".dat"
        
    'Ensure the given file exists
    On Local Error GoTo DSMUSICERROR
    If Dir(App.Path & "\Music\" & strFileName & ".dat") = "" Then
        Log "DSound", "LoadMusic", "Music file not found!  Disabling music..."
        gblnMusic = False
        Exit Sub
        'TerminateSpecific True, True, True, True
        'MsgBox "Music file " & strFileName & " not found.  Ensure music DAT files are in a sub-directory called 'Music', or else disable music playback."
        'End
    End If

    'Erase the previous buffer
    Set mobjMusicBuffer = Nothing
    
    'Modify file
    Open App.Path & "\Music\" & strFileName & ".dat" For Binary Access Read Write Lock Write As 1
        'Make it playable
        Put 1, 1, "RIFF"
    Close #1
    
    'Load file
    udtBufferDesc.lFlags = DSBCAPS_CTRLFREQUENCY Or DSBCAPS_CTRLPAN Or DSBCAPS_CTRLVOLUME
    Set mobjMusicBuffer = mobjDS.CreateSoundBufferFromFile(App.Path & "\Music\" & strFileName & ".dat", udtBufferDesc, udtFormat)
    
    'Modify file
    Open App.Path & "\Music\" & strFileName & ".dat" For Binary Access Read Write Lock Write As 1
        'Make it unplayable
        Put 1, 1, "LUCK"
    Close #1

    'Exit before error code
    On Error GoTo 0
    Exit Sub
    
'Error handlers
DSMUSICERROR:
    TerminateSpecific True, True, True, True
    Log "DSound", "LoadMusic", "Error loading music file! (" & strFileName & ")"
    MsgBox "Error loading music.  Please report this to moml31@hotmail.com along with your system specifications, the log.txt file, and any information you think may be helpful."
    End

End Sub

Public Sub DeleteSound(lngSound As Long)

    'Exit if no sound
    If gblnSound = False Then Exit Sub

    'Delete a sound buffer
    On Local Error GoTo DSDELETEERROR
    mudtBuffer(lngSound).blnExists = False
    Set mudtBuffer(lngSound).objBuffer = Nothing

    'Exit before error code
    On Error GoTo 0
    Exit Sub
    
'Error handlers
DSDELETEERROR:
    TerminateSpecific True, True, True, True
    Log "DSound", "LoadSound", "Error deleting sound file! (" & lngSound & ")"
    MsgBox "Error deleting a sound buffer.  Please report this to moml31@hotmail.com along with your system specifications, the log.txt file, and any information you think may be helpful."
    End

End Sub

Public Function PlaySound(ByVal lngSound As Long, Optional blnInterrupt As Boolean = True, Optional blnDuplicate As Boolean = False, Optional blnLooping As Boolean = False, Optional lngPan As Long = 0, Optional lngVolume As Long = 0) As Long

Dim blnFound As Boolean
Dim i As Long

    'Exit if no sound
    If gblnSound = False Then Exit Function

    'Interrupt?
    On Local Error GoTo DSPLAYERROR
    If blnInterrupt = True Then StopSound lngSound
    
    'If the buffer is playing, find another identical one, or duplicate
    If (mudtBuffer(lngSound).objBuffer.GetStatus = DSBSTATUS_PLAYING) And (blnDuplicate = True) Then
        'Try to find another..
        blnFound = False
        For i = 0 To UBound(mudtBuffer)
            'Is this the same res string?
            If (mudtBuffer(i).strResName = mudtBuffer(lngSound).strResName) And (mudtBuffer(i).objBuffer.GetStatus <> DSBSTATUS_PLAYING) Then
                'It is!  Use it!
                blnFound = True
                'Set it as the sound to play
                lngSound = i
                Exit For
            End If
        Next i
        'If we haven't found one, make another!
        If blnFound = False Then
            'Get a new spot
            i = FreeSpot
            'Fill it with data
            mudtBuffer(i).blnExists = True
            mudtBuffer(i).blnFile = mudtBuffer(lngSound).blnFile
            mudtBuffer(i).strResName = mudtBuffer(lngSound).strResName
            Set mudtBuffer(i).objBuffer = mobjDS.DuplicateSoundBuffer(mudtBuffer(lngSound).objBuffer)
            'Set it as the sound to play
            lngSound = i
        End If
    End If
    
    'Set vol + pan
    SetVolume lngSound, lngVolume
    SetPan lngSound, lngPan
    
    'Loop?
    If blnLooping = True Then
        mudtBuffer(lngSound).objBuffer.Play DSBPLAY_LOOPING
    Else
        mudtBuffer(lngSound).objBuffer.Play DSBPLAY_DEFAULT
    End If
    
    'Return the buffer handle
    PlaySound = lngSound
    
    'Exit before error code
    On Error GoTo 0
    Exit Function
    
'Error handlers
DSPLAYERROR:
    TerminateSpecific True, True, True, True
    Log "DSound", "PlaySound", "Error playing sound buffer! (" & lngSound & ", " & blnInterrupt & ", " & blnDuplicate & ", " & blnLooping & ", " & lngPan & ", " & lngVolume & ")"
    MsgBox "Error playing a sound buffer.  Please report this to moml31@hotmail.com along with your system specifications, the log.txt file, and any information you think may be helpful."
    End

End Function

Public Sub PlayMusic()

    'Exit if no sound
    If gblnSound = False Then Exit Sub

    'If we're playing music, then go ahead
    On Local Error GoTo DSMUSICERROR
    If gblnMusic And Not (mobjMusicBuffer Is Nothing) Then mobjMusicBuffer.Play DSBPLAY_DEFAULT

    'Exit before error code
    On Error GoTo 0
    Exit Sub
    
'Error handlers
DSMUSICERROR:
    TerminateSpecific True, True, True, True
    Log "DSound", "PlayMusic", "Error playing music!"
    MsgBox "Error playing music.  Please report this to moml31@hotmail.com along with your system specifications, the log.txt file, and any information you think may be helpful."
    End

End Sub

Public Sub StopMusic()

    'Exit if no sound
    If gblnSound = False Then Exit Sub
    
    'If we're playing music, then stop it!
    On Local Error GoTo DSMUSICSTOPERROR
    If gblnMusic And Not (mobjMusicBuffer Is Nothing) Then mobjMusicBuffer.Stop

    'Exit before error code
    On Error GoTo 0
    Exit Sub
    
'Error handlers
DSMUSICSTOPERROR:
    TerminateSpecific True, True, True, True
    Log "DSound", "StopMusic", "Error stopping music playback!"
    MsgBox "Error stopping music playback.  Please report this to moml31@hotmail.com along with your system specifications, the log.txt file, and any information you think may be helpful."
    End

End Sub

Public Sub StopSound(lngIndex As Long)

    'Exit if no sound
    If gblnSound = False Then Exit Sub
    
    'Stop the sound!
    On Local Error GoTo DSSOUNDSTOPERROR
    mudtBuffer(lngIndex).objBuffer.Stop
    mudtBuffer(lngIndex).objBuffer.SetCurrentPosition 0

    'Exit before error code
    On Error GoTo 0
    Exit Sub
    
'Error handlers
DSSOUNDSTOPERROR:
    TerminateSpecific True, True, True, True
    Log "DSound", "StopSound", "Error stopping sound playback! (" & lngIndex & ")"
    MsgBox "Error stopping sound playback.  Please report this to moml31@hotmail.com along with your system specifications, the log.txt file, and any information you think may be helpful."
    End

End Sub

Private Function FreeSpot() As Long

Dim i As Long

    'Exit if no sound
    If gblnSound = False Then Exit Function

    'Check for empty slots
    For i = 0 To UBound(mudtBuffer)
        If mudtBuffer(i).blnExists = False Then
            FreeSpot = i
            Exit Function
        End If
    Next i
    
    'Make a new slot
    ReDim Preserve mudtBuffer(UBound(mudtBuffer) + 1)
    FreeSpot = UBound(mudtBuffer)

End Function

Public Sub Main()

    'Exit if no sound
    If gblnSound = False Then Exit Sub

    'Handle music fading
    Fade

End Sub

Private Sub Fade()

    'Exit if no sound
    If gblnSound = False Then Exit Sub
    
    'Music fade
    On Local Error GoTo DSFADEERROR
    If gblnMusic = True And gblnMusicFadeComplete = False Then
        'Set the new values
        If plngFadeValue > mlngFadeTarget Then plngFadeValue = plngFadeValue - FADE_MUSIC_STEP
        If plngFadeValue < mlngFadeTarget Then plngFadeValue = plngFadeValue + FADE_MUSIC_STEP
        'Set the volume
        SetMusicVolume plngFadeValue
        'Check for completion
        If plngFadeValue = mlngFadeTarget Then
            gblnMusicFadeComplete = True
            If mlngFadeTarget <= FADE_OUT_MUSIC Then StopMusic
        End If
    End If

    'Exit before error code
    On Error GoTo 0
    Exit Sub
    
'Error handlers
DSFADEERROR:
    TerminateSpecific True, True, True, True
    Log "DSound", "Fade", "Error fading sound buffer!"
    MsgBox "Error fading sound buffer.  Please report this to moml31@hotmail.com along with your system specifications, the log.txt file, and any information you think may be helpful."
    End

End Sub

Public Sub SetFade(lngValue As Long)

    'Exit if no sound
    If gblnSound = False Then Exit Sub
    
    'Set the fade value
    On Local Error GoTo DSSETFADEERROR
    mlngFadeTarget = lngValue
    gblnMusicFadeComplete = False

    'Exit before error code
    On Error GoTo 0
    Exit Sub
    
'Error handlers
DSSETFADEERROR:
    TerminateSpecific True, True, True, True
    Log "DSound", "SetFade", "Error setting sound fade value! (" & lngValue & ")"
    MsgBox "Error setting sound fade value.  Please report this to moml31@hotmail.com along with your system specifications, the log.txt file, and any information you think may be helpful."
    End

End Sub

Public Sub SetMusicVolume(lngValue As Long)

    'Exit if no sound
    If gblnSound = False Then Exit Sub

    'Set the music buffer volume
    On Local Error GoTo DSMUSICVOLUMEERROR
    plngFadeValue = lngValue
    mobjMusicBuffer.SetVolume lngValue

    'Exit before error code
    On Error GoTo 0
    Exit Sub
    
'Error handlers
DSMUSICVOLUMEERROR:
    TerminateSpecific True, True, True, True
    Log "DSound", "SetMusicVolume", "Error setting music volume! (" & lngValue & ")"
    MsgBox "Error setting music volume.  Please report this to moml31@hotmail.com along with your system specifications, the log.txt file, and any information you think may be helpful."
    End

End Sub

Public Sub SetPan(lngBuffer As Long, lngValue As Long)

    'Exit if no sound
    If gblnSound = False Then Exit Sub

    'Set the buffer's pan value
    On Local Error GoTo DSPANERROR
    mudtBuffer(lngBuffer).objBuffer.SetPan lngValue
    
    'Exit before error code
    On Error GoTo 0
    Exit Sub
    
'Error handlers
DSPANERROR:
    TerminateSpecific True, True, True, True
    Log "DSound", "SetPan", "Error setting buffer pan! (" & lngBuffer & ", " & lngValue & ")"
    MsgBox "Error setting buffer pan.  Please report this to moml31@hotmail.com along with your system specifications, the log.txt file, and any information you think may be helpful."
    End

End Sub

Public Function GetPan(lngBuffer As Long) As Long

    'Exit if no sound
    If gblnSound = False Then Exit Function

    'Get the buffer's pan value
    On Local Error GoTo DSGETPANERROR
    GetPan = mudtBuffer(lngBuffer).objBuffer.GetPan

    'Exit before error code
    On Error GoTo 0
    Exit Function
    
'Error handlers
DSGETPANERROR:
    TerminateSpecific True, True, True, True
    Log "DSound", "GetPan", "Error getting buffer pan! (" & lngBuffer & ")"
    MsgBox "Error getting buffer pan.  Please report this to moml31@hotmail.com along with your system specifications, the log.txt file, and any information you think may be helpful."
    End

End Function

Public Sub SetVolume(lngBuffer As Long, lngValue As Long)

    'Exit if no sound
    If gblnSound = False Then Exit Sub

    'Set the buffer's volume value
    On Local Error GoTo DSSETVOLERROR
    mudtBuffer(lngBuffer).objBuffer.SetVolume lngValue
    
    'Exit before error code
    On Error GoTo 0
    Exit Sub
    
'Error handlers
DSSETVOLERROR:
    TerminateSpecific True, True, True, True
    Log "DSound", "SetVolume", "Error setting buffer volume! (" & lngBuffer & ", " & lngValue & ")"
    MsgBox "Error setting buffer volume.  Please report this to moml31@hotmail.com along with your system specifications, the log.txt file, and any information you think may be helpful."
    End

End Sub

Public Function GetVolume(lngBuffer As Long) As Long

    'Exit if no sound
    If gblnSound = False Then Exit Function

    'Get the buffer's volume value
    On Local Error GoTo DSGETVOLERROR
    GetVolume = mudtBuffer(lngBuffer).objBuffer.GetVolume

    'Exit before error code
    On Error GoTo 0
    Exit Function
    
'Error handlers
DSGETVOLERROR:
    TerminateSpecific True, True, True, True
    Log "DSound", "GetVolume", "Error getting buffer volume! (" & lngBuffer & ")"
    MsgBox "Error getting buffer volume.  Please report this to moml31@hotmail.com along with your system specifications, the log.txt file, and any information you think may be helpful."
    End

End Function

Public Sub SetFrequency(lngBuffer As Long, lngValue As Long)

    'Exit if no sound
    If gblnSound = False Then Exit Sub

    'Set the buffer's frequency value
    On Local Error GoTo DSSETFREQERROR
    mudtBuffer(lngBuffer).objBuffer.SetFrequency lngValue

    'Exit before error code
    On Error GoTo 0
    Exit Sub
    
'Error handlers
DSSETFREQERROR:
    TerminateSpecific True, True, True, True
    Log "DSound", "SetFrequency", "Error setting buffer frequency! (" & lngBuffer & ", " & lngValue & ")"
    MsgBox "Error setting buffer frequency.  Please report this to moml31@hotmail.com along with your system specifications, the log.txt file, and any information you think may be helpful."
    End

End Sub

Public Function GetFrequency(lngBuffer As Long) As Long

    'Exit if no sound
    If gblnSound = False Then Exit Function

    'Get the buffer's frequency value
    On Local Error GoTo DSGETFREQERROR
    GetFrequency = mudtBuffer(lngBuffer).objBuffer.GetFrequency

    'Exit before error code
    On Error GoTo 0
    Exit Function
    
'Error handlers
DSGETFREQERROR:
    TerminateSpecific True, True, True, True
    Log "DSound", "GetFrequency", "Error getting buffer frequency! (" & lngBuffer & ")"
    MsgBox "Error getting buffer frequency.  Please report this to moml31@hotmail.com along with your system specifications, the log.txt file, and any information you think may be helpful."
    End

End Function

Public Sub Beep1()

    'Exit if no sound
    If gblnSound = False Then Exit Sub

    'Play the computer beep1 sound
    On Local Error GoTo DSBEEP1ERROR
    PlaySound mlngBeep1, False, True, False

    'Exit before error code
    On Error GoTo 0
    Exit Sub
    
'Error handlers
DSBEEP1ERROR:
    TerminateSpecific True, True, True, True
    Log "DSound", "Beep1", "Error playing BEEP1 buffer!"
    MsgBox "Error playing BEEP1 buffer.  Please report this to moml31@hotmail.com along with your system specifications, the log.txt file, and any information you think may be helpful."
    End

End Sub

Public Sub Beep2()

    'Exit if no sound
    If gblnSound = False Then Exit Sub

    'Play the computer beep2 sound
    On Local Error GoTo DSBEEP2ERROR
    PlaySound mlngBeep2, False, True, False

    'Exit before error code
    On Error GoTo 0
    Exit Sub
    
'Error handlers
DSBEEP2ERROR:
    TerminateSpecific True, True, True, True
    Log "DSound", "Beep2", "Error playing BEEP2 buffer!"
    MsgBox "Error playing BEEP2 buffer.  Please report this to moml31@hotmail.com along with your system specifications, the log.txt file, and any information you think may be helpful."
    End

End Sub

Public Sub DeleteAllBuffers()

Dim i As Long

    'Exit if no sound
    If gblnSound = False Then Exit Sub

    'Terminate every buffer created
    On Local Error GoTo DSDELBUFFERERROR
    For i = 0 To UBound(mudtBuffer)
        Set mudtBuffer(i).objBuffer = Nothing
        mudtBuffer(i).blnExists = False
    Next i
    Set mobjMusicBuffer = Nothing

    'Exit before error code
    On Error GoTo 0
    Exit Sub
    
'Error handlers
DSDELBUFFERERROR:
    TerminateSpecific False, True, True, True
    Log "DSound", "DeleteAllBuffers", "Error deleting all sound buffers! (" & i & ")"
    MsgBox "Error deleting all sound buffers.  Please report this to moml31@hotmail.com along with your system specifications, the log.txt file, and any information you think may be helpful."
    End

End Sub

Public Sub Terminate()

    'Exit if no sound
    If gblnSound = False Then Exit Sub

    'Terminate DirectSound
    On Local Error GoTo DSTERMERROR
    DeleteAllBuffers
    Set mobjDS = Nothing

    'Exit before error code
    On Error GoTo 0
    Exit Sub
    
'Error handlers
DSTERMERROR:
    TerminateSpecific False, True, True, True
    Log "DSound", "Terminate", "Error terminating DSound!"
    MsgBox "Error terminating DirectSound.  Please report this to moml31@hotmail.com along with your system specifications, the log.txt file, and any information you think may be helpful."
    End

End Sub
