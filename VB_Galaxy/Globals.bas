Attribute VB_Name = "Globals"
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

'API Calls
Private Declare Function SetBkMode Lib "gdi32" (ByVal hdc As Long, ByVal nBkMode As Long) As Long
Private Declare Function CreateFont Lib "gdi32" Alias "CreateFontA" (ByVal H As Long, ByVal W As Long, ByVal E As Long, ByVal O As Long, ByVal W As Long, ByVal i As Long, ByVal u As Long, ByVal s As Long, ByVal c As Long, ByVal OP As Long, ByVal CP As Long, ByVal Q As Long, ByVal PAF As Long, ByVal F As String) As Long
Private Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Private Declare Function SetBkColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function LineTo Lib "gdi32.dll" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function MoveToEx Lib "gdi32.dll" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, lpPoint As POINT_TYPE) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Declare Function IntersectRect Lib "user32" (lpDestRect As RECT, lpSrc1Rect As RECT, lpSrc2Rect As RECT) As Long
Declare Function StretchDIBits Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal dx As Long, ByVal dy As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal wSrcWidth As Long, ByVal wSrcHeight As Long, lpBits As Any, lpBitsInfo As BITMAPINFO, ByVal wUsage As Long, ByVal dwRop As Long) As Long
Type POINT_TYPE
  x As Long
  y As Long
End Type
Global Const SRCCOPY = &HCC0020
Global Const DIB_RGB_COLORS = 0

'Bitmap file format structures
Type BITMAPFILEHEADER
    bfType As Integer
    bfSize As Long
    bfReserved1 As Integer
    bfReserved2 As Integer
    bfOffBits As Long
End Type
Type BITMAPINFOHEADER
    biSize As Long
    biWidth As Long
    biHeight As Long
    biPlanes As Integer
    biBitCount As Integer
    biCompression As Long
    biSizeImage As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed As Long
    biClrImportant As Long
End Type
Type RGBQUAD
    rgbBlue As Byte
    rgbGreen As Byte
    rgbRed As Byte
    rgbReserved As Byte
End Type
Type BITMAPINFO
    bmiHeader As BITMAPINFOHEADER
    bmiColors(0 To 255) As RGBQUAD
End Type

Global gudtBMPFileHeader As BITMAPFILEHEADER   'Holds the file header
Global gudtBMPInfo As BITMAPINFO               'Holds the bitmap info
Global gudtBMPData() As Byte                   'Holds the pixel data
'Our footer struct
Type FOOTERTYPE
    strFileName() As String
    lngFileLocation() As Long
End Type
Global gudtFooter As FOOTERTYPE

'Wave file format info
Private Type WAVETYPE
    strHead As String * 12
    strFormatID As String * 4
    lngChunkSize As Long
    intFormat As Integer
    intChannels As Integer
    lngSamplesPerSec As Long
    lngAvgBytesPerSec As Long
    intBlockAlign As Integer
    intBitsPerSample As Integer
End Type
Global gudtHeader As WAVETYPE
Global glngChunkSize As Long
Global gbytData() As Byte
'Our footer struct
Global gudtWavFooter As FOOTERTYPE


'Program flow variables
Global gblnRunning As Boolean                   'Is the game still running?
Global gbytDisplay As Byte                      'Which screen is currently being displayed?
Global Const DISPLAY_TITLE = 0                  'Display the title screen
Global Const DISPLAY_TACTICAL = 1               'Display the tactical screen
Global Const DISPLAY_NONE = 255                 'Display nothing!

'DirectX
Global gobjDX As DirectX7                       'The main directx object

'Screen info
Global Const SCREEN_WIDTH = 800
Global Const SCREEN_HEIGHT = 600
Global Const SCREEN_BITDEPTH = 16
Global Const TACTICAL_BOTTOM = 492
Global Const TACTICAL_TOP = 9
Global Const TACTICAL_RIGHT = 652
Global Const TACTICAL_LEFT = 9
Global Const BAR_WIDTH = 125

'Sound
Global Const VOL_MAX = 0
Global Const VOL_MIN = -10000
Global Const VOL_ATTENUATION = 0.1          'How does volume fade with distance?
Global Const PAN_LEFT = -10000
Global Const PAN_CENTER = 0
Global Const PAN_RIGHT = 10000
Global Const PAN_MAX_DIST = 500
Global Const PAN_ATTENUATION = 0.2

'Fade
Global gblnFadeComplete As Boolean          'Is the gamma fade complete?
Global gblnMusicFadeComplete As Boolean     'Is the music fade complete?
Global Const FADE_OUT_GAMMA = -100          'Gamma fade out value
Global Const FADE_OUT_MUSIC = -4000         'Music fade out value
Global Const FADE_MUSIC_STEP = 40           'Rate of music volume fade (must be a factor of FADE_OUT_MUSIC or we get infinite loop)

'FPS Counter/Timer
Global glngTimer As Long            'Holds system time since last frame was displayed
Global glngElapsed As Long          'MS of elapsed time since last frame
Global glngFrameTimer As Long       'Stores the time since the last FPS display
Global gintFPS As Long              'An FPS storage variable
Global gintFPSCounter As Long       'An FPS counter
Global gblnDisplayFPS As Boolean    'Should we be displaying the fps?

'Used in Pen Setting
Dim mlngPenRef As Long
Dim mlngOldPen As Long

'Used in Font Setting
Dim mlngNewFont As Long
Dim mlngOldFont As Long
Dim mlngRetVal As Long

'Font constants
Public Const FW_NORMAL = 400
Public Const DEFAULT_CHARSET = 1
Public Const OUT_TT_ONLY_PRECIS = 7
Public Const CLIP_DEFAULT_PRECIS = 0
Public Const CLIP_LH_ANGLES = &H10
Public Const PROOF_QUALITY = 2
Public Const TRUETYPE_FONTTYPE = &H4

'Physics constants
Global Const Pi = 3.14159

'Size of ship pictures
Global Const SHIPPIC_WIDTH = 105
Global Const SHIPPIC_HEIGHT = 105

'Game Options
Global gblnMusic As Boolean     'Are we playing music?
Global gblnSound As Boolean     'Are we playing sounds?
Global gblnVSYNC As Boolean     'Disable VSYNC?

Sub ExtractData(strFileName As String, strFile As String)

Dim lngOffset As Long
Dim intBMPFile As Integer
Dim i As Integer
Dim j As Integer

    'Ensure the file exists
    If Dir(strFileName) = "" Then
        DDraw.RestoreDisplay
        MsgBox "Error.  " & strFileName & " not found."
        End
    End If

    'Open the bitmap
    intBMPFile = FreeFile()
    Open strFileName For Binary Access Read Lock Write As intBMPFile
        'Find this bitmap in the footer
        lngOffset = -1
        For j = 0 To UBound(gudtFooter.strFileName)
            If LCase(gudtFooter.strFileName(j)) = LCase(strFile) Then
                lngOffset = gudtFooter.lngFileLocation(j)
                Exit For
            End If
        Next j
        'If there is no such file, then raise error
        If lngOffset = -1 Then
            DDraw.RestoreDisplay
            MsgBox strFile & " not found.  Resource may be outdated."
            End
        End If
        'Fill the File Header structure
        Get intBMPFile, lngOffset, gudtBMPFileHeader
        'Fill the Info structure
        Get intBMPFile, , gudtBMPInfo
        'Size the BMPData array
        ReDim gudtBMPData(gudtBMPInfo.bmiHeader.biSizeImage)
        'Fill the BMPData array
        Get intBMPFile, , gudtBMPData
    Close intBMPFile
    
End Sub

Sub ExtractWaveData(strFileName As String, strFile As String)

Dim lngOffset As Long
Dim intWAVFile As Integer
Dim i As Integer
Dim j As Integer

    'Ensure the file exists
    If Dir(strFileName) = "" Then
        DDraw.RestoreDisplay
        MsgBox "Error.  " & strFileName & " not found."
        End
    End If

    'Open the wave
    intWAVFile = FreeFile()
    Open strFileName For Binary Access Read Lock Write As intWAVFile
        'Find this wave in the footer
        lngOffset = -1
        For j = 0 To UBound(gudtWavFooter.strFileName)
            If LCase(gudtWavFooter.strFileName(j)) = LCase(strFile) Then
                lngOffset = gudtWavFooter.lngFileLocation(j)
                Exit For
            End If
        Next j
        'If there is no such file, then raise error
        If lngOffset = -1 Then
            DDraw.RestoreDisplay
            MsgBox strFile & " not found.  Resource may be outdated."
            End
        End If
        'Fill the File Header structure
        Get intWAVFile, lngOffset, gudtHeader
        'Fill the Info structure
        Get intWAVFile, , glngChunkSize
        'Size the BMPData array
        ReDim gbytData(glngChunkSize)
        'Fill the BMPData array
        Get intWAVFile, , gbytData
    Close intWAVFile
    
End Sub

Function ExtractFilename(strFileName As String) As String

Dim strTemp As String

    'Remove the path from the filename
    strTemp = strFileName
    Do While InStr(1, strTemp, "\") <> 0
        strTemp = Right(strTemp, Len(strTemp) - InStr(1, strTemp, "\"))
    Loop
    ExtractFilename = Left(strTemp, InStr(1, strTemp, ".") - 1)

End Function

Private Function FileSize(lngWidth As Long, lngHeight As Long) As Long

    'Return the size of the image portion of the bitmap
    If lngWidth Mod 4 > 0 Then
        FileSize = ((lngWidth \ 4) + 1) * 4 * lngHeight - 1
    Else
        FileSize = lngWidth * lngHeight - 1
    End If

End Function

Sub LoadFooter(strFileName As String)

Dim lngOffset As Long
Dim intBMPFile As Integer

    'Ensure the file exists
    If Dir(strFileName) = "" Then
        DDraw.RestoreDisplay
        MsgBox "Error.  " & strFileName & " not found."
        End
    End If

    'Load the footer struct
    intBMPFile = FreeFile()
    Open strFileName For Binary Access Read Lock Write As intBMPFile
        Get intBMPFile, , lngOffset
        Get intBMPFile, lngOffset, gudtFooter
    Close intBMPFile

End Sub

Sub LoadWavFooter(strFileName As String)

Dim lngOffset As Long
Dim intWAVFile As Integer

    'Ensure the file exists
    If Dir(strFileName) = "" Then
        DDraw.RestoreDisplay
        MsgBox "Error.  " & strFileName & " not found."
        End
    End If

    'Load the footer struct
    intWAVFile = FreeFile()
    Open strFileName For Binary Access Read Lock Write As intWAVFile
        Get intWAVFile, , lngOffset
        Get intWAVFile, lngOffset, gudtWavFooter
    Close intWAVFile

End Sub

Sub AddVectors(ByVal sngMag1 As Single, ByVal sngDir1 As Single, ByVal sngMag2 As Single, ByVal sngDir2 As Single, ByRef sngMagResult As Single, ByRef sngDirResult As Single)

Dim dblXComp As Double
Dim dblYComp As Double

    'Determine the X and Y components of the resultant vector
    dblXComp = sngMag1 * Sin(sngDir1) + sngMag2 * Sin(sngDir2)
    dblYComp = sngMag1 * Cos(sngDir1) + sngMag2 * Cos(sngDir2)
    'Determine the resultant magnitude
    sngMagResult = Sqr(dblXComp ^ 2 + dblYComp ^ 2)
    'Calculate the resultant heading
    sngDirResult = FindAngleFromComps(dblXComp, dblYComp)

End Sub

Function FindAngleFromComps(dblX As Double, dblY As Double) As Single

    'Determine the angle to the given set of coords
    If Sgn(dblY) > 0 Then FindAngleFromComps = Atn(dblX / dblY)
    If Sgn(dblY) < 0 Then FindAngleFromComps = Atn(dblX / dblY) + Pi
    If dblY = 0 And Sgn(dblX) < 0 Then FindAngleFromComps = 3 * Pi / 2
    If dblY = 0 And Sgn(dblX) > 0 Then FindAngleFromComps = Pi / 2

End Function

Function FixAngle(ByVal sngAngle As Single) As Single

Dim blnFixed As Boolean 'Has the angle been fixed?

    'Loop until the angle is between 0 and 2Pi
    blnFixed = False
    Do While Not (blnFixed)
        If sngAngle > 2 * Pi Then
            sngAngle = sngAngle - (2 * Pi)
        ElseIf sngAngle < 0 Then
            sngAngle = sngAngle + (2 * Pi)
        Else
            blnFixed = True
        End If
    Loop
    
    'Return the value
    FixAngle = sngAngle

End Function

Function FindAngle(dblX1 As Double, dblY1 As Double, dblX2 As Double, dblY2 As Double) As Double

Dim dblXComp As Double
Dim dblYComp As Double

    'Find the angle between the 2 coords
    dblXComp = dblX2 - dblX1
    dblYComp = dblY1 - dblY2
    If Sgn(dblYComp) > 0 Then FindAngle = Atn(dblXComp / dblYComp)
    If Sgn(dblYComp) < 0 Then FindAngle = Atn(dblXComp / dblYComp) + Pi
    If dblYComp = 0 And Sgn(dblXComp) < 0 Then FindAngle = 3 * Pi / 2
    If dblYComp = 0 And Sgn(dblXComp) > 0 Then FindAngle = Pi / 2

End Function

Function GetDist(dblX1 As Double, dblY1 As Double, dblX2 As Double, dblY2 As Double) As Double

    'Return the distance between the two podbls (I love you, Mr. Pythagoras)
    GetDist = Sqr((dblX1 - dblX2) ^ 2 + (dblY1 - dblY2) ^ 2)

End Function

Function ConvToSignedValue(lngValue As Long) As Integer

    'Cheezy method for converting to signed integer
    If lngValue <= 32767 Then
        ConvToSignedValue = CInt(lngValue)
        Exit Function
    End If
    
    ConvToSignedValue = CInt(lngValue - 65535)

End Function

Function ConvToUnSignedValue(intValue As Integer) As Long

    'Cheezy method for converting to unsigned integer
    If intValue >= 0 Then
        ConvToUnSignedValue = intValue
        Exit Function
    End If
    
    ConvToUnSignedValue = intValue + 65535

End Function

Sub Timer()

    'Determine the time that has elapsed since the last frame was displayed
    glngElapsed = gobjDX.TickCount() - glngTimer
    'Add to game timer
    glngGameTime = glngGameTime + glngElapsed
    'Reset the general timer
    glngTimer = gobjDX.TickCount()
    'Check if one second has elapsed
    If gobjDX.TickCount() - glngFrameTimer >= 1000 Then
        'Set the FPS storage var, and reset the FPS counter/timer
        gintFPS = gintFPSCounter + 1
        gintFPSCounter = 0
        glngFrameTimer = gobjDX.TickCount()
    Else
        'If a second hasn't elapsed, add to the FPS counter
        gintFPSCounter = gintFPSCounter + 1
    End If

End Sub

Sub Motion(ByRef dblX As Double, ByRef dblY As Double, ByVal sngSpeed As Single, ByVal sngHeading As Single)

    'Move an object w.r.t. its speed
    dblX = dblX + sngSpeed * Sin(sngHeading) * glngElapsed
    dblY = dblY - sngSpeed * Cos(sngHeading) * glngElapsed
    
End Sub

Function SeekTarget(sngObjectSpeed As Single, sngObjectAccel As Single, sngObjectHeading As Single, sngObjectFacing As Single, dblObjectX As Double, dblObjectY As Double, sngTargetSpeed As Single, sngTargetHeading As Single, dblTargetX As Double, dblTargetY As Double, ByRef sngDesiredFacing As Single, Optional sngMinDist As Single = 0, Optional sngSeekDist As Single = 0, Optional sngTargetBias As Single = 0, Optional sngCannonSpeed As Single = 0) As Byte

Dim sngMagDiff As Single
Dim sngDirDiff As Single
Dim sngTarDist As Single
Dim sngTarAngle As Single
Dim sngTemp As Single
Dim sngSig As Single

    'Zero the function
    SeekTarget = 0

    'Calc distance
    sngTarDist = CSng(GetDist(dblObjectX, dblObjectY, dblTargetX, dblTargetY))
    
    'Calc angle
    sngTarAngle = CSng(FindAngle(dblObjectX, dblObjectY, dblTargetX, dblTargetY))
    
    'Subtract velocity vectors
    AddVectors sngTargetSpeed, sngTargetHeading, sngObjectSpeed, sngObjectHeading + Pi, sngMagDiff, sngDirDiff
    
    'If we're far away, seek and thrust!
    If sngTarDist > sngMinDist Then
        'Add approach vector, determine desired facing
        If sngMagDiff <> 0 Then AddVectors sngObjectAccel * (sngTarDist - sngSeekDist) / sngMagDiff, sngTarAngle + sngTargetBias, sngMagDiff, sngDirDiff, sngSig, sngDesiredFacing
        'Ensure we have a significant difference
        If sngSig < MIN_VECTOR_SPEED_DIFF And sngObjectSpeed > MIN_VECTOR_SPEED_DIFF Then Exit Function
        'Set thrust direction
        If (FixAngle(sngObjectFacing - sngTarAngle + Pi / 2) < Pi) And (FixAngle(sngDesiredFacing - sngTarAngle + Pi / 2) >= Pi) Then
            'Thrust reverse
            SeekTarget = SeekTarget Or ACTION_REVERSETHRUST
            sngDesiredFacing = sngDesiredFacing + Pi
        Else
            'Thrust forward
            SeekTarget = SeekTarget Or ACTION_THRUST
        End If
        'Adjust facing
        sngTemp = FixAngle(sngDesiredFacing - sngObjectFacing)
        If sngTemp <= 2 * Pi / (FRAME_NUM + 1) Or sngTemp >= 2 * Pi - (2 * Pi / (FRAME_NUM + 1)) Then
            SeekTarget = SeekTarget Or ACTION_NOTURN
        ElseIf sngTemp >= Pi Then
            SeekTarget = SeekTarget Or ACTION_LEFT
        ElseIf sngTemp <= Pi Then
            SeekTarget = SeekTarget Or ACTION_RIGHT
        End If
    'Otherwise, face the target
    Else
        SeekTarget = ACTION_NOTHRUST
        AccurateShot dblTargetX, dblTargetY, sngTargetSpeed, sngTargetHeading, dblObjectX, dblObjectY, sngObjectSpeed, sngObjectHeading, sngCannonSpeed, sngTarAngle
        sngTemp = FixAngle(sngTarAngle - sngObjectFacing)
        If sngTemp <= 2 * Pi / (FRAME_NUM + 1) Or sngTemp >= 2 * Pi - (2 * Pi / (FRAME_NUM + 1)) Then
            SeekTarget = SeekTarget Or ACTION_NOTURN
            sngDesiredFacing = sngTarAngle
        ElseIf sngTemp >= Pi Then
            SeekTarget = SeekTarget Or ACTION_LEFT
        ElseIf sngTemp <= Pi Then
            SeekTarget = SeekTarget Or ACTION_RIGHT
        End If
    End If

End Function

Sub AccurateShot(dblTargetX As Double, dblTargetY As Double, sngTargetSpeed As Single, sngTargetHeading As Single, dblSourceX As Double, dblSourceY As Double, sngSourceSpeed As Single, sngSourceHeading As Single, sngProjectileSpeed As Single, ByRef sngAccurateHeading As Single)

Dim dblDeltaX As Double
Dim dblDeltaY As Double
Dim sngDeltaSpeed As Single
Dim sngDeltaHeading As Single
Dim dblResultX As Double
Dim dblResultY As Double
Dim sngTResult As Single
Dim blnPossible As Boolean

Dim a As Single
Dim b As Single
Dim c As Single
Dim sq As Single
Dim t1 As Single
Dim t2 As Single

    'Assume it's possible
    blnPossible = True

    'Determine the relative location of the target
    dblDeltaX = dblTargetX - dblSourceX
    dblDeltaY = dblTargetY - dblSourceY
    
    'Subtract the velocity vectors to find the relative velocity
    AddVectors sngTargetSpeed, sngTargetHeading, sngSourceSpeed, sngSourceHeading + Pi, sngDeltaSpeed, sngDeltaHeading

    'Set up the quadratic equation's variables
    a = (sngProjectileSpeed ^ 2 - sngDeltaSpeed ^ 2)
    b = -(2 * sngDeltaSpeed * (dblDeltaX * Sin(sngDeltaHeading) - dblDeltaY * Cos(sngDeltaHeading)))
    c = -(dblDeltaX ^ 2 + dblDeltaY ^ 2)
    
    'Ensure there's no problem with the square root, and no divide by zero
    sq = (b ^ 2) - (4 * a * c)
    If (sq < 0) Or (a = 0) Then
        blnPossible = False
    Else
        'We're good to go, get the two results of the quadratic
        t1 = (-b - Sqr(sq)) / (2 * a)
        t2 = (-b + Sqr(sq)) / (2 * a)
        'Is the first Time value the optimal one?
        If t1 > 0 And t1 < t2 Then
            sngTResult = t1
        ElseIf t2 > 0 Then
            sngTResult = t2
        Else
            blnPossible = False
        End If
    End If
    
    'Is there a solution?
    If blnPossible Then
        'Where will the target be, in sngTResult seconds?
        dblResultX = dblTargetX + sngDeltaSpeed * Sin(sngDeltaHeading) * sngTResult
        dblResultY = dblTargetY - sngDeltaSpeed * Cos(sngDeltaHeading) * sngTResult
        'Return the angle to hit the target
        sngAccurateHeading = FindAngle(dblSourceX, dblSourceY, dblResultX, dblResultY)
    Else
        'It's not possible, just shoot straight at 'em
        sngAccurateHeading = FindAngle(dblSourceX, dblSourceY, dblTargetX, dblTargetY)
    End If

End Sub

Function SeekTargetNoRev(sngObjectSpeed As Single, sngObjectAccel As Single, sngObjectHeading As Single, sngObjectFacing As Single, dblObjectX As Double, dblObjectY As Double, sngTargetSpeed As Single, sngTargetHeading As Single, dblTargetX As Double, dblTargetY As Double, ByRef sngDesiredFacing As Single, Optional sngMinDist As Single = 0, Optional sngSeekDist As Single = 0, Optional sngTargetBias As Single = 0) As Byte

Dim sngMagDiff As Single
Dim sngDirDiff As Single
Dim sngTarDist As Single
Dim sngTarAngle As Single
Dim sngTemp As Single
Dim sngSig As Single

    'Zero the function
    SeekTargetNoRev = 0

    'Calc distance
    sngTarDist = CSng(GetDist(dblObjectX, dblObjectY, dblTargetX, dblTargetY))
    
    'Calc angle
    sngTarAngle = CSng(FindAngle(dblObjectX, dblObjectY, dblTargetX, dblTargetY))
    
    'Subtract velocity vectors
    AddVectors sngTargetSpeed, sngTargetHeading, sngObjectSpeed, sngObjectHeading + Pi, sngMagDiff, sngDirDiff
    
    'If we're far away, seek and thrust!
    If sngTarDist > sngMinDist Then
        'Add approach vector, determine desired facing
        If sngMagDiff <> 0 Then AddVectors sngObjectAccel * (sngTarDist - sngSeekDist) / sngMagDiff, sngTarAngle + sngTargetBias, sngMagDiff, sngDirDiff, sngSig, sngDesiredFacing
        'Ensure we have a significant difference
        If sngSig < MIN_VECTOR_SPEED_DIFF And sngObjectSpeed > MIN_VECTOR_SPEED_DIFF Then Exit Function
        'Set thrust direction
        SeekTargetNoRev = SeekTargetNoRev Or ACTION_THRUST
        'Adjust facing
        sngTemp = FixAngle(sngDesiredFacing - sngObjectFacing)
        If sngTemp <= 2 * Pi / (FRAME_NUM + 1) Or sngTemp >= 2 * Pi - (2 * Pi / (FRAME_NUM + 1)) Then
            SeekTargetNoRev = SeekTargetNoRev Or ACTION_NOTURN
        ElseIf sngTemp >= Pi Then
            SeekTargetNoRev = SeekTargetNoRev Or ACTION_LEFT
        ElseIf sngTemp <= Pi Then
            SeekTargetNoRev = SeekTargetNoRev Or ACTION_RIGHT
        End If
    'Otherwise, face the target
    Else
        SeekTargetNoRev = ACTION_NOTHRUST
        sngTemp = FixAngle(sngTarAngle - sngObjectFacing)
        If sngTemp <= 2 * Pi / (FRAME_NUM + 1) Or sngTemp >= 2 * Pi - (2 * Pi / (FRAME_NUM + 1)) Then
            SeekTargetNoRev = SeekTargetNoRev Or ACTION_NOTURN
            sngDesiredFacing = sngTarAngle
        ElseIf sngTemp >= Pi Then
            SeekTargetNoRev = SeekTargetNoRev Or ACTION_LEFT
        ElseIf sngTemp <= Pi Then
            SeekTargetNoRev = SeekTargetNoRev Or ACTION_RIGHT
        End If
    End If

End Function

Public Function FaceTarget(dblObjectX As Double, dblObjectY As Double, dblTargetX As Double, dblTargetY As Double, sngFacing As Single, ByRef sngDesiredFacing As Single) As Byte

Dim sngTemp As Single

    'Find the angle between the coords
    sngTemp = FixAngle(CSng(FindAngle(dblObjectX, dblObjectY, dblTargetX, dblTargetY)) - sngFacing)
    sngDesiredFacing = FixAngle(CSng(FindAngle(dblObjectX, dblObjectY, dblTargetX, dblTargetY)))
    
    'Determine action
    If sngTemp <= 2 * Pi / (FRAME_NUM + 1) Or sngTemp >= 2 * Pi - (2 * Pi / (FRAME_NUM + 1)) Then
        FaceTarget = FaceTarget Or ACTION_NOTURN
    ElseIf sngTemp >= Pi Then
        FaceTarget = FaceTarget Or ACTION_LEFT
    ElseIf sngTemp <= Pi Then
        FaceTarget = FaceTarget Or ACTION_RIGHT
    End If

End Function

Sub ShowText(strText As String, intX As Integer, intY As Integer, lngColour As Long, lngDC As Long, Optional intOpaque As Integer)

    'This function writes text to the backbuffer
    Call SetBkColor(lngDC, 0)
    If intOpaque = 1 Then Call SetBkMode(lngDC, intOpaque)
    Call SetTextColor(lngDC, lngColour)
    Call TextOut(lngDC, intX, intY, strText, Len(strText))

End Sub

Sub LineDraw(lngX1 As Long, lngY1 As Long, lngX2 As Long, lngY2 As Long, lngDC As Long)

Dim udtPoint As POINT_TYPE
Dim lngPenRef As Long
Dim lngOldPen As Long

    'This routine draws a box of specific colour on the display
    Call MoveToEx(lngDC, lngX1, lngY1, udtPoint)    'Move current pen x,y
    Call LineTo(lngDC, lngX2, lngY2)                'Draw line from current x,y to given x,y

End Sub

Public Sub SetPen(lngDC As Long, intWidth As Integer, lngColour As Long)

    'Create the pen object and apply it
    mlngPenRef = CreatePen(0, intWidth, lngColour)
    mlngOldPen = SelectObject(lngDC, mlngPenRef)

End Sub

Public Sub RemovePen(lngDC As Long)

    'Remove the pen object and delete it
    DeleteObject SelectObject(lngDC, mlngOldPen)
    DeleteObject mlngPenRef

End Sub

Public Sub SetFont(lngDC As Long, strFontName As String, intFontSize As Integer)

Dim nHeight As Long
Dim nWidth As Long
Dim nEscapement As Long
Dim fnWeight As Long
Dim fbItalic As Long
Dim fbUnderline As Long
Dim fbStrikeOut As Long
Dim fbCharSet As Long
Dim fbOutputPrecision As Long
Dim fbClipPrecision As Long
Dim fbQuality As Long
Dim fbPitchAndFamily As Long
Dim sFont As String

    'Sets up the new font
    sFont = strFontName
    fnWeight = FW_NORMAL
    nHeight = intFontSize
    nWidth = 0
    nEscapement = 0
    fbItalic = 0
    fbUnderline = 0
    fbStrikeOut = 0
    fbCharSet = DEFAULT_CHARSET
    fbOutputPrecision = OUT_TT_ONLY_PRECIS
    fbClipPrecision = CLIP_LH_ANGLES Or CLIP_DEFAULT_PRECIS
    fbQuality = PROOF_QUALITY
    fbPitchAndFamily = TRUETYPE_FONTTYPE
    
    'Makes the new font
    mlngNewFont = CreateFont(nHeight, nWidth, nEscapement, 0, fnWeight, fbItalic, fbUnderline, fbStrikeOut, fbCharSet, fbOutputPrecision, fbClipPrecision, fbQuality, fbPitchAndFamily, sFont)
    
    'Selects the font onto the surface
    mlngOldFont = SelectObject(lngDC, mlngNewFont)

End Sub

Public Sub RemoveFont(lngDC As Long)

    'Returns the old font
    If mlngOldFont <> 0 Then mlngNewFont = SelectObject(lngDC, mlngOldFont)
    mlngRetVal = DeleteObject(mlngNewFont)
    
End Sub

Function RaceName(bytRace As Byte) As String

    'Return the name of given race
    Select Case bytRace
        Case RACE_TERRAN
            RaceName = "Terran"
        Case RACE_KALE
            RaceName = "Kale"
        Case RACE_PRAEMALI
            RaceName = "Praemali"
        Case RACE_HANTAKAS
            RaceName = "Hantakas"
        Case RACE_ALTAIRIAN
            RaceName = "Altairian"
        Case RACE_GRAME
            RaceName = "Grame"
        Case RACE_VEGAN
            RaceName = "Vegan"
        Case RACE_ULWAR
            RaceName = "Ulwar"
        Case RACE_TULONI
            RaceName = "Tuloni"
        Case RACE_SICARIUS
            RaceName = "Sicarius"
        Case RACE_INDEPENDENT
            RaceName = "Independent"
        Case RACE_PLAYER
            RaceName = "Independent"
        Case RACE_PLANET
            RaceName = "Planet/Star"
    End Select

End Function

Sub DistressCall(dblX As Double, dblY As Double, strName As String, bytRace As Byte)

Dim strTemp As String

    'Ensure there's no ship with an active jammer nearby

    'Scan through all ships within range of transmission and allow them a chance to take action
    
    'Is player within range?
    If GetDist(gudtPlayer.udtPhysics.dblX, gudtPlayer.udtPhysics.dblY, dblX, dblY) <= DISTRESS_RANGE Then
        'Set coords
        gblnDistress = True
        gdblDistressX = dblX
        gdblDistressY = dblY
        'Display message
        Tactical.AddComm "Distress Call: This is the " & RaceName(bytRace) & " ship " & strName & ".  Request assist!"
    End If

End Sub

Function Trunc(sngValue As Single, lngDigits As Long) As String

Dim strTemp As String
Dim lngNumZeros As Long
Dim i As Long


    'Truncate to two decimal places
    strTemp = ((CStr(sngValue) * (10 ^ lngDigits)) \ 1) / (10 ^ lngDigits)
    If InStr(1, strTemp, ".") = 0 Then
        lngNumZeros = lngDigits
        strTemp = strTemp & "."
    ElseIf Len(strTemp) - InStr(1, strTemp, ".") < lngDigits Then
        lngNumZeros = lngDigits - (Len(strTemp) - InStr(1, strTemp, "."))
    End If
    
    'Add training zeros
    If lngNumZeros > 0 Then
        For i = 1 To lngNumZeros
            strTemp = strTemp & "0"
        Next i
    End If
    
    'Return the value!
    Trunc = strTemp
    
End Function

Function NormalizeDistance(sngDistance As Single) As String

Dim sngTemp As Single
Dim blnLY As Boolean

    'Convert to AU
    sngTemp = sngDistance / NORMALIZE_DISTANCE_AU
    
    'Check if we should convert to LY
    If sngTemp >= 99.99 Then
        blnLY = True
        sngTemp = sngTemp / NORMALIZE_DISTANCE_LY
    End If

    'Return formatted data
    If blnLY Then
        NormalizeDistance = Trunc(sngTemp, 4) & " ly"
    Else
        NormalizeDistance = Trunc(sngTemp, 2) & " AU"
    End If

End Function

Function NormalizeSpeed(sngSpeed As Single) As String

Dim sngTemp As Single

    'Convert to "c"
    sngTemp = sngSpeed / NORMALIZE_SPEED
    
    'Return formatted data
    NormalizeSpeed = Trunc(sngTemp, 2) & "c"
    
End Function

Sub BoxDraw(lngX1 As Long, lngY1 As Long, lngX2 As Long, lngY2 As Long, lngDC As Long)

    'Draw the box!
    LineDraw lngX1, lngY1, lngX2, lngY1, lngDC
    LineDraw lngX2, lngY1, lngX2, lngY2, lngDC
    LineDraw lngX2, lngY2, lngX1, lngY2, lngDC
    LineDraw lngX1, lngY2, lngX1, lngY1, lngDC

End Sub

Function AngleDifference(sngAngle1 As Single, sngAngle2 As Single) As Single

    'Find the absolute difference between two angles
    AngleDifference = Abs(FixAngle(sngAngle1) - FixAngle(sngAngle2))

End Function

Function ConvByte(dblValue As Double) As Byte

    'Return the byte value
    If dblValue > 255 Then
        ConvByte = 255
    ElseIf dblValue < 0 Then
        ConvByte = 0
    Else
        ConvByte = CByte(dblValue)
    End If

End Function

Function NormalizeLogScale(lngValue As Long, lngScaleSize As Long, sngBase As Single, lngMin As Long, lngMax As Long) As Long

    'If the lngValue is wacky, set to min
    If lngValue < 0 Then
        NormalizeLogScale = lngMin
        Exit Function
    End If

    'Normalize the log scale based on a range (ie. 0-100) a value within that range (ie. 52) and a log base (since base 10 gives nasty results)
    NormalizeLogScale = (lngMax - lngMin) * ((lngValue / lngScaleSize) ^ sngBase) + lngMin

End Function

Function FindClosestEnemy(lngObject As Long) As Long

Dim i As Long
Dim lngEnemy As Long
Dim dblDist As Double
Dim dblTemp As Double

    'Init the vars
    dblDist = 9E+99
    lngEnemy = TARGET_NONE
        
    'Check for nearby enemies
    For i = 0 To UBound(gudtObject)
        'Ensure this isn't the same object!
        If lngObject <> i Then
            'Is this a planet, or a nonexistant?
            If (gudtObject(i).udtInfo.bytRace <> RACE_PLANET) And (gudtObject(i).blnExists = True) Then
                'Is this an enemy?
                If gudtRace(gudtObject(lngObject).udtInfo.bytRace).intRelations(gudtObject(i).udtInfo.bytRace) <= RELATIONS_BAD Then
                    'Is this closer?
                    dblTemp = GetDist(gudtObject(i).udtPhysics.dblX, gudtObject(i).udtPhysics.dblY, gudtObject(lngObject).udtPhysics.dblX, gudtObject(lngObject).udtPhysics.dblY)
                    If dblDist > dblTemp Then
                        dblDist = dblTemp
                        lngEnemy = i
                    End If
                End If
            End If
        End If
    Next i
    
    'Is the player an enemy?
    If (gudtRace(gudtObject(lngObject).udtInfo.bytRace).intRelations(RACE_PLAYER) <= RELATIONS_BAD) And (gblnPlayerDead = False) Then
        'Is the player closer?
        If dblDist > gudtObject(lngObject).udtInfo.dblDistance Then lngEnemy = -1
    End If
    
    'Return the value
    FindClosestEnemy = lngEnemy

End Function

Sub TerminateSpecific(blnSound As Boolean, blnDraw As Boolean, blnInput As Boolean, blnDirectX As Boolean)

    'Kill the objects
    If blnSound = True Then DSound.Terminate
    If blnDraw = True Then DDraw.Terminate frmMain
    If blnInput = True Then DInput.Terminate
    If blnDirectX = True Then Set gobjDX = Nothing

End Sub

Sub EndProgram(frmTerm As Form)

    'Kill everything
    DSound.Terminate
    DDraw.Terminate frmTerm
    DInput.Terminate
    Set gobjDX = Nothing

End Sub

Public Sub KillFile(strFileName As String)

    'Destroy the given file, if it exists
    If Dir(strFileName) <> "" Then Kill strFileName

End Sub

Public Sub PointOnLine(dblX As Double, dblY As Double, sngDirection As Single, sngDistance As Single, ByRef dblResultX As Double, ByRef dblResultY As Double)

    'Find the coordinates of a point, given the direction and distance
    dblResultX = dblX + sngDistance * Sin(sngDirection)
    dblResultY = dblY - sngDistance * Cos(sngDirection)

End Sub
