Attribute VB_Name = "Globals"
'**************************************************************
'
' THIS WORK, INCLUDING THE SOURCE CODE, DOCUMENTATION
' AND RELATED MEDIA AND DATA, IS PLACED INTO THE PUBLIC DOMAIN.
'
' THE ORIGINAL AUTHOR IS RYAN CLARK.
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

'API
Declare Function StretchDIBits Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal dx As Long, ByVal dy As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal wSrcWidth As Long, ByVal wSrcHeight As Long, lpBits As Any, lpBitsInfo As BITMAPINFO, ByVal wUsage As Long, ByVal dwRop As Long) As Long
Global Const SRCCOPY = &HCC0020
Global Const DIB_RGB_COLORS = 0

Global Const HideReadOnly = &H4         'Common Dialog Constants
Global Const OverWritePrompt = &H2
Global Const FileMustExist = &H1000
Global Const AllowMultiSelect = &H200

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

'Our current BNK file
Global gstrFileName As String

'Our footer struct
Type FOOTERTYPE
    strFileName() As String
    lngFileLocation() As Long
End Type
Global gudtFooter As FOOTERTYPE

Function BitRoll(bytData As Byte, bytRollAmount As Integer) As Byte

    'Roll the byte data over
    If bytData + bytRollAmount > 255 Then
        BitRoll = bytData + bytRollAmount - 256
    ElseIf bytData + bytRollAmount < 0 Then
        BitRoll = bytData + bytRollAmount + 256
    Else
        BitRoll = bytData + bytRollAmount
    End If

End Function

Sub ExtractData(strFileName As String, Optional lngOffset As Long = 1)

Dim intBMPFile As Integer
Dim i As Integer

    'Init variables
    Erase gudtBMPInfo.bmiColors

    'Open the bitmap
    intBMPFile = FreeFile()
    Open strFileName For Binary Access Read Lock Write As intBMPFile
        'Fill the File Header structure
        Get intBMPFile, lngOffset, gudtBMPFileHeader
        'Fill the Info structure
        Get intBMPFile, , gudtBMPInfo.bmiHeader
        If gudtBMPInfo.bmiHeader.biClrUsed <> 0 Then
            For i = 0 To gudtBMPInfo.bmiHeader.biClrUsed - 1
                Get intBMPFile, , gudtBMPInfo.bmiColors(i).rgbBlue
                Get intBMPFile, , gudtBMPInfo.bmiColors(i).rgbGreen
                Get intBMPFile, , gudtBMPInfo.bmiColors(i).rgbRed
                Get intBMPFile, , gudtBMPInfo.bmiColors(i).rgbReserved
            Next i
        Else
            Get intBMPFile, , gudtBMPInfo.bmiColors
        End If
        'Size the BMPData array
        ReDim gudtBMPData(FileSize(gudtBMPInfo.bmiHeader.biWidth, gudtBMPInfo.bmiHeader.biHeight))
        'Fill the BMPData array
        Get intBMPFile, , gudtBMPData
        'Ensure info is correct
        gudtBMPFileHeader.bfOffBits = 1078
        gudtBMPInfo.bmiHeader.biSizeImage = FileSize(gudtBMPInfo.bmiHeader.biWidth, gudtBMPInfo.bmiHeader.biHeight)
        gudtBMPInfo.bmiHeader.biClrUsed = 0
        gudtBMPInfo.bmiHeader.biClrImportant = 0
        gudtBMPInfo.bmiHeader.biXPelsPerMeter = 0
        gudtBMPInfo.bmiHeader.biYPelsPerMeter = 0
    Close intBMPFile
    
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

Public Sub KillFile(strFileName As String)

    'Destroy the given file, if it exists
    If Dir(strFileName) <> "" Then Kill strFileName

End Sub
