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

Global Const HideReadOnly = &H4         'Common Dialog Constants
Global Const OverWritePrompt = &H2
Global Const FileMustExist = &H1000
Global Const AllowMultiSelect = &H200

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

Sub ExtractData(strFileName As String)

Dim intWAVFile As Integer
Dim i As Long
Dim strTemp As String * 4
Dim blnFound As Boolean

    'Open the wave
    intWAVFile = FreeFile()
    Open strFileName For Binary Access Read Lock Write As intWAVFile
        'Get the header info
        Get intWAVFile, 1, gudtHeader
        'Find the "data" portion of the file
        For i = 1 To LOF(intWAVFile)
            Get intWAVFile, i, strTemp
            If strTemp = "data" Then
                blnFound = True
                Exit For
            End If
        Next i
        'Ensure this is a wave file
        If blnFound = False Then
            MsgBox strFileName & " is not a valid WAV file.", vbCritical, "Invalid File"
            Close intWAVFile
            Exit Sub
        End If
        'Get the data information
        Get intWAVFile, , glngChunkSize
        ReDim gbytData(glngChunkSize)
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

Public Sub KillFile(strFileName As String)

    'Destroy the given file, if it exists
    If Dir(strFileName) <> "" Then Kill strFileName

End Sub
