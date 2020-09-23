Attribute VB_Name = "Logging"
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

Const LOG_NAME = "log.txt"
Global gblnLogging As Boolean

Public Sub Log(strModule As String, strProcedure As String, strMessage As String)

Dim intLogFile As Integer
Dim strTemp As String
    
    'Is logging enabled?
    If gblnLogging = False Then Exit Sub
    
    'Open/create the file
    intLogFile = FreeFile()
    Open LOG_NAME For Binary Access Read Write Lock Write As intLogFile
    
    'Assemble the data/time message
    strTemp = Now() & " - (" & strModule & ": " & strProcedure & ") " & strMessage & vbCrLf
    
    'Write the message
    If LOF(intLogFile) = 0 Then
        Put intLogFile, , strTemp
    Else
        Put intLogFile, LOF(intLogFile) + 1, strTemp
    End If
    
    'Write to disk
    Close intLogFile
    
End Sub
