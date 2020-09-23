VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   Caption         =   "Bank File Formatter"
   ClientHeight    =   7560
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   9240
   LinkTopic       =   "Form1"
   ScaleHeight     =   504
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   616
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog cmdDialog 
      Left            =   10920
      Top             =   4200
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox picBMP 
      AutoRedraw      =   -1  'True
      Height          =   7560
      Left            =   1680
      ScaleHeight     =   500
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   500
      TabIndex        =   1
      Top             =   0
      Width           =   7560
   End
   Begin VB.ListBox lstFiles 
      Height          =   7560
      IntegralHeight  =   0   'False
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1695
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileNew 
         Caption         =   "&New"
      End
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Open"
      End
      Begin VB.Menu sep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileAdd 
         Caption         =   "&Add"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuFileRecursive 
         Caption         =   "Add Re&cursive"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuFileRemove 
         Caption         =   "&Remove"
         Enabled         =   0   'False
         Shortcut        =   ^R
      End
      Begin VB.Menu mnuFileRename 
         Caption         =   "Rena&me"
         Shortcut        =   ^N
      End
      Begin VB.Menu sep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "&About"
   End
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

Private Sub Form_Load()

    'Set the common dialog directory
    cmdDialog.InitDir = App.Path

End Sub

Private Sub Form_Resize()

    'Resize lstbox and picbox
    lstFiles.Height = frmMain.ScaleHeight
    picBMP.Height = frmMain.ScaleHeight
    picBMP.Width = frmMain.ScaleWidth - lstFiles.Width

End Sub

Private Sub mnuFileNew_Click()

Dim lngTemp As Long

    'Make a new bank file
    cmdDialog.FileName = ""
    cmdDialog.Flags = HideReadOnly Or OverWritePrompt
    cmdDialog.Filter = "Bank Files (*.bnk)|*.bnk"
    On Error Resume Next
    cmdDialog.ShowSave
    'If the user canceled out of the dialog, exit sub
    If Err.Number = cdlCancel Or cmdDialog.FileName = "" Then
        Exit Sub
    End If
    On Error GoTo 0
    
    'Create the file
    lngTemp = 5
    gstrFileName = cmdDialog.FileName
    KillFile gstrFileName
    Open gstrFileName For Binary Access Read Write Lock Write As #1
    Put 1, 1, lngTemp
    Close #1
    
    'Clear the footer
    ReDim gudtFooter.lngFileLocation(0)
    ReDim gudtFooter.strFileName(0)
    
    'Clear the display
    lstFiles.Clear

End Sub

Private Sub mnuFileAdd_Click()

Dim lngTemp As Long
Dim strPath As String
Dim strFileList As String
Dim strFiles() As String
Dim i As Integer

    'If there is no "current" file, then make a new one first
    If gstrFileName = "" Then mnuFileNew_Click

    'Add a bitmap to the bank file
    cmdDialog.FileName = ""
    cmdDialog.Flags = HideReadOnly Or FileMustExist Or AllowMultiSelect
    cmdDialog.Filter = "Bitmap Files (*.bmp)|*.bmp"
    On Error Resume Next
    cmdDialog.ShowOpen
    'If the user canceled out of the dialog, exit sub
    If Err.Number = cdlCancel Or cmdDialog.FileName = "" Then
        Exit Sub
    End If
    On Error GoTo 0
    
    'Handle multiple selection
    ReDim strFiles(0)
    strPath = ""
    If InStr(1, cmdDialog.FileName, " ") > 0 Then
        strPath = Left(cmdDialog.FileName, InStr(1, cmdDialog.FileName, " ") - 1) & "\"
        cmdDialog.FileName = Right(cmdDialog.FileName, Len(cmdDialog.FileName) - InStr(1, cmdDialog.FileName, " "))
    End If
    cmdDialog.InitDir = strPath
    'Extract the filenames from the selection
    strFileList = cmdDialog.FileName & " "
    strFiles(0) = Left(strFileList, InStr(1, strFileList, " ") - 1)
    strFileList = Right(strFileList, Len(strFileList) - InStr(1, strFileList, " "))
    Do While InStr(1, strFileList, " ")
        ReDim Preserve strFiles(UBound(strFiles) + 1)
        strFiles(UBound(strFiles)) = Left(strFileList, InStr(1, strFileList, " ") - 1)
        strFileList = Right(strFileList, Len(strFileList) - InStr(1, strFileList, " "))
    Loop
    
    'Place the data in the resource
    Open gstrFileName For Binary Access Read Write Lock Write As #1
    'Loop through each selected file
    For i = 0 To UBound(strFiles)
        'Extract the bitmap data
        ExtractData strPath & strFiles(i)
        Get 1, 1, lngTemp
        'If this file is empty, init the footer
        If lngTemp = 5 Then
            ReDim gudtFooter.lngFileLocation(0)
            ReDim gudtFooter.strFileName(0)
        'Otherwise just add to the end
        Else
            ReDim Preserve gudtFooter.lngFileLocation(UBound(gudtFooter.lngFileLocation) + 1)
            ReDim Preserve gudtFooter.strFileName(UBound(gudtFooter.strFileName) + 1)
        End If
        'Place the data
        gudtFooter.lngFileLocation(UBound(gudtFooter.lngFileLocation)) = lngTemp
        gudtFooter.strFileName(UBound(gudtFooter.strFileName)) = ExtractFilename(strFiles(i))
        Put 1, lngTemp, gudtBMPFileHeader
        Put 1, , gudtBMPInfo
        Put 1, , gudtBMPData
        lngTemp = Seek(1)
        Put 1, , gudtFooter
        Put 1, 1, lngTemp
    Next i
    Close #1
    
    'Update the display
    UpdateDisplay
    
End Sub

Private Sub mnuFileOpen_Click()

Dim lngTemp As Long

    'Open a bank file
    cmdDialog.FileName = ""
    cmdDialog.Flags = HideReadOnly Or FileMustExist
    cmdDialog.Filter = "Bank Files (*.bnk)|*.bnk"
    On Error Resume Next
    cmdDialog.ShowSave
    'If the user canceled out of the dialog, exit sub
    If Err.Number = cdlCancel Or cmdDialog.FileName = "" Then
        Exit Sub
    End If
    On Error GoTo 0
    
    'Extract the data
    gstrFileName = cmdDialog.FileName
    Open gstrFileName For Binary Access Read Write Lock Write As #1
    Get 1, 1, lngTemp
    'If this file is empty, clear the footer
    If lngTemp = 5 Then
        ReDim gudtFooter.lngFileLocation(0)
        ReDim gudtFooter.strFileName(0)
    'Otherwise extract the footer
    Else
        Get 1, lngTemp, gudtFooter
    End If
    Close #1
    
    'Update the display
    UpdateDisplay

End Sub

Private Sub mnuFileRecursive_Click()

Dim strPath As String
Dim strDir() As String
Dim intDirNum As Integer
Dim strFile() As String
Dim blnBMPS As Boolean
Dim blnBMPFirst As Boolean
Dim blnSubDirs As Boolean
Dim blnSubDirFirst As Boolean
Dim blnComplete As Boolean
Dim strTemp As String
Dim i As Integer
Dim lngTemp As Long

    'If there is no "current" file, then make a new one first
    If gstrFileName = "" Then mnuFileNew_Click
    
    'Find a directory
    strPath = InputBox("Enter directory:", "Directory", App.Path)
    If strPath = "" Then Exit Sub
    If Dir(strPath, vbDirectory) = "" Then
        MsgBox "Directory does not exist.", vbOKOnly, "No Directory"
        Exit Sub
    End If
    
    'Recurse
    ReDim strDir(0)
    ReDim strFile(0)
    strDir(0) = strPath
    intDirNum = 0
    blnBMPS = True
    blnBMPFirst = True
    blnSubDirs = True
    blnSubDirFirst = True
    blnComplete = False
    Do While Not (blnComplete)
        'Search for bitmaps
        If blnBMPS Then
            If blnBMPFirst Then
                strTemp = Dir(strDir(intDirNum) & "\*.BMP")
                blnBMPFirst = False
            Else
                strTemp = Dir()
            End If
            If strTemp = "" Then
                'No more bitmaps in this subdir
                blnBMPS = False
            Else
                'Otherwise, add it to the list
                If strFile(0) = "" Then
                    strFile(0) = strDir(intDirNum) & "\" & strTemp
                Else
                    ReDim Preserve strFile(UBound(strFile) + 1)
                    strFile(UBound(strFile)) = strDir(intDirNum) & "\" & strTemp
                End If
            End If
        ElseIf blnSubDirs Then
            If blnSubDirFirst Then
                strTemp = Dir(strDir(intDirNum) & "\*.*", vbDirectory)
                blnSubDirFirst = False
            Else
                strTemp = Dir()
            End If
            If strTemp = "" Then
                'No more subdirs
                blnSubDirs = False
            ElseIf strTemp <> "." And strTemp <> ".." And GetAttr(strDir(intDirNum) & "\" & strTemp) = vbDirectory Then
                'Otherwise, add it to the list
                ReDim Preserve strDir(UBound(strDir) + 1)
                strDir(UBound(strDir)) = strDir(intDirNum) & "\" & strTemp
            End If
        Else
            'Check if this was the last subdir
            If intDirNum = UBound(strDir) Then
                blnComplete = True
            'Otherwise, just start with the next one!
            Else
                blnBMPS = True
                blnBMPFirst = True
                blnSubDirs = True
                blnSubDirFirst = True
                intDirNum = intDirNum + 1
            End If
        End If
    Loop
    
    'Place the data in the resource
    Open gstrFileName For Binary Access Read Write Lock Write As #1
    'Loop through each selected file
    For i = 0 To UBound(strFile)
        'Extract the bitmap data
        ExtractData strFile(i)
        Get 1, 1, lngTemp
        'If this file is empty, init the footer
        If lngTemp = 5 Then
            ReDim gudtFooter.lngFileLocation(0)
            ReDim gudtFooter.strFileName(0)
        'Otherwise just add to the end
        Else
            ReDim Preserve gudtFooter.lngFileLocation(UBound(gudtFooter.lngFileLocation) + 1)
            ReDim Preserve gudtFooter.strFileName(UBound(gudtFooter.strFileName) + 1)
        End If
        'Place the data
        gudtFooter.lngFileLocation(UBound(gudtFooter.lngFileLocation)) = lngTemp
        gudtFooter.strFileName(UBound(gudtFooter.strFileName)) = ExtractFilename(strFile(i))
        Put 1, lngTemp, gudtBMPFileHeader
        Put 1, , gudtBMPInfo
        Put 1, , gudtBMPData
        lngTemp = Seek(1)
        Put 1, , gudtFooter
        Put 1, 1, lngTemp
    Next i
    Close #1
    
    'Update the display
    UpdateDisplay

End Sub

Private Sub UpdateDisplay()

Dim i As Integer

    'Update the display
    lstFiles.Clear
    For i = 0 To UBound(gudtFooter.strFileName)
        lstFiles.AddItem gudtFooter.strFileName(i)
    Next i

End Sub

Private Sub lstFiles_Click()

    'Show the bitmap selected
    DisplayImage

End Sub

Private Sub DisplayImage()

    'If there isn't an image selected, exit sub
    If lstFiles.ListIndex = -1 Then Exit Sub
    
    'Clear the picturebox
    picBMP.Cls
    
    'Otherwise, display the bitmap in the picturebox
    ExtractData gstrFileName, gudtFooter.lngFileLocation(lstFiles.ListIndex)
    StretchDIBits picBMP.hdc, 0, 0, gudtBMPInfo.bmiHeader.biWidth, gudtBMPInfo.bmiHeader.biHeight, 0, 0, gudtBMPInfo.bmiHeader.biWidth, gudtBMPInfo.bmiHeader.biHeight, gudtBMPData(0), gudtBMPInfo, DIB_RGB_COLORS, vbSrcCopy
    
    'Show the bitmap
    picBMP.Refresh

End Sub

Private Sub mnuFileExit_Click()

    'Exit the program
    End

End Sub

Private Sub mnuAbout_Click()

    'Display the about box
    frmAbout.Show vbModal, Me

End Sub

Private Sub mnuFileRename_Click()

Dim lngTemp As Long
Dim strTemp As String

    'If there isn't an image selected, exit sub
    If lstFiles.ListIndex = -1 Then Exit Sub
    
    'Get the new name
    strTemp = InputBox("Enter new filename:", "Rename")
    If strTemp = "" Then Exit Sub
    
    'Modify footer
    gudtFooter.strFileName(lstFiles.ListIndex) = strTemp
    Open gstrFileName For Binary Access Read Write Lock Write As #1
    Get #1, 1, lngTemp
    Put #1, lngTemp, gudtFooter
    Close #1
    
    'Update the display
    UpdateDisplay

End Sub
