Attribute VB_Name = "Globals"
Option Explicit

'For the boolean dropdowns
Global Const FALSE_DROP = 0
Global Const TRUE_DROP = 1

'Common Dialog Constants
Global Const HideReadOnly = &H4
Global Const OverWritePrompt = &H2
Global Const FileMustExist = &H1000
Global Const AllowMultiSelect = &H200

Sub ClearMemory()
    
    'Clear out all of the arrays
    Erase gudtArmour
    glngArmourNum = 0
    Erase gudtCannon
    glngCannonNum = 0
    Erase gudtEngine
    glngEngineNum = 0
    Erase gudtGenerator
    glngGeneratorNum = 0
    Erase gudtHull
    glngHullNum = 0
    'Erase the gudtJammer data somehow
    Erase gudtLaser
    glngLaserNum = 0
    Erase gudtMissile
    glngMissileNum = 0
    Erase gudtScanner
    glngScannerNum = 0
    Erase gudtShield
    glngShieldNum = 0
    Erase gudtRace
    Erase gudtObject
    'Erase the gudtPlayer data somehow

End Sub

Sub InitBoolCombo(ByRef cmbBool As ComboBox)

    'Initialize the combo box with boolean values
    cmbBool.Clear
    cmbBool.AddItem "False", FALSE_DROP
    cmbBool.AddItem "True", TRUE_DROP

End Sub

Function SetBoolCombo(blnVal As Boolean) As Integer

    'Is this a true, or false?
    If blnVal Then
        SetBoolCombo = TRUE_DROP
    Else
        SetBoolCombo = FALSE_DROP
    End If

End Function

Function GetBoolCombo(ByRef cmbBool As ComboBox)

    'Is the selected index true or false?
    If cmbBool.ListIndex = TRUE_DROP Then
        GetBoolCombo = True
    Else
        GetBoolCombo = False
    End If

End Function

Public Sub KillFile(strFileName As String)

    'Destroy the given file, if it exists
    If Dir(strFileName) <> "" Then Kill strFileName

End Sub
