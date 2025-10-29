VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CleanupToolForm 
   Caption         =   "REYDM - Basic Clean"
   ClientHeight    =   3630
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   4188
   OleObjectBlob   =   "CleanupToolForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "CleanupToolForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub btnRunCleanup_Click()
    If chkSectionBreaks.Value Then Call Remove_Section_Breaks_And_PageBreaks
    If chkCharClean.Value Then Call Char_Clean
    If chkParaClean.Value Then Call Para_Clean
    If chkNumbering.Value Then Call NumeringToText
    If chkExtraParas.Value Then Call Remove_Extra_Paragraphs

    MsgBox "Selected cleanup tasks completed!", vbInformation
    Unload Me
End Sub

Private Sub btnCancel_Click()
    Unload Me
End Sub

Private Sub UserForm_Initialize()
    chkSectionBreaks.Value = True
    chkCharClean.Value = True
    chkParaClean.Value = True
    chkNumbering.Value = False
    chkExtraParas.Value = True
End Sub

