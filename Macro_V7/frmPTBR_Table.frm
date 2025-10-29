VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmPTBR_Table 
   Caption         =   "PTBR Table Format"
   ClientHeight    =   2715
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   5172
   OleObjectBlob   =   "frmPTBR_Table.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmPTBR_Table"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub btnOK_Click()
    ' Run selected macros
    If chkMerge.Value Then Call MergeBasedOnAceito
    If chkClean.Value Then Call CleanBreaksInSelectedTableOnly
    If chkFormat.Value Then Call FormatSelectedTable

    Unload Me
End Sub


Private Sub btnCancel_Click()
    Unload Me
End Sub
Private Sub UserForm_Initialize()
    chkMerge.Value = True
    chkClean.Value = True
    chkFormat.Value = True
End Sub

