VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmHideTags 
   Caption         =   "Hide and Highlight Tags"
   ClientHeight    =   7020
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   6768
   OleObjectBlob   =   "frmHideTags.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmHideTags"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub btnOK_Click()
    ' At least one action must be selected
    If Not (optHide.Value Or optUnhide.Value Or optHighlight.Value Or optRemoveHighlight.Value) Then
        MsgBox "Please select Hide/Unhide or Highlight/Remove Highlight.", vbExclamation
        Exit Sub
    End If
    Me.tag = "OK"
    Me.Hide
End Sub

Private Sub btnCancel_Click()
    Me.tag = "Cancel"
    Me.Hide
End Sub

Private Sub UserForm_Initialize()
    ' Set default checked tags
    chkSIGNATURE.Value = True
    chkLOGO.Value = True
    ' Add others as needed
     chkEMBLEM.Value = True
     chkDIGITALSIGNATURE.Value = True
     chkREDACTION.Value = True
     chkSTAMP.Value = True

    ' Optional: Set default for action buttons
    'optHide.Value = True  ' default to Hide
    'optHighlight.Value = False
End Sub
