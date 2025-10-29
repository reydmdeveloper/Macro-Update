VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmGreekPicker 
   Caption         =   "Insert Greek Character"
   ClientHeight    =   1815
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   3672
   OleObjectBlob   =   "frmGreekPicker.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmGreekPicker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UserForm_Initialize()
    With cboGreek
        .Clear
        .ColumnCount = 2
        .ColumnWidths = "60 pt;0 pt" ' 2nd col hidden (actual char)
        .MatchEntry = fmMatchEntryComplete
    End With

    ' ---- Lowercase ----
    AddGreek "a  (alpha, lowercase)", &H3B1
    AddGreek "ß  (beta, lowercase)", &H3B2
    AddGreek "?  (gamma, lowercase)", &H3B3
    AddGreek "d  (delta, lowercase)", &H3B4
    AddGreek "e  (epsilon, lowercase)", &H3B5
    AddGreek "?  (zeta, lowercase)", &H3B6
    AddGreek "?  (eta, lowercase)", &H3B7
    AddGreek "?  (theta, lowercase)", &H3B8
    AddGreek "?  (iota, lowercase)", &H3B9
    AddGreek "?  (kappa, lowercase)", &H3BA
    AddGreek "?  (lambda, lowercase)", &H3BB
    AddGreek "µ  (mu, lowercase)", &H3BC
    AddGreek "?  (nu, lowercase)", &H3BD
    AddGreek "?  (xi, lowercase)", &H3BE
    AddGreek "?  (omicron, lowercase)", &H3BF
    AddGreek "p  (pi, lowercase)", &H3C0
    AddGreek "?  (rho, lowercase)", &H3C1
    AddGreek "s  (sigma, lowercase)", &H3C3
    AddGreek "?  (final sigma, lowercase)", &H3C2
    AddGreek "t  (tau, lowercase)", &H3C4
    AddGreek "?  (upsilon, lowercase)", &H3C5
    AddGreek "f  (phi, lowercase)", &H3C6
    AddGreek "?  (chi, lowercase)", &H3C7
    AddGreek "?  (psi, lowercase)", &H3C8
    AddGreek "?  (omega, lowercase)", &H3C9

    ' ---- Uppercase ----
    AddGreek "?  (Alpha, uppercase)", &H391
    AddGreek "?  (Beta, uppercase)", &H392
    AddGreek "G  (Gamma, uppercase)", &H393
    AddGreek "?  (Delta, uppercase)", &H394
    AddGreek "?  (Epsilon, uppercase)", &H395
    AddGreek "?  (Zeta, uppercase)", &H396
    AddGreek "?  (Eta, uppercase)", &H397
    AddGreek "T  (Theta, uppercase)", &H398
    AddGreek "?  (Iota, uppercase)", &H399
    AddGreek "?  (Kappa, uppercase)", &H39A
    AddGreek "?  (Lambda, uppercase)", &H39B
    AddGreek "?  (Mu, uppercase)", &H39C
    AddGreek "?  (Nu, uppercase)", &H39D
    AddGreek "?  (Xi, uppercase)", &H39E
    AddGreek "?  (Omicron, uppercase)", &H39F
    AddGreek "?  (Pi, uppercase)", &H3A0
    AddGreek "?  (Rho, uppercase)", &H3A1
    AddGreek "S  (Sigma, uppercase)", &H3A3
    AddGreek "?  (Tau, uppercase)", &H3A4
    AddGreek "?  (Upsilon, uppercase)", &H3A5
    AddGreek "F  (Phi, uppercase)", &H3A6
    AddGreek "?  (Chi, uppercase)", &H3A7
    AddGreek "?  (Psi, uppercase)", &H3A8
    AddGreek "O  (Omega, uppercase)", &H3A9

    If cboGreek.ListCount > 0 Then cboGreek.ListIndex = 0
End Sub

Private Sub AddGreek(ByVal labelText As String, ByVal codePoint As Long)
    With cboGreek
        .AddItem labelText
        .list(.ListCount - 1, 1) = ChrW$(codePoint)
    End With
End Sub

Private Sub cmdInsert_Click()
    If cboGreek.ListIndex < 0 Then
        MsgBox "Please pick a letter.", vbInformation
        Exit Sub
    End If
    Dim ch As String
    ch = cboGreek.list(cboGreek.ListIndex, 1) ' hidden column
    Selection.TypeText ch
    Unload Me
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

