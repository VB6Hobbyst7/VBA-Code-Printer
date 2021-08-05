VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} uPrinter 
   Caption         =   "CodePrinter"
   ClientHeight    =   4644
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   2328
   OleObjectBlob   =   "uPrinter.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "uPrinter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_Initialize()
    FormatColourFormatters

End Sub

Private Sub cColourCode_Click()
    ColorPaletteDialog _
        ThisWorkbook.Sheets("TXTColour").Range("GeneralFontBackground"), _
        uPrinter.LBLcolourCode
End Sub

Private Sub cColourComments_Click()
    ColorPaletteDialog _
        ThisWorkbook.Sheets("TXTColour").Range("ColourComments"), _
        uPrinter.LBLcolourComment
End Sub

Private Sub cColourKeywords_Click()
    ColorPaletteDialog _
        ThisWorkbook.Sheets("TXTColour").Range("ColourKeywords"), _
        uPrinter.LBLcolourKey
End Sub

Private Sub cColourOddLines_Click()
    ColorPaletteDialog _
        ThisWorkbook.Sheets("TXTColour").Range("OddLine"), _
        uPrinter.LBLcolourOdd
End Sub

Private Sub cInfo_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    uDEV.Show
End Sub

Private Sub Image1_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Me.Caption = "Magic storm"
    PrintProject
    Me.Caption = "CodePrinter"
End Sub

Sub ColorPaletteDialog(rng As Range, lbl As MSForms.Label)
    If Application.Dialogs(xlDialogEditColor).Show(10, 0, 125, 125) = True Then
        'user pressed OK
        Lcolor = ActiveWorkbook.Colors(10)
        rng.Value = Lcolor
        rng.Offset(0, 1).Interior.Color = Lcolor
        lbl.ForeColor = Lcolor
    End If
    ActiveWorkbook.ResetColors
End Sub


