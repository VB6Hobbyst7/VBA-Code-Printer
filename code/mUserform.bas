Attribute VB_Name = "mUserform"
Sub ShowPrinter()
    '    If Not IsLoaded("uPrinter") Then
    uPrinter.Show
    '    End If
End Sub

Function IsLoaded(formName As String) As Boolean
    Dim frm As Object
    For Each frm In VBA.UserForms
        If frm.Name = formName Then
            IsLoaded = True
            Exit Function
        End If
    Next frm
    IsLoaded = False
End Function

Sub AddCommandbar()
    DeleteCommandBar
    Dim cControl As Office.CommandBarButton
    Set cControl = Application.CommandBars("Worksheet Menu Bar").Controls.Add(Temporary:=True)

    With cControl
        .Caption = "CodePrinter"
        .Tag = "CodePrinter"
        .Style = msoButtonIconAndCaption
        .OnAction = "ShowPrinter"                'Macro stored in a Standard Module
        .FaceId = 246
    End With
    On Error GoTo 0
End Sub

Sub DeleteCommandBar()
    Dim bar As CommandBarControl
    Set bar = Application.CommandBars("Worksheet Menu Bar").FindControl(, , "CodePrinter")
    Do While Not bar Is Nothing
        bar.Delete
        Set bar = Application.CommandBars("Worksheet Menu Bar").FindControl(, , "CodePrinter")
    Loop
End Sub

