Attribute VB_Name = "mDEV"
Sub MailDev()
    'For Tips see: http://www.rondebruin.nl/win/winmail/Outlook/tips.htm
    'Working in Office 2000-2016
    Dim OutApp As Object
    Dim OutMail As Object
    Dim strbody As String
    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(0)
    '    strbody = "Hi there" & vbNewLine & vbNewLine & _
    "This is line 1" & vbNewLine & _
    "This is line 2" & vbNewLine & _
    "This is line 3" & vbNewLine & _
    "This is line 4"
    On Error Resume Next
    With OutMail
        .To = "anastasioualex@gmail.com"
        .CC = vbNullString
        .BCC = vbNullString
        .Subject = "DEV REQUEST OR FEEDBACK FOR -CODE ARCHIVE-"
        .body = strbody
        'You can add a file like this
        '.Attachments.Add ("C:\test.txt")
        '.Send
        .Display
    End With
    On Error GoTo 0
    Set OutMail = Nothing
    Set OutApp = Nothing
End Sub

Function OutlookCheck() As Boolean
    'is outlook installed?
    Dim xOLApp As Object
    '    On Error GoTo L1
    Set xOLApp = CreateObject("Outlook.Application")
    If Not xOLApp Is Nothing Then
        OutlookCheck = True
        'MsgBox "Outlook " & xOLApp.Version & " installed", vbExclamation
        Set xOLApp = Nothing
        Exit Function
    End If
    OutlookCheck = False
    'L1: MsgBox "Outlook not installed", vbExclamation, "Kutools for Outlook"
End Function

Function Clipboard(Optional StoreText As String) As String
    'PURPOSE: Read/Write to Clipboard
    'Source: ExcelHero.com (Daniel Ferry)
    Dim X As Variant
    'Store as variant for 64-bit VBA support
    X = StoreText
    'Create HTMLFile Object
    With CreateObject("htmlfile")
        With .parentWindow.clipboardData
            Select Case True
            Case Len(StoreText)
                'Write to the clipboard
                .SetData "text", X
            Case Else
                'Read from the clipboard (no variable passed through)
                Clipboard = .GetData("text")
            End Select
        End With
    End With
End Function


