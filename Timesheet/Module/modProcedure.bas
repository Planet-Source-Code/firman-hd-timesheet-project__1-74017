Attribute VB_Name = "modProcedure"
'Procedure used to highlight text when focus
Public StatusForm As Boolean
Public Sub HLText(ByRef sText)
    On Error Resume Next
    With sText
        .SelStart = 0
        .SelLength = Len(sText.Text)
    End With
End Sub
'

'Procedure used to clear the text content
Public Sub clearText(ByRef sForm As Form)
    Dim CONTROL As CONTROL
    For Each CONTROL In sForm.Controls
        If (TypeOf CONTROL Is TextBox) Then CONTROL = vbNullString
    Next CONTROL
    Set CONTROL = Nothing
End Sub
Public Sub LoadForm(ByRef srcForm As Form)
    srcForm.Show
    If StatusForm = False Then srcForm.WindowState = vbMaximized
    srcForm.SetFocus
    StatusForm = False
End Sub
