Attribute VB_Name = "modMain"
Public Sub Number_Only(KeyAscii As Integer, Text As String)
    Dim strvalid As String * 11
    If InStr(Text, ".") = 0 Then
        strvalid = "-0123456789."
    Else
        strvalid = "-0123456789"
    End If
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    ElseIf InStr(strvalid, Chr(KeyAscii)) = 0 And KeyAscii > 26 Then
        KeyAscii = 0
    End If
End Sub

