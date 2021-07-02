Attribute VB_Name = "Module1"
Public Function UTF16(ByVal sText As String) As String
Dim i As Integer, sTemp As String, sTemp2 As String, sTemp3 As String
Dim sHead As String
Dim sBody As String
For i = 1 To Len(sText)
 sTemp = Asc(Mid$(sText, i, 1&))
 If sTemp < 0 Then
  sTemp2 = Right$(Hex$(Asc(Mid$(sText, i, 1&))), 4&)
  sHead = Left$(sTemp2, 2&)
  sBody = Right$(sTemp2, 2&)
  sTemp3 = sTemp3 & "%" & sHead & "%" & sBody
  UTF16 = sTemp3
 Else
  sTemp3 = sTemp3 & "%" & Hex$(Asc(Mid$(sText, i, 1&)))
  UTF16 = sTemp3
 End If
Next i
sTemp2 = ""
sTemp3 = ""
End Function

