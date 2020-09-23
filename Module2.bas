Attribute VB_Name = "Module2"
Function FileExists(strPath As String) As Integer
Dim lngRetVal As Long
On Error Resume Next
lngRetVal = Len(Dir$(strPath))
If Err Or lngRetVal = 0 Then
FileExists = False
Else
FileExists = True
End If
End Function
