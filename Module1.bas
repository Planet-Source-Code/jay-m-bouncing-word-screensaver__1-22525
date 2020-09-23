Attribute VB_Name = "Module1"
Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
Public iniexits As Boolean
Sub Main()

If App.PrevInstance Then End

If LCase(Left(Command, 2)) = "/p" Then
If FileExists("C:\windows\wordscreen.ini") Then
GoTo exists:
Else
Open "C:\windows\wordscreen.ini" For Append As #1
Write #1, "a"
Close #1
Kill "c:\windows\wordscreen.ini"
Open "C:\windows\wordscreen.ini" For Append As #1
Print #1, "Hello"
Print #1, 3;
Close #1
End If
End If

If LCase(Left(Command, 2)) = "/c" Then
  frmLogin.Show
  Exit Sub
End If
    Hid$ = ShowCursor(False)
Form1.Show
exists:
End Sub
