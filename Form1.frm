VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   11505
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15000
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1.14
   ScaleMode       =   0  'User
   ScaleWidth      =   1.25
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   2880
      Top             =   2280
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000008&
      Caption         =   "Hello"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   555
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1080
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim MouseMoveCount As Integer
Dim VelX As String
Dim VelY As String
Dim MaxX As Integer
Dim MaxY As Integer
Dim X As Integer
Dim Y As Integer
Dim Speed As String
Dim text As String
Private Sub Form_KeyPress(KeyAscii As Integer)
Hid$ = ShowCursor(True)
End
End Sub

Private Sub Form_Load()
Open "C:\windows\wordscreen.ini" For Input As #1
Input #1, text
Input #1, Speed
Close #1
If (Speed = 0) Then
Speed = 3
End If
Me.ScaleMode = 3
VelX = Speed
VelY = Speed
Label1.Caption = text
MaxX = Me.ScaleWidth - Label1.Width
MaxY = Me.ScaleHeight - Label1.Height
X = 0
Y = 0

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Hid$ = ShowCursor(True)
End
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  MouseMoveCount = MouseMoveCount + 1
  If MouseMoveCount <= 4 Then Exit Sub
  Hid$ = ShowCursor(True)
  End
End Sub

Private Sub Timer1_Timer()
X = X + VelX
Y = Y + VelY
If X <= 0 Or X >= MaxX Then
VelX = -VelX
End If
If Y <= 0 Or Y >= MaxY Then
VelY = -VelY
End If
Label1.Move X, Y
End Sub
