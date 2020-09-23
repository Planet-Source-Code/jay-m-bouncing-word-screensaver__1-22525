VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Options"
   ClientHeight    =   1785
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   3750
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1054.637
   ScaleMode       =   0  'User
   ScaleWidth      =   3521.047
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox text 
      Height          =   345
      Left            =   1290
      TabIndex        =   1
      Top             =   135
      Width           =   2325
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   390
      Left            =   480
      TabIndex        =   4
      Top             =   1200
      Width           =   1140
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   390
      Left            =   2100
      TabIndex        =   5
      Top             =   1200
      Width           =   1140
   End
   Begin VB.TextBox Speed 
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   1290
      TabIndex        =   3
      Top             =   525
      Width           =   405
   End
   Begin VB.Label Label1 
      Caption         =   "Written by Jay Mykytiyk  (C) Jay Inc. 2001            All Rights Reserved"
      Height          =   615
      Left            =   1800
      TabIndex        =   6
      Top             =   480
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "&Text to use:"
      Height          =   270
      Index           =   0
      Left            =   105
      TabIndex        =   0
      Top             =   150
      Width           =   1080
   End
   Begin VB.Label lblLabels 
      Caption         =   "&Speed:"
      Height          =   270
      Index           =   1
      Left            =   105
      TabIndex        =   2
      Top             =   540
      Width           =   1080
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public LoginSucceeded As Boolean
Dim newspeed As String
Dim textf As String
Dim speedf As String


Private Sub cmdCancel_Click()
    'set the global var to false
    'to denote a failed login
    LoginSucceeded = False
    Unload Me
End Sub

Private Sub cmdOK_Click()
    'check for correct speed
    If IsNumeric(Speed.text) = True Then
   GoTo valid:
    Else
        MsgBox "Speed must be a number", vbCritical, "Speed"
        Speed.SetFocus
        SendKeys "{Home}+{End}"
        GoTo esub:
    End If
valid:
If (Speed = 0) Then
MsgBox "Speed must not be zero.", vbCritical, "Speed"
GoTo esub:
End If
newspeed = Speed
newspeed = newspeed * 3
Open "C:\windows\wordscreen.ini" For Append As #1
Write #1, "a"
Close #1
Kill "c:\windows\wordscreen.ini"
Open "C:\windows\wordscreen.ini" For Append As #1
Print #1, text
Print #1, newspeed;
Close #1
Unload Me
esub:
End Sub

Private Sub Form_Load()
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
exists:
Open "C:\windows\wordscreen.ini" For Input As #1
Input #1, textf
Input #1, speedf
Close #1
If (speedf = 0) Then
MsgBox "Speed must not be zero. Speed will now be set to one.", vbCritical, "Speed"
speedf = 3
End If
text.text = textf
Speed.text = speedf / 3
End Sub

