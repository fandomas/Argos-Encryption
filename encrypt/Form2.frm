VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H00000080&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Login"
   ClientHeight    =   3975
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8925
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3975
   ScaleWidth      =   8925
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "LOGIN FIRST"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   5400
      TabIndex        =   0
      Top             =   120
      Width           =   3255
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   120
      Picture         =   "Form2.frx":030A
      Top             =   120
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   255
      Left            =   0
      Top             =   3720
      Width           =   255
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Dim a As Currency
Dim b As Currency
Dim c As Currency
Dim d As Currency
Const p = 6

a = Val(InputBox("PASSWORD 1", "PS1"))
b = Val(InputBox("PASSWORD 2", "PS2"))
c = a * b
a = Val(InputBox("PS1"))
b = Val(InputBox("PS2"))
a = a & b
If a = c Then
c = Left(a, 2)
d = Right(c, 2)
d = c + d
If d <> 0 Then
d = d / p
End If
If d = p Then
Load Form1
Form1.Show
Form2.Hide
Unload Form2
Else
End
End If
Else
End
End If
End Sub
