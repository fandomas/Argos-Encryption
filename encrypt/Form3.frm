VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Form3 
   BackColor       =   &H8000000C&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ascii"
   ClientHeight    =   4380
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10290
   Icon            =   "Form3.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4380
   ScaleWidth      =   10290
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   6600
      Top             =   1800
   End
   Begin RichTextLib.RichTextBox tr2 
      Height          =   1935
      Left            =   120
      TabIndex        =   1
      Top             =   2160
      Width           =   10095
      _ExtentX        =   17806
      _ExtentY        =   3413
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"Form3.frx":030A
   End
   Begin RichTextLib.RichTextBox tr1 
      Height          =   1935
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10095
      _ExtentX        =   17806
      _ExtentY        =   3413
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"Form3.frx":038E
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
tr1.Text = Form1.TexT2.Text

a = Len(tr1.Text)
Do While a > 1
a = Len(tr1.Text)
b = Left(tr1.Text, 1)

If b = "." Then

tr1.Text = Right(tr1.Text, a - 1)
tr2.Text = tr2.Text & Chr(ch)
ch = ""

Else

ch = ch & Left(tr1.Text, 1)
tr1.Text = Right(tr1.Text, a - 1)
End If

Loop

Form1.TexT2.Text = tr2.Text

Timer1.Enabled = True


End Sub

Private Sub Timer1_Timer()
Unload Form3
End Sub
