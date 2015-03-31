VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Form2 
   BackColor       =   &H80000012&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ASCII CONVERTER"
   ClientHeight    =   7320
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9660
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7320
   ScaleWidth      =   9660
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   7680
      Top             =   360
   End
   Begin RichTextLib.RichTextBox Rt2 
      Height          =   3135
      Left            =   120
      TabIndex        =   2
      Top             =   4080
      Width           =   9495
      _ExtentX        =   16748
      _ExtentY        =   5530
      _Version        =   393217
      BackColor       =   -2147483629
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      TextRTF         =   $"Form2.frx":030A
   End
   Begin RichTextLib.RichTextBox RT1 
      Height          =   3135
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   9495
      _ExtentX        =   16748
      _ExtentY        =   5530
      _Version        =   393217
      BackColor       =   8388608
      Enabled         =   -1  'True
      TextRTF         =   $"Form2.frx":038E
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Convert to ASCII Code"
      ForeColor       =   &H8000000B&
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   7095
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()



End Sub

Private Sub Form_Load()
Form2.Hide

RT1.Text = Form1.txt1.Text

a = Len(RT1.Text)
Do While a > 1
a = Len(RT1.Text)
Rt2.Text = Rt2.Text & Asc(Left(RT1.Text, 1)) & "."
If a > 0 Then
RT1.Text = Right(RT1.Text, a - 1)
End If
Loop

Form1.txt1.Text = Rt2.Text
Timer1.Enabled = True



End Sub

Private Sub Timer1_Timer()

Unload Form2
End Sub
