VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Argos - Decoder"
   ClientHeight    =   6945
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   14010
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6945
   ScaleWidth      =   14010
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command11 
      Caption         =   "ASCII"
      Height          =   1455
      Left            =   8400
      TabIndex        =   211
      Top             =   480
      Width           =   255
   End
   Begin VB.TextBox Textc1 
      Height          =   285
      Left            =   7680
      MaxLength       =   1
      TabIndex        =   210
      Text            =   "1"
      Top             =   2280
      Width           =   375
   End
   Begin VB.TextBox Textc2 
      Height          =   285
      Left            =   7680
      MaxLength       =   1
      TabIndex        =   209
      Text            =   "1"
      Top             =   2880
      Width           =   375
   End
   Begin VB.TextBox Textc3 
      Height          =   285
      Left            =   7680
      MaxLength       =   1
      TabIndex        =   208
      Text            =   "1"
      Top             =   3480
      Width           =   375
   End
   Begin VB.ListBox llr 
      BackColor       =   &H00000080&
      ForeColor       =   &H0000FFFF&
      Height          =   1425
      ItemData        =   "Form1.frx":030A
      Left            =   120
      List            =   "Form1.frx":0311
      TabIndex        =   173
      Top             =   480
      Width           =   8175
   End
   Begin VB.TextBox Text6 
      BackColor       =   &H80000006&
      ForeColor       =   &H80000005&
      Height          =   285
      Left            =   7200
      MaxLength       =   1
      TabIndex        =   172
      Top             =   3480
      Width           =   375
   End
   Begin VB.TextBox Text5 
      BackColor       =   &H80000006&
      ForeColor       =   &H80000005&
      Height          =   285
      Left            =   7200
      MaxLength       =   1
      TabIndex        =   171
      Top             =   2880
      Width           =   375
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H80000006&
      ForeColor       =   &H80000005&
      Height          =   285
      Left            =   7200
      MaxLength       =   1
      TabIndex        =   170
      Top             =   2280
      Width           =   375
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   8520
      Top             =   5040
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   8520
      Top             =   5640
   End
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   8520
      Top             =   6120
   End
   Begin VB.CommandButton Command10 
      Caption         =   "->"
      Height          =   255
      Left            =   8520
      TabIndex        =   169
      Top             =   6240
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton Command9 
      Caption         =   "->"
      Height          =   255
      Left            =   8520
      TabIndex        =   168
      Top             =   5640
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton Command8 
      Caption         =   "->"
      Height          =   255
      Left            =   8520
      TabIndex        =   167
      Top             =   5040
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton Command7 
      Caption         =   "<-"
      Height          =   255
      Left            =   9120
      TabIndex        =   122
      Top             =   6240
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton Command6 
      Caption         =   "<-"
      Height          =   255
      Left            =   9120
      TabIndex        =   121
      Top             =   5640
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton Command5 
      Caption         =   "<-"
      Height          =   255
      Left            =   9120
      TabIndex        =   120
      Top             =   5040
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton Command4 
      Caption         =   "PASTE TEXT"
      Height          =   255
      Left            =   960
      TabIndex        =   111
      Top             =   4560
      Width           =   4935
   End
   Begin VB.CommandButton Command3 
      Caption         =   " CLS"
      Height          =   255
      Left            =   7440
      TabIndex        =   110
      Top             =   4560
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "CLS"
      Height          =   255
      Left            =   0
      TabIndex        =   109
      Top             =   4560
      Width           =   855
   End
   Begin VB.Timer Timereng1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   4800
      Top             =   480
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   2400
      Top             =   4560
   End
   Begin RichTextLib.RichTextBox TEMP 
      Height          =   135
      Left            =   3360
      TabIndex        =   107
      Top             =   4680
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   238
      _Version        =   393217
      Enabled         =   0   'False
      ReadOnly        =   -1  'True
      TextRTF         =   $"Form1.frx":031C
   End
   Begin RichTextLib.RichTextBox TexT2 
      Height          =   4815
      Left            =   8400
      TabIndex        =   106
      Top             =   2040
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   8493
      _Version        =   393217
      BackColor       =   14737632
      ReadOnly        =   -1  'True
      TextRTF         =   $"Form1.frx":03A0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin RichTextLib.RichTextBox TexT1 
      Height          =   1935
      Left            =   0
      TabIndex        =   105
      Top             =   4920
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   3413
      _Version        =   393217
      BackColor       =   8421504
      TextRTF         =   $"Form1.frx":0426
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Decoding"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   13.5
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   6000
      TabIndex        =   104
      Top             =   4920
      Width           =   2295
   End
   Begin RichTextLib.RichTextBox labelspacewrong 
      Height          =   375
      Left            =   6840
      TabIndex        =   195
      Top             =   120
      Visible         =   0   'False
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   661
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"Form1.frx":04AA
   End
   Begin VB.Timer Timer5 
      Interval        =   1
      Left            =   8160
      Top             =   120
   End
   Begin VB.Frame TEXT101 
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      Height          =   2295
      Left            =   0
      TabIndex        =   174
      Top             =   1680
      Width           =   11655
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         BackColor       =   &H80000006&
         ForeColor       =   &H00008000&
         Height          =   285
         Left            =   5760
         TabIndex        =   198
         Text            =   "1"
         Top             =   1890
         Width           =   375
      End
      Begin VB.Shape Shape3 
         BorderColor     =   &H8000000D&
         BorderWidth     =   3
         Height          =   615
         Left            =   4680
         Top             =   1680
         Width           =   1695
      End
      Begin VB.Label LBLMS 
         BackStyle       =   0  'Transparent
         Caption         =   " KEY :"
         ForeColor       =   &H0000FF00&
         Height          =   255
         Left            =   4920
         TabIndex        =   199
         Top             =   1920
         Width           =   735
      End
      Begin VB.Line Line3 
         BorderColor     =   &H80000009&
         X1              =   6600
         X2              =   7920
         Y1              =   1080
         Y2              =   1080
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000009&
         X1              =   6600
         X2              =   7920
         Y1              =   1680
         Y2              =   1680
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H8000000D&
         BorderWidth     =   3
         Height          =   1935
         Left            =   6360
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label A27 
         BackColor       =   &H000000FF&
         Caption         =   " "
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   6720
         TabIndex        =   177
         Top             =   600
         Width           =   255
      End
      Begin VB.Label b27 
         BackColor       =   &H000000FF&
         Caption         =   " "
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   6720
         TabIndex        =   176
         Top             =   1200
         Width           =   255
      End
      Begin VB.Label c27 
         BackColor       =   &H000000FF&
         Caption         =   " "
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   6720
         TabIndex        =   175
         Top             =   1800
         Width           =   255
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H80000009&
         Height          =   1695
         Left            =   6600
         Top             =   480
         Visible         =   0   'False
         Width           =   495
      End
   End
   Begin VB.Image Image3 
      Height          =   1380
      Left            =   9840
      Picture         =   "Form1.frx":0532
      Top             =   0
      Width           =   3120
   End
   Begin VB.Label Label46 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   12000
      TabIndex        =   207
      Top             =   1680
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.Label A46 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   12000
      TabIndex        =   206
      Top             =   2160
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.Label B46 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   12000
      TabIndex        =   205
      Top             =   2760
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.Label C46 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   12000
      TabIndex        =   204
      Top             =   3360
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.Image Image1 
      Height          =   435
      Left            =   7320
      Picture         =   "Form1.frx":E5B4
      Top             =   120
      Width           =   1065
   End
   Begin VB.Label Label45 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   11760
      TabIndex        =   203
      Top             =   1680
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Label A45 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   11760
      TabIndex        =   202
      Top             =   2160
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Label B45 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   11760
      TabIndex        =   201
      Top             =   2760
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Label C45 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   11760
      TabIndex        =   200
      Top             =   3360
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Label Label46asd 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Encrypted Text:"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   2160
      TabIndex        =   197
      Top             =   4200
      Width           =   1695
   End
   Begin VB.Label Label47asfds 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Decrypted Text:"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   6720
      TabIndex        =   196
      Top             =   4200
      Width           =   1695
   End
   Begin VB.Label Label45sadfg 
      BackStyle       =   0  'Transparent
      Caption         =   "V.3.2"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   0
      TabIndex        =   194
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Label41 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "@"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   10320
      TabIndex        =   193
      Top             =   1680
      Visible         =   0   'False
      Width           =   165
   End
   Begin VB.Label A41 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "@"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   10320
      TabIndex        =   192
      Top             =   2160
      Visible         =   0   'False
      Width           =   165
   End
   Begin VB.Label B41 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "@"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   10320
      TabIndex        =   191
      Top             =   2760
      Visible         =   0   'False
      Width           =   165
   End
   Begin VB.Label C41 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "@"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   10320
      TabIndex        =   190
      Top             =   3360
      Visible         =   0   'False
      Width           =   165
   End
   Begin VB.Label C42 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "%"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   10680
      TabIndex        =   189
      Top             =   3360
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.Label B42 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "%"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   10680
      TabIndex        =   188
      Top             =   2760
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.Label A42 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "%"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   10680
      TabIndex        =   187
      Top             =   2160
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.Label Label42 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "%"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   10680
      TabIndex        =   186
      Top             =   1680
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.Label Label43 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "$"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   11040
      TabIndex        =   185
      Top             =   1680
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Label A43 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "$"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   11040
      TabIndex        =   184
      Top             =   2160
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Label B43 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "$"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   11040
      TabIndex        =   183
      Top             =   2760
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Label C43 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "$"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   11040
      TabIndex        =   182
      Top             =   3360
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Label C44 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "_"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   11400
      TabIndex        =   181
      Top             =   3360
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Label B44 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "_"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   11400
      TabIndex        =   180
      Top             =   2760
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Label A44 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "_"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   11400
      TabIndex        =   179
      Top             =   2160
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Label Label44 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "_"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   11400
      TabIndex        =   178
      Top             =   1680
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Label C40 
      BackStyle       =   0  'Transparent
      Caption         =   "?"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   9960
      TabIndex        =   166
      Top             =   3360
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label C39 
      BackStyle       =   0  'Transparent
      Caption         =   "."
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   9720
      TabIndex        =   165
      Top             =   3360
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label C38 
      BackStyle       =   0  'Transparent
      Caption         =   "!"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   9480
      TabIndex        =   164
      Top             =   3360
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label C37 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   9240
      TabIndex        =   163
      Top             =   3360
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label C36 
      BackStyle       =   0  'Transparent
      Caption         =   "9"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   9000
      TabIndex        =   162
      Top             =   3360
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label C35 
      BackStyle       =   0  'Transparent
      Caption         =   "8"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   8760
      TabIndex        =   161
      Top             =   3360
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label C34 
      BackStyle       =   0  'Transparent
      Caption         =   "7"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   8520
      TabIndex        =   160
      Top             =   3360
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label C33 
      BackStyle       =   0  'Transparent
      Caption         =   "6"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   8280
      TabIndex        =   159
      Top             =   3360
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label C32 
      BackStyle       =   0  'Transparent
      Caption         =   "5"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   8040
      TabIndex        =   158
      Top             =   3360
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label C31 
      BackStyle       =   0  'Transparent
      Caption         =   "4"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   7800
      TabIndex        =   157
      Top             =   3360
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label C30 
      BackStyle       =   0  'Transparent
      Caption         =   "3"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   7560
      TabIndex        =   156
      Top             =   3360
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label B40 
      BackStyle       =   0  'Transparent
      Caption         =   "?"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   9960
      TabIndex        =   155
      Top             =   2760
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label B39 
      BackStyle       =   0  'Transparent
      Caption         =   "."
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   9720
      TabIndex        =   154
      Top             =   2760
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label B38 
      BackStyle       =   0  'Transparent
      Caption         =   "!"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   9480
      TabIndex        =   153
      Top             =   2760
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label B37 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   9240
      TabIndex        =   152
      Top             =   2760
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label B36 
      BackStyle       =   0  'Transparent
      Caption         =   "9"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   9000
      TabIndex        =   151
      Top             =   2760
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label B35 
      BackStyle       =   0  'Transparent
      Caption         =   "8"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   8760
      TabIndex        =   150
      Top             =   2760
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label B34 
      BackStyle       =   0  'Transparent
      Caption         =   "7"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   8520
      TabIndex        =   149
      Top             =   2760
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label B33 
      BackStyle       =   0  'Transparent
      Caption         =   "6"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   8280
      TabIndex        =   148
      Top             =   2760
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label B32 
      BackStyle       =   0  'Transparent
      Caption         =   "5"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   8040
      TabIndex        =   147
      Top             =   2760
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label B31 
      BackStyle       =   0  'Transparent
      Caption         =   "4"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   7800
      TabIndex        =   146
      Top             =   2760
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label B30 
      BackStyle       =   0  'Transparent
      Caption         =   "3"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   7560
      TabIndex        =   145
      Top             =   2760
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label A40 
      BackStyle       =   0  'Transparent
      Caption         =   "?"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   9960
      TabIndex        =   144
      Top             =   2160
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label A39 
      BackStyle       =   0  'Transparent
      Caption         =   "."
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   9720
      TabIndex        =   143
      Top             =   2160
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label A38 
      BackStyle       =   0  'Transparent
      Caption         =   "!"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   9480
      TabIndex        =   142
      Top             =   2160
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label A37 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   9240
      TabIndex        =   141
      Top             =   2160
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label A36 
      BackStyle       =   0  'Transparent
      Caption         =   "9"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   9000
      TabIndex        =   140
      Top             =   2160
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label A35 
      BackStyle       =   0  'Transparent
      Caption         =   "8"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   8760
      TabIndex        =   139
      Top             =   2160
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label A34 
      BackStyle       =   0  'Transparent
      Caption         =   "7"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   8520
      TabIndex        =   138
      Top             =   2160
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label A33 
      BackStyle       =   0  'Transparent
      Caption         =   "6"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   8280
      TabIndex        =   137
      Top             =   2160
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label A32 
      BackStyle       =   0  'Transparent
      Caption         =   "5"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   8040
      TabIndex        =   136
      Top             =   2160
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label A31 
      BackStyle       =   0  'Transparent
      Caption         =   "4"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   7800
      TabIndex        =   135
      Top             =   2160
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label A30 
      BackStyle       =   0  'Transparent
      Caption         =   "3"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   7560
      TabIndex        =   134
      Top             =   2160
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label Label40 
      BackStyle       =   0  'Transparent
      Caption         =   "?"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   9960
      TabIndex        =   133
      Top             =   1680
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label Label39 
      BackStyle       =   0  'Transparent
      Caption         =   "."
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   9720
      TabIndex        =   132
      Top             =   1680
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label Label38 
      BackStyle       =   0  'Transparent
      Caption         =   "!"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   9480
      TabIndex        =   131
      Top             =   1680
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label Label37 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   9240
      TabIndex        =   130
      Top             =   1680
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label Label36 
      BackStyle       =   0  'Transparent
      Caption         =   "9"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   9000
      TabIndex        =   129
      Top             =   1680
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label Label35 
      BackStyle       =   0  'Transparent
      Caption         =   "8"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   8760
      TabIndex        =   128
      Top             =   1680
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label Label34 
      BackStyle       =   0  'Transparent
      Caption         =   "7"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   8520
      TabIndex        =   127
      Top             =   1680
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label Label33 
      BackStyle       =   0  'Transparent
      Caption         =   "6"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   8280
      TabIndex        =   126
      Top             =   1680
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label Label32 
      BackStyle       =   0  'Transparent
      Caption         =   "5"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   8040
      TabIndex        =   125
      Top             =   1680
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label Label31 
      BackStyle       =   0  'Transparent
      Caption         =   "4"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   7800
      TabIndex        =   124
      Top             =   1680
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label Label30 
      BackStyle       =   0  'Transparent
      Caption         =   "3"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   7560
      TabIndex        =   123
      Top             =   1680
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label C29 
      BackStyle       =   0  'Transparent
      Caption         =   "2"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   7320
      TabIndex        =   119
      Top             =   3360
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label A29 
      BackStyle       =   0  'Transparent
      Caption         =   "2"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   7320
      TabIndex        =   118
      Top             =   2160
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label B29 
      BackStyle       =   0  'Transparent
      Caption         =   "2"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   7320
      TabIndex        =   117
      Top             =   2760
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label Label29 
      BackStyle       =   0  'Transparent
      Caption         =   "2"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   7320
      TabIndex        =   116
      Top             =   1680
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label C28 
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   7080
      TabIndex        =   115
      Top             =   3360
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label B28 
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   7080
      TabIndex        =   114
      Top             =   2760
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label A28 
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   7080
      TabIndex        =   113
      Top             =   2160
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label Label28 
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   7080
      TabIndex        =   112
      Top             =   1680
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label Label27 
      BackColor       =   &H00FF0000&
      Caption         =   " "
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   6720
      TabIndex        =   108
      Top             =   1680
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000004&
      Visible         =   0   'False
      X1              =   120
      X2              =   10200
      Y1              =   2040
      Y2              =   2040
   End
   Begin VB.Label Label26 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "F"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   6360
      TabIndex        =   103
      Top             =   1680
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Label Label25 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "G"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   6120
      TabIndex        =   102
      Top             =   1680
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.Label Label24 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "H"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   5880
      TabIndex        =   101
      Top             =   1680
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.Label Label23 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "J"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   5640
      TabIndex        =   100
      Top             =   1680
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.Label Label22 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "K"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   5400
      TabIndex        =   99
      Top             =   1680
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label Label21 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "L"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   5160
      TabIndex        =   98
      Top             =   1680
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Z"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   4920
      TabIndex        =   97
      Top             =   1680
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   4680
      TabIndex        =   96
      Top             =   1680
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "C"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   4440
      TabIndex        =   95
      Top             =   1680
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "V"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   4200
      TabIndex        =   94
      Top             =   1680
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "B"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   3960
      TabIndex        =   93
      Top             =   1680
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "N"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   3720
      TabIndex        =   92
      Top             =   1680
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "M"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   3480
      TabIndex        =   91
      Top             =   1680
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "D"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   3240
      TabIndex        =   90
      Top             =   1680
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "S"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   3000
      TabIndex        =   89
      Top             =   1680
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "A"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   2760
      TabIndex        =   88
      Top             =   1680
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "P"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   2520
      TabIndex        =   87
      Top             =   1680
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "O"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   2280
      TabIndex        =   86
      Top             =   1680
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "I"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   2040
      TabIndex        =   85
      Top             =   1680
      Visible         =   0   'False
      Width           =   165
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "U"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   1800
      TabIndex        =   84
      Top             =   1680
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Y"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   1560
      TabIndex        =   83
      Top             =   1680
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "T"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   1320
      TabIndex        =   82
      Top             =   1680
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "R"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   1080
      TabIndex        =   81
      Top             =   1680
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "E"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   840
      TabIndex        =   80
      Top             =   1680
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "W"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   600
      TabIndex        =   79
      Top             =   1680
      Visible         =   0   'False
      Width           =   165
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Q"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   360
      TabIndex        =   78
      Top             =   1680
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.Label C26 
      BackStyle       =   0  'Transparent
      Caption         =   "Q"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   6360
      TabIndex        =   77
      Top             =   3360
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label C25 
      BackStyle       =   0  'Transparent
      Caption         =   "W"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   6120
      TabIndex        =   76
      Top             =   3360
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label C24 
      BackStyle       =   0  'Transparent
      Caption         =   "E"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   5880
      TabIndex        =   75
      Top             =   3360
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label C23 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "R"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   5640
      TabIndex        =   74
      Top             =   3360
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.Label C22 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "T"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   5400
      TabIndex        =   73
      Top             =   3360
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label C21 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Y"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   5160
      TabIndex        =   72
      Top             =   3360
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label C20 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "U"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   4920
      TabIndex        =   71
      Top             =   3360
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.Label C19 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "I"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   4680
      TabIndex        =   70
      Top             =   3360
      Visible         =   0   'False
      Width           =   165
   End
   Begin VB.Label C18 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "O"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   4440
      TabIndex        =   69
      Top             =   3360
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.Label C17 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "P"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   4200
      TabIndex        =   68
      Top             =   3360
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label C16 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "A"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   3960
      TabIndex        =   67
      Top             =   3360
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label C15 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "S"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   3720
      TabIndex        =   66
      Top             =   3360
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label C14 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "D"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   3480
      TabIndex        =   65
      Top             =   3360
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.Label C13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "F"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   3240
      TabIndex        =   64
      Top             =   3360
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Label C12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "G"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   3000
      TabIndex        =   63
      Top             =   3360
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.Label C11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "H"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   2760
      TabIndex        =   62
      Top             =   3360
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.Label C10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "J"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   2520
      TabIndex        =   61
      Top             =   3360
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.Label C9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "K"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   2280
      TabIndex        =   60
      Top             =   3360
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label C8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "L"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   2040
      TabIndex        =   59
      Top             =   3360
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Label C7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Z"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   1800
      TabIndex        =   58
      Top             =   3360
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label C6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   1560
      TabIndex        =   57
      Top             =   3360
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label C5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "C"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   1320
      TabIndex        =   56
      Top             =   3360
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label C4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "V"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   1080
      TabIndex        =   55
      Top             =   3360
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label C3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "B"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   840
      TabIndex        =   54
      Top             =   3360
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label C2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "N"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   600
      TabIndex        =   53
      Top             =   3360
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.Label C1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "M"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   360
      TabIndex        =   52
      Top             =   3360
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label B26 
      BackStyle       =   0  'Transparent
      Caption         =   "M"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   6360
      TabIndex        =   51
      Top             =   2760
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label B25 
      BackStyle       =   0  'Transparent
      Caption         =   "N"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   6120
      TabIndex        =   50
      Top             =   2760
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label B24 
      BackStyle       =   0  'Transparent
      Caption         =   "B"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   5880
      TabIndex        =   49
      Top             =   2760
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label B23 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "V"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   5640
      TabIndex        =   48
      Top             =   2760
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label B22 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "C"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   5400
      TabIndex        =   47
      Top             =   2760
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label B21 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   5160
      TabIndex        =   46
      Top             =   2760
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label B20 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Z"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   4920
      TabIndex        =   45
      Top             =   2760
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label B19 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "L"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   4680
      TabIndex        =   44
      Top             =   2760
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Label B18 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "K"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   4440
      TabIndex        =   43
      Top             =   2760
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label B17 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "J"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   4200
      TabIndex        =   42
      Top             =   2760
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.Label B16 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "H"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   3960
      TabIndex        =   41
      Top             =   2760
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.Label B15 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "G"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   3720
      TabIndex        =   40
      Top             =   2760
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.Label B14 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "F"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   3480
      TabIndex        =   39
      Top             =   2760
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Label B13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "D"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   3240
      TabIndex        =   38
      Top             =   2760
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.Label B12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "A"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   3000
      TabIndex        =   37
      Top             =   2760
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label B11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "S"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   2760
      TabIndex        =   36
      Top             =   2760
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label B10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "P"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   2520
      TabIndex        =   35
      Top             =   2760
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label B9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "O"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   2280
      TabIndex        =   34
      Top             =   2760
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.Label B8 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "I"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   2070
      TabIndex        =   33
      Top             =   2760
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.Label B7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "U"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   1800
      TabIndex        =   32
      Top             =   2760
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.Label B6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Y"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   1560
      TabIndex        =   31
      Top             =   2760
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label B5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "T"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   1320
      TabIndex        =   30
      Top             =   2760
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label B4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "R"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   1080
      TabIndex        =   29
      Top             =   2760
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.Label B3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "E"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   840
      TabIndex        =   28
      Top             =   2760
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label B2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "W"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   600
      TabIndex        =   27
      Top             =   2760
      Visible         =   0   'False
      Width           =   165
   End
   Begin VB.Label B1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Q"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   360
      TabIndex        =   26
      Top             =   2760
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.Label A26 
      BackStyle       =   0  'Transparent
      Caption         =   "M"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   6360
      TabIndex        =   25
      Top             =   2160
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label A25 
      BackStyle       =   0  'Transparent
      Caption         =   "N"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   6120
      TabIndex        =   24
      Top             =   2160
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label A24 
      BackStyle       =   0  'Transparent
      Caption         =   "B"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   5880
      TabIndex        =   23
      Top             =   2160
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label A23 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "V"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   5640
      TabIndex        =   22
      Top             =   2160
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label A22 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "C"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   5400
      TabIndex        =   21
      Top             =   2160
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label A21 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   5160
      TabIndex        =   20
      Top             =   2160
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label A20 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "J"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   4920
      TabIndex        =   19
      Top             =   2160
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.Label A19 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "L"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   4680
      TabIndex        =   18
      Top             =   2160
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Label A18 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Z"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   4440
      TabIndex        =   17
      Top             =   2160
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label A17 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "K"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   4200
      TabIndex        =   16
      Top             =   2160
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label A16 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "H"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   3960
      TabIndex        =   15
      Top             =   2160
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.Label A15 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "G"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   3720
      TabIndex        =   14
      Top             =   2160
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.Label A14 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "F"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   3480
      TabIndex        =   13
      Top             =   2160
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Label A13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "D"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   3240
      TabIndex        =   12
      Top             =   2160
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.Label A12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "S"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   3000
      TabIndex        =   11
      Top             =   2160
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label A11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "A"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   2760
      TabIndex        =   10
      Top             =   2160
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label A10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "P"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   2520
      TabIndex        =   9
      Top             =   2160
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label A9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "O"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   2280
      TabIndex        =   8
      Top             =   2160
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.Label A8 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "I"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   2070
      TabIndex        =   7
      Top             =   2160
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.Label A7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "U"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   1800
      TabIndex        =   6
      Top             =   2160
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.Label A6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Y"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   1560
      TabIndex        =   5
      Top             =   2160
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label A5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "T"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   1320
      TabIndex        =   4
      Top             =   2160
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label A4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "R"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   1080
      TabIndex        =   3
      Top             =   2160
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.Label A3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "E"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   840
      TabIndex        =   2
      Top             =   2160
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label A2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "W"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   600
      TabIndex        =   1
      Top             =   2160
      Visible         =   0   'False
      Width           =   165
   End
   Begin VB.Label A1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Q"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   360
      TabIndex        =   0
      Top             =   2160
      Visible         =   0   'False
      Width           =   120
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim k As String, p As String
Dim lng As Integer
Dim tempa, tempb, tempc
Dim tempAA As String
Dim tempBB As String
Dim tempCC As String
Dim hh

Dim EnG As String
Dim t As Integer

Private Sub Command1_Click()
TexT2.Text = ""

t = Val(Text3.Text)
A27.ForeColor = &HFF&
b27.ForeColor = &HFF&
c27.ForeColor = &HFF&
TexT1.Text = UCase(TexT1.Text)
Timer1.Enabled = True
Text3.Locked = True
llr.List(hh) = (A27 & " " & b27 & " " & c27)
hh = hh + 1
End Sub

Private Sub Command10_Click()
tempCC = C40
C40 = C39
C39 = C38
C38 = C37
C37 = C36
C36 = C35
C35 = C34
C34 = C33
C33 = C32
C32 = C31
C31 = C30
C30 = C29
C29 = C28
C28 = c27
c27 = C26
C26 = C25
C25 = C24
C24 = C23
C23 = C22
C22 = C21
C21 = C20
C20 = C19
C19 = C18
C18 = C17
C17 = C16
C16 = C15
C15 = C14
C14 = C13
C13 = C12
C12 = C11
C11 = C10
C10 = C9
C9 = C8
C8 = C7
C7 = C6
C6 = C5
C5 = C4
C4 = C3
C3 = C2
C2 = C1
C1 = tempCC
End Sub







Private Sub Command12_Click()

End Sub

Private Sub Command13_Click()

End Sub

Private Sub Command11_Click()
Load Form3

End Sub

Private Sub Command2_Click()
TexT1.Text = ""
End Sub

Private Sub Command3_Click()
TexT2.Text = ""

End Sub

Private Sub Command4_Click()
TexT1.Text = Clipboard.GetText()

End Sub

Private Sub Command5_Click()
tempa = A1
A1 = A2
A2 = A3
A3 = A4
A4 = A5
A5 = A6
A6 = A7
A7 = A8
A8 = A9
A9 = A10
A10 = A11
A11 = A12
A12 = A13
A13 = A14
A14 = A15
A15 = A16
A16 = A17
A17 = A18
A18 = A19
A19 = A20
A20 = A21
A21 = A22
A22 = A23
A23 = A24
A24 = A25
A25 = A26
A26 = A27
A27 = A28
A28 = A29
A29 = A30
A30 = A31
A31 = A32
A32 = A33
A33 = A34
A34 = A35
A35 = A36
A36 = A37
A37 = A38
A38 = A39
A39 = A40
A40 = tempa
End Sub

Private Sub Command6_Click()
tempb = B1
B1 = B2
B2 = B3
B3 = B4
B4 = B5
B5 = B6
B6 = B7
B7 = B8
B8 = B9
B9 = B10
B10 = B11
B11 = B12
B12 = B13
B13 = B14
B14 = B15
B15 = B16
B16 = B17
B17 = B18
B18 = B19
B19 = B20
B20 = B21
B21 = B22
B22 = B23
B23 = B24
B24 = B25
B25 = B26
B26 = b27
b27 = B28
B28 = B29
B29 = B30
B30 = B31
B31 = B32
B32 = B33
B33 = B34
B34 = B35
B35 = B36
B36 = B37
B37 = B38
B38 = B39
B39 = B40
B40 = tempb
End Sub

Private Sub Command7_Click()
tempc = C1
C1 = C2
C2 = C3
C3 = C4
C4 = C5
C5 = C6
C6 = C7
C7 = C8
C8 = C9
C9 = C10
C10 = C11
C11 = C12
C12 = C13
C13 = C14
C14 = C15
C15 = C16
C16 = C17
C17 = C18
C18 = C19
C19 = C20
C20 = C21
C21 = C22
C22 = C23
C23 = C24
C24 = C25
C25 = C26
C26 = c27
c27 = C28
C28 = C29
C29 = C30
C30 = C31
C31 = C32
C32 = C33
C33 = C34
C34 = C35
C35 = C36
C36 = C37
C37 = C38
C38 = C39
C39 = C40
C40 = tempc
End Sub

Private Sub Command8_Click()
tempAA = A40
A40 = A39
A39 = A38
A38 = A37
A37 = A36
A36 = A35
A35 = A34
A34 = A33
A33 = A32
A32 = A31
A31 = A30
A30 = A29
A29 = A28
A28 = A27
A27 = A26
A26 = A25
A25 = A24
A24 = A23
A23 = A22
A22 = A21
A21 = A20
A20 = A19
A19 = A18
A18 = A17
A17 = A16
A16 = A15
A15 = A14
A14 = A13
A13 = A12
A12 = A11
A11 = A10
A10 = A9
A9 = A8
A8 = A7
A7 = A6
A6 = A5
A5 = A4
A4 = A3
A3 = A2
A2 = A1
A1 = tempAA
End Sub

Private Sub Command9_Click()
tempBB = B40
B40 = B39
B39 = B38
B38 = B37
B37 = B36
B36 = B35
B35 = B34
B34 = B33
B33 = B32
B32 = B31
B31 = B30
B30 = B29
B29 = B28
B28 = b27
b27 = B26
B26 = B25
B25 = B24
B24 = B23
B23 = B22
B22 = B21
B21 = B20
B20 = B19
B19 = B18
B18 = B17
B17 = B16
B16 = B15
B15 = B14
B14 = B13
B13 = B12
B12 = B11
B11 = B10
B10 = B9
B9 = B8
B8 = B7
B7 = B6
B6 = B5
B5 = B4
B4 = B3
B3 = B2
B2 = B1
B1 = tempBB
End Sub

Private Sub Form_Load()
hh = 1
End Sub

Private Sub llr_Click()

Dim ires As Integer

If llr.ListIndex = 0 Then
ires = MsgBox("Are you sure are you want to set this key?", vbQuestion + vbYesNo + vbDefaultButton1, "Question")
If ires = 6 Then

Dim setkeys As String
Dim a As String, b As String, c As String

setkeys = llr.List(llr.ListIndex)

If setkeys <> "" Or setkeys <> "KEYS:" Then
setkeys = Replace(setkeys, " ", "")

a = Left(setkeys, 1)
Text4 = a
A27.ForeColor = &HFF&
Timer2.Enabled = True
''''''''''''''''''''''''
b = Left(setkeys, 2)
b = Right(b, 1)
Text5 = b
b27.ForeColor = &HFF&
Timer3.Enabled = True

c = Right(setkeys, 1)
Text6 = c
c27.ForeColor = &HFF&
Timer4.Enabled = True

End If

End If

Else



setkeys = llr.List(llr.ListIndex)

If setkeys <> "" Or setkeys <> "KEYS:" Then
setkeys = Replace(setkeys, " ", "")

a = Left(setkeys, 1)
Text4 = a
A27.ForeColor = &HFF&
Timer2.Enabled = True
''''''''''''''''''''''''
b = Left(setkeys, 2)
b = Right(b, 1)
Text5 = b
b27.ForeColor = &HFF&
Timer3.Enabled = True

c = Right(setkeys, 1)
Text6 = c
c27.ForeColor = &HFF&
Timer4.Enabled = True

End If

End If


End Sub

Private Sub TexT1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
TexT2.Text = ""
t = Val(Text3.Text)
A27.ForeColor = &HFF&
b27.ForeColor = &HFF&
c27.ForeColor = &HFF&
TexT1.Text = UCase(TexT1.Text)
Timer1.Enabled = True
Text3.Locked = True
llr.List(hh) = (A27 & " " & b27 & " " & c27)
hh = hh + 1
End If
End Sub

Private Sub Text3_Change()
If Text3.Text = "" Then
Text3 = ""
ElseIf Val(Text3.Text) <= 0 Or Val(Text3.Text) > 3 Then
Text3 = "1"
End If
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
If Len(Text4.Text) = 0 Then
MsgBox "SET CHARACTER", vbCritical, "ERROR"
Text4.SetFocus
Else
A27.ForeColor = &HFF&
Timer2.Enabled = True
End If
Text5.SetFocus
End If

End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
If Len(Text5.Text) = 0 Then
MsgBox "SET CHARACTER", vbCritical, "ERROR"
Text5.SetFocus
Else
b27.ForeColor = &HFF&
Timer3.Enabled = True
End If
Text6.SetFocus
End If

End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
If Len(Text6.Text) = 0 Then
MsgBox "SET CHARACTER", vbCritical, "ERROR"
Text6.SetFocus
Else
Timer4.Enabled = True
c27.ForeColor = &HFF&
End If
TexT1.SetFocus
End If

End Sub

Private Sub Timer1_Timer()
Dim metr As Integer

If Timer5.Enabled = False Then
Timer1.Enabled = False '}anticracking - ignore
End If

If Len(TexT1.Text) <> 0 Then

p = Left(TexT1.Text, 1)
lng = Len(TexT1.Text)
lng = lng - 1
TEMP.Text = Right(TexT1.Text, lng)
TexT1.Text = TEMP.Text

'If p = labelspacewrong.Text Then
'TexT1.Text = Right(TexT1.Text, Len(TexT1.Text) - 1)
'GoTo CONT2
'End If

If t = 1 Then
If p = Label1 Then
TexT2.Text = TexT2.Text + A1
EnG = "1"
ElseIf p = Label2 Then
TexT2.Text = TexT2.Text + A2
EnG = "1"
ElseIf p = Label3 Then
TexT2.Text = TexT2.Text + A3
EnG = "1"
ElseIf p = Label4 Then
TexT2.Text = TexT2.Text + A4
EnG = "1"
ElseIf p = Label5 Then
TexT2.Text = TexT2.Text + A5
EnG = "1"
ElseIf p = Label6 Then
TexT2.Text = TexT2.Text + A6
EnG = "1"
ElseIf p = Label7 Then
TexT2.Text = TexT2.Text + A7
EnG = "1"
ElseIf p = Label8 Then
TexT2.Text = TexT2.Text + A8
EnG = "1"
ElseIf p = Label9 Then
TexT2.Text = TexT2.Text + A9
EnG = "1"
ElseIf p = Label10 Then
TexT2.Text = TexT2.Text + A10
EnG = "1"
ElseIf p = Label11 Then
TexT2.Text = TexT2.Text + A11
EnG = "1"

ElseIf p = Label12 Then
TexT2.Text = TexT2.Text + A12
EnG = "1"

ElseIf p = Label13 Then
TexT2.Text = TexT2.Text + A13
EnG = "1"
ElseIf p = Label14 Then
TexT2.Text = TexT2.Text + A14
EnG = "1"
ElseIf p = Label15 Then
TexT2.Text = TexT2.Text + A15
EnG = "1"
ElseIf p = Label16 Then
TexT2.Text = TexT2.Text + A16
EnG = "1"
ElseIf p = Label17 Then
TexT2.Text = TexT2.Text + A17
EnG = "1"
ElseIf p = Label18 Then
TexT2.Text = TexT2.Text + A18
EnG = "1"
ElseIf p = Label19 Then
TexT2.Text = TexT2.Text + A19
EnG = "1"
ElseIf p = Label20 Then
TexT2.Text = TexT2.Text + A20
EnG = "1"
ElseIf p = Label21 Then
TexT2.Text = TexT2.Text + A21
EnG = "1"
ElseIf p = Label22 Then
TexT2.Text = TexT2.Text + A22
EnG = "1"
ElseIf p = Label23 Then
TexT2.Text = TexT2.Text + A23
EnG = "1"
ElseIf p = Label24 Then
TexT2.Text = TexT2.Text + A24
EnG = "1"
ElseIf p = Label25 Then
TexT2.Text = TexT2.Text + A25
EnG = "1"
ElseIf p = Label26 Then
TexT2.Text = TexT2.Text + A26
EnG = "1"
ElseIf p = Label27 Then
TexT2.Text = TexT2.Text + A27
EnG = "1"
ElseIf p = Label28 Then
TexT2.Text = TexT2.Text + A28
EnG = "1"
ElseIf p = Label29 Then
TexT2.Text = TexT2.Text + A29
EnG = "1"
ElseIf p = Label30 Then
TexT2.Text = TexT2.Text + A30
EnG = "1"
ElseIf p = Label31 Then
TexT2.Text = TexT2.Text + A31
EnG = "1"
ElseIf p = Label32 Then
TexT2.Text = TexT2.Text + A32
EnG = "1"
ElseIf p = Label33 Then
TexT2.Text = TexT2.Text + A33
EnG = "1"
ElseIf p = Label34 Then
TexT2.Text = TexT2.Text + A34
EnG = "1"
ElseIf p = Label35 Then
TexT2.Text = TexT2.Text + A35
EnG = "1"
ElseIf p = Label36 Then
TexT2.Text = TexT2.Text + A36
EnG = "1"
ElseIf p = Label37 Then
TexT2.Text = TexT2.Text + A37
EnG = "1"
ElseIf p = Label38 Then
TexT2.Text = TexT2.Text + A38
EnG = "1"
ElseIf p = Label39 Then
TexT2.Text = TexT2.Text + A39
EnG = "1"
ElseIf p = Label40 Then
TexT2.Text = TexT2.Text + A40
EnG = "1"
ElseIf p = Label41 Then
TexT2.Text = TexT2.Text + A41
EnG = "1"
ElseIf p = Label42 Then
TexT2.Text = TexT2.Text + A42
EnG = "1"
ElseIf p = Label43 Then
TexT2.Text = TexT2.Text + A43
EnG = "1"
ElseIf p = Label44 Then
TexT2.Text = TexT2.Text + A44
EnG = "1"
ElseIf p = Label45 Then
TexT2.Text = TexT2.Text + A45
EnG = "1"
ElseIf p = Label46 Then
TexT2.Text = TexT2.Text + A46
EnG = "1"
End If
ElseIf t = 2 Then
If p = Label1 Then
TexT2.Text = TexT2.Text + B1
EnG = "2"
ElseIf p = Label2 Then
TexT2.Text = TexT2.Text + B2
EnG = "2"
ElseIf p = Label3 Then
TexT2.Text = TexT2.Text + B3
EnG = "2"
ElseIf p = Label4 Then
TexT2.Text = TexT2.Text + B4
EnG = "2"
ElseIf p = Label5 Then
TexT2.Text = TexT2.Text + B5
EnG = "2"
ElseIf p = Label6 Then
TexT2.Text = TexT2.Text + B6
EnG = "2"
ElseIf p = Label7 Then
TexT2.Text = TexT2.Text + B7
EnG = "2"
ElseIf p = Label8 Then
TexT2.Text = TexT2.Text + B8
EnG = "2"
ElseIf p = Label9 Then
TexT2.Text = TexT2.Text + B9
EnG = "2"
ElseIf p = Label10 Then
TexT2.Text = TexT2.Text + B10
EnG = "2"
ElseIf p = Label11 Then
TexT2.Text = TexT2.Text + B11
EnG = "2"
ElseIf p = Label12 Then
TexT2.Text = TexT2.Text + B12
EnG = "2"
ElseIf p = Label13 Then
TexT2.Text = TexT2.Text + B13
EnG = "2"
ElseIf p = Label14 Then
TexT2.Text = TexT2.Text + B14
EnG = "2"
ElseIf p = Label15 Then
TexT2.Text = TexT2.Text + B15
EnG = "2"
ElseIf p = Label16 Then
TexT2.Text = TexT2.Text + B16
EnG = "2"
ElseIf p = Label17 Then
TexT2.Text = TexT2.Text + B17
EnG = "2"
ElseIf p = Label18 Then
TexT2.Text = TexT2.Text + B18
EnG = "2"
ElseIf p = Label19 Then
TexT2.Text = TexT2.Text + B19
EnG = "2"
ElseIf p = Label20 Then
TexT2.Text = TexT2.Text + B20
EnG = "2"
ElseIf p = Label21 Then
TexT2.Text = TexT2.Text + B21
EnG = "2"
ElseIf p = Label22 Then
TexT2.Text = TexT2.Text + B22
EnG = "2"
ElseIf p = Label23 Then
TexT2.Text = TexT2.Text + B23
EnG = "2"
ElseIf p = Label24 Then
TexT2.Text = TexT2.Text + B24
EnG = "2"
ElseIf p = Label25 Then
TexT2.Text = TexT2.Text + B25
EnG = "2"
ElseIf p = Label26 Then
TexT2.Text = TexT2.Text + B26
EnG = "2"
ElseIf p = Label27 Then
TexT2.Text = TexT2.Text + b27
EnG = "2"
ElseIf p = Label28 Then
TexT2.Text = TexT2.Text + B28
EnG = "2"
ElseIf p = Label29 Then
TexT2.Text = TexT2.Text + B29
EnG = "2"
ElseIf p = Label30 Then
TexT2.Text = TexT2.Text + B30
EnG = "2"
ElseIf p = Label31 Then
TexT2.Text = TexT2.Text + B31
EnG = "2"
ElseIf p = Label32 Then
TexT2.Text = TexT2.Text + B32
EnG = "2"
ElseIf p = Label33 Then
TexT2.Text = TexT2.Text + B33
EnG = "2"
ElseIf p = Label34 Then
TexT2.Text = TexT2.Text + B34
EnG = "2"
ElseIf p = Label35 Then
TexT2.Text = TexT2.Text + B35
EnG = "2"
ElseIf p = Label36 Then
TexT2.Text = TexT2.Text + B36
EnG = "2"
ElseIf p = Label37 Then
TexT2.Text = TexT2.Text + B37
EnG = "2"
ElseIf p = Label38 Then
TexT2.Text = TexT2.Text + B38
EnG = "2"
ElseIf p = Label39 Then
TexT2.Text = TexT2.Text + B39
EnG = "2"
ElseIf p = Label40 Then
TexT2.Text = TexT2.Text + B40
EnG = "2"
ElseIf p = Label41 Then
TexT2.Text = TexT2.Text + B41
EnG = "2"
ElseIf p = Label42 Then
TexT2.Text = TexT2.Text + B42
EnG = "2"
ElseIf p = Label43 Then
TexT2.Text = TexT2.Text + B43
EnG = "2"
ElseIf p = Label44 Then
TexT2.Text = TexT2.Text + B44
EnG = "2"
ElseIf p = Label45 Then
TexT2.Text = TexT2.Text + B45
EnG = "2"
ElseIf p = Label46 Then
TexT2.Text = TexT2.Text + B46
EnG = "2"
End If

ElseIf t = 3 Then

If p = Label1 Then
TexT2.Text = TexT2.Text + C1
EnG = "3"
ElseIf p = Label2 Then
TexT2.Text = TexT2.Text + C2
EnG = "3"

ElseIf p = Label3 Then
TexT2.Text = TexT2.Text + C3
EnG = "3"
ElseIf p = Label4 Then
TexT2.Text = TexT2.Text + C4
EnG = "3"
ElseIf p = Label5 Then
TexT2.Text = TexT2.Text + C5
EnG = "3"
ElseIf p = Label6 Then
TexT2.Text = TexT2.Text + C6
EnG = "3"
ElseIf p = Label7 Then
TexT2.Text = TexT2.Text + C7
EnG = "3"
ElseIf p = Label8 Then
TexT2.Text = TexT2.Text + C8
EnG = "3"
ElseIf p = Label9 Then
TexT2.Text = TexT2.Text + C9
EnG = "3"
ElseIf p = Label10 Then
TexT2.Text = TexT2.Text + C10
EnG = "3"
ElseIf p = Label11 Then
TexT2.Text = TexT2.Text + C11
EnG = "3"
ElseIf p = Label12 Then
TexT2.Text = TexT2.Text + C12
EnG = "3"
ElseIf p = Label13 Then
TexT2.Text = TexT2.Text + C13
EnG = "3"
ElseIf p = Label14 Then
TexT2.Text = TexT2.Text + C14
EnG = "3"

ElseIf p = Label15 Then
TexT2.Text = TexT2.Text + C15
EnG = "3"
ElseIf p = Label16 Then
TexT2.Text = TexT2.Text + C16
EnG = "3"
ElseIf p = Label17 Then
TexT2.Text = TexT2.Text + C17
EnG = "3"
ElseIf p = Label18 Then
TexT2.Text = TexT2.Text + C18
EnG = "3"
ElseIf p = Label19 Then
TexT2.Text = TexT2.Text + C19
EnG = "3"
ElseIf p = Label20 Then
TexT2.Text = TexT2.Text + C20
EnG = "3"
ElseIf p = Label21 Then
TexT2.Text = TexT2.Text + C21
EnG = "3"
ElseIf p = Label22 Then
TexT2.Text = TexT2.Text + C22
EnG = "3"
ElseIf p = Label23 Then
TexT2.Text = TexT2.Text + C23
EnG = "3"
ElseIf p = Label24 Then
TexT2.Text = TexT2.Text + C24
EnG = "3"

ElseIf p = Label25 Then
TexT2.Text = TexT2.Text + C25
EnG = "3"
ElseIf p = Label26 Then
TexT2.Text = TexT2.Text + C26
EnG = "3"
ElseIf p = Label27 Then
TexT2.Text = TexT2.Text + c27
EnG = "3"
ElseIf p = Label28 Then
TexT2.Text = TexT2.Text + C28
EnG = "3"
ElseIf p = Label29 Then
TexT2.Text = TexT2.Text + C29
EnG = "3"
ElseIf p = Label30 Then
TexT2.Text = TexT2.Text + C30
EnG = "3"
ElseIf p = Label31 Then
TexT2.Text = TexT2.Text + C31
EnG = "3"
ElseIf p = Label32 Then
TexT2.Text = TexT2.Text + C32
EnG = "3"
ElseIf p = Label33 Then
TexT2.Text = TexT2.Text + C33
EnG = "3"
ElseIf p = Label34 Then
TexT2.Text = TexT2.Text + C34
EnG = "3"
ElseIf p = Label35 Then
TexT2.Text = TexT2.Text + C35
EnG = "3"
ElseIf p = Label36 Then
TexT2.Text = TexT2.Text + C36
EnG = "3"
ElseIf p = Label37 Then
TexT2.Text = TexT2.Text + C37
EnG = "3"
ElseIf p = Label38 Then
TexT2.Text = TexT2.Text + C38
EnG = "3"
ElseIf p = Label39 Then
TexT2.Text = TexT2.Text + C39
EnG = "3"
ElseIf p = Label40 Then
TexT2.Text = TexT2.Text + C40
EnG = "3"
ElseIf p = Label41 Then
TexT2.Text = TexT2.Text + C41
EnG = "3"
ElseIf p = Label42 Then
TexT2.Text = TexT2.Text + C42
EnG = "3"
ElseIf p = Label43 Then
TexT2.Text = TexT2.Text + C43
EnG = "3"
ElseIf p = Label44 Then
TexT2.Text = TexT2.Text + C44
EnG = "3"
ElseIf p = Label45 Then
TexT2.Text = TexT2.Text + C45
EnG = "3"
ElseIf p = Label46 Then
TexT2.Text = TexT2.Text + C46
EnG = "3"
End If
End If

If EnG = "1" Then

t = t + 1
For metr = 1 To Val(Textc1)

tempa = A1
A1 = A2
A2 = A3
A3 = A4
A4 = A5
A5 = A6
A6 = A7
A7 = A8
A8 = A9
A9 = A10
A10 = A11
A11 = A12
A12 = A13
A13 = A14
A14 = A15
A15 = A16
A16 = A17
A17 = A18
A18 = A19
A19 = A20
A20 = A21
A21 = A22
A22 = A23
A23 = A24
A24 = A25
A25 = A26
A26 = A27
A27 = A28
A28 = A29
A29 = A30
A30 = A31
A31 = A32
A32 = A33
A33 = A34
A34 = A35
A35 = A36
A36 = A37
A37 = A38
A38 = A39
A39 = A40
A40 = A41
A41 = A42
A42 = A43
A43 = A44
A44 = A45
A45 = A46
A46 = tempa
Next

EnG = ""
ElseIf EnG = "2" Then
t = t + 1
For metr = 1 To Val(Textc2)
tempb = B1
B1 = B2
B2 = B3
B3 = B4
B4 = B5
B5 = B6
B6 = B7
B7 = B8
B8 = B9
B9 = B10
B10 = B11
B11 = B12
B12 = B13
B13 = B14
B14 = B15
B15 = B16
B16 = B17
B17 = B18
B18 = B19
B19 = B20
B20 = B21
B21 = B22
B22 = B23
B23 = B24
B24 = B25
B25 = B26
B26 = b27
b27 = B28
B28 = B29
B29 = B30
B30 = B31
B31 = B32
B32 = B33
B33 = B34
B34 = B35
B35 = B36
B36 = B37
B37 = B38
B38 = B39
B39 = B40
B40 = B41
B41 = B42
B42 = B43
B43 = B44
B44 = B45
B45 = B46
B46 = tempb
Next

EnG = ""
ElseIf EnG = "3" Then
t = 1
For metr = 1 To Val(Textc3)
tempc = C1
C1 = C2
C2 = C3
C3 = C4
C4 = C5
C5 = C6
C6 = C7
C7 = C8
C8 = C9
C9 = C10
C10 = C11
C11 = C12
C12 = C13
C13 = C14
C14 = C15
C15 = C16
C16 = C17
C17 = C18
C18 = C19
C19 = C20
C20 = C21
C21 = C22
C22 = C23
C23 = C24
C24 = C25
C25 = C26
C26 = c27
c27 = C28
C28 = C29
C29 = C30
C30 = C31
C31 = C32
C32 = C33
C33 = C34
C34 = C35
C35 = C36
C36 = C37
C37 = C38
C38 = C39
C39 = C40
C40 = C41
C41 = C42
C42 = C43
C43 = C44
C44 = C45
C45 = C46
C46 = tempc
Next

EnG = ""
End If



Else
Timer1.Enabled = False
Text3.Locked = False
A27.ForeColor = &H0&
b27.ForeColor = &H0&
c27.ForeColor = &H0&
Clipboard.Clear
Clipboard.SetText (TexT2.Text)
TexT2.Text = Replace(TexT2.Text, "-", Chr$(13))
End If
CONT2:
End Sub

Private Sub Timer2_Timer()
Dim metr As Integer

If A27.Caption = UCase(Text4.Text) Then
Timer2.Enabled = False
A27.ForeColor = vbBlack
Text4.Text = ""
Else
For metr = 1 To Val(Textc1)
tempAA = A46
A46 = A45
A45 = A44
A44 = A43
A43 = A42
A42 = A41
A41 = A40
A40 = A39
A39 = A38
A38 = A37
A37 = A36
A36 = A35
A35 = A34
A34 = A33
A33 = A32
A32 = A31
A31 = A30
A30 = A29
A29 = A28
A28 = A27
A27 = A26
A26 = A25
A25 = A24
A24 = A23
A23 = A22
A22 = A21
A21 = A20
A20 = A19
A19 = A18
A18 = A17
A17 = A16
A16 = A15
A15 = A14
A14 = A13
A13 = A12
A12 = A11
A11 = A10
A10 = A9
A9 = A8
A8 = A7
A7 = A6
A6 = A5
A5 = A4
A4 = A3
A3 = A2
A2 = A1
A1 = tempAA
Next metr

End If


End Sub

Private Sub Timer3_Timer()
Dim metr As Integer

If b27.Caption = UCase(Text5.Text) Then
Timer3.Enabled = False
b27.ForeColor = vbBlack
Text5.Text = ""
Else
For metr = 1 To Val(Textc2)
tempBB = B46
B46 = B45
B45 = B44
B44 = B43
B43 = B42
B42 = B41
B41 = B40
B40 = B39
B39 = B38
B38 = B37
B37 = B36
B36 = B35
B35 = B34
B34 = B33
B33 = B32
B32 = B31
B31 = B30
B30 = B29
B29 = B28
B28 = b27
b27 = B26
B26 = B25
B25 = B24
B24 = B23
B23 = B22
B22 = B21
B21 = B20
B20 = B19
B19 = B18
B18 = B17
B17 = B16
B16 = B15
B15 = B14
B14 = B13
B13 = B12
B12 = B11
B11 = B10
B10 = B9
B9 = B8
B8 = B7
B7 = B6
B6 = B5
B5 = B4
B4 = B3
B3 = B2
B2 = B1
B1 = tempBB
Next
End If


End Sub

Private Sub Timer4_Timer()
Dim metr As Integer

If c27.Caption = UCase(Text6.Text) Then
Timer4.Enabled = False
c27.ForeColor = vbBlack
Text6.Text = ""
Else
For metr = 1 To Val(Textc3)

tempCC = C46
C46 = C45
C45 = C44
C44 = C43
C43 = C42
C42 = C41
C41 = C40
C40 = C39
C39 = C38
C38 = C37
C37 = C36
C36 = C35
C35 = C34
C34 = C33
C33 = C32
C32 = C31
C31 = C30
C30 = C29
C29 = C28
C28 = c27
c27 = C26
C26 = C25
C25 = C24
C24 = C23
C23 = C22
C22 = C21
C21 = C20
C20 = C19
C19 = C18
C18 = C17
C17 = C16
C16 = C15
C15 = C14
C14 = C13
C13 = C12
C12 = C11
C11 = C10
C10 = C9
C9 = C8
C8 = C7
C7 = C6
C6 = C5
C5 = C4
C4 = C3
C3 = C2
C2 = C1
C1 = tempCC
Next metr
End If

End Sub

Private Sub Timer5_Timer()
If Timer1.Enabled = True Then
TexT2.Visible = False
Else
TexT2.Visible = True
End If
End Sub
