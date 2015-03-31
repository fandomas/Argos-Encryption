VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Form1 
   BackColor       =   &H80000012&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Argos - Encoder"
   ClientHeight    =   6900
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13170
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6900
   ScaleWidth      =   13170
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command12 
      Caption         =   "ASCII"
      Height          =   495
      Left            =   2520
      TabIndex        =   218
      Top             =   6000
      Width           =   735
   End
   Begin VB.ListBox List1 
      BackColor       =   &H80000007&
      ForeColor       =   &H000080FF&
      Height          =   3375
      ItemData        =   "Form1.frx":030A
      Left            =   1680
      List            =   "Form1.frx":0311
      TabIndex        =   200
      Top             =   240
      Width           =   1575
   End
   Begin RichTextLib.RichTextBox DEL 
      Height          =   255
      Left            =   2520
      TabIndex        =   199
      Top             =   6120
      Visible         =   0   'False
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   450
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"Form1.frx":0325
   End
   Begin VB.ListBox llt 
      BackColor       =   &H00000080&
      ForeColor       =   &H0000FFFF&
      Height          =   3375
      ItemData        =   "Form1.frx":03A9
      Left            =   480
      List            =   "Form1.frx":03B0
      TabIndex        =   174
      Top             =   240
      Width           =   1095
   End
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   12480
      Top             =   2280
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   12480
      Top             =   1800
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   12480
      Top             =   1320
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H80000006&
      ForeColor       =   &H80000005&
      Height          =   285
      Left            =   7320
      MaxLength       =   1
      TabIndex        =   16
      Top             =   2400
      Width           =   375
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000006&
      ForeColor       =   &H80000005&
      Height          =   285
      Left            =   7320
      MaxLength       =   1
      TabIndex        =   15
      Top             =   1800
      Width           =   375
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H80000006&
      ForeColor       =   &H80000005&
      Height          =   285
      Left            =   7320
      MaxLength       =   1
      TabIndex        =   14
      Top             =   1200
      Width           =   375
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      BackColor       =   &H80000006&
      ForeColor       =   &H8000000A&
      Height          =   285
      Left            =   6480
      MaxLength       =   1
      TabIndex        =   13
      Text            =   "1"
      Top             =   2880
      Width           =   495
   End
   Begin VB.CommandButton Command11 
      Caption         =   "<-"
      Height          =   255
      Left            =   10080
      TabIndex        =   12
      Top             =   5160
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton Command6 
      Caption         =   "<-"
      Height          =   255
      Left            =   10080
      TabIndex        =   11
      Top             =   5760
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton Command7 
      Caption         =   "<-"
      Height          =   255
      Left            =   10080
      TabIndex        =   10
      Top             =   6360
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton Command8 
      Caption         =   "->"
      Height          =   255
      Left            =   9480
      TabIndex        =   9
      Top             =   5160
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton Command9 
      Caption         =   "->"
      Height          =   255
      Left            =   9480
      TabIndex        =   8
      Top             =   5760
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton Command10 
      Caption         =   "->"
      Height          =   255
      Left            =   9480
      TabIndex        =   7
      Top             =   6360
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton Command5 
      Caption         =   "CLS ENC TXT"
      Height          =   615
      Left            =   8400
      TabIndex        =   6
      Top             =   6000
      Width           =   975
   End
   Begin VB.CommandButton Command4 
      Caption         =   "SHOW ENCRYPTED MESSAGE"
      Height          =   735
      Left            =   8040
      TabIndex        =   5
      Top             =   2880
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      Caption         =   "COPY RESULT"
      Height          =   375
      Left            =   10920
      TabIndex        =   4
      Top             =   6360
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "CLS REAL TEXT"
      Height          =   615
      Left            =   7320
      TabIndex        =   3
      Top             =   6000
      Width           =   975
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   6720
      Top             =   6120
   End
   Begin RichTextLib.RichTextBox txt1 
      Height          =   2175
      Left            =   480
      TabIndex        =   1
      Top             =   3720
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   3836
      _Version        =   393217
      BackColor       =   8421504
      TextRTF         =   $"Form1.frx":03BB
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ENCODING"
      Height          =   495
      Left            =   480
      TabIndex        =   0
      Top             =   6000
      Width           =   1935
   End
   Begin RichTextLib.RichTextBox txt2 
      Height          =   5895
      Left            =   9600
      TabIndex        =   2
      Top             =   360
      Visible         =   0   'False
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   10398
      _Version        =   393217
      BackColor       =   16777088
      ReadOnly        =   -1  'True
      TextRTF         =   $"Form1.frx":043F
   End
   Begin VB.Frame TEXT100 
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      Height          =   3015
      Left            =   120
      TabIndex        =   175
      Top             =   360
      Width           =   9375
      Begin VB.TextBox Textc3 
         Height          =   285
         Left            =   7680
         MaxLength       =   1
         TabIndex        =   216
         Text            =   "1"
         Top             =   2040
         Width           =   375
      End
      Begin VB.TextBox Textc2 
         Height          =   285
         Left            =   7680
         MaxLength       =   1
         TabIndex        =   215
         Text            =   "1"
         Top             =   1440
         Width           =   375
      End
      Begin VB.TextBox Textc1 
         Height          =   285
         Left            =   7680
         MaxLength       =   1
         TabIndex        =   214
         Text            =   "1"
         Top             =   840
         Width           =   375
      End
      Begin VB.Label Label47 
         BackStyle       =   0  'Transparent
         Caption         =   "Step:"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   7680
         TabIndex        =   217
         Top             =   480
         Width           =   615
      End
      Begin VB.Line Line3 
         BorderColor     =   &H80000001&
         X1              =   6360
         X2              =   7320
         Y1              =   1200
         Y2              =   1200
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000001&
         X1              =   6360
         X2              =   7320
         Y1              =   1800
         Y2              =   1800
      End
      Begin VB.Label Label47asgsfsdg 
         BackStyle       =   0  'Transparent
         Caption         =   "Column 3:"
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   5280
         TabIndex        =   198
         Top             =   2040
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Label46asfdgadsg 
         BackStyle       =   0  'Transparent
         Caption         =   "Column 2:"
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   5280
         TabIndex        =   197
         Top             =   1440
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Label45sgfdgsd 
         BackStyle       =   0  'Transparent
         Caption         =   "Column 1:"
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   5280
         TabIndex        =   196
         Top             =   840
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label LBLMS 
         BackStyle       =   0  'Transparent
         Caption         =   " KEY:"
         ForeColor       =   &H0000C000&
         Height          =   255
         Left            =   5280
         TabIndex        =   179
         Top             =   2520
         Width           =   855
      End
      Begin VB.Label c27 
         BackColor       =   &H000000FF&
         Caption         =   " "
         ForeColor       =   &H8000000E&
         Height          =   195
         Left            =   6480
         TabIndex        =   178
         Top             =   2040
         Width           =   285
      End
      Begin VB.Label b27 
         BackColor       =   &H000000FF&
         Caption         =   " "
         ForeColor       =   &H8000000E&
         Height          =   195
         Left            =   6480
         TabIndex        =   177
         Top             =   1440
         Width           =   285
      End
      Begin VB.Label A27 
         BackColor       =   &H000000FF&
         Caption         =   " "
         ForeColor       =   &H8000000E&
         Height          =   195
         Left            =   6480
         TabIndex        =   176
         Top             =   840
         Width           =   285
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H80000001&
         BorderWidth     =   2
         Height          =   1695
         Left            =   6360
         Top             =   720
         Width           =   495
      End
      Begin VB.Image Image1 
         Height          =   1380
         Left            =   3000
         Picture         =   "Form1.frx":04C3
         Top             =   600
         Width           =   3120
      End
   End
   Begin RichTextLib.RichTextBox labelspacewrong 
      Height          =   375
      Left            =   12480
      TabIndex        =   205
      Top             =   120
      Visible         =   0   'False
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   661
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"Form1.frx":E545
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
      Left            =   11760
      TabIndex        =   213
      Top             =   720
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
      Left            =   11760
      TabIndex        =   212
      Top             =   1200
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
      Left            =   11760
      TabIndex        =   211
      Top             =   1800
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
      Left            =   11760
      TabIndex        =   210
      Top             =   2400
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.Image Image5 
      Height          =   435
      Left            =   0
      Picture         =   "Form1.frx":E5CD
      Top             =   6480
      Width           =   1065
   End
   Begin VB.Image Image3 
      Height          =   135
      Left            =   240
      Top             =   6600
      Width           =   375
   End
   Begin VB.Label C45 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   11520
      TabIndex        =   209
      Top             =   2400
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Label B45 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   11520
      TabIndex        =   208
      Top             =   1800
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Label A45 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   11520
      TabIndex        =   207
      Top             =   1200
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Label Label45 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   11520
      TabIndex        =   206
      Top             =   720
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Label Label47sfg 
      BackStyle       =   0  'Transparent
      Caption         =   "V.3.2"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   0
      TabIndex        =   204
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Label46fghgj 
      BackStyle       =   0  'Transparent
      Caption         =   "Char. Support:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3360
      TabIndex        =   203
      Top             =   65
      Width           =   1095
   End
   Begin VB.Label stgdfgdfhdfh 
      BackStyle       =   0  'Transparent
      Caption         =   "1.224.024.534.000"
      ForeColor       =   &H8000000B&
      Height          =   255
      Left            =   1200
      TabIndex        =   202
      Top             =   6555
      Width           =   4695
   End
   Begin VB.Label Label45345 
      BackStyle       =   0  'Transparent
      Caption         =   "! $ :"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   11400
      TabIndex        =   201
      Top             =   60
      Width           =   735
   End
   Begin VB.Image Image4 
      Height          =   360
      Left            =   4320
      Picture         =   "Form1.frx":FE87
      Top             =   0
      Width           =   7185
   End
   Begin VB.Label Label44 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "_"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   11160
      TabIndex        =   195
      Top             =   720
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Label A44 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "_"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   11160
      TabIndex        =   194
      Top             =   1200
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Label B44 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "_"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   11160
      TabIndex        =   193
      Top             =   1800
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Label C44 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "_"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   11160
      TabIndex        =   192
      Top             =   2400
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Label C43 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "$"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   10800
      TabIndex        =   191
      Top             =   2400
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Label B43 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "$"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   10800
      TabIndex        =   190
      Top             =   1800
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Label A43 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "$"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   10800
      TabIndex        =   189
      Top             =   1200
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Label Label43 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "$"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   10800
      TabIndex        =   188
      Top             =   720
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Label Label42 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "%"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   10440
      TabIndex        =   187
      Top             =   720
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.Label A42 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "%"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   10440
      TabIndex        =   186
      Top             =   1200
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.Label B42 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "%"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   10440
      TabIndex        =   185
      Top             =   1800
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.Label C42 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "%"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   10440
      TabIndex        =   184
      Top             =   2400
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.Label C41 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "@"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   10080
      TabIndex        =   183
      Top             =   2400
      Visible         =   0   'False
      Width           =   165
   End
   Begin VB.Label B41 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "@"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   10080
      TabIndex        =   182
      Top             =   1800
      Visible         =   0   'False
      Width           =   165
   End
   Begin VB.Label A41 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "@"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   10080
      TabIndex        =   181
      Top             =   1200
      Visible         =   0   'False
      Width           =   165
   End
   Begin VB.Label Label41 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "@"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   10080
      TabIndex        =   180
      Top             =   720
      Visible         =   0   'False
      Width           =   165
   End
   Begin VB.Image Image2 
      Height          =   1380
      Left            =   9720
      Picture         =   "Form1.frx":185C9
      Top             =   3000
      Width           =   3120
   End
   Begin VB.Label A1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Q"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   240
      TabIndex        =   173
      Top             =   1200
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.Label A2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "W"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   480
      TabIndex        =   172
      Top             =   1200
      Visible         =   0   'False
      Width           =   165
   End
   Begin VB.Label A3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "E"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   720
      TabIndex        =   171
      Top             =   1200
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label A4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "R"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   960
      TabIndex        =   170
      Top             =   1200
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.Label A5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "T"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   1200
      TabIndex        =   169
      Top             =   1200
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label A6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Y"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   1440
      TabIndex        =   168
      Top             =   1200
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label A7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "U"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   1680
      TabIndex        =   167
      Top             =   1200
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
      Left            =   1950
      TabIndex        =   166
      Top             =   1200
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.Label A9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "O"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   2160
      TabIndex        =   165
      Top             =   1200
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.Label A10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "P"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   2400
      TabIndex        =   164
      Top             =   1200
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label A11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "A"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   2640
      TabIndex        =   163
      Top             =   1200
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label A12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "S"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   2880
      TabIndex        =   162
      Top             =   1200
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label A13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "D"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   3120
      TabIndex        =   161
      Top             =   1200
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.Label A14 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "F"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   3360
      TabIndex        =   160
      Top             =   1200
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Label A15 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "G"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   3600
      TabIndex        =   159
      Top             =   1200
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.Label A16 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "H"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   3840
      TabIndex        =   158
      Top             =   1200
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.Label A17 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "K"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   4080
      TabIndex        =   157
      Top             =   1200
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label A18 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Z"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   4320
      TabIndex        =   156
      Top             =   1200
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label A19 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "L"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   4560
      TabIndex        =   155
      Top             =   1200
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Label A20 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "J"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   4800
      TabIndex        =   154
      Top             =   1200
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.Label A21 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   5040
      TabIndex        =   153
      Top             =   1200
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label A22 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "C"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   5280
      TabIndex        =   152
      Top             =   1200
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label A23 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "V"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   5520
      TabIndex        =   151
      Top             =   1200
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label A24 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "B"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   5760
      TabIndex        =   150
      Top             =   1200
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label A25 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "N"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   6000
      TabIndex        =   149
      Top             =   1200
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.Label A26 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "M"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   6240
      TabIndex        =   148
      Top             =   1200
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label B1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Q"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   240
      TabIndex        =   147
      Top             =   1800
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.Label B2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "W"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   480
      TabIndex        =   146
      Top             =   1800
      Visible         =   0   'False
      Width           =   165
   End
   Begin VB.Label B3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "E"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   720
      TabIndex        =   145
      Top             =   1800
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label B4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "R"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   960
      TabIndex        =   144
      Top             =   1800
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.Label B5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "T"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   1200
      TabIndex        =   143
      Top             =   1800
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label B6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Y"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   1440
      TabIndex        =   142
      Top             =   1800
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label B7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "U"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   1680
      TabIndex        =   141
      Top             =   1800
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
      Left            =   1950
      TabIndex        =   140
      Top             =   1800
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.Label B9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "O"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   2160
      TabIndex        =   139
      Top             =   1800
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.Label B10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "P"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   2400
      TabIndex        =   138
      Top             =   1800
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label B11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "S"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   2640
      TabIndex        =   137
      Top             =   1800
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label B12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "A"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   2880
      TabIndex        =   136
      Top             =   1800
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label B13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "D"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   3120
      TabIndex        =   135
      Top             =   1800
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.Label B14 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "F"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   3360
      TabIndex        =   134
      Top             =   1800
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Label B15 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "G"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   3600
      TabIndex        =   133
      Top             =   1800
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.Label B16 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "H"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   3840
      TabIndex        =   132
      Top             =   1800
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.Label B17 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "J"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   4080
      TabIndex        =   131
      Top             =   1800
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.Label B18 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "K"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   4320
      TabIndex        =   130
      Top             =   1800
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label B19 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "L"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   4560
      TabIndex        =   129
      Top             =   1800
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Label B20 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Z"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   4800
      TabIndex        =   128
      Top             =   1800
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label B21 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   5040
      TabIndex        =   127
      Top             =   1800
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label B22 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "C"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   5280
      TabIndex        =   126
      Top             =   1800
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label B23 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "V"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   5520
      TabIndex        =   125
      Top             =   1800
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label B24 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "B"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   5760
      TabIndex        =   124
      Top             =   1800
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label B25 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "N"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   6000
      TabIndex        =   123
      Top             =   1800
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.Label B26 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "M"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   6240
      TabIndex        =   122
      Top             =   1800
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label C1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "M"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   240
      TabIndex        =   121
      Top             =   2400
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label C2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "N"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   480
      TabIndex        =   120
      Top             =   2400
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.Label C3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "B"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   720
      TabIndex        =   119
      Top             =   2400
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label C4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "V"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   960
      TabIndex        =   118
      Top             =   2400
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label C5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "C"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   1200
      TabIndex        =   117
      Top             =   2400
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label C6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   1440
      TabIndex        =   116
      Top             =   2400
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label C7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Z"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   1680
      TabIndex        =   115
      Top             =   2400
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label C8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "L"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   1920
      TabIndex        =   114
      Top             =   2400
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Label C9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "K"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   2160
      TabIndex        =   113
      Top             =   2400
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label C10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "J"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   2400
      TabIndex        =   112
      Top             =   2400
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.Label C11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "H"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   2640
      TabIndex        =   111
      Top             =   2400
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.Label C12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "G"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   2880
      TabIndex        =   110
      Top             =   2400
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.Label C13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "F"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   3120
      TabIndex        =   109
      Top             =   2400
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Label C14 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "D"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   3360
      TabIndex        =   108
      Top             =   2400
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.Label C15 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "S"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   3600
      TabIndex        =   107
      Top             =   2400
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label C16 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "A"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   3840
      TabIndex        =   106
      Top             =   2400
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label C17 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "P"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   4080
      TabIndex        =   105
      Top             =   2400
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label C18 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "O"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   4320
      TabIndex        =   104
      Top             =   2400
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
      Left            =   4605
      TabIndex        =   103
      Top             =   2400
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.Label C20 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "U"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   4800
      TabIndex        =   102
      Top             =   2400
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.Label C21 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Y"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   5040
      TabIndex        =   101
      Top             =   2400
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label C22 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "T"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   5280
      TabIndex        =   100
      Top             =   2400
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label C23 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "R"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   5520
      TabIndex        =   99
      Top             =   2400
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.Label C24 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "E"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   5760
      TabIndex        =   98
      Top             =   2400
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label C25 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "W"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   6000
      TabIndex        =   97
      Top             =   2400
      Visible         =   0   'False
      Width           =   165
   End
   Begin VB.Label C26 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Q"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   6240
      TabIndex        =   96
      Top             =   2400
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Q"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   240
      TabIndex        =   95
      Top             =   720
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "W"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   480
      TabIndex        =   94
      Top             =   720
      Visible         =   0   'False
      Width           =   165
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "E"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   720
      TabIndex        =   93
      Top             =   720
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "R"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   960
      TabIndex        =   92
      Top             =   720
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "T"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   1200
      TabIndex        =   91
      Top             =   720
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Y"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   1440
      TabIndex        =   90
      Top             =   720
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "U"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   1680
      TabIndex        =   89
      Top             =   720
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
      Left            =   1965
      TabIndex        =   88
      Top             =   720
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "O"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   2160
      TabIndex        =   87
      Top             =   720
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "P"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   2400
      TabIndex        =   86
      Top             =   720
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "A"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   2640
      TabIndex        =   85
      Top             =   720
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "S"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   2880
      TabIndex        =   84
      Top             =   720
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "D"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   3120
      TabIndex        =   83
      Top             =   720
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "M"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   3360
      TabIndex        =   82
      Top             =   720
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "N"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   3600
      TabIndex        =   81
      Top             =   720
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "B"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   3840
      TabIndex        =   80
      Top             =   720
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "V"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   4080
      TabIndex        =   79
      Top             =   720
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "C"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   4320
      TabIndex        =   78
      Top             =   720
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   4560
      TabIndex        =   77
      Top             =   720
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Z"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   4800
      TabIndex        =   76
      Top             =   720
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label Label21 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "L"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   5040
      TabIndex        =   75
      Top             =   720
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Label Label22 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "K"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   5280
      TabIndex        =   74
      Top             =   720
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label Label23 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "J"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   5520
      TabIndex        =   73
      Top             =   720
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.Label Label24 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "H"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   5760
      TabIndex        =   72
      Top             =   720
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.Label Label25 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "G"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   6000
      TabIndex        =   71
      Top             =   720
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.Label Label26 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "F"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   6240
      TabIndex        =   70
      Top             =   720
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Label Label27 
      BackColor       =   &H00FF0000&
      Caption         =   " "
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   6600
      TabIndex        =   69
      Top             =   720
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Label Label28 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   6960
      TabIndex        =   68
      Top             =   720
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Label A28 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   6960
      TabIndex        =   67
      Top             =   1200
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Label B28 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   6960
      TabIndex        =   66
      Top             =   1800
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Label C28 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   6960
      TabIndex        =   65
      Top             =   2400
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Label Label29 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "2"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   7200
      TabIndex        =   64
      Top             =   720
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Label B29 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "2"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   7200
      TabIndex        =   63
      Top             =   1800
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Label A29 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "2"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   7200
      TabIndex        =   62
      Top             =   1200
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Label C29 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "2"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   7200
      TabIndex        =   61
      Top             =   2400
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Label Label30 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "3"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   7440
      TabIndex        =   60
      Top             =   720
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Label Label31 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "4"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   7680
      TabIndex        =   59
      Top             =   720
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Label Label32 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "5"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   7920
      TabIndex        =   58
      Top             =   720
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Label Label33 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "6"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   8160
      TabIndex        =   57
      Top             =   720
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Label Label34 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "7"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   8400
      TabIndex        =   56
      Top             =   720
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Label Label35 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "8"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   8640
      TabIndex        =   55
      Top             =   720
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Label Label36 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "9"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   8880
      TabIndex        =   54
      Top             =   720
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Label Label37 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   9120
      TabIndex        =   53
      Top             =   720
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Label Label38 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "!"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   9360
      TabIndex        =   52
      Top             =   720
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Label Label39 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "."
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   9600
      TabIndex        =   51
      Top             =   720
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Label Label40 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "?"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   9840
      TabIndex        =   50
      Top             =   720
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Label A30 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "3"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   7440
      TabIndex        =   49
      Top             =   1200
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Label A31 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "4"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   7680
      TabIndex        =   48
      Top             =   1200
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Label A32 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "5"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   7920
      TabIndex        =   47
      Top             =   1200
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Label A33 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "6"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   8160
      TabIndex        =   46
      Top             =   1200
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Label A34 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "7"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   8400
      TabIndex        =   45
      Top             =   1200
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Label A35 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "8"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   8640
      TabIndex        =   44
      Top             =   1200
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Label A36 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "9"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   8880
      TabIndex        =   43
      Top             =   1200
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Label A37 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   9120
      TabIndex        =   42
      Top             =   1200
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Label A38 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "!"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   9360
      TabIndex        =   41
      Top             =   1200
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Label A39 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "."
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   9600
      TabIndex        =   40
      Top             =   1200
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Label A40 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "?"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   9840
      TabIndex        =   39
      Top             =   1200
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Label B30 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "3"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   7440
      TabIndex        =   38
      Top             =   1800
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Label B31 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "4"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   7680
      TabIndex        =   37
      Top             =   1800
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Label B32 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "5"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   7920
      TabIndex        =   36
      Top             =   1800
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Label B33 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "6"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   8160
      TabIndex        =   35
      Top             =   1800
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Label B34 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "7"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   8400
      TabIndex        =   34
      Top             =   1800
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Label B35 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "8"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   8640
      TabIndex        =   33
      Top             =   1800
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Label B36 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "9"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   8880
      TabIndex        =   32
      Top             =   1800
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Label B37 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   9120
      TabIndex        =   31
      Top             =   1800
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Label B38 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "!"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   9360
      TabIndex        =   30
      Top             =   1800
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Label B39 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "."
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   9600
      TabIndex        =   29
      Top             =   1800
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Label B40 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "?"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   9840
      TabIndex        =   28
      Top             =   1800
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Label C30 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "3"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   7440
      TabIndex        =   27
      Top             =   2400
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Label C31 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "4"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   7680
      TabIndex        =   26
      Top             =   2400
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Label C32 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "5"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   7920
      TabIndex        =   25
      Top             =   2400
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Label C33 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "6"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   8160
      TabIndex        =   24
      Top             =   2400
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Label C34 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "7"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   8400
      TabIndex        =   23
      Top             =   2400
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Label C35 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "8"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   8640
      TabIndex        =   22
      Top             =   2400
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Label C36 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "9"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   8880
      TabIndex        =   21
      Top             =   2400
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Label C37 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   9120
      TabIndex        =   20
      Top             =   2400
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Label C38 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "!"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   9360
      TabIndex        =   19
      Top             =   2400
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Label C39 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "."
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   9600
      TabIndex        =   18
      Top             =   2400
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Label C40 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "?"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   9840
      TabIndex        =   17
      Top             =   2400
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000005&
      Visible         =   0   'False
      X1              =   240
      X2              =   8640
      Y1              =   1080
      Y2              =   1080
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strngforlist1 As Integer
Dim showagainmsbox As Integer ' rwtoumen ton xristin an theli na tou xanafkalei msgbox
Dim t, p, lng, ll, tempa, tempb, tempc, tempaa, tempcc, tempbb, hh



Private Sub Command1_Click()
txt1.Visible = False

Dim resp As Integer ' apantisi an the xanafkali tin erwtisin gia not valid character
Dim char1 As String ' vlepoume an o xaraktiras einai egiros gia tin grammatosiran mas
DEL.Text = ""
Do
char1 = Left(txt1.Text, 1)
char1 = UCase(char1)


If char1 <> "A" And char1 <> "B" And char1 <> "C" And char1 <> "D" And char1 <> "E" And char1 <> "F" And char1 <> "G" And char1 <> "H" And char1 <> "I" And char1 <> "J" And char1 <> "K" And char1 <> "L" And char1 <> "M" And char1 <> "N" And char1 <> "O" And char1 <> "P" And char1 <> "Q" And char1 <> "R" And char1 <> "S" And char1 <> "T" And char1 <> "V" And char1 <> "U" And char1 <> "W" And char1 <> "X" And char1 <> "Y" And char1 <> "Z" And char1 <> " " And char1 <> "0" And char1 <> "1" And char1 <> "2" And char1 <> "3" And char1 <> "4" And char1 <> "5" And char1 <> "6" And char1 <> "7" And char1 <> "8" And char1 <> "9" And char1 <> "10" And char1 <> "." And char1 <> "?" And char1 <> "@" And char1 <> "%" And char1 <> "_" And char1 <> "!" And char1 <> "$" And char1 <> Chr$(13) And char1 <> ":" Then




If showagainmsbox = 1 Then

'If char1 = labelspacewrong.Text Then
'char1 = "|AD.ENTER|" 'andiniable?? prepei na to koitaxeis simainei undiniable
'End If

resp = MsgBox(("The following character( " & char1 & " )will replace with null for full functional encoding | | ASK NEXT TIME?"), vbInformation + vbYesNo + vbDefaultButton2, "INFORMATION")

If resp <> 6 Then
showagainmsbox = 0
End If
End If

List1.List(strngforlist1) = char1
strngforlist1 = strngforlist1 + 1
If Len(txt1.Text) > 0 Then
txt1.Text = Right(txt1.Text, Len(txt1.Text) - 1)
End If

Else
DEL.Text = DEL.Text & Left(txt1.Text, 1)
txt1.Text = Right(txt1.Text, Len(txt1.Text) - 1)
End If

Loop Until Len(txt1.Text) <= 0
txt1.Text = DEL.Text
txt1.Text = Replace(txt1.Text, Chr$(13), "-")
txt2.Text = ""
t = Val(Text3.Text)
Timer1.Enabled = True
A27.Visible = False
b27.Visible = False
c27.Visible = False
Text1.Locked = True
Text2.Locked = True
Text4.Locked = True
txt2.Visible = False
Command3.Visible = False
Command4.Enabled = False
llt.List(hh) = (A27 & " " & b27 & " " & c27)
hh = hh + 1
End Sub

Private Sub Command10_Click()
tempcc = C40
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
C1 = tempcc
End Sub

Private Sub Command11_Click()
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

Private Sub Command13_Click()

End Sub

Private Sub Command14_Click()

End Sub

Private Sub Command12_Click()
Load Form2
End Sub

Private Sub Command2_Click()
txt1.Text = ""

End Sub

Private Sub Command3_Click()
Clipboard.Clear
Clipboard.SetText txt2.Text

End Sub



Private Sub Command4_Click()
If txt2.Visible = False Then
txt2.Visible = True
Command3.Visible = True
ElseIf txt2.Visible = True Then
txt2.Visible = False
Command3.Visible = False
End If
End Sub

Private Sub Command5_Click()
txt2.Text = ""

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
tempaa = A40
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
A1 = tempaa
End Sub

Private Sub Command9_Click()
tempbb = B40
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
B1 = tempbb
End Sub

Private Sub Form_Load()
strngforlist1 = 1
hh = 1
showagainmsbox = 1
End Sub

Private Sub llt_Click()
Dim ires As Integer

If llt.ListIndex = 0 Then
ires = MsgBox("Are you sure are you want to set this key?", vbQuestion + vbYesNo + vbDefaultButton1, "Question")
If ires = 6 Then

Dim setkeys As String
Dim a As String, b As String, c As String

setkeys = llt.List(llt.ListIndex)

If setkeys <> "" Or setkeys <> "KEYS:" Then
setkeys = Replace(setkeys, " ", "")

a = Left(setkeys, 1)
Text1 = a
A27.ForeColor = &HFF&
Timer2.Enabled = True
''''''''''''''''''''''''
b = Left(setkeys, 2)
b = Right(b, 1)
Text2 = b
b27.ForeColor = &HFF&
Timer3.Enabled = True

c = Right(setkeys, 1)
Text4 = c
c27.ForeColor = &HFF&
Timer4.Enabled = True

End If

End If

Else



setkeys = llt.List(llt.ListIndex)

If setkeys <> "" Or setkeys <> "KEYS:" Then
setkeys = Replace(setkeys, " ", "")

a = Left(setkeys, 1)
Text1 = a
A27.ForeColor = &HFF&
Timer2.Enabled = True
''''''''''''''''''''''''
b = Left(setkeys, 2)
b = Right(b, 1)
Text2 = b
b27.ForeColor = &HFF&
Timer3.Enabled = True

c = Right(setkeys, 1)
Text4 = c
c27.ForeColor = &HFF&
Timer4.Enabled = True

End If

End If


End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then

If Len(Text1.Text) = 0 Then
MsgBox "SET CHARACTER", vbCritical, "ERROR"
Text1.SetFocus
Else
A27.ForeColor = &HFF&
Timer2.Enabled = True
End If
Text2.SetFocus

End If
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
If Len(Text2.Text) = 0 Then
MsgBox "SET CHARACTER", vbCritical, "ERROR"
Text2.SetFocus
Else
b27.ForeColor = &HFF&
Timer3.Enabled = True
End If
Text4.SetFocus
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
Text2.SetFocus
Else
c27.ForeColor = &HFF&
Timer4.Enabled = True
End If
txt1.SetFocus
End If
End Sub

Private Sub Timer1_Timer()
Dim metr As Integer

If Len(txt1.Text) <> 0 Then
txt1.Text = UCase(txt1.Text)
p = Left(txt1.Text, 1)
lng = Len(txt1.Text)
lng = lng - 1
txt1.Text = Right(txt1.Text, lng)

'If p = labelspacewrong.Text Then
'txt1.Text = Right(txt1.Text, Len(txt1.Text) - 1)
'GoTo CONT
'End If



If t = 1 Then
If A1 = p Then
txt2.Text = txt2.Text + Label1
ll = 1
ElseIf A2 = p Then
txt2.Text = txt2.Text + Label2
ll = 1
ElseIf A3 = p Then
txt2.Text = txt2.Text + Label3
ll = 1
ElseIf A4 = p Then
txt2.Text = txt2.Text + Label4
ll = 1
ElseIf A5 = p Then
txt2.Text = txt2.Text + Label5
ll = 1
ElseIf A6 = p Then
txt2.Text = txt2.Text + Label6
ll = 1
ElseIf A7 = p Then
txt2.Text = txt2.Text + Label7
ll = 1
ElseIf A8 = p Then
txt2.Text = txt2.Text + Label8
ll = 1
ElseIf A9 = p Then
txt2.Text = txt2.Text + Label9
ll = 1
ElseIf A10 = p Then
txt2.Text = txt2.Text + Label10
ll = 1
ElseIf A11 = p Then
txt2.Text = txt2.Text + Label11
ll = 1
ElseIf A12 = p Then
txt2.Text = txt2.Text + Label12
ll = 1
ElseIf A13 = p Then
txt2.Text = txt2.Text + Label13
ll = 1
ElseIf A14 = p Then
txt2.Text = txt2.Text + Label14
ll = 1
ElseIf A15 = p Then
txt2.Text = txt2.Text + Label15
ll = 1
ElseIf A16 = p Then
txt2.Text = txt2.Text + Label16
ll = 1
ElseIf A17 = p Then
txt2.Text = txt2.Text + Label17
ll = 1
ElseIf A18 = p Then
txt2.Text = txt2.Text + Label18
ll = 1
ElseIf A19 = p Then
txt2.Text = txt2.Text + Label19
ll = 1
ElseIf A20 = p Then
txt2.Text = txt2.Text + Label20
ll = 1
ElseIf A21 = p Then
txt2.Text = txt2.Text + Label21
ll = 1
ElseIf A22 = p Then
txt2.Text = txt2.Text + Label22
ll = 1
ElseIf A23 = p Then
txt2.Text = txt2.Text + Label23
ll = 1
ElseIf A24 = p Then
txt2.Text = txt2.Text + Label24
ll = 1
ElseIf A25 = p Then
txt2.Text = txt2.Text + Label25
ll = 1
ElseIf A26 = p Then
txt2.Text = txt2.Text + Label26
ll = 1
ElseIf A27 = p Then
txt2.Text = txt2.Text + Label27
ll = 1
ElseIf A28 = p Then
txt2.Text = txt2.Text + Label28
ll = 1
ElseIf A29 = p Then
txt2.Text = txt2.Text + Label29
ll = 1
ElseIf A30 = p Then
txt2.Text = txt2.Text + Label30
ll = 1
ElseIf A31 = p Then
txt2.Text = txt2.Text + Label31
ll = 1
ElseIf A32 = p Then
txt2.Text = txt2.Text + Label32
ll = 1
ElseIf A33 = p Then
txt2.Text = txt2.Text + Label33
ll = 1
ElseIf A34 = p Then
txt2.Text = txt2.Text + Label34
ll = 1
ElseIf A35 = p Then
txt2.Text = txt2.Text + Label35
ll = 1
ElseIf A36 = p Then
txt2.Text = txt2.Text + Label36
ll = 1
ElseIf A37 = p Then
txt2.Text = txt2.Text + Label37
ll = 1
ElseIf A38 = p Then
txt2.Text = txt2.Text + Label38
ll = 1
ElseIf A39 = p Then
txt2.Text = txt2.Text + Label39
ll = 1
ElseIf A40 = p Then
txt2.Text = txt2.Text + Label40
ll = 1
ElseIf A41 = p Then
txt2.Text = txt2.Text + Label41
ll = 1
ElseIf A42 = p Then
txt2.Text = txt2.Text + Label42
ll = 1
ElseIf A43 = p Then
txt2.Text = txt2.Text + Label43
ll = 1
ElseIf A44 = p Then
txt2.Text = txt2.Text + Label44
ll = 1
ElseIf A45 = p Then
txt2.Text = txt2.Text + Label45
ll = 1
ElseIf A46 = p Then
txt2.Text = txt2.Text + Label46
ll = 1
End If


ElseIf t = 2 Then


If B1 = p Then
txt2.Text = txt2.Text + Label1
ll = 2
ElseIf B2 = p Then
txt2.Text = txt2.Text + Label2
ll = 2
ElseIf B3 = p Then
txt2.Text = txt2.Text + Label3
ll = 2
ElseIf B4 = p Then
txt2.Text = txt2.Text + Label4
ll = 2
ElseIf B5 = p Then
txt2.Text = txt2.Text + Label5
ll = 2
ElseIf B6 = p Then
txt2.Text = txt2.Text + Label6
ll = 2
ElseIf B7 = p Then
txt2.Text = txt2.Text + Label7
ll = 2
ElseIf B8 = p Then
txt2.Text = txt2.Text + Label8
ll = 2
ElseIf B9 = p Then
txt2.Text = txt2.Text + Label9
ll = 2
ElseIf B10 = p Then
txt2.Text = txt2.Text + Label10
ll = 2
ElseIf B11 = p Then
txt2.Text = txt2.Text + Label11
ll = 2
ElseIf B12 = p Then
txt2.Text = txt2.Text + Label12
ll = 2
ElseIf B13 = p Then
txt2.Text = txt2.Text + Label13
ll = 2
ElseIf B14 = p Then
txt2.Text = txt2.Text + Label14
ll = 2
ElseIf B15 = p Then
txt2.Text = txt2.Text + Label15
ll = 2
ElseIf B16 = p Then
txt2.Text = txt2.Text + Label16
ll = 2
ElseIf B17 = p Then
txt2.Text = txt2.Text + Label17
ll = 2
ElseIf B18 = p Then
txt2.Text = txt2.Text + Label18
ll = 2
ElseIf B19 = p Then
txt2.Text = txt2.Text + Label19
ll = 2
ElseIf B20 = p Then
txt2.Text = txt2.Text + Label20
ll = 2
ElseIf B21 = p Then
txt2.Text = txt2.Text + Label21
ll = 2
ElseIf B22 = p Then
txt2.Text = txt2.Text + Label22
ll = 2
ElseIf B23 = p Then
txt2.Text = txt2.Text + Label23
ll = 2
ElseIf B24 = p Then
txt2.Text = txt2.Text + Label24
ll = 2
ElseIf B25 = p Then
txt2.Text = txt2.Text + Label25
ll = 2
ElseIf B26 = p Then
txt2.Text = txt2.Text + Label26
ll = 2
ElseIf b27 = p Then
txt2.Text = txt2.Text + Label27
ll = 2
ElseIf B28 = p Then
txt2.Text = txt2.Text + Label28
ll = 2
ElseIf B29 = p Then
txt2.Text = txt2.Text + Label29
ll = 2
ElseIf B30 = p Then
txt2.Text = txt2.Text + Label30
ll = 2
ElseIf B31 = p Then
txt2.Text = txt2.Text + Label31
ll = 2
ElseIf B32 = p Then
txt2.Text = txt2.Text + Label32
ll = 2
ElseIf B33 = p Then
txt2.Text = txt2.Text + Label33
ll = 2
ElseIf B34 = p Then
txt2.Text = txt2.Text + Label34
ll = 2
ElseIf B35 = p Then
txt2.Text = txt2.Text + Label35
ll = 2
ElseIf B36 = p Then
txt2.Text = txt2.Text + Label36
ll = 2
ElseIf B37 = p Then
txt2.Text = txt2.Text + Label37
ll = 2
ElseIf B38 = p Then
txt2.Text = txt2.Text + Label38
ll = 2
ElseIf B39 = p Then
txt2.Text = txt2.Text + Label39
ll = 2
ElseIf B40 = p Then
txt2.Text = txt2.Text + Label40
ll = 2
ElseIf B41 = p Then
txt2.Text = txt2.Text + Label41
ll = 2
ElseIf B42 = p Then
txt2.Text = txt2.Text + Label42
ll = 2
ElseIf B43 = p Then
txt2.Text = txt2.Text + Label43
ll = 2
ElseIf B44 = p Then
txt2.Text = txt2.Text + Label44
ll = 2
ElseIf B45 = p Then
txt2.Text = txt2.Text + Label45
ll = 2
ElseIf B46 = p Then
txt2.Text = txt2.Text + Label46
ll = 2
End If

ElseIf t = 3 Then


If C1 = p Then
txt2.Text = txt2.Text + Label1
ll = 3
ElseIf C2 = p Then
txt2.Text = txt2.Text + Label2
ll = 3
ElseIf C3 = p Then
txt2.Text = txt2.Text + Label3
ll = 3
ElseIf C4 = p Then
txt2.Text = txt2.Text + Label4
ll = 3
ElseIf C5 = p Then
txt2.Text = txt2.Text + Label5
ll = 3
ElseIf C6 = p Then
txt2.Text = txt2.Text + Label6
ll = 3
ElseIf C7 = p Then
txt2.Text = txt2.Text + Label7
ll = 3
ElseIf C8 = p Then
txt2.Text = txt2.Text + Label8
ll = 3
ElseIf C9 = p Then
txt2.Text = txt2.Text + Label9
ll = 3
ElseIf C10 = p Then
txt2.Text = txt2.Text + Label10
ll = 3
ElseIf C11 = p Then
txt2.Text = txt2.Text + Label11
ll = 3
ElseIf C12 = p Then
txt2.Text = txt2.Text + Label12
ll = 3
ElseIf C13 = p Then
txt2.Text = txt2.Text + Label13
ll = 3
ElseIf C14 = p Then
txt2.Text = txt2.Text + Label14
ll = 3
ElseIf C15 = p Then
txt2.Text = txt2.Text + Label15
ll = 3
ElseIf C16 = p Then
txt2.Text = txt2.Text + Label16
ll = 3
ElseIf C17 = p Then
txt2.Text = txt2.Text + Label17
ll = 3
ElseIf C18 = p Then
txt2.Text = txt2.Text + Label18
ll = 3
ElseIf C19 = p Then
txt2.Text = txt2.Text + Label19
ll = 3
ElseIf C20 = p Then
txt2.Text = txt2.Text + Label20
ll = 3
ElseIf C21 = p Then
txt2.Text = txt2.Text + Label21
ll = 3
ElseIf C22 = p Then
txt2.Text = txt2.Text + Label22
ll = 3
ElseIf C23 = p Then
txt2.Text = txt2.Text + Label23
ll = 3
ElseIf C24 = p Then
txt2.Text = txt2.Text + Label24
ll = 3
ElseIf C25 = p Then
txt2.Text = txt2.Text + Label25
ll = 3
ElseIf C26 = p Then
txt2.Text = txt2.Text + Label26
ll = 3
ElseIf c27 = p Then
txt2.Text = txt2.Text + Label27
ll = 3
ElseIf C28 = p Then
txt2.Text = txt2.Text + Label28
ll = 3
ElseIf C29 = p Then
txt2.Text = txt2.Text + Label29
ll = 3
ElseIf C30 = p Then
txt2.Text = txt2.Text + Label30
ll = 3
ElseIf C31 = p Then
txt2.Text = txt2.Text + Label31
ll = 3
ElseIf C32 = p Then
txt2.Text = txt2.Text + Label32
ll = 3
ElseIf C33 = p Then
txt2.Text = txt2.Text + Label33
ll = 3
ElseIf C34 = p Then
txt2.Text = txt2.Text + Label34
ll = 3
ElseIf C35 = p Then
txt2.Text = txt2.Text + Label35
ll = 3
ElseIf C36 = p Then
txt2.Text = txt2.Text + Label36
ll = 3
ElseIf C37 = p Then
txt2.Text = txt2.Text + Label37
ll = 3
ElseIf C38 = p Then
txt2.Text = txt2.Text + Label38
ll = 3
ElseIf C39 = p Then
txt2.Text = txt2.Text + Label39
ll = 3
ElseIf C40 = p Then
txt2.Text = txt2.Text + Label40
ll = 3
ElseIf C41 = p Then
txt2.Text = txt2.Text + Label41
ll = 3
ElseIf C42 = p Then
txt2.Text = txt2.Text + Label42
ll = 3
ElseIf C43 = p Then
txt2.Text = txt2.Text + Label43
ll = 3
ElseIf C44 = p Then
txt2.Text = txt2.Text + Label44
ll = 3
ElseIf C45 = p Then
txt2.Text = txt2.Text + Label45
ll = 3
ElseIf C46 = p Then
txt2.Text = txt2.Text + Label46
ll = 3
End If
End If

If ll = 1 Then
ll = 0
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
Next metr
ElseIf ll = 2 Then
ll = 0
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
Next metr
ElseIf ll = 3 Then
ll = 0
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
Next metr
End If

Else
Timer1.Enabled = False
Clipboard.Clear
Clipboard.SetText (txt2.Text)
Text1.Locked = False
Text2.Locked = False
Text4.Locked = False
A27.Visible = True
b27.Visible = True
c27.Visible = True
Command4.Enabled = True
txt1.Visible = True

Clipboard.Clear
Clipboard.SetText (txt2.Text)

End If

CONT:



End Sub

Private Sub Timer2_Timer()
Dim metr As Integer

If A27.Caption = UCase(Text1.Text) Then
Timer2.Enabled = False
A27.ForeColor = &H8000000E
Text1.Text = ""
Else

For metr = 1 To Textc1

tempaa = A46
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
A1 = tempaa

Next metr
End If


End Sub

Private Sub Timer3_Timer()
Dim metr As Integer

If b27.Caption = UCase(Text2.Text) Then
Timer3.Enabled = False
b27.ForeColor = &H8000000E
Text2.Text = ""
Else
For metr = 1 To Textc2

tempbb = B46
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
B1 = tempbb
Next metr

End If

End Sub

Private Sub Timer4_Timer()
Dim metr As Integer

If c27.Caption = UCase(Text4.Text) Then
Timer4.Enabled = False
c27.ForeColor = &H8000000E
Text4.Text = ""
Else
For metr = 1 To Textc3

tempcc = C46
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
C1 = tempcc
Next metr

End If

End Sub

