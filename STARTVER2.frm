VERSION 5.00
Begin VB.Form START 
   BackColor       =   &H00000000&
   Caption         =   "HORSE CHAMPIONSHIP 2.0"
   ClientHeight    =   5565
   ClientLeft      =   2505
   ClientTop       =   1890
   ClientWidth     =   7530
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5565
   ScaleWidth      =   7530
   Begin VB.Timer Timer5 
      Left            =   6000
      Top             =   3240
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00808000&
      Caption         =   "READY TO PLAY"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   5160
      Visible         =   0   'False
      Width           =   7575
   End
   Begin VB.Timer Timer3 
      Left            =   7080
      Top             =   4440
   End
   Begin VB.Timer Timer2 
      Left            =   5400
      Top             =   6000
   End
   Begin VB.Timer Timer4 
      Interval        =   500
      Left            =   6120
      Top             =   360
   End
   Begin VB.Timer Timer1 
      Left            =   3600
      Top             =   2040
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   ".................................................................................."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   735
      Left            =   720
      TabIndex        =   10
      Top             =   600
      Width           =   15
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   99.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   15
      Left            =   4920
      TabIndex        =   9
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00008000&
      BorderWidth     =   4
      Height          =   1455
      Left            =   7800
      Top             =   3480
      Width           =   1695
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H000000FF&
      BorderWidth     =   3
      Height          =   1455
      Left            =   120
      Top             =   -1500
      Width           =   1695
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "e"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   735
      Index           =   4
      Left            =   4800
      TabIndex        =   7
      Top             =   5760
      Width           =   495
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "s"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   735
      Index           =   3
      Left            =   4200
      TabIndex        =   6
      Top             =   5760
      Width           =   495
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "r"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   735
      Index           =   2
      Left            =   3600
      TabIndex        =   5
      Top             =   5760
      Width           =   495
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "o"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   735
      Index           =   0
      Left            =   3000
      TabIndex        =   4
      Top             =   5760
      Width           =   495
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "H"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   735
      Index           =   1
      Left            =   2400
      TabIndex        =   3
      Top             =   5760
      Width           =   495
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "MCPOWER COMPUTERS PRESENTS"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   1080
      TabIndex        =   2
      Top             =   360
      Width           =   5655
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "WITH JOHN MCCRIRICK"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   855
      Left            =   7320
      TabIndex        =   1
      Top             =   3840
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   1455
      Left            =   120
      Picture         =   "STARTVER2.frx":0000
      Stretch         =   -1  'True
      Top             =   3480
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "CHAMPIONSHIP"
      BeginProperty Font 
         Name            =   "Lucida Handwriting"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   135
      Left            =   480
      TabIndex        =   0
      Top             =   2160
      Width           =   135
   End
End
Attribute VB_Name = "START"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim LEFTVAL, TOPVAL, LEFT1, leftval1, COUNT1, upwards
Dim X As Integer
Dim y, x2
Dim leftimage
Dim proced
Private Sub Command1_Click()
Form1.Visible = True
START.Visible = False
Load Form1
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Form1.Visible = True
START.Visible = False
Load Form1
End If
End Sub

Private Sub Form_Load()
LEFTVAL = 100
leftval1 = 150
TOPVAL = 100
LEFT1 = 50
COUNT1 = 0
upwards = 400
y = 90
x2 = 120
leftimage = 300
proced = 1
End Sub
Private Sub Timer1_Timer()
Label6.Width = Label6.Width + leftval1

If Label6.Width > 6135 Then
leftval1 = 0
proced = 2
End If

If proced = 2 Then
Label1.Height = Label1.Height + TOPVAL
Label1.Width = Label1.Width + LEFTVAL
If Timer1.Interval = 1 Then
Timer4.Interval = 0
End If


If Label1.Height > 975 Then
TOPVAL = 0
End If
If Label1.Width > 6855 Then
LEFTVAL = 0
Timer1.Interval = 0
Timer2.Interval = 1
End If
End If
End Sub
Private Sub Timer2_Timer()
For X = 0 To 4
Label2(X).Top = Label2(X).Top - upwards
Next X

For X = 0 To 4
If Label2(X).Top < 1560 Then
upwards = 0
Timer4.Interval = 0
Timer3.Interval = 1
End If
Next X

End Sub

Private Sub Timer3_Timer()
Dim proced
proced = 1
If proced = 1 Then
Timer2.Interval = 0

Shape1.Top = Shape1.Top + y
Shape2.Left = Shape2.Left - x2

If Shape1.Top > 3480 Then
y = 0
End If

If Shape2.Left < 120 Then
x2 = 0
Label3.Visible = True
proced = 2
End If
End If

If proced = 2 Then
Label3.Left = Label3.Left - leftimage
If Label3.Left < 3600 Then
leftimage = 0
Image1.Visible = True
Timer3.Interval = 0
Timer5.Interval = 1
'Command1.Visible = True
End If
End If
End Sub

Private Sub Timer4_Timer()
COUNT1 = COUNT1 + 1
If COUNT1 = 8 Then
Timer1.Interval = 1
Timer4.Interval = 0
End If


If Label4.Visible = False Then
Label4.Visible = True
Exit Sub
End If
If Label4.Visible = True Then
Label4.Visible = False
Exit Sub
End If
End Sub

Private Sub Timer5_Timer()
Label5.Height = Label5.Height + 50

If Label5.Height > 2055 Then
Timer5.Interval = 0
Command1.Visible = True
End If
End Sub
