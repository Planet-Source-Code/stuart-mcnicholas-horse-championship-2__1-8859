VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H0000C000&
   Caption         =   "HORSE CHAMPIONSHIP"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7605
   LinkTopic       =   "Form1"
   ScaleHeight     =   8595
   ScaleWidth      =   7605
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer3 
      Left            =   4800
      Top             =   0
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   9600
      TabIndex        =   44
      Top             =   8160
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Timer FALL 
      Left            =   1080
      Top             =   4080
   End
   Begin VB.Timer JOHN_TIM 
      Left            =   9360
      Top             =   240
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00E0E0E0&
      Height          =   375
      Index           =   7
      Left            =   2520
      TabIndex        =   43
      Top             =   8040
      Width           =   1575
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   6
      Left            =   4800
      TabIndex        =   42
      Top             =   7440
      Width           =   1575
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00008000&
      Height          =   375
      Index           =   5
      Left            =   2520
      TabIndex        =   41
      Top             =   7440
      Width           =   1575
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00C000C0&
      Height          =   375
      Index           =   4
      Left            =   240
      TabIndex        =   40
      Top             =   7440
      Width           =   1575
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00FF0000&
      Height          =   375
      Index           =   3
      Left            =   4800
      TabIndex        =   39
      Top             =   6840
      Width           =   1575
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H0000FFFF&
      Height          =   375
      Index           =   2
      Left            =   2520
      TabIndex        =   38
      Top             =   6840
      Width           =   1575
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H000000FF&
      Height          =   375
      Index           =   1
      Left            =   240
      TabIndex        =   37
      Top             =   6840
      Width           =   1575
   End
   Begin VB.Timer HORSE1 
      Left            =   1080
      Top             =   2760
   End
   Begin VB.Timer Timer2 
      Left            =   9360
      Top             =   6240
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00808080&
      Caption         =   "READY TO BET"
      Height          =   495
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   6240
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00008000&
      Caption         =   "START"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4320
      Width           =   1455
   End
   Begin VB.Timer Timer_start 
      Left            =   5880
      Top             =   6360
   End
   Begin VB.TextBox Text4 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   7200
      TabIndex        =   6
      Top             =   7440
      Width           =   1935
   End
   Begin VB.TextBox Text2 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   9600
      TabIndex        =   1
      Text            =   "2000"
      Top             =   7440
      Width           =   2055
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H0080FF80&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   0
      TabIndex        =   5
      Top             =   6240
      Width           =   9975
   End
   Begin VB.Line Line12 
      BorderColor     =   &H80000005&
      BorderWidth     =   2
      Index           =   0
      X1              =   9840
      X2              =   9840
      Y1              =   5880
      Y2              =   6240
   End
   Begin VB.Line Line12 
      BorderColor     =   &H80000005&
      BorderWidth     =   2
      Index           =   20
      X1              =   8880
      X2              =   8880
      Y1              =   5880
      Y2              =   6240
   End
   Begin VB.Line Line12 
      BorderColor     =   &H80000005&
      BorderWidth     =   2
      Index           =   19
      X1              =   8400
      X2              =   8400
      Y1              =   5880
      Y2              =   6240
   End
   Begin VB.Line Line12 
      BorderColor     =   &H80000005&
      BorderWidth     =   2
      Index           =   18
      X1              =   7920
      X2              =   7920
      Y1              =   5880
      Y2              =   6240
   End
   Begin VB.Line Line12 
      BorderColor     =   &H80000005&
      BorderWidth     =   2
      Index           =   17
      X1              =   7440
      X2              =   7440
      Y1              =   5880
      Y2              =   6240
   End
   Begin VB.Line Line12 
      BorderColor     =   &H80000005&
      BorderWidth     =   2
      Index           =   16
      X1              =   6960
      X2              =   6960
      Y1              =   5880
      Y2              =   6240
   End
   Begin VB.Line Line12 
      BorderColor     =   &H80000005&
      BorderWidth     =   2
      Index           =   15
      X1              =   6480
      X2              =   6480
      Y1              =   5880
      Y2              =   6240
   End
   Begin VB.Line Line12 
      BorderColor     =   &H80000005&
      BorderWidth     =   2
      Index           =   14
      X1              =   6000
      X2              =   6000
      Y1              =   5880
      Y2              =   6240
   End
   Begin VB.Line Line12 
      BorderColor     =   &H80000005&
      BorderWidth     =   2
      Index           =   13
      X1              =   5520
      X2              =   5520
      Y1              =   5880
      Y2              =   6240
   End
   Begin VB.Line Line12 
      BorderColor     =   &H80000005&
      BorderWidth     =   2
      Index           =   12
      X1              =   5040
      X2              =   5040
      Y1              =   5880
      Y2              =   6240
   End
   Begin VB.Line Line12 
      BorderColor     =   &H80000005&
      BorderWidth     =   2
      Index           =   11
      X1              =   4560
      X2              =   4560
      Y1              =   5880
      Y2              =   6240
   End
   Begin VB.Line Line12 
      BorderColor     =   &H80000005&
      BorderWidth     =   2
      Index           =   10
      X1              =   4080
      X2              =   4080
      Y1              =   5880
      Y2              =   6240
   End
   Begin VB.Line Line12 
      BorderColor     =   &H80000005&
      BorderWidth     =   2
      Index           =   9
      X1              =   3600
      X2              =   3600
      Y1              =   5880
      Y2              =   6240
   End
   Begin VB.Line Line12 
      BorderColor     =   &H80000005&
      BorderWidth     =   2
      Index           =   8
      X1              =   3120
      X2              =   3120
      Y1              =   5880
      Y2              =   6240
   End
   Begin VB.Line Line12 
      BorderColor     =   &H80000005&
      BorderWidth     =   2
      Index           =   7
      X1              =   2640
      X2              =   2640
      Y1              =   5880
      Y2              =   6240
   End
   Begin VB.Line Line12 
      BorderColor     =   &H80000005&
      BorderWidth     =   2
      Index           =   6
      X1              =   2160
      X2              =   2160
      Y1              =   5880
      Y2              =   6240
   End
   Begin VB.Line Line12 
      BorderColor     =   &H80000005&
      BorderWidth     =   2
      Index           =   5
      X1              =   1680
      X2              =   1680
      Y1              =   5880
      Y2              =   6240
   End
   Begin VB.Line Line12 
      BorderColor     =   &H80000005&
      BorderWidth     =   2
      Index           =   4
      X1              =   1200
      X2              =   1200
      Y1              =   5880
      Y2              =   6240
   End
   Begin VB.Line Line12 
      BorderColor     =   &H80000005&
      BorderWidth     =   2
      Index           =   3
      X1              =   720
      X2              =   720
      Y1              =   5880
      Y2              =   6240
   End
   Begin VB.Line Line12 
      BorderColor     =   &H80000005&
      BorderWidth     =   2
      Index           =   2
      X1              =   240
      X2              =   240
      Y1              =   5880
      Y2              =   6240
   End
   Begin VB.Line Line12 
      BorderColor     =   &H80000005&
      BorderWidth     =   2
      Index           =   1
      X1              =   9360
      X2              =   9360
      Y1              =   5880
      Y2              =   6240
   End
   Begin VB.Line Line11 
      BorderColor     =   &H80000005&
      BorderWidth     =   3
      X1              =   0
      X2              =   9840
      Y1              =   6000
      Y2              =   6000
   End
   Begin VB.Line Line10 
      BorderColor     =   &H80000005&
      BorderWidth     =   2
      X1              =   0
      X2              =   9840
      Y1              =   6120
      Y2              =   6120
   End
   Begin VB.Shape Shape8 
      BackStyle       =   1  'Opaque
      Height          =   255
      Index           =   1
      Left            =   8160
      Top             =   1920
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Shape Shape8 
      BackStyle       =   1  'Opaque
      Height          =   255
      Index           =   0
      Left            =   3600
      Top             =   1920
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image HORSE 
      Height          =   375
      Index           =   7
      Left            =   120
      Picture         =   "RACE.frx":0000
      Stretch         =   -1  'True
      Top             =   5640
      Width           =   615
   End
   Begin VB.Image HORSE 
      Height          =   375
      Index           =   6
      Left            =   120
      Picture         =   "RACE.frx":08CA
      Stretch         =   -1  'True
      Top             =   5160
      Width           =   615
   End
   Begin VB.Image HORSE 
      Height          =   375
      Index           =   5
      Left            =   120
      Picture         =   "RACE.frx":1194
      Stretch         =   -1  'True
      Top             =   4680
      Width           =   615
   End
   Begin VB.Image HORSE 
      Height          =   375
      Index           =   4
      Left            =   120
      Picture         =   "RACE.frx":1A5E
      Stretch         =   -1  'True
      Top             =   4200
      Width           =   615
   End
   Begin VB.Image HORSE 
      Height          =   375
      Index           =   3
      Left            =   120
      Picture         =   "RACE.frx":2328
      Stretch         =   -1  'True
      Top             =   3720
      Width           =   615
   End
   Begin VB.Image HORSE 
      Height          =   375
      Index           =   2
      Left            =   120
      Picture         =   "RACE.frx":2BF2
      Stretch         =   -1  'True
      Top             =   3240
      Width           =   615
   End
   Begin VB.Image HORSE 
      Height          =   375
      Index           =   1
      Left            =   120
      Picture         =   "RACE.frx":34BC
      Stretch         =   -1  'True
      Top             =   2760
      Width           =   615
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "HORSE CHAMPIONSHIP"
      BeginProperty Font 
         Name            =   "Rockwell Condensed"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   975
      Left            =   1440
      TabIndex        =   0
      Top             =   0
      Width           =   7335
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000005&
      BorderWidth     =   2
      Index           =   12
      X1              =   6000
      X2              =   6000
      Y1              =   2520
      Y2              =   2040
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000005&
      BorderWidth     =   2
      Index           =   11
      X1              =   5520
      X2              =   5520
      Y1              =   2520
      Y2              =   2040
   End
   Begin VB.Line Line5 
      BorderColor     =   &H80000005&
      BorderWidth     =   2
      X1              =   9960
      X2              =   0
      Y1              =   2280
      Y2              =   2280
   End
   Begin VB.Line Line4 
      BorderWidth     =   2
      X1              =   1440
      X2              =   8280
      Y1              =   960
      Y2              =   960
   End
   Begin VB.OLE OLE2 
      Class           =   "SoundRec"
      Height          =   375
      Left            =   840
      OleObjectBlob   =   "RACE.frx":3D86
      SourceDoc       =   "C:\ICONS\SOUNDS\Gallops.wav"
      TabIndex        =   36
      Top             =   240
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.OLE OLE1 
      Class           =   "SoundRec"
      Height          =   375
      Left            =   360
      OleObjectBlob   =   "RACE.frx":1939E
      SourceDoc       =   "C:\ICONS\SOUNDS\Horses.wav"
      TabIndex        =   35
      Top             =   240
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H00E0E0E0&
      BorderWidth     =   2
      Height          =   375
      Left            =   2520
      Top             =   8160
      Width           =   2175
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Height          =   375
      Index           =   2
      Left            =   4800
      Top             =   6960
      Width           =   2175
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H0000FFFF&
      BorderWidth     =   2
      Height          =   375
      Index           =   1
      Left            =   2520
      Top             =   6960
      Width           =   2175
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      Height          =   375
      Index           =   0
      Left            =   240
      Top             =   6960
      Width           =   2175
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      FillColor       =   &H0000FF00&
      Height          =   375
      Index           =   2
      Left            =   4800
      Top             =   7560
      Width           =   2175
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00C000C0&
      BorderWidth     =   2
      FillColor       =   &H0000FF00&
      Height          =   375
      Index           =   1
      Left            =   240
      Top             =   7560
      Width           =   2175
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00008000&
      BorderWidth     =   2
      FillColor       =   &H0000FF00&
      Height          =   375
      Index           =   0
      Left            =   2520
      Top             =   7560
      Width           =   2175
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "1/"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   6360
      TabIndex        =   34
      Top             =   7560
      Width           =   255
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "1/"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   4080
      TabIndex        =   33
      Top             =   7560
      Width           =   255
   End
   Begin VB.Label odd 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   6
      Left            =   6600
      TabIndex        =   32
      Top             =   7560
      Width           =   255
   End
   Begin VB.Label odd 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   5
      Left            =   4320
      TabIndex        =   31
      Top             =   7560
      Width           =   255
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "BLUE BOY"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   12
      Left            =   10080
      TabIndex        =   30
      Top             =   6255
      Width           =   1695
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "NIJINSKI"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   11
      Left            =   10080
      TabIndex        =   29
      Top             =   5880
      Width           =   1695
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "CHANSITY"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   10
      Left            =   10080
      TabIndex        =   28
      Top             =   5520
      Width           =   1695
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "BOBBY JO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   9
      Left            =   10080
      TabIndex        =   27
      Top             =   5160
      Width           =   1695
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "MONACLE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   8
      Left            =   10080
      TabIndex        =   26
      Top             =   4800
      Width           =   1695
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "SUNY BAY"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   7
      Left            =   10080
      TabIndex        =   25
      Top             =   4440
      Width           =   1695
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "RED DEVIL"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   6
      Left            =   10080
      TabIndex        =   24
      Top             =   4080
      Width           =   1695
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "GREEN FLY"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   5
      Left            =   10080
      TabIndex        =   23
      Top             =   3720
      Width           =   1695
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "FORMULA 1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   4
      Left            =   10080
      TabIndex        =   22
      Top             =   3360
      Width           =   1695
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "EXCALABUR"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   3
      Left            =   10080
      TabIndex        =   21
      Top             =   3000
      Width           =   1695
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "RED RUM"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   2
      Left            =   10080
      TabIndex        =   20
      Top             =   2640
      Width           =   1695
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "LUCKY DAY"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   1
      Left            =   10080
      TabIndex        =   19
      Top             =   2280
      Width           =   1695
   End
   Begin VB.Line Line9 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Visible         =   0   'False
      X1              =   9720
      X2              =   9720
      Y1              =   6240
      Y2              =   2520
   End
   Begin VB.Label odd 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   7
      Left            =   4320
      TabIndex        =   18
      Top             =   8160
      Width           =   255
   End
   Begin VB.Label odd 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   2160
      TabIndex        =   17
      Top             =   7560
      Width           =   255
   End
   Begin VB.Label odd 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   6600
      TabIndex        =   16
      Top             =   6960
      Width           =   255
   End
   Begin VB.Label odd 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   4320
      TabIndex        =   15
      Top             =   6960
      Width           =   375
   End
   Begin VB.Label odd 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   2160
      TabIndex        =   14
      Top             =   6960
      Width           =   255
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "1/"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   6
      Left            =   4080
      TabIndex        =   13
      Top             =   8160
      Width           =   255
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "1/"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   1920
      TabIndex        =   12
      Top             =   7560
      Width           =   255
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "1/"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   6360
      TabIndex        =   11
      Top             =   6960
      Width           =   255
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "1/"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   4080
      TabIndex        =   10
      Top             =   6960
      Width           =   255
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "1/"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   1920
      TabIndex        =   9
      Top             =   6960
      Width           =   255
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000005&
      BorderWidth     =   2
      Index           =   20
      X1              =   9840
      X2              =   9840
      Y1              =   2520
      Y2              =   2040
   End
   Begin VB.Line Line8 
      BorderColor     =   &H80000009&
      BorderWidth     =   2
      Index           =   1
      Visible         =   0   'False
      X1              =   9600
      X2              =   9600
      Y1              =   2520
      Y2              =   6240
   End
   Begin VB.Line Line8 
      BorderColor     =   &H80000009&
      BorderWidth     =   2
      Index           =   0
      Visible         =   0   'False
      X1              =   9120
      X2              =   9120
      Y1              =   2520
      Y2              =   6240
   End
   Begin VB.Label Label4 
      BackColor       =   &H00808000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " HORSES"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10200
      TabIndex        =   4
      Top             =   1800
      Width           =   1455
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "  PLACE BET"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7200
      TabIndex        =   3
      Top             =   7080
      Width           =   1935
   End
   Begin VB.Line Line7 
      BorderWidth     =   2
      X1              =   7080
      X2              =   7080
      Y1              =   8640
      Y2              =   6720
   End
   Begin VB.Line Line6 
      BorderWidth     =   2
      X1              =   9360
      X2              =   9360
      Y1              =   8640
      Y2              =   6720
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "       KITTY"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9600
      TabIndex        =   2
      Top             =   7080
      Width           =   2055
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000005&
      BorderWidth     =   2
      Index           =   19
      X1              =   9360
      X2              =   9360
      Y1              =   2520
      Y2              =   2040
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000005&
      BorderWidth     =   2
      Index           =   18
      X1              =   8880
      X2              =   8880
      Y1              =   2520
      Y2              =   2040
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000005&
      BorderWidth     =   2
      Index           =   17
      X1              =   8400
      X2              =   8400
      Y1              =   2520
      Y2              =   2040
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000005&
      BorderWidth     =   2
      Index           =   16
      X1              =   7920
      X2              =   7920
      Y1              =   2520
      Y2              =   2040
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000005&
      BorderWidth     =   2
      Index           =   15
      X1              =   7440
      X2              =   7440
      Y1              =   2520
      Y2              =   2040
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000005&
      BorderWidth     =   2
      Index           =   14
      X1              =   6960
      X2              =   6960
      Y1              =   2520
      Y2              =   2040
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000005&
      BorderWidth     =   2
      Index           =   13
      X1              =   6480
      X2              =   6480
      Y1              =   2520
      Y2              =   2040
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000005&
      BorderWidth     =   2
      Index           =   10
      X1              =   5040
      X2              =   5040
      Y1              =   2520
      Y2              =   2040
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000005&
      BorderWidth     =   2
      Index           =   9
      X1              =   4560
      X2              =   4560
      Y1              =   2520
      Y2              =   2040
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000005&
      BorderWidth     =   2
      Index           =   8
      X1              =   4080
      X2              =   4080
      Y1              =   2520
      Y2              =   2040
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000005&
      BorderWidth     =   2
      Index           =   7
      X1              =   3600
      X2              =   3600
      Y1              =   2520
      Y2              =   2040
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000005&
      BorderWidth     =   2
      Index           =   6
      X1              =   3120
      X2              =   3120
      Y1              =   2520
      Y2              =   2040
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000005&
      BorderWidth     =   2
      Index           =   5
      X1              =   2640
      X2              =   2640
      Y1              =   2520
      Y2              =   2040
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000005&
      BorderWidth     =   2
      Index           =   4
      X1              =   2160
      X2              =   2160
      Y1              =   2520
      Y2              =   2040
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000005&
      BorderWidth     =   2
      Index           =   3
      X1              =   1680
      X2              =   1680
      Y1              =   2520
      Y2              =   2040
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000005&
      BorderWidth     =   2
      Index           =   2
      X1              =   1200
      X2              =   1200
      Y1              =   2520
      Y2              =   2040
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000005&
      BorderWidth     =   2
      Index           =   1
      X1              =   240
      X2              =   240
      Y1              =   2520
      Y2              =   2040
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000005&
      BorderWidth     =   2
      Index           =   0
      X1              =   720
      X2              =   720
      Y1              =   2520
      Y2              =   2040
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000005&
      BorderWidth     =   3
      X1              =   0
      X2              =   9960
      Y1              =   2400
      Y2              =   2400
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000005&
      BorderWidth     =   3
      X1              =   9960
      X2              =   0
      Y1              =   2160
      Y2              =   2160
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   1695
      Left            =   9960
      Picture         =   "RACE.frx":52BB6
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1980
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H8000000D&
      BackStyle       =   1  'Opaque
      BorderWidth     =   3
      Height          =   5055
      Left            =   9960
      Top             =   1680
      Width           =   1935
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00808000&
      BackStyle       =   1  'Opaque
      BorderWidth     =   3
      Height          =   1935
      Left            =   0
      Top             =   6720
      Width           =   11895
   End
   Begin VB.Shape Shape7 
      BackColor       =   &H00008000&
      BackStyle       =   1  'Opaque
      Height          =   3735
      Index           =   0
      Left            =   4570
      Top             =   2520
      Width           =   120
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H00004080&
      BackStyle       =   1  'Opaque
      Height          =   3735
      Index           =   0
      Left            =   4530
      Top             =   2520
      Width           =   215
   End
   Begin VB.Shape Shape7 
      BackColor       =   &H00008000&
      BackStyle       =   1  'Opaque
      Height          =   3735
      Index           =   1
      Left            =   9270
      Top             =   2520
      Width           =   120
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H00004080&
      BackStyle       =   1  'Opaque
      Height          =   3735
      Index           =   1
      Left            =   9230
      Top             =   2520
      Width           =   215
   End
   Begin VB.Image Image2 
      Height          =   2295
      Left            =   0
      Picture         =   "RACE.frx":54C60
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9975
   End
   Begin VB.Shape Shape9 
      BackColor       =   &H00008000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00008000&
      Height          =   2655
      Left            =   0
      Top             =   0
      Width           =   9975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim X As Integer
Dim odds(1 To 7) As Integer
Dim BET, COUNT1, count2, count3, count4, count5, count6, count7
Dim RESULT, counter, JOHN1, finished
Dim FALL1, JUMP
Const Finish = 9120
Const end1 = 0
Dim end2 As Byte

Private Sub Command1_Click()
If RESULT = BET Then
MsgBox " WELL DONE"
If BET = 1 Then
Text1.Text = Val(Text4.Text) * Val(odd(1).Caption)
End If
If BET = 2 Then
Text1.Text = Val(Text4.Text) * Val(odd(2).Caption)
End If
If BET = 3 Then
Text1.Text = Val(Text4.Text) * Val(odd(3).Caption)
End If
If BET = 4 Then
Text1.Text = Val(Text4.Text) * Val(odd(4).Caption)
End If
If BET = 5 Then
Text1.Text = Val(Text4.Text) * Val(odd(5).Caption)
End If
If BET = 6 Then
Text1.Text = Val(Text4.Text) * Val(odd(6).Caption)
End If
If BET = 7 Then
Text1.Text = Val(Text4.Text) * Val(odd(7).Caption)
End If
End If

Text2.Text = Val(Text1.Text) + Val(Text2.Text)
Text1.Text = ""

Call Form_Load
Command3.Visible = False

For X = 1 To 7
    Option1(X).Enabled = True
    Option1(X).Value = False
Next X

For X = 1 To 7
  HORSE(X).Visible = True
  Next X
If Text2.Text < 200 Then
   Text2.BackColor = &HC0&
   Else
   Text2.BackColor = &HFFFFFF
End If
If Text2.Text < 1 Then
End
End If

BET = 0
RESULT = 0
Line8(1).Visible = False
Line9.Visible = False
Text4.Text = "0"
Shape6(0).Left = 4480
Shape6(1).Left = 9230
Shape7(0).Left = 4550
Shape7(1).Left = 9350

End Sub
Private Sub Command2_Click()
Dim X
end2 = 0
Timer2.Interval = 1
Timer_start.Interval = 1
JOHN_TIM.Interval = 100
FALL.Interval = 400
Command2.Visible = False
counter = 1
End Sub

Private Sub Command3_Click()
Dim X As Byte
HORSE(1).Left = 120
HORSE(2).Left = 120
HORSE(3).Left = 120
HORSE(4).Left = 120
HORSE(5).Left = 120
HORSE(6).Left = 120
HORSE(7).Left = 120
Command3.Visible = False
Command2.Visible = True
Command2.Enabled = True

'...............
For X = 1 To 7
    If Option1(X).Value = False Then
    Option1(X).Enabled = False
    End If
Next X

'...............
Text4.Enabled = False
Text2.Text = Text2.Text - Text4.Text
If Text4.Text < 0 Then
MsgBox "fuck you! fuck you! fuck you! fuck you! fuck you!"
End
End If

If Text2.Text < 0 Then
MsgBox "UNLUCKY YOU BLEW IT, YOUR BANKRUPT, TRY AGAIN", vbCritical
End
End If
End Sub
Private Sub FALL_Timer()
For X = 0 To 1
If HORSE(1).Left > Shape8(X).Left And HORSE(1).Left < (Shape8(X).Left + Shape8(X).Width) Then
If FALL1 = 22 Then
HORSE(1).Visible = False
End If
End If
    If HORSE(2).Left > Shape8(X).Left And HORSE(2).Left < (Shape8(X).Left + Shape8(X).Width) Then
    If FALL1 = 12 Then
    HORSE(2).Visible = False
    End If
    End If
        If HORSE(3).Left > Shape8(X).Left And HORSE(3).Left < (Shape8(X).Left + Shape8(X).Width) Then
        If FALL1 = 1 Then
        HORSE(3).Visible = False
        End If
        End If
            If HORSE(4).Left > Shape8(X).Left And HORSE(4).Left < (Shape8(X).Left + Shape8(X).Width) Then
            If FALL1 = 22 Then
            HORSE(4).Visible = False
            End If
            End If
                If HORSE(5).Left > Shape8(X).Left And HORSE(5).Left < (Shape8(X).Left + Shape8(X).Width) Then
                If FALL1 = 7 Then
                HORSE(5).Visible = False
                End If
                End If
                If HORSE(6).Left > Shape8(X).Left And HORSE(6).Left < (Shape8(X).Left + Shape8(X).Width) Then
                    If FALL1 = 45 Then
                    HORSE(6).Visible = False
                    End If
                    End If
                        If HORSE(7).Left > Shape8(X).Left And HORSE(7).Left < (Shape8(X).Left + Shape8(X).Width) Then
                        If FALL1 = 44 Then
                        HORSE(7).Visible = False
                        End If
                        End If
 
    Next X
'...................................................
If HORSE(1).Visible = True Then
If HORSE(1).Left > HORSE(2).Left And HORSE(1).Left > HORSE(3).Left And HORSE(1).Left > HORSE(4).Left And HORSE(1).Left > HORSE(5).Left And HORSE(1).Left > HORSE(6).Left And HORSE(1).Left > HORSE(7).Left Then
Text3.Text = "                   '" & Option1(1).Caption & "' IS AHEAD OF THE REST"
End If
End If
       
    If HORSE(2).Visible = True Then
    If HORSE(2).Left > HORSE(1).Left And HORSE(2).Left > HORSE(3).Left And HORSE(2).Left > HORSE(4).Left And HORSE(2).Left > HORSE(5).Left And HORSE(2).Left > HORSE(6).Left And HORSE(2).Left > HORSE(7).Left Then
    Text3.Text = "                   CAN " & "'" & Option1(2).Caption & "' BE CAUGHT"
    End If
    End If
      
        If HORSE(3).Visible = True Then
        If HORSE(3).Left > HORSE(2).Left And HORSE(3).Left > HORSE(1).Left And HORSE(3).Left > HORSE(4).Left And HORSE(3).Left > HORSE(5).Left And HORSE(3).Left > HORSE(6).Left And HORSE(3).Left > HORSE(7).Left Then
        Text3.Text = "                  '" & Option1(3).Caption & "' IS AHEAD OF THE REST"
        End If
        End If
            
            If HORSE(4).Visible = True Then
            If HORSE(4).Left > HORSE(2).Left And HORSE(4).Left > HORSE(3).Left And HORSE(4).Left > HORSE(1).Left And HORSE(4).Left > HORSE(5).Left And HORSE(4).Left > HORSE(6).Left And HORSE(4).Left > HORSE(7).Left Then
            Text3.Text = "                    '" & Option1(4).Caption & "' HEAD'S JUST INFRONT"
            End If
            End If
                
                If HORSE(5).Visible = True Then
                If HORSE(5).Left > HORSE(2).Left And HORSE(5).Left > HORSE(3).Left And HORSE(5).Left > HORSE(4).Left And HORSE(5).Left > HORSE(1).Left And HORSE(5).Left > HORSE(6).Left And HORSE(5).Left > HORSE(7).Left Then
                Text3.Text = "                   '" & Option1(5).Caption & "' LOOKS LIKE WINNING NOW"
                End If
                End If
                   
                    If HORSE(6).Visible = True Then
                    If HORSE(6).Left > HORSE(2).Left And HORSE(6).Left > HORSE(3).Left And HORSE(6).Left > HORSE(4).Left And HORSE(6).Left > HORSE(5).Left And HORSE(6).Left > HORSE(1).Left And HORSE(6).Left > HORSE(7).Left Then
                    Text3.Text = "                      '" & Option1(6).Caption & "' IS PULLING AWAY"
                    End If
                    End If
                        
                        If HORSE(7).Visible = True Then
                        If HORSE(7).Left > HORSE(2).Left And HORSE(7).Left > HORSE(3).Left And HORSE(7).Left > HORSE(4).Left And HORSE(7).Left > HORSE(5).Left And HORSE(7).Left > HORSE(6).Left And HORSE(7).Left > HORSE(1).Left Then
                        Text3.Text = "                      THE REST NEED TO CATCH UP WITH '" & Option1(7).Caption & "'"
                        End If
                        End If

End Sub

Private Sub Form_Load()
'................................(randomize odds)
Randomize
For X = 1 To 7
  odds(X) = Rnd * 8 + 2
  Next X
For X = 1 To 7
  odd(X).Caption = odds(X)
  Next X
'..................................(picks horse)
Dim indexnum As Byte
For X = 1 To 12
Label6(X).Tag = 0
Next X
X = 1
For X = 1 To 7
redo:
indexnum = Int(Rnd * 12) + 1
If Label6(indexnum).Tag = 0 Then
Option1(X).Caption = Label6(indexnum).Caption
Label6(indexnum).Tag = 1
Else
GoTo redo
End If
Next X
'..................................(LINES STUFF)
Line3(0).X1 = 720
                Line3(0).x2 = 720
Line3(1).X1 = 240
                Line3(1).x2 = 240
Line3(2).X1 = 1200
                Line3(2).x2 = 1200
Line3(3).X1 = 1680
                Line3(3).x2 = 1680
Line3(4).X1 = 2160
                Line3(4).x2 = 2160
Line3(5).X1 = 2640
                Line3(5).x2 = 2640
Line3(6).X1 = 3120
                Line3(6).x2 = 3120
Line3(7).X1 = 3600
                Line3(7).x2 = 3600
Line3(8).X1 = 4080
                Line3(8).x2 = 4080
Line3(9).X1 = 4560
                Line3(9).x2 = 4560
Line3(10).X1 = 5040
                Line3(10).x2 = 5040
Line3(11).X1 = 5520
                Line3(11).x2 = 5520
Line3(12).X1 = 6000
                Line3(12).x2 = 6000
Line3(13).X1 = 6480
                Line3(13).x2 = 6480
Line3(14).X1 = 6960
                Line3(14).x2 = 6960
Line3(15).X1 = 7440
                Line3(15).x2 = 7440
Line3(16).X1 = 7920
                Line3(16).x2 = 7920
Line3(17).X1 = 8400
                Line3(17).x2 = 8400
Line3(18).X1 = 8880
                Line3(18).x2 = 8880
Line3(19).X1 = 9360
                Line3(19).x2 = 9360
Line3(20).X1 = 9840
                Line3(20).x2 = 9840
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
finished = False
COUNT1 = 1
count2 = 1
count3 = 1
count4 = 1
count5 = 1
count6 = 1
count7 = 1
JOHN1 = 1
counter = 1

Timer3.Interval = 1
Load FINISHED1
FINISHED1.Top = -6000
FINISHED1.Left = 3300
Form1.Enabled = False
End Sub

Private Sub Form_Terminate()
End
End Sub

Private Sub JOHN_TIM_Timer()
If JOHN1 = 1 Then
    Image1.Picture = LoadResPicture("JOHN1", vbResBitmap)
    JOHN1 = 2
    Exit Sub
End If
If JOHN1 = 2 Then
    Image1.Picture = LoadResPicture("JOHN2", vbResBitmap)
    JOHN1 = 3
    Exit Sub
End If
If JOHN1 = 3 Then
    Image1.Picture = LoadResPicture("JOHN3", vbResBitmap)
    JOHN1 = 4
    Exit Sub
End If
If JOHN1 = 4 Then
    Image1.Picture = LoadResPicture("JOHN4", vbResBitmap)
    JOHN1 = 1
    Exit Sub
End If
End Sub
Private Sub Option1_Click(Index As Integer)
Text4.Enabled = True
BET = Index
End Sub

Private Sub Picture2_Click()

End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
Command3.Visible = True
End Sub
Private Sub Timer_start_Timer()
If end2 = 0 Then
If counter = 1 Then
OLE1.DoVerb
counter = 2
Exit Sub
End If

If counter = 2 Then
Timer_start.Interval = 10
HORSE1.Interval = 150
End If

'......................................(horse pace)
Randomize
For X = 1 To 7
HORSE(X).Left = HORSE(X).Left + Int((35 * Rnd) + 1) - (odds(X) / 4)
Next X
For X = 1 To 7
If HORSE(X).Left > 8880 Then
Line9.Visible = True
Line8(1).Visible = True
End If
Next X
Call FENCES
End If
End Sub

Private Sub Timer2_Timer()
Randomize
FALL1 = Int((Rnd * 99) + 1)
'...................................
Shape6(0).Left = Line3(7).X1
Shape6(1).Left = Line3(16).X1

Shape7(0).Left = Line3(7).X1 + 50
Shape7(1).Left = Line3(16).X1 + 50
'.....................................
JUMP = 1

Shape8(0).Left = Shape7(0).Left - 700
Shape8(1).Left = Shape7(1).Left - 700
'.....................................

If HORSE(1).Visible = True Then
If HORSE(1).Left > Finish Then
Timer_start.Interval = end1
end2 = 1
HORSE1.Interval = 0
FALL.Interval = 0
MsgBox "HORSE 1 WINS"
RESULT = 1
finished = True
Text3.Text = ""
Timer2.Interval = 0
End If
End If

           If HORSE(2).Visible = True Then
           If HORSE(2).Left > Finish Then
           Timer_start.Interval = end1
           end2 = 1
           HORSE1.Interval = 0
           FALL.Interval = 0
           MsgBox "HORSE 2 WINS"
           RESULT = 2
           finished = True
           Text3.Text = ""
           Timer2.Interval = 0
           End If
           End If
           
 If HORSE(3).Visible = True Then
If HORSE(3).Left > Finish Then
Timer_start.Interval = end1
end2 = 1
HORSE1.Interval = 0
FALL.Interval = 0
MsgBox "HORSE 3 WINS"
RESULT = 3
finished = True
Text3.Text = ""
Timer2.Interval = 0
End If
End If
                    If HORSE(4).Visible = True Then
                  If HORSE(4).Left > Finish Then
                  Timer_start.Interval = end1
                  end2 = 1
                  HORSE1.Interval = 0
                  FALL.Interval = 0
                  MsgBox "HORSE 4 WINS"
                  RESULT = 4
                  finished = True
                  Text3.Text = ""
                  Timer2.Interval = 0
                  End If
                  End If
                  
If HORSE(5).Visible = True Then
If HORSE(5).Left > Finish Then
Timer_start.Interval = end1
end2 = 1
HORSE1.Interval = 0
FALL.Interval = 0
MsgBox "HORSE 5 WINS"
RESULT = 5
finished = True
Text3.Text = ""
Timer2.Interval = 0
End If
End If
                    If HORSE(6).Visible = True Then
                    If HORSE(6).Left > Finish Then
                    Timer_start.Interval = end1
                    end2 = 1
                    HORSE1.Interval = 0
                    FALL.Interval = 0
                    MsgBox "HORSE 6 WINS"
                    RESULT = 6
                    finished = True
                    Text3.Text = ""
                    Timer2.Interval = 0
                    End If
                    End If
If HORSE(7).Visible = True Then
If HORSE(7).Left > Finish Then
Timer_start.Interval = end1
end2 = 1
HORSE1.Interval = 0
FALL.Interval = 0
MsgBox "HORSE 7 WINS"
RESULT = 7
finished = True
Text3.Text = ""
Timer2.Interval = 0
End If
End If
           
'.......................................(FINISHED)
If finished = True Then
Call Command1_Click
'Command1.Visible = True
HORSE(1).Left = 120
HORSE(2).Left = 120
HORSE(3).Left = 120
HORSE(4).Left = 120
HORSE(5).Left = 120
HORSE(6).Left = 120
HORSE(7).Left = 120
JOHN_TIM.Interval = 0
End If
    End Sub
Public Sub FENCES()
Dim X
For X = 0 To 20
Line3(X).X1 = Line3(X).X1 - 40
Line3(X).x2 = Line3(X).x2 - 40
    If Line3(X).X1 < 0 Then
    Line3(X).X1 = 10000
    Line3(X).x2 = 10000
    End If
Next X

For X = 0 To 20
Line12(X).X1 = Line12(X).X1 - 40
Line12(X).x2 = Line12(X).x2 - 40
    If Line12(X).X1 < 0 Then
    Line12(X).X1 = 10000
    Line12(X).x2 = 10000
End If
Next X

End Sub
Private Sub HORSE1_Timer()
'.......................HORSE1
If COUNT1 = 1 Then
HORSE(1).Picture = LoadResPicture(112, vbResIcon)
COUNT1 = 2
GoTo ct2
End If
If COUNT1 = 2 Then
HORSE(1).Picture = LoadResPicture(12, vbResIcon)
COUNT1 = 1
GoTo ct2
End If
'............................
ct2:
If count2 = 1 Then
HORSE(2).Picture = LoadResPicture(22, vbResIcon)
count2 = 2
GoTo ct3
End If
If count2 = 2 Then
HORSE(2).Picture = LoadResPicture(2, vbResIcon)
count2 = 1
GoTo ct3
End If
'..............................
ct3:
If count3 = 1 Then
HORSE(3).Picture = LoadResPicture(33, vbResIcon)
count3 = 2
GoTo ct4
End If
If count3 = 2 Then
HORSE(3).Picture = LoadResPicture(3, vbResIcon)
count3 = 1
GoTo ct4
End If
'................................
ct4:
If count4 = 1 Then
HORSE(4).Picture = LoadResPicture(44, vbResIcon)
count4 = 2
GoTo ct5
End If
If count4 = 2 Then
HORSE(4).Picture = LoadResPicture(4, vbResIcon)
count4 = 1
GoTo ct5
End If
'.........................
ct5:
If count5 = 1 Then
HORSE(5).Picture = LoadResPicture(55, vbResIcon)
count5 = 2
GoTo ct6
End If
If count5 = 2 Then
HORSE(5).Picture = LoadResPicture(5, vbResIcon)
count5 = 1
GoTo ct6
End If
'..............................
ct6:
If count6 = 1 Then
HORSE(6).Picture = LoadResPicture(66, vbResIcon)
count6 = 2
GoTo ct7
End If
If count6 = 2 Then
HORSE(6).Picture = LoadResPicture(6, vbResIcon)
count6 = 1
GoTo ct7
End If
'..............................
ct7:
If count7 = 1 Then
HORSE(7).Picture = LoadResPicture(77, vbResIcon)
count7 = 2
GoTo ct8
End If
If count7 = 2 Then
HORSE(7).Picture = LoadResPicture(7, vbResIcon)
count7 = 1
GoTo ct8
End If
'................
ct8:
If Form1.Visible = True Then
End If
'..................................................................................
'..................................................................................
For X = 0 To 1
If HORSE(1).Left > Shape8(X).Left And HORSE(1).Left < (Shape8(X).Left + Shape8(X).Width) Then
HORSE(1).Picture = LoadResPicture(111, vbResIcon)
End If
    If HORSE(2).Left > Shape8(X).Left And HORSE(2).Left < (Shape8(X).Left + Shape8(X).Width) Then
    HORSE(2).Picture = LoadResPicture(222, vbResIcon)
    End If
        If HORSE(3).Left > Shape8(X).Left And HORSE(3).Left < (Shape8(X).Left + Shape8(X).Width) Then
        HORSE(3).Picture = LoadResPicture(333, vbResIcon)
        End If
            If HORSE(4).Left > Shape8(X).Left And HORSE(4).Left < (Shape8(X).Left + Shape8(X).Width) Then
            HORSE(4).Picture = LoadResPicture(444, vbResIcon)
            End If
                If HORSE(5).Left > Shape8(X).Left And HORSE(5).Left < (Shape8(X).Left + Shape8(X).Width) Then
                HORSE(5).Picture = LoadResPicture(555, vbResIcon)
                End If
                   If HORSE(6).Left > Shape8(X).Left And HORSE(6).Left < (Shape8(X).Left + Shape8(X).Width) Then
                    HORSE(6).Picture = LoadResPicture(666, vbResIcon)
                    End If
                        If HORSE(7).Left > Shape8(X).Left And HORSE(7).Left < (Shape8(X).Left + Shape8(X).Width) Then
                        HORSE(7).Picture = LoadResPicture(777, vbResIcon)
                        End If

Next X
End Sub

Private Sub Timer3_Timer()
Load FINISHED1
FINISHED1.Show

FINISHED1.Top = FINISHED1.Top + 270

If FINISHED1.Top > 1040 Then
FINISHED1.Command1.Enabled = True
Timer3.Interval = 0
End If

End Sub

