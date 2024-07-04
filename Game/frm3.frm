VERSION 5.00
Begin VB.Form frm3 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Game   Created by   Pollob.C.Roy"
   ClientHeight    =   10770
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   20250
   Icon            =   "frm3.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10770
   ScaleWidth      =   20250
   Visible         =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer9 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   11880
      Top             =   9720
   End
   Begin VB.Timer Timer8 
      Interval        =   1
      Left            =   11280
      Top             =   9720
   End
   Begin VB.Timer Timer7 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   3600
      Top             =   3600
   End
   Begin VB.Timer Timer6 
      Interval        =   1
      Left            =   5880
      Top             =   4800
   End
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   1920
      Top             =   4800
   End
   Begin VB.Timer Timer5 
      Interval        =   1
      Left            =   3840
      Top             =   9000
   End
   Begin VB.Timer Timer3 
      Interval        =   100
      Left            =   960
      Top             =   8760
   End
   Begin VB.Timer Timer2 
      Interval        =   1
      Left            =   1320
      Top             =   7800
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   600
      Top             =   7800
   End
   Begin VB.Image Image5 
      Height          =   1095
      Left            =   3360
      Picture         =   "frm3.frx":0A02
      Stretch         =   -1  'True
      Top             =   8520
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label30 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Pollob .C. Roy"
      BeginProperty Font 
         Name            =   "Pristina"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   735
      Left            =   8880
      TabIndex        =   93
      Top             =   10200
      Width           =   2535
   End
   Begin VB.Label Label29 
      BackColor       =   &H0080FFFF&
      Caption         =   "Developer...."
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   8880
      TabIndex        =   92
      Top             =   9720
      Width           =   2535
   End
   Begin VB.Image Image8 
      Height          =   495
      Left            =   19680
      Picture         =   "frm3.frx":5B1C
      Stretch         =   -1  'True
      Top             =   240
      Width           =   495
   End
   Begin VB.Image Image7 
      Height          =   495
      Left            =   19200
      Picture         =   "frm3.frx":26A59
      Stretch         =   -1  'True
      Top             =   240
      Width           =   495
   End
   Begin VB.Label Label28 
      BackStyle       =   0  'Transparent
      Caption         =   "Robber :"
      BeginProperty Font 
         Name            =   "Perpetua Titling MT"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   735
      Left            =   120
      TabIndex        =   91
      Top             =   3480
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.Label Label27 
      BackStyle       =   0  'Transparent
      Caption         =   "Thief     :"
      BeginProperty Font 
         Name            =   "Perpetua Titling MT"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   735
      Left            =   120
      TabIndex        =   90
      Top             =   2640
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.Label Label26 
      BackStyle       =   0  'Transparent
      Caption         =   "Police   :"
      BeginProperty Font 
         Name            =   "Perpetua Titling MT"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   735
      Left            =   120
      TabIndex        =   89
      Top             =   1800
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.Label Label25 
      BackStyle       =   0  'Transparent
      Caption         =   "King     :"
      BeginProperty Font 
         Name            =   "Perpetua Titling MT"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   735
      Left            =   240
      TabIndex        =   88
      Top             =   960
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.Line Line2 
      BorderWidth     =   4
      X1              =   16320
      X2              =   16320
      Y1              =   240
      Y2              =   9480
   End
   Begin VB.Line Line4 
      BorderWidth     =   4
      X1              =   10800
      X2              =   10800
      Y1              =   240
      Y2              =   9480
   End
   Begin VB.Line Line5 
      BorderWidth     =   4
      X1              =   8040
      X2              =   8040
      Y1              =   240
      Y2              =   9480
   End
   Begin VB.Label Label24 
      Height          =   375
      Left            =   6360
      TabIndex        =   87
      Top             =   3600
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Label23 
      Height          =   375
      Left            =   6360
      TabIndex        =   86
      Top             =   2880
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Label22 
      Height          =   375
      Left            =   6360
      TabIndex        =   85
      Top             =   2040
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Label21 
      Height          =   375
      Left            =   6360
      TabIndex        =   84
      Top             =   1080
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Label19 
      Caption         =   "0"
      Height          =   735
      Left            =   0
      TabIndex        =   82
      Top             =   8640
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label Label18 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   735
      Left            =   120
      TabIndex        =   81
      Top             =   4680
      Visible         =   0   'False
      Width           =   4695
   End
   Begin VB.Image Image6 
      Height          =   975
      Left            =   5400
      Picture         =   "frm3.frx":2B31D
      Stretch         =   -1  'True
      Top             =   4560
      Width           =   1215
   End
   Begin VB.Label Label17 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   735
      Left            =   3120
      TabIndex        =   80
      Top             =   3480
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.Label Label16 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   735
      Left            =   3120
      TabIndex        =   79
      Top             =   2640
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.Label Label15 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   735
      Left            =   3120
      TabIndex        =   78
      Top             =   1800
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.Label Label14 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   735
      Left            =   3120
      TabIndex        =   77
      Top             =   960
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.Image Image4 
      Height          =   2655
      Left            =   6120
      Picture         =   "frm3.frx":2EECA
      Stretch         =   -1  'True
      Top             =   5760
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Image Image3 
      Height          =   2655
      Left            =   4080
      Picture         =   "frm3.frx":11949D
      Stretch         =   -1  'True
      Top             =   5760
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Image Image2 
      Height          =   2655
      Left            =   2040
      Picture         =   "frm3.frx":203A70
      Stretch         =   -1  'True
      Top             =   5760
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Image Image1 
      Height          =   2655
      Left            =   0
      Picture         =   "frm3.frx":2EE043
      Stretch         =   -1  'True
      Top             =   5760
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Total  :"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   735
      Left            =   6120
      TabIndex        =   76
      Top             =   8880
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Line Line1 
      BorderWidth     =   4
      X1              =   19080
      X2              =   19080
      Y1              =   240
      Y2              =   9480
   End
   Begin VB.Line Line3 
      BorderWidth     =   4
      X1              =   13440
      X2              =   13440
      Y1              =   240
      Y2              =   9480
   End
   Begin VB.Line Line6 
      BorderColor     =   &H00008080&
      BorderWidth     =   4
      Index           =   17
      X1              =   8040
      X2              =   19080
      Y1              =   9000
      Y2              =   9000
   End
   Begin VB.Line Line6 
      BorderWidth     =   2
      Index           =   16
      X1              =   8040
      X2              =   19080
      Y1              =   8520
      Y2              =   8520
   End
   Begin VB.Line Line6 
      BorderWidth     =   2
      Index           =   15
      X1              =   8040
      X2              =   19080
      Y1              =   8040
      Y2              =   8040
   End
   Begin VB.Line Line6 
      BorderWidth     =   2
      Index           =   14
      X1              =   8040
      X2              =   19080
      Y1              =   7560
      Y2              =   7560
   End
   Begin VB.Line Line6 
      BorderWidth     =   2
      Index           =   13
      X1              =   8040
      X2              =   19080
      Y1              =   7080
      Y2              =   7080
   End
   Begin VB.Line Line6 
      BorderWidth     =   2
      Index           =   12
      X1              =   8040
      X2              =   19080
      Y1              =   6600
      Y2              =   6600
   End
   Begin VB.Line Line6 
      BorderWidth     =   2
      Index           =   11
      X1              =   8040
      X2              =   19080
      Y1              =   6120
      Y2              =   6120
   End
   Begin VB.Line Line6 
      BorderWidth     =   2
      Index           =   10
      X1              =   8040
      X2              =   19080
      Y1              =   5640
      Y2              =   5640
   End
   Begin VB.Line Line6 
      BorderWidth     =   2
      Index           =   9
      X1              =   8040
      X2              =   19080
      Y1              =   5160
      Y2              =   5160
   End
   Begin VB.Line Line6 
      BorderWidth     =   2
      Index           =   8
      X1              =   8040
      X2              =   19080
      Y1              =   4680
      Y2              =   4680
   End
   Begin VB.Line Line6 
      BorderWidth     =   2
      Index           =   7
      X1              =   8040
      X2              =   19080
      Y1              =   4200
      Y2              =   4200
   End
   Begin VB.Line Line6 
      BorderWidth     =   2
      Index           =   6
      X1              =   8040
      X2              =   19080
      Y1              =   3720
      Y2              =   3720
   End
   Begin VB.Line Line6 
      BorderWidth     =   2
      Index           =   5
      X1              =   8040
      X2              =   19080
      Y1              =   3240
      Y2              =   3240
   End
   Begin VB.Line Line6 
      BorderWidth     =   2
      Index           =   4
      X1              =   8040
      X2              =   19080
      Y1              =   2760
      Y2              =   2760
   End
   Begin VB.Line Line6 
      BorderWidth     =   2
      Index           =   3
      X1              =   8040
      X2              =   19080
      Y1              =   2280
      Y2              =   2280
   End
   Begin VB.Line Line6 
      BorderWidth     =   2
      Index           =   2
      X1              =   8040
      X2              =   19080
      Y1              =   1800
      Y2              =   1800
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00FF80FF&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   22.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   495
      Index           =   12
      Left            =   8160
      TabIndex        =   16
      Top             =   6600
      Width           =   2535
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00FF80FF&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   22.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   495
      Index           =   11
      Left            =   8160
      TabIndex        =   15
      Top             =   6120
      Width           =   2535
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00FF80FF&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   22.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   495
      Index           =   10
      Left            =   8160
      TabIndex        =   14
      Top             =   5640
      Width           =   2535
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00FF80FF&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   22.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   495
      Index           =   9
      Left            =   8160
      TabIndex        =   13
      Top             =   5160
      Width           =   2535
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00FF80FF&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   22.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   495
      Index           =   8
      Left            =   8160
      TabIndex        =   12
      Top             =   4680
      Width           =   2535
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00FF80FF&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   22.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   495
      Index           =   7
      Left            =   8160
      TabIndex        =   11
      Top             =   4200
      Width           =   2535
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00FF80FF&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   22.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   495
      Index           =   6
      Left            =   8160
      TabIndex        =   10
      Top             =   3720
      Width           =   2535
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00FF80FF&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   22.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   495
      Index           =   5
      Left            =   8160
      TabIndex        =   9
      Top             =   3240
      Width           =   2535
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00FF80FF&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   22.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   495
      Index           =   4
      Left            =   8160
      TabIndex        =   8
      Top             =   2760
      Width           =   2535
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00FF80FF&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   22.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   495
      Index           =   3
      Left            =   8160
      TabIndex        =   7
      Top             =   2280
      Width           =   2535
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00FF80FF&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   22.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   495
      Index           =   2
      Left            =   8160
      TabIndex        =   6
      Top             =   1800
      Width           =   2535
   End
   Begin VB.Line Line6 
      BorderWidth     =   2
      Index           =   1
      X1              =   8040
      X2              =   19080
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Line Line6 
      BorderWidth     =   4
      Index           =   0
      X1              =   8040
      X2              =   19080
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   615
      Left            =   16320
      TabIndex        =   3
      Top             =   240
      Width           =   2775
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   615
      Left            =   13560
      TabIndex        =   2
      Top             =   240
      Width           =   2775
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   615
      Left            =   10800
      TabIndex        =   1
      Top             =   240
      Width           =   2775
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   615
      Left            =   8040
      TabIndex        =   0
      Top             =   240
      Width           =   2775
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   22.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   495
      Index           =   0
      Left            =   10920
      TabIndex        =   21
      Top             =   840
      Width           =   2295
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00FF80FF&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   22.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   495
      Index           =   0
      Left            =   8160
      TabIndex        =   4
      Top             =   840
      Width           =   2535
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00FF80FF&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   22.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   495
      Index           =   1
      Left            =   8160
      TabIndex        =   5
      Top             =   1320
      Width           =   2535
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00FF80FF&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   22.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   495
      Index           =   13
      Left            =   8160
      TabIndex        =   17
      Top             =   7080
      Width           =   2535
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00FF80FF&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   22.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   495
      Index           =   14
      Left            =   8160
      TabIndex        =   18
      Top             =   7560
      Width           =   2535
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00FF80FF&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   22.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   495
      Index           =   15
      Left            =   8160
      TabIndex        =   19
      Top             =   8040
      Width           =   2535
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00FF80FF&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   22.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   495
      Index           =   16
      Left            =   8160
      TabIndex        =   20
      Top             =   8520
      Width           =   2535
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   22.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   495
      Index           =   1
      Left            =   10920
      TabIndex        =   24
      Top             =   1320
      Width           =   2295
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   22.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   495
      Index           =   2
      Left            =   10920
      TabIndex        =   25
      Top             =   1800
      Width           =   2295
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   22.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   495
      Index           =   3
      Left            =   10920
      TabIndex        =   26
      Top             =   2280
      Width           =   2295
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   22.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   495
      Index           =   4
      Left            =   10920
      TabIndex        =   27
      Top             =   2760
      Width           =   2295
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   22.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   495
      Index           =   5
      Left            =   10920
      TabIndex        =   28
      Top             =   3240
      Width           =   2295
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   22.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   495
      Index           =   6
      Left            =   10920
      TabIndex        =   29
      Top             =   3720
      Width           =   2295
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   22.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   495
      Index           =   7
      Left            =   10920
      TabIndex        =   30
      Top             =   4200
      Width           =   2295
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   22.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   495
      Index           =   8
      Left            =   10920
      TabIndex        =   31
      Top             =   4680
      Width           =   2295
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   22.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   495
      Index           =   9
      Left            =   10920
      TabIndex        =   32
      Top             =   5160
      Width           =   2295
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   22.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   495
      Index           =   10
      Left            =   10920
      TabIndex        =   33
      Top             =   5640
      Width           =   2295
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   22.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   495
      Index           =   11
      Left            =   10920
      TabIndex        =   34
      Top             =   6120
      Width           =   2295
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   22.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   495
      Index           =   12
      Left            =   10920
      TabIndex        =   35
      Top             =   6600
      Width           =   2295
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   22.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   495
      Index           =   13
      Left            =   10920
      TabIndex        =   36
      Top             =   7080
      Width           =   2295
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   22.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   495
      Index           =   14
      Left            =   10920
      TabIndex        =   37
      Top             =   7560
      Width           =   2295
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   22.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   495
      Index           =   15
      Left            =   10920
      TabIndex        =   38
      Top             =   8040
      Width           =   2295
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   22.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   495
      Index           =   16
      Left            =   10920
      TabIndex        =   39
      Top             =   8520
      Width           =   2295
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   22.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   495
      Index           =   0
      Left            =   13800
      TabIndex        =   22
      Top             =   840
      Width           =   2295
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   22.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   495
      Index           =   1
      Left            =   13800
      TabIndex        =   40
      Top             =   1320
      Width           =   2295
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   22.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   495
      Index           =   2
      Left            =   13800
      TabIndex        =   41
      Top             =   1800
      Width           =   2295
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   22.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   495
      Index           =   3
      Left            =   13800
      TabIndex        =   42
      Top             =   2280
      Width           =   2295
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   22.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   495
      Index           =   4
      Left            =   13800
      TabIndex        =   43
      Top             =   2760
      Width           =   2295
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   22.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   495
      Index           =   5
      Left            =   13800
      TabIndex        =   44
      Top             =   3240
      Width           =   2295
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   22.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   495
      Index           =   6
      Left            =   13800
      TabIndex        =   45
      Top             =   3720
      Width           =   2295
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   22.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   495
      Index           =   7
      Left            =   13800
      TabIndex        =   46
      Top             =   4200
      Width           =   2295
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   22.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   495
      Index           =   8
      Left            =   13800
      TabIndex        =   47
      Top             =   4680
      Width           =   2295
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   22.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   495
      Index           =   9
      Left            =   13800
      TabIndex        =   48
      Top             =   5160
      Width           =   2295
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   22.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   495
      Index           =   10
      Left            =   13800
      TabIndex        =   49
      Top             =   5640
      Width           =   2295
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   22.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   495
      Index           =   11
      Left            =   13800
      TabIndex        =   50
      Top             =   6120
      Width           =   2295
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   22.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   495
      Index           =   12
      Left            =   13800
      TabIndex        =   51
      Top             =   6600
      Width           =   2295
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   22.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   495
      Index           =   13
      Left            =   13800
      TabIndex        =   52
      Top             =   7080
      Width           =   2295
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   22.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   495
      Index           =   14
      Left            =   13800
      TabIndex        =   53
      Top             =   7560
      Width           =   2295
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   22.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   495
      Index           =   15
      Left            =   13800
      TabIndex        =   54
      Top             =   8040
      Width           =   2295
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   22.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   495
      Index           =   16
      Left            =   13800
      TabIndex        =   55
      Top             =   8520
      Width           =   2295
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   22.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   495
      Index           =   0
      Left            =   16560
      TabIndex        =   23
      Top             =   840
      Width           =   2295
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   22.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   495
      Index           =   1
      Left            =   16560
      TabIndex        =   56
      Top             =   1320
      Width           =   2295
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   22.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   495
      Index           =   2
      Left            =   16560
      TabIndex        =   57
      Top             =   1800
      Width           =   2295
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   22.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   495
      Index           =   3
      Left            =   16560
      TabIndex        =   58
      Top             =   2280
      Width           =   2295
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   22.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   495
      Index           =   4
      Left            =   16560
      TabIndex        =   59
      Top             =   2760
      Width           =   2295
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   22.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   495
      Index           =   5
      Left            =   16560
      TabIndex        =   60
      Top             =   3240
      Width           =   2295
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   22.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   495
      Index           =   6
      Left            =   16560
      TabIndex        =   61
      Top             =   3720
      Width           =   2295
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   22.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   495
      Index           =   7
      Left            =   16560
      TabIndex        =   62
      Top             =   4200
      Width           =   2295
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   22.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   495
      Index           =   8
      Left            =   16560
      TabIndex        =   63
      Top             =   4680
      Width           =   2295
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   22.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   495
      Index           =   9
      Left            =   16560
      TabIndex        =   64
      Top             =   5160
      Width           =   2295
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   22.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   495
      Index           =   10
      Left            =   16560
      TabIndex        =   65
      Top             =   5640
      Width           =   2295
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   22.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   495
      Index           =   11
      Left            =   16560
      TabIndex        =   66
      Top             =   6120
      Width           =   2295
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   22.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   495
      Index           =   12
      Left            =   16560
      TabIndex        =   67
      Top             =   6600
      Width           =   2295
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   22.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   495
      Index           =   13
      Left            =   16560
      TabIndex        =   68
      Top             =   7080
      Width           =   2295
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   22.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   495
      Index           =   14
      Left            =   16560
      TabIndex        =   69
      Top             =   7560
      Width           =   2295
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   22.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   495
      Index           =   15
      Left            =   16560
      TabIndex        =   70
      Top             =   8040
      Width           =   2295
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   22.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   495
      Index           =   16
      Left            =   16560
      TabIndex        =   71
      Top             =   8520
      Width           =   2295
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FF8080&
      BackStyle       =   1  'Opaque
      Height          =   8775
      Left            =   8040
      Top             =   240
      Width           =   11055
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   8040
      TabIndex        =   72
      Top             =   9000
      Width           =   2775
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   13440
      TabIndex        =   74
      Top             =   9000
      Width           =   2895
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   16320
      TabIndex        =   75
      Top             =   9000
      Width           =   2775
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   10800
      TabIndex        =   73
      Top             =   9000
      Width           =   2655
   End
   Begin VB.Label Label20 
      Caption         =   "0"
      Height          =   615
      Left            =   4920
      TabIndex        =   83
      Top             =   8880
      Visible         =   0   'False
      Width           =   2535
   End
End
Attribute VB_Name = "frm3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a, b, c, d, t, x, y As Integer
Public Sub pollob1()
    If a = 100 Then
        Label14.Caption = Label1.Caption
        Label14.Visible = True
        Label21.Caption = 1
        Label25.Visible = True

    ElseIf a = 80 Then
        Label15.Caption = Label1.Caption
        Label15.Visible = True
        Label22.Caption = 1
        Label26.Visible = True

    ElseIf a = 60 Then
        Label16.Caption = Label1.Caption
        Label23.Caption = 1

    ElseIf a = 40 Then
        Label17.Caption = Label1.Caption
        Label24.Caption = 1

    Else

    End If

End Sub


Public Sub pollob2()
If b = 100 Then
    Label14.Caption = Label2.Caption
    Label14.Visible = True
    Label21.Caption = 2
    Label25.Visible = True

    ElseIf b = 80 Then
    Label15.Caption = Label2.Caption
    Label15.Visible = True
    Label22.Caption = 2
    Label26.Visible = True

    ElseIf b = 60 Then
    Label16.Caption = Label2.Caption
    Label23.Caption = 2

    ElseIf b = 40 Then
    Label17.Caption = Label2.Caption
    Label24.Caption = 2

    Else

    End If

End Sub


Public Sub pollob3()
    If c = 100 Then
    Label14.Caption = Label3.Caption
    Label14.Visible = True
    Label21.Caption = 3
    Label25.Visible = True

    ElseIf c = 80 Then
    Label15.Caption = Label3.Caption
    Label15.Visible = True
    Label22.Caption = 3
    Label26.Visible = True

    ElseIf c = 60 Then
    Label16.Caption = Label3.Caption
    Label23.Caption = 3

    ElseIf c = 40 Then
    Label17.Caption = Label3.Caption
    Label24.Caption = 3

    Else

    End If

End Sub

Public Sub pollob4()
If d = 100 Then
    Label14.Caption = Label4.Caption
    Label14.Visible = True
    Label21.Caption = 4
    Label25.Visible = True

    ElseIf d = 80 Then
    Label15.Caption = Label4.Caption
    Label15.Visible = True
    Label22.Caption = 4
    Label26.Visible = True

    ElseIf d = 60 Then
    Label16.Caption = Label4.Caption
    Label23.Caption = 4

    ElseIf d = 40 Then
    Label17.Caption = Label4.Caption
    Label24.Caption = 4

    Else

    End If
End Sub


Private Sub Form_Load()
x = 0
a = 0
b = 0
c = 0
d = 0
y = 1
t = 0

Label1.Caption = frm2.Text1.Text
Label2.Caption = frm2.Text2.Text
Label3.Caption = frm2.Text3.Text
Label4.Caption = frm2.Text4.Text

End Sub

Private Sub Image1_Click()
Image1.Visible = False

End Sub

Private Sub Image2_Click()
Image2.Visible = False


End Sub

Private Sub Image3_Click()
Image3.Visible = False


End Sub

Private Sub Image4_Click()
Image4.Visible = False



End Sub

Private Sub Image5_Click()
frm5.Show
Timer5.Enabled = False
Timer7.Enabled = False
Label20.Caption = t
t = t + 1
Image5.Visible = False


End Sub

Private Sub Image6_Click()
Timer1.Enabled = False
Timer2.Enabled = False
Image1.Visible = True
Image2.Visible = True
Image3.Visible = True
Image4.Visible = True
'Image5.Visible = True
Image6.Visible = False
Timer3.Enabled = True
Label14.Visible = False
Label15.Visible = False
Label16.Visible = False
Label17.Visible = False

Label25.Visible = False
Label26.Visible = False
Label27.Visible = False
Label28.Visible = False

Timer5.Enabled = True
Timer4.Enabled = True
Timer3.Enabled = False
Timer7.Enabled = True

Label14.Caption = ""
Label15.Caption = ""
Label16.Caption = ""
Label17.Caption = ""
Label21.Caption = ""
Label22.Caption = ""
Label23.Caption = ""
Label24.Caption = ""

End Sub

Private Sub Image7_Click()
frm4.Show
End Sub

Private Sub Image8_Click()
End
End Sub

Private Sub Timer1_Timer()

If x < 0 And 5 >= x Then '1
a = 100
b = 80
c = 60
d = 40

ElseIf x < 5 And 10 >= x Then '2
b = 100
a = 80
c = 60
d = 40

ElseIf x < 10 And 15 >= x Then '3
c = 100
b = 80
a = 60
d = 40

ElseIf x < 15 And 20 >= x Then '4////

d = 100
b = 80
c = 60
a = 40


ElseIf x < 20 And 25 >= x Then '5

a = 100
b = 80
d = 60
c = 40



ElseIf x < 25 And 30 >= x Then '6
b = 100
a = 80
d = 60
c = 40

ElseIf x < 30 And 35 >= x Then '7
c = 100
b = 80
d = 60
a = 40

ElseIf x < 35 And 40 >= x Then '8
d = 100
b = 80
a = 60
c = 40

ElseIf x < 40 And 45 >= x Then '9
a = 100
c = 80
b = 60
d = 40

ElseIf x < 45 And 50 >= x Then '10
b = 100
c = 80
a = 60
d = 40

ElseIf x < 50 And 55 >= x Then '11
c = 100
a = 80
b = 60
d = 40

ElseIf x < 55 And 60 >= x Then '12
d = 100
c = 80
b = 40
a = 60

ElseIf x < 60 And 65 >= x Then '13
a = 100
c = 80
d = 60
b = 40

ElseIf x < 65 And 70 >= x Then '14
b = 100
c = 80
d = 60
a = 40

ElseIf x < 70 And 75 >= x Then '15
c = 100
a = 80
d = 60
b = 40

ElseIf x < 75 And 80 >= x Then '16
d = 100
c = 80
a = 60
b = 40


ElseIf x < 80 And 85 >= x Then '17
a = 100
d = 80
b = 60
c = 40

ElseIf x < 85 And 90 >= x Then '18
b = 100
d = 80
a = 60
c = 40

ElseIf x < 90 And 95 >= x Then '19
c = 100
d = 80
b = 60
a = 40

ElseIf x < 95 And 100 >= x Then '20
d = 100
a = 80
b = 60
c = 40

ElseIf x < 100 And 105 >= x Then '21ppppp
a = 100
d = 80
c = 60
b = 40

ElseIf x < 105 And 110 >= x Then '22
b = 100
d = 80
c = 60
a = 40


ElseIf x < 110 And 115 >= x Then '23
c = 100
d = 80
a = 60
b = 40

ElseIf x < 115 And 120 >= x Then '24
d = 100
a = 80
c = 60
b = 40


Else

End If
End Sub

Private Sub Timer2_Timer()
x = x + 1
'Text1.Text = x
If x > 120 Then
x = 0
End If
End Sub

Private Sub Timer3_Timer()
Label19.Caption = y
y = y + 1
If y > 100 Then
y = 1
End If
End Sub



Private Sub Timer4_Timer()
If Image1.Visible = True And Image2.Visible = True And Image3.Visible = True And Image4.Visible = True Then
Label18.Visible = True
Label18.Caption = Label1.Caption & "  Your Try"


ElseIf Image1.Visible = False And Image2.Visible = True And Image3.Visible = True And Image4.Visible = True Then
Label18.Visible = True
Label18.Caption = Label2.Caption & "  Your Try"

ElseIf Image2.Visible = False And Image1.Visible = True And Image3.Visible = True And Image4.Visible = True Then
Label18.Visible = True
Label18.Caption = Label2.Caption & "  Your Try"

ElseIf Image3.Visible = False And Image2.Visible = True And Image1.Visible = True And Image4.Visible = True Then
Label18.Visible = True
Label18.Caption = Label2.Caption & "  Your Try"

ElseIf Image4.Visible = False And Image2.Visible = True And Image3.Visible = True And Image1.Visible = True Then
Label18.Visible = True
Label18.Caption = Label2.Caption & "  Your Try"



ElseIf Image1.Visible = False And Image2.Visible = False Then
    Label18.Visible = True
    Label18.Caption = Label3.Caption & "  Your Try"

ElseIf Image1.Visible = False And Image3.Visible = False Then
Label18.Visible = True
Label18.Caption = Label3.Caption & "  Your Try"

ElseIf Image1.Visible = False And Image4.Visible = False Then
Label18.Visible = True
Label18.Caption = Label3.Caption & "  Your Try"

ElseIf Image2.Visible = False And Image1.Visible = False Then
Label18.Visible = True
Label18.Caption = Label3.Caption & "  Your Try"


ElseIf Image2.Visible = False And Image3.Visible = False Then
Label18.Visible = True
Label18.Caption = Label3.Caption & "  Your Try"

ElseIf Image2.Visible = False And Image4.Visible = False Then
Label18.Visible = True
Label18.Caption = Label3.Caption & "  Your Try"

ElseIf Image3.Visible = False And Image1.Visible = False Then
Label18.Visible = True
Label18.Caption = Label3.Caption & "  Your Try"

ElseIf Image3.Visible = False And Image2.Visible = False Then
Label18.Visible = True
Label18.Caption = Label3.Caption & "  Your Try"

ElseIf Image3.Visible = False And Image4.Visible = False Then
Label18.Visible = True
Label18.Caption = Label3.Caption & "  Your Try"

ElseIf Image4.Visible = False And Image1.Visible = False Then
Label18.Visible = True
Label18.Caption = Label3.Caption & "  Your Try"

ElseIf Image4.Visible = False And Image2.Visible = False Then
Label18.Visible = True
Label18.Caption = Label3.Caption & "  Your Try"

ElseIf Image4.Visible = False And Image3.Visible = False Then
Label18.Visible = True
Label18.Caption = Label3.Caption & "  Your Try"

Else
Label18.Visible = False

End If


If Image1.Visible = True And Image2.Visible = False And Image3.Visible = False And Image4.Visible = False Then
Label18.Visible = True
Label18.Caption = Label4.Caption & "  Your Try"

ElseIf Image1.Visible = False And Image2.Visible = True And Image3.Visible = False And Image4.Visible = False Then
Label18.Visible = True
Label18.Caption = Label4.Caption & "  Your Try"

ElseIf Image1.Visible = False And Image2.Visible = False And Image3.Visible = True And Image4.Visible = False Then
Label18.Visible = True
Label18.Caption = Label4.Caption & "  Your Try"

ElseIf Image1.Visible = False And Image2.Visible = False And Image3.Visible = False And Image4.Visible = True Then
Label18.Visible = True
Label18.Caption = Label4.Caption & "  Your Try"

Else

End If
If Image1.Visible = False And Image2.Visible = False And Image3.Visible = False And Image4.Visible = False Then
Label18.Visible = False
Label18.Caption = ""
Timer4.Enabled = False

End If
End Sub

Private Sub Timer5_Timer()
If Image1.Visible = False And Image2.Visible = False And Image3.Visible = False And Image4.Visible = False And Image6.Visible = False Then

Image5.Visible = True

Else
Image5.Visible = False
End If



End Sub

Private Sub Timer6_Timer()
Dim a0, a1, a2, a3, a4, a5, a6, a7, a8, a9, a10, a11, a12, a13, a14, a15, a16, at As Integer
Dim b0, b1, b2, b3, b4, b5, b6, b7, b8, b9, b10, b11, b12, b13, b14, b15, b16, bt As Integer
Dim c0, c1, c2, c3, c4, c5, c6, c7, c8, c9, c10, c11, c12, c13, c14, c15, c16, ct As Integer
Dim d0, d1, d2, d3, d4, d5, d6, d7, d8, d9, d10, d11, d12, d13, d14, d15, d16, dt As Integer

If Label5(16).Caption <> "" And Label6(16).Caption <> "" And Label7(16).Caption <> "" And Label8(16).Caption <> "" Then

Image6.Visible = False
Timer3.Enabled = False
Label13.Visible = True

    a0 = Val(Label5(0))
    a1 = Val(Label5(1))
    a2 = Val(Label5(2))
    a3 = Val(Label5(3))
    a4 = Val(Label5(4))
    a5 = Val(Label5(5))
    a6 = Val(Label5(6))
    a7 = Val(Label5(7))
    a8 = Val(Label5(8))
    a9 = Val(Label5(9))
    a10 = Val(Label5(10))
    a11 = Val(Label5(11))
    a12 = Val(Label5(12))
    a13 = Val(Label5(13))
    a14 = Val(Label5(14))
    a15 = Val(Label5(15))
    a16 = Val(Label5(16))
    
    at = a0 + a1 + a2 + a3 + a4 + a5 + a6 + a7 + a8 + a9 + a10 + a11 + a12 + a13 + a14 + a15 + a16
    Label9.Caption = at
    
     b0 = Val(Label6(0))
     b1 = Val(Label6(1))
     b2 = Val(Label6(2))
     b3 = Val(Label6(3))
     b4 = Val(Label6(4))
     b5 = Val(Label6(5))
     b6 = Val(Label6(6))
     b7 = Val(Label6(7))
     b8 = Val(Label6(8))
     b9 = Val(Label6(9))
     b10 = Val(Label6(10))
     b11 = Val(Label6(11))
     b12 = Val(Label6(12))
     b13 = Val(Label6(13))
     b14 = Val(Label6(14))
     b15 = Val(Label6(15))
     b16 = Val(Label6(16))
    
    bt = b0 + b1 + b2 + b3 + b4 + b5 + b6 + b7 + b8 + b9 + b10 + b11 + b12 + b13 + b14 + b15 + b16
     Label10.Caption = bt
    
    c0 = Val(Label7(0))
    c1 = Val(Label7(1))
    c2 = Val(Label7(2))
    c3 = Val(Label7(3))
    c4 = Val(Label7(4))
    c5 = Val(Label7(5))
    c6 = Val(Label7(6))
    c7 = Val(Label7(7))
    c8 = Val(Label7(8))
    c9 = Val(Label7(9))
    c10 = Val(Label7(10))
    c11 = Val(Label7(11))
    c12 = Val(Label7(12))
    c13 = Val(Label7(13))
    c14 = Val(Label7(14))
    c15 = Val(Label7(15))
    c16 = Val(Label7(16))
    
    ct = c0 + c1 + c2 + c3 + c4 + c5 + c6 + c7 + c8 + c9 + c10 + c11 + c12 + c13 + c14 + c15 + c16
     Label11.Caption = ct
    
    d0 = Val(Label8(0))
    d1 = Val(Label8(1))
    d2 = Val(Label8(2))
    d3 = Val(Label8(3))
    d4 = Val(Label8(4))
    d5 = Val(Label8(5))
    d6 = Val(Label8(6))
    d7 = Val(Label8(7))
    d8 = Val(Label8(8))
    d9 = Val(Label8(9))
    d10 = Val(Label8(10))
    d11 = Val(Label8(11))
    d12 = Val(Label8(12))
    d13 = Val(Label8(13))
    d14 = Val(Label8(14))
    d15 = Val(Label8(15))
    d16 = Val(Label8(16))
    
    dt = d0 + d1 + d2 + d3 + d4 + d5 + d6 + d7 + d8 + d9 + d10 + d11 + d12 + d13 + d14 + d15 + d16
     Label12.Caption = dt
     
     
  frm6.Show
End If


End Sub

Private Sub Timer7_Timer()

If Image1.Visible = False And Image2.Visible = True And Image3.Visible = True And Image4.Visible = True Then
pollob1

ElseIf Image2.Visible = False And Image1.Visible = True And Image3.Visible = True And Image4.Visible = True Then
pollob1

ElseIf Image3.Visible = False And Image2.Visible = True And Image1.Visible = True And Image4.Visible = True Then
pollob1

ElseIf Image4.Visible = False And Image2.Visible = True And Image3.Visible = True And Image1.Visible = True Then
pollob1

Else
End If

If Image1.Visible = False And Image2.Visible = False Then
pollob2

ElseIf Image1.Visible = False And Image3.Visible = False Then
pollob2

ElseIf Image1.Visible = False And Image4.Visible = False Then
pollob2

ElseIf Image2.Visible = False And Image1.Visible = False Then
pollob2

ElseIf Image2.Visible = False And Image3.Visible = False Then
pollob2

ElseIf Image2.Visible = False And Image4.Visible = False Then
pollob2

ElseIf Image3.Visible = False And Image1.Visible = False Then
pollob2

ElseIf Image3.Visible = False And Image2.Visible = False Then
pollob2

ElseIf Image3.Visible = False And Image4.Visible = False Then
pollob2

ElseIf Image4.Visible = False And Image1.Visible = False Then
pollob2

ElseIf Image4.Visible = False And Image2.Visible = False Then
pollob2

ElseIf Image4.Visible = False And Image3.Visible = False Then
pollob2

Else
End If

If Image1.Visible = True And Image2.Visible = False And Image3.Visible = False And Image4.Visible = False Then
pollob3

ElseIf Image1.Visible = False And Image2.Visible = True And Image3.Visible = False And Image4.Visible = False Then
pollob3

ElseIf Image1.Visible = False And Image2.Visible = False And Image3.Visible = True And Image4.Visible = False Then
pollob3

ElseIf Image1.Visible = False And Image2.Visible = False And Image3.Visible = False And Image4.Visible = True Then
pollob3

Else
End If


If Image1.Visible = False And Image2.Visible = False And Image3.Visible = False And Image4.Visible = False Then
pollob4

Else
End If

End Sub


Private Sub Timer8_Timer()
Label29.Left = Label29.Left - 10
Label30.Left = Label30.Left + 10
If Label30.Left >= Me.Width Then
Timer9.Enabled = True
Timer8.Enabled = False
End If

End Sub

Private Sub Timer9_Timer()
Label29.Left = Label29.Left + 10
Label30.Left = Label30.Left - 10
If Label29.Left >= Me.Width Then
Timer9.Enabled = False
Timer8.Enabled = True
End If
End Sub
