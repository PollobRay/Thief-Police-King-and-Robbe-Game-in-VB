VERSION 5.00
Begin VB.Form frm2 
   BackColor       =   &H00404000&
   Caption         =   "Name"
   ClientHeight    =   10215
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15945
   Icon            =   "name.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10215
   ScaleWidth      =   15945
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer2 
      Interval        =   1
      Left            =   17520
      Top             =   3720
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   1080
      Top             =   3000
   End
   Begin VB.TextBox Text4 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   735
      Left            =   10680
      MaxLength       =   8
      TabIndex        =   7
      Top             =   6360
      Width           =   5175
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   735
      Left            =   10680
      MaxLength       =   8
      TabIndex        =   6
      Top             =   5400
      Width           =   5175
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   735
      Left            =   10680
      MaxLength       =   8
      TabIndex        =   5
      Top             =   4440
      Width           =   5175
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   735
      Left            =   10680
      MaxLength       =   8
      TabIndex        =   4
      Top             =   3480
      Width           =   5175
   End
   Begin VB.Image Image4 
      Height          =   855
      Left            =   120
      Picture         =   "name.frx":0A02
      Stretch         =   -1  'True
      Top             =   120
      Width           =   855
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   16800
      Picture         =   "name.frx":52C6
      Stretch         =   -1  'True
      Top             =   10080
      Width           =   3615
   End
   Begin VB.Image Image2 
      Height          =   975
      Left            =   16800
      Picture         =   "name.frx":BBCB
      Stretch         =   -1  'True
      Top             =   0
      Width           =   3615
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Please Enter 4 Player Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   735
      Left            =   5400
      TabIndex        =   8
      Top             =   7440
      Visible         =   0   'False
      Width           =   8175
   End
   Begin VB.Image Image1 
      Height          =   1215
      Left            =   9360
      Picture         =   "name.frx":121AB
      Stretch         =   -1  'True
      Top             =   8760
      Width           =   1575
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Enter Fourth Player Name :"
      BeginProperty Font 
         Name            =   "Poor Richard"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   975
      Left            =   2760
      TabIndex        =   3
      Top             =   6240
      Width           =   8055
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Enter Third Player Name :"
      BeginProperty Font 
         Name            =   "Poor Richard"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   975
      Left            =   3120
      TabIndex        =   2
      Top             =   5280
      Width           =   7335
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Enter Second Player Name :"
      BeginProperty Font 
         Name            =   "Poor Richard"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   975
      Left            =   2640
      TabIndex        =   1
      Top             =   4320
      Width           =   7935
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Enter First Player Name :"
      BeginProperty Font 
         Name            =   "Poor Richard"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   975
      Left            =   3480
      TabIndex        =   0
      Top             =   3360
      Width           =   7335
   End
End
Attribute VB_Name = "frm2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Image1_Click()

If Text1.Text <> "" And Text2.Text <> "" And Text3.Text <> "" And Text4.Text <> "" Then

frm3.Show
Unload Me


Else
Label5.Visible = True
Timer1.Enabled = False
End If
End Sub

Private Sub Image4_Click()
frm4.Show
End Sub

Private Sub Timer1_Timer()

If Text1.Text <> "" Then
Label5.Visible = True
ElseIf Text2.Text <> "" Then
Label5.Visible = True
ElseIf Text3.Text <> "" Then
Label5.Visible = True
ElseIf Text4.Text <> "" Then
Label5.Visible = True
Else
Label5.Visible = False
End If
End Sub

Private Sub Timer2_Timer()

If Image2.Top = 4440 Then
Image2.Top = Image2.Top + 0
Else
Image2.Top = Image2.Top + 1
End If
If Image3.Top = 5400 Then
Image3.Top = Image3.Top - 0
Else
Image3.Top = Image3.Top - 1
End If
End Sub
