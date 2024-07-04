VERSION 5.00
Begin VB.Form frm6 
   BackColor       =   &H00400040&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Position"
   ClientHeight    =   5415
   ClientLeft      =   690
   ClientTop       =   1995
   ClientWidth     =   10395
   Icon            =   "frm6.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5415
   ScaleWidth      =   10395
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   9000
      Top             =   3960
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   9600
      Top             =   3960
   End
   Begin VB.Image Image2 
      Height          =   855
      Left            =   3600
      Picture         =   "frm6.frx":0A02
      Stretch         =   -1  'True
      Top             =   4560
      Width           =   1215
   End
   Begin VB.Label Label11 
      BackColor       =   &H00C0FFC0&
      Caption         =   "P . C . Roy"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   8760
      TabIndex        =   10
      Top             =   4920
      Width           =   1695
   End
   Begin VB.Label Label10 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Developer..."
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7200
      TabIndex        =   9
      Top             =   4440
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   855
      Left            =   5040
      Picture         =   "frm6.frx":3322
      Stretch         =   -1  'True
      Top             =   4560
      Width           =   1335
   End
   Begin VB.Label Label9 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3960
      TabIndex        =   8
      Top             =   3720
      Width           =   4815
   End
   Begin VB.Label Label8 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3960
      TabIndex        =   7
      Top             =   2880
      Width           =   4815
   End
   Begin VB.Label Label7 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3960
      TabIndex        =   6
      Top             =   2040
      Width           =   4815
   End
   Begin VB.Label Label6 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3960
      TabIndex        =   5
      Top             =   1200
      Width           =   4815
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "ROBBER :"
      BeginProperty Font 
         Name            =   "Perpetua Titling MT"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   735
      Left            =   120
      TabIndex        =   4
      Top             =   3600
      Width           =   3855
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "POLICE   :"
      BeginProperty Font 
         Name            =   "Perpetua Titling MT"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   735
      Left            =   120
      TabIndex        =   3
      Top             =   1920
      Width           =   3975
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "THIEF     :"
      BeginProperty Font 
         Name            =   "Perpetua Titling MT"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   735
      Left            =   120
      TabIndex        =   2
      Top             =   2760
      Width           =   3735
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "KING      :"
      BeginProperty Font 
         Name            =   "Perpetua Titling MT"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   735
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   4095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0FF&
      Caption         =   "Position of Player"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   855
      Left            =   1680
      TabIndex        =   0
      Top             =   120
      Width           =   7095
   End
End
Attribute VB_Name = "frm6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Dim a, b, c, d As Integer

a = Val(frm3.Label9)
b = Val(frm3.Label10)
c = Val(frm3.Label11)
d = Val(frm3.Label12)

If a > b And a > c And a > d Then
Label6.Caption = frm3.Label1.Caption

ElseIf b > a And b > c And b > d Then
Label6.Caption = frm3.Label2.Caption

ElseIf c > a And c > b And c > d Then
Label6.Caption = frm3.Label3.Caption

Else
Label6.Caption = frm3.Label4.Caption
End If

If a < b And a > c And a > d Then
Label7.Caption = frm3.Label1.Caption
ElseIf a < c And a > b And a > d Then
Label7.Caption = frm3.Label1.Caption
ElseIf a < d And a > c And a > b Then
Label7.Caption = frm3.Label1.Caption

ElseIf b < a And b > c And b > d Then
Label7.Caption = frm3.Label2.Caption
ElseIf b < c And b > a And b > d Then
Label7.Caption = frm3.Label2.Caption
ElseIf b < d And b > c And b > a Then
Label7.Caption = frm3.Label2.Caption

ElseIf c < a And c > a And c > d Then
Label7.Caption = frm3.Label3.Caption
ElseIf c < b And c > a And c > d Then
Label7.Caption = frm3.Label3.Caption
ElseIf c < d And c > b And c > a Then
Label7.Caption = frm3.Label3.Caption

Else
Label7.Caption = frm3.Label4.Caption
End If

If a < b And a < c And a > d Then
Label8.Caption = frm3.Label1.Caption
ElseIf a < b And a < d And a > c Then
Label8.Caption = frm3.Label1.Caption
ElseIf a < c And a < d And a > b Then
Label8.Caption = frm3.Label1.Caption

ElseIf b < a And b < c And b > d Then
Label8.Caption = frm3.Label2.Caption
ElseIf b < a And b < d And b > c Then
Label8.Caption = frm3.Label2.Caption
ElseIf b < c And b < d And b > a Then
Label8.Caption = frm3.Label2.Caption

ElseIf c < a And c < b And c > d Then
Label8.Caption = frm3.Label3.Caption
ElseIf c < a And c < d And c > b Then
Label8.Caption = frm3.Label3.Caption
ElseIf c < b And c < d And c > a Then
Label8.Caption = frm3.Label3.Caption

Else
Label8.Caption = frm3.Label4.Caption
End If


If a < b And a < c And a < d Then
Label9.Caption = frm3.Label1.Caption
ElseIf b < a And b < c And b < d Then
Label9.Caption = frm3.Label2.Caption
ElseIf c < a And c < b And c < d Then
Label9.Caption = frm3.Label3.Caption
Else
Label9.Caption = frm3.Label4.Caption
End If
End Sub

Private Sub Image1_Click()
main.Show
Unload Me
Unload frm3

End Sub


Private Sub Image2_Click()
frm2.Show
Unload Me
Unload frm3
End Sub

Private Sub Timer1_Timer()
Label11.Left = Label11.Left - 1
If Label11.Left <= 6360 Then
Timer1.Enabled = False
Timer2.Enabled = True
End If
End Sub

Private Sub Timer2_Timer()
Label11.Left = Label11.Left + 1
If Label11.Left >= 8760 Then
Timer1.Enabled = False
Timer2.Enabled = True
End If
End Sub


