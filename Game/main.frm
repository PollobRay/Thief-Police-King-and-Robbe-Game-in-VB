VERSION 5.00
Begin VB.Form main 
   Caption         =   "Thief Police King & Robber"
   ClientHeight    =   8895
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14355
   Icon            =   "main.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "main.frx":0A02
   ScaleHeight     =   8895
   ScaleWidth      =   14355
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   5160
      Top             =   9480
   End
   Begin VB.Image Image4 
      Height          =   855
      Left            =   0
      Picture         =   "main.frx":B1D5E
      Stretch         =   -1  'True
      Top             =   0
      Width           =   855
   End
   Begin VB.Image Image3 
      Height          =   855
      Left            =   0
      Picture         =   "main.frx":B6622
      Stretch         =   -1  'True
      Top             =   9840
      Width           =   3735
   End
   Begin VB.Image Image1 
      Height          =   1815
      Left            =   9000
      Picture         =   "main.frx":BD3A8
      Stretch         =   -1  'True
      Top             =   8760
      Width           =   2415
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00400000&
      Caption         =   "Thief Police King And Robber"
      BeginProperty Font 
         Name            =   "Pristina"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0C0&
      Height          =   1335
      Left            =   4560
      TabIndex        =   0
      Top             =   360
      Width           =   11895
   End
   Begin VB.Image Image2 
      Height          =   11055
      Left            =   0
      Picture         =   "main.frx":C0FFE
      Stretch         =   -1  'True
      Top             =   -120
      Width           =   20655
   End
End
Attribute VB_Name = "main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Image1_Click()
frm2.Show
Unload Me
End Sub

Private Sub Image4_Click()
frm4.Show
End Sub

Private Sub Timer1_Timer()
Image3.Left = Image3.Left + 10


If Image3.Left >= 5000 Then
Image3.Left = -Image3.Width

End If
End Sub
