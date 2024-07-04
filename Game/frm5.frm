VERSION 5.00
Begin VB.Form frm5 
   BackColor       =   &H00404000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5595
   ClientLeft      =   7365
   ClientTop       =   315
   ClientWidth     =   10305
   DrawStyle       =   1  'Dash
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5595
   ScaleWidth      =   10305
   ShowInTaskbar   =   0   'False
   Begin VB.OptionButton Option2 
      BackColor       =   &H8000000D&
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
      Left            =   2040
      TabIndex        =   2
      Top             =   3720
      Width           =   6015
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H8000000D&
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
      Left            =   2040
      TabIndex        =   1
      Top             =   3000
      Width           =   6015
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00808000&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C0C0&
      Height          =   1095
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   10335
   End
   Begin VB.Label Label2 
      Height          =   495
      Left            =   720
      TabIndex        =   3
      Top             =   3600
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   4560
      Picture         =   "frm5.frx":0000
      Stretch         =   -1  'True
      Top             =   4560
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackColor       =   &H00808000&
      Caption         =   "Who is Thief .??"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   72
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   2055
      Left            =   0
      TabIndex        =   0
      Top             =   960
      Width           =   10335
   End
End
Attribute VB_Name = "frm5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim b, c As Integer
Private Sub Form_Load()
Dim a As Integer
Label2.Caption = frm3.Label19.Caption
Label3.Caption = frm3.Label15.Caption
a = Val(frm3.Label19)


b = a Mod 2

If b = 0 Then

Option1.Caption = frm3.Label16.Caption
Option2.Caption = frm3.Label17.Caption

Else
Option1.Caption = frm3.Label17.Caption
Option2.Caption = frm3.Label16.Caption

End If


End Sub

Private Sub Image1_Click()

frm3.Timer1.Enabled = True
frm3.Timer2.Enabled = True
frm3.Timer3.Enabled = True
frm3.Image6.Visible = True
frm3.Label16.Visible = True
frm3.Label17.Visible = True

frm3.Label25.Visible = True
frm3.Label26.Visible = True
frm3.Label27.Visible = True
frm3.Label28.Visible = True
frm3.Timer3.Enabled = True

c = Val(frm3.Label20)

'/////////////

If frm3.Label21.Caption = 1 Then
frm3.Label5(c).Caption = 100

ElseIf frm3.Label21.Caption = 2 Then
frm3.Label6(c).Caption = 100

ElseIf frm3.Label21.Caption = 3 Then
frm3.Label7(c).Caption = 100

ElseIf frm3.Label21.Caption = 4 Then
frm3.Label8(c).Caption = 100

Else
End If
'////////////

If frm3.Label24.Caption = 1 Then
frm3.Label5(c).Caption = 40

ElseIf frm3.Label24.Caption = 2 Then
frm3.Label6(c).Caption = 40

ElseIf frm3.Label24.Caption = 3 Then
frm3.Label7(c).Caption = 40

ElseIf frm3.Label24.Caption = 4 Then
frm3.Label8(c).Caption = 40

Else
End If
'////////////


If b = 0 Then

     If Option1 = True Then
    
         If frm3.Label22.Caption = 1 Then
            frm3.Label5(c).Caption = 80

         ElseIf frm3.Label22.Caption = 2 Then
             frm3.Label6(c).Caption = 80

        ElseIf frm3.Label22.Caption = 3 Then
             frm3.Label7(c).Caption = 80

        ElseIf frm3.Label22.Caption = 4 Then
             frm3.Label8(c).Caption = 80

         Else
         
         End If
    
    '\\\\\\\\\\\\\\\\
    
        If frm3.Label23.Caption = 1 Then
            frm3.Label5(c).Caption = 0

        ElseIf frm3.Label23.Caption = 2 Then
            frm3.Label6(c).Caption = 0

        ElseIf frm3.Label23.Caption = 3 Then
             frm3.Label7(c).Caption = 0

        ElseIf frm3.Label23.Caption = 4 Then
            frm3.Label8(c).Caption = 0

        Else
        
        End If
    
    
    ElseIf Option2 = True Then


        If frm3.Label23.Caption = 1 Then
            frm3.Label5(c).Caption = 60 'uuiyuyu

        ElseIf frm3.Label23.Caption = 2 Then
            frm3.Label6(c).Caption = 60

        ElseIf frm3.Label23.Caption = 3 Then
            frm3.Label7(c).Caption = 60

        ElseIf frm3.Label23.Caption = 4 Then
            frm3.Label8(c).Caption = 60

        Else
        
        End If

    
    
        If frm3.Label22.Caption = 1 Then
            frm3.Label5(c).Caption = 0

        ElseIf frm3.Label22.Caption = 2 Then
            frm3.Label6(c).Caption = 0

        ElseIf frm3.Label22.Caption = 3 Then
            frm3.Label7(c).Caption = 0

        ElseIf frm3.Label22.Caption = 4 Then
            frm3.Label8(c).Caption = 0

        Else
        
        End If
    
    
    
    Else
    End If
'end modiule


'//////////////////////////


Else
    
    If Option2 = True Then


        If frm3.Label22.Caption = 1 Then
            frm3.Label5(c).Caption = 80

        ElseIf frm3.Label22.Caption = 2 Then
            frm3.Label6(c).Caption = 80

        ElseIf frm3.Label22.Caption = 3 Then
            frm3.Label7(c).Caption = 80

        ElseIf frm3.Label22.Caption = 4 Then
            frm3.Label8(c).Caption = 80

        Else
        
        End If
    
    
        If frm3.Label23.Caption = 1 Then
            frm3.Label5(c).Caption = 0

        ElseIf frm3.Label23.Caption = 2 Then
            frm3.Label6(c).Caption = 0

        ElseIf frm3.Label23.Caption = 3 Then
            frm3.Label7(c).Caption = 0

        ElseIf frm3.Label23.Caption = 4 Then
            frm3.Label8(c).Caption = 0

        Else
        
        End If
'////
    ElseIf Option1 = True Then
    

        If frm3.Label23.Caption = 1 Then
            frm3.Label5(c).Caption = 60

        ElseIf frm3.Label23.Caption = 2 Then
            frm3.Label6(c).Caption = 60

        ElseIf frm3.Label23.Caption = 3 Then
            frm3.Label7(c).Caption = 60

        ElseIf frm3.Label23.Caption = 4 Then
            frm3.Label8(c).Caption = 60

        Else
        
        End If
    
    
        If frm3.Label22.Caption = 1 Then
            frm3.Label5(c).Caption = 0

        ElseIf frm3.Label22.Caption = 2 Then
            frm3.Label6(c).Caption = 0

        ElseIf frm3.Label22.Caption = 3 Then
            frm3.Label7(c).Caption = 0

        ElseIf frm3.Label22.Caption = 4 Then
            frm3.Label8(c).Caption = 0

        Else
        
        End If
    
    
    Else
    
    End If
'end modioule



End If

Unload Me
End Sub
