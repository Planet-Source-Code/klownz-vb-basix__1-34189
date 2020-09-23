VERSION 5.00
Begin VB.Form frmLoop 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form7"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Stop!"
      Height          =   495
      Left            =   1680
      TabIndex        =   8
      Top             =   840
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Loop"
      Height          =   495
      Left            =   1680
      TabIndex        =   1
      Top             =   360
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1680
      TabIndex        =   0
      Text            =   "KLoWnZ"
      Top             =   0
      Width           =   1215
   End
   Begin VB.Label Label6 
      BackColor       =   &H000000FF&
      Height          =   495
      Left            =   3360
      TabIndex        =   7
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Label Label5 
      BackColor       =   &H80000007&
      Height          =   495
      Left            =   3360
      TabIndex        =   6
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   3360
      TabIndex        =   5
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFF00&
      Height          =   495
      Left            =   0
      TabIndex        =   4
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000FF00&
      Height          =   495
      Left            =   0
      TabIndex        =   3
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FF0000&
      Height          =   495
      Left            =   0
      TabIndex        =   2
      Top             =   1440
      Width           =   1215
   End
End
Attribute VB_Name = "frmLoop"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
   Do 'Ok, this is the hardest line...it means....Do.
        If Text1.Text = "KLoWnZ" Then 'If the text says my name
            Label6.Visible = False 'Hide label 6
            Label1.Visible = True 'show label1
                Pause 1 'Pauses
            Label1.Visible = False 'Hides label1
            Label2.Visible = True 'shows label2
                Pause 1 'Pauses
            Label2.Visible = False 'hide label2
            Label3.Visible = True 'show label3
                Pause 1 'Pauses
            Label3.Visible = False 'hide label3
            Label4.Visible = True 'showlabel4
                Pause 1 'Pauses
            Label4.Visible = False 'hide label4
            Label5.Visible = True 'show label 5
                Pause 1 'Pauses
            Label5.Visible = False 'hide label5
            Label6.Visible = True 'show label 6
    End If 'must put this after every if statement
    Loop Until Text1.Text = "Looping Stopped" 'Keep looping until text1's text says 'Looping Stopped'
End Sub

Private Sub Command2_Click()
    Text1.Text = "Looping Stopped" 'Stop the looping...but the loop has to finish it's loop, I don't know why and it annoys me =\
End Sub

Private Sub Form_Load()
    Label1.Visible = False 'hides it
    Label2.Visible = False 'hides it
    Label3.Visible = False 'hides it
    Label4.Visible = False 'hides it
    Label5.Visible = False 'hides it
    Label6.Visible = False 'hides it
End Sub
