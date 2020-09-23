VERSION 5.00
Begin VB.Form frmTimer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "form2"
   ClientHeight    =   1650
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1650
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "Go back!"
      Height          =   615
      Left            =   600
      TabIndex        =   7
      Top             =   480
      Width           =   735
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   2880
      TabIndex        =   5
      Text            =   "0"
      Top             =   0
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Stop"
      Height          =   495
      Left            =   2040
      TabIndex        =   4
      Top             =   1200
      Width           =   615
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Start"
      Height          =   195
      Left            =   1680
      TabIndex        =   2
      Top             =   960
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Set Interval"
      Height          =   255
      Left            =   1680
      TabIndex        =   1
      Top             =   360
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   480
      TabIndex        =   0
      Text            =   "Set Interval here."
      Top             =   0
      Width           =   1335
   End
   Begin VB.Timer Timer1 
      Left            =   0
      Top             =   0
   End
   Begin VB.Label Label3 
      Caption         =   "After started --->"
      Height          =   255
      Left            =   840
      TabIndex        =   6
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "<--After interval set."
      Height          =   375
      Left            =   3000
      TabIndex        =   3
      Top             =   960
      Width           =   1215
   End
End
Attribute VB_Name = "frmTimer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Timer1.Interval = Text1.Text 'The timer's interval is what text1's text is, please note that the interval is a number..lol
    MsgBox "Interval set to " & Text1.Text & " !" ''a messagebox pops up saying what the interval is set to.
End Sub

Private Sub Command2_Click()
    Timer1.Enabled = True 'enables timer1
End Sub

Private Sub Command3_Click()
    Timer1.Enabled = False 'shuts off timer1, sets the value to zero.
    MsgBox "You killed the timer!" 'a messagebox
    MsgBox "What did the timer ever do to you?" 'a messagebox
End Sub

Private Sub Form_Load()
    Timer1.Enabled = False 'the damn annoying timer will start going for some odd reason...so if you have a timer make sure you put this to be sure it is NOT currently active.
End Sub

Private Sub Label2_Click()
    frmTimer.Hide 'hide it
    frmMain.Show 'show it
End Sub

Private Sub Timer1_Timer()
    Text3.Text = Text3.Text + 1 'text3's text will keep on adding one until the timer is shut off.
End Sub
