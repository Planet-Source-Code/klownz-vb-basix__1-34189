VERSION 5.00
Begin VB.Form frmColor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Color Form!"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command14 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Remember"
      Height          =   375
      Left            =   2760
      TabIndex        =   15
      Top             =   960
      Width           =   1215
   End
   Begin VB.CommandButton Command12 
      Caption         =   "and their text!"
      Height          =   375
      Left            =   1080
      TabIndex        =   14
      Top             =   2640
      Width           =   1215
   End
   Begin VB.CommandButton Command11 
      Caption         =   "Text Boxes"
      Height          =   375
      Left            =   0
      TabIndex        =   13
      Top             =   2640
      Width           =   1095
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Labels Text"
      Height          =   375
      Left            =   2880
      TabIndex        =   12
      Top             =   2280
      Width           =   975
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Labels"
      Height          =   375
      Left            =   2280
      TabIndex        =   11
      Top             =   2280
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   2280
      TabIndex        =   9
      Text            =   "klownz2k2.topcities.com"
      Top             =   2640
      Width           =   1935
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Green"
      Height          =   375
      Left            =   840
      TabIndex        =   8
      Top             =   1320
      Width           =   615
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Magenta"
      Height          =   375
      Left            =   0
      TabIndex        =   7
      Top             =   1320
      Width           =   855
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Yellow"
      Height          =   375
      Left            =   720
      TabIndex        =   6
      Top             =   960
      Width           =   735
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Blue"
      Height          =   375
      Left            =   720
      TabIndex        =   5
      Top             =   600
      Width           =   735
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Red"
      Height          =   375
      Left            =   720
      TabIndex        =   4
      Top             =   240
      Width           =   735
   End
   Begin VB.CommandButton Command3 
      Caption         =   "White"
      Height          =   375
      Left            =   0
      TabIndex        =   3
      Top             =   960
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cyan"
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   600
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Black"
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   240
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "Note: This can also be used on"
      Height          =   255
      Left            =   0
      TabIndex        =   10
      Top             =   2400
      Width           =   2415
   End
   Begin VB.Line Line2 
      X1              =   0
      X2              =   1560
      Y1              =   1800
      Y2              =   1800
   End
   Begin VB.Line Line1 
      X1              =   1560
      X2              =   1560
      Y1              =   1800
      Y2              =   0
   End
   Begin VB.Label Label1 
      Caption         =   "Form Color:"
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   855
   End
End
Attribute VB_Name = "frmColor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    frmColor.BackColor = vbBlack 'Sets the forms backround color to Black.
End Sub

Private Sub Command10_Click()
    Label2.ForeColor = vbGreen 'Sets Label 2's Text color to green.
End Sub

Private Sub Command11_Click()
    Text1.BackColor = vbMagenta 'Sets the backround color of text1 to magenta.
End Sub

Private Sub Command12_Click()
    Text1.ForeColor = vbRed 'Forcolor is the texts color.
End Sub

Private Sub Command13_Click()
    frmColor.Hide 'Hides form
    frmMain.Show 'Shows form
End Sub

Private Sub Command14_Click()
    MsgBox "You do NOT have to use vbRed, blue ect... you can also use HEX codes." 'Pops up a lame looking message box.
End Sub

Private Sub Command2_Click()
    frmColor.BackColor = vbCyan 'backcolor is the backround color which this sets the forms backround color to cyan.
End Sub

Private Sub Command3_Click()
    frmColor.BackColor = vbWhite '&H00FFFFFF& HEX code for white, can be used instead of vbWhite
End Sub

Private Sub Command4_Click()
    frmColor.BackColor = vbRed '&H000000FF& HEX code for red, could be used instead of vbRed.
End Sub

Private Sub Command5_Click()
    frmColor.BackColor = vbBlue 'Sets forms backround color to...blue!
End Sub

Private Sub Command6_Click()
    frmColor.BackColor = vbMagenta 'Sets forms backround color to...Magenta.
End Sub

Private Sub Command7_Click()
    frmColor.BackColor = vbGreen 'Sets forms backround color to...green.
End Sub

Private Sub Command8_Click()
    frmColor.BackColor = vbYellow 'Sets forms backround color to...yellow.
End Sub

Private Sub Command9_Click()
    Label2.BackColor = vbBlack 'Sets the forms backround color to black.
End Sub
