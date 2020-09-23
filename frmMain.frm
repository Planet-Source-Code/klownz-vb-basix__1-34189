VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "KLoWnZ Project for stupid people"
   ClientHeight    =   1500
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4920
   LinkTopic       =   "Form1"
   ScaleHeight     =   1500
   ScaleWidth      =   4920
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command11 
      Caption         =   "About"
      Height          =   495
      Left            =   0
      TabIndex        =   11
      Top             =   960
      Width           =   1215
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Variables"
      Height          =   495
      Left            =   0
      TabIndex        =   10
      Top             =   360
      Width           =   1215
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Menus"
      Height          =   495
      Left            =   3720
      TabIndex        =   9
      Top             =   960
      Width           =   1215
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Special!"
      Height          =   495
      Left            =   2520
      TabIndex        =   8
      Top             =   960
      Width           =   1215
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Select Case"
      Height          =   495
      Left            =   1320
      TabIndex        =   7
      Top             =   960
      Width           =   1215
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Looping"
      Height          =   495
      Left            =   3720
      TabIndex        =   6
      Top             =   480
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      Caption         =   "ElseIf"
      Height          =   495
      Left            =   2520
      TabIndex        =   5
      Top             =   480
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "If/Then"
      Height          =   495
      Left            =   1320
      TabIndex        =   4
      Top             =   480
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Captions"
      Height          =   495
      Left            =   3720
      TabIndex        =   3
      Top             =   0
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Color"
      Height          =   495
      Left            =   2520
      TabIndex        =   2
      Top             =   0
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Timers"
      Height          =   495
      Left            =   1320
      TabIndex        =   1
      Top             =   0
      Width           =   1215
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   0
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   0
      Width           =   1215
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Must I comment this? Oh well
Private Sub Command1_Click()
    frmTimer.Visible = True 'Shows it
End Sub

Private Sub Command10_Click()
    frmVariables.Show 'Shows it
End Sub

Private Sub Command11_Click()
    frmAbout.Show 'Shows it
End Sub

Private Sub Command2_Click()
    frmColor.Show 'Shows it
End Sub

Private Sub Command3_Click()
    frmCaption.Show 'Shows it
End Sub

Private Sub Command4_Click()
    frmIfThen.Show 'Shows it
End Sub

Private Sub Command5_Click()
    frmElseIf.Show 'Shows it
End Sub

Private Sub Command6_Click()
    frmLoop.Show 'Shows it
End Sub

Private Sub Command7_Click()
    frmSelectCase.Show
End Sub

Private Sub Command8_Click()
    frmSpecial.Show 'Shows it
End Sub

Private Sub Command9_Click()
    frmMenu.Show 'Shows it
End Sub

Private Sub Form_Load()
    Combo1.AddItem "KLoWnZ" 'adds my name to combo1
    Combo1.AddItem "2k2" 'adds 2k2 to combo1
End Sub

