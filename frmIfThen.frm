VERSION 5.00
Begin VB.Form frmIfThen 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "If/Then"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Back!"
      Height          =   375
      Left            =   0
      TabIndex        =   4
      Top             =   2880
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   1680
      TabIndex        =   2
      Text            =   "Password here"
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Check"
      Height          =   495
      Left            =   1680
      TabIndex        =   0
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Enter the correct password!"
      Height          =   495
      Left            =   1680
      TabIndex        =   3
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "This can be used in boring programs like this one...or fun games"
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   4695
   End
End
Attribute VB_Name = "frmIfThen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Rem Don't want the password case sensitive?
Rem Add 'Option Compare Text' here.
Private Sub Command1_Click()
    If Text1.Text = "KLoWnZ" Then 'if the text is KLoWnZ then a message box will pop up displaing my wesites name.
        MsgBox "klownz2k2.topcities.com" 'this is that messagebox.
        Label3.Caption = "Correct!" 'changes label3's caption to  'Correct!'.
    Else 'if the text does not say KLoWnZ the do this (look down)
        Label3.Caption = "Enter the correct password!" 'It changes label3's caption to...read it
    End If 'You must put this after every if statement
End Sub

Private Sub Command3_Click()
    frmMain.Show 'show the main form
    Unload Me 'unload this project!
End Sub
