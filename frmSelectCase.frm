VERSION 5.00
Begin VB.Form frmSelectCase 
   Caption         =   "I like to eat cherries"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   0
      TabIndex        =   1
      Text            =   "I like cherries"
      Top             =   0
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Check It"
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Passwords: Lamer, Cherries, Green"
      Height          =   615
      Left            =   1320
      TabIndex        =   2
      Top             =   0
      Width           =   1215
   End
End
Attribute VB_Name = "frmSelectCase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Select Case Text1.Text 'The case is text1's text
    
    Case "Lamer" 'if it says Lamer then...
        MsgBox "Hey lamer." 'this pops up
    Case "Cherries" 'if it says Cherries then...
        MsgBox "Yummy" 'this pops up
    Case "Green" 'if it says Green then...
        MsgBox "Green is cool." 'this pops up
    Case Else 'if it is none of them, then
        MsgBox "Are you stupid? The passwords are right there..." 'this pops up
    End Select 'just like end if...but end select...must have this
End Sub
