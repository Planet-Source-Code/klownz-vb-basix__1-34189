VERSION 5.00
Begin VB.Form frmVariables 
   Caption         =   "Variable Stuff"
   ClientHeight    =   1410
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3195
   LinkTopic       =   "Form1"
   ScaleHeight     =   1410
   ScaleWidth      =   3195
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "Multiply"
      Height          =   375
      Left            =   1200
      TabIndex        =   6
      Top             =   960
      Width           =   855
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Subtract"
      Height          =   375
      Left            =   1200
      TabIndex        =   5
      Top             =   480
      Width           =   855
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   2400
      TabIndex        =   4
      Text            =   "0"
      Top             =   120
      Width           =   855
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1200
      TabIndex        =   1
      Text            =   "0"
      Top             =   120
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   0
      TabIndex        =   0
      Text            =   "0"
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "="
      Height          =   255
      Left            =   2160
      TabIndex        =   3
      Top             =   120
      Width           =   135
   End
   Begin VB.Label Label1 
      Height          =   255
      Left            =   960
      TabIndex        =   2
      Top             =   120
      Width           =   135
   End
End
Attribute VB_Name = "frmVariables"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command3_Click()
    Dim Answer As Integer 'To keep this shit simple...Byte = small number Integer = fairly large number Long = big ass number
    Answer = Text1.Text - Text2.Text 'Our variable is text1 subtracted by text 2
    
    Text3.Text = Answer 'text3 is our answer
    Label1.Caption = "-" 'changes caption to -
End Sub

Private Sub Command4_Click()
    Dim Answer As Integer 'To keep this shit simple...Byte = small number Integer = fairly large number Long = big ass number
    Answer = Text1.Text * Text2.Text 'Our variable is text1 multiplied by text 2
    
    Text3.Text = Answer 'text3 is our answer
    Label1.Caption = "X" 'changes caption to X
End Sub
