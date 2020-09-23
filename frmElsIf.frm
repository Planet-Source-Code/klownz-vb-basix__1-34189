VERSION 5.00
Begin VB.Form frmElseIf 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ElseIf"
   ClientHeight    =   975
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   1725
   LinkTopic       =   "Form6"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   975
   ScaleWidth      =   1725
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "?"
      Height          =   495
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Check"
      Height          =   495
      Left            =   0
      TabIndex        =   1
      Top             =   480
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   480
      TabIndex        =   0
      Text            =   "Stuff"
      Top             =   0
      Width           =   1215
   End
End
Attribute VB_Name = "frmElseIf"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Add Option Compare text to kill case sensitivity
Private Sub Command1_Click()
    If Text1.Text = "KLoWnZ" Then 'When you click command 1, if text1's text value is KLoWnZ (capitals matter unless you put Option compare text at the top) a message box will pop up.
        MsgBox "klownz2k2.topcities.com" 'and this is that message box
    ElseIf Text1.Text = "Stuff" Then 'Elseif....I don't know why you can't just put if, but if you want to try it then go ahead, it won't work.
        MsgBox "Stuff" 'A...you guessed it, another messagebox!
    ElseIf Text1.Text = "Hi" Then 'If text1's text is Hi theen you are going to get a message box.
        MsgBox "Hello." 'And this is that message box.
    ElseIf Text1.Text = "lamer" Then 'lamer hahaha
        MsgBox "" 'a blank messagebox...
End Sub

Private Sub Command2_Click()
    frmElseIf.Hide 'hides the form can also use me.hide
    frmMain.Show 'shows the form
End Sub

Private Sub Command3_Click()
    MsgBox "Password is KLoWnZ, Stuff, Hi, and lamer." 'a messagebox that says stuff
    MsgBox "Remember, if you DO NOT want them to be case sensative then add 'Option Compare Text' at the top of the code window." 'a messagebox that says stuff
End Sub

