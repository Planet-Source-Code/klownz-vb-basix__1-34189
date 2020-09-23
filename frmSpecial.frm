VERSION 5.00
Begin VB.Form frmSpecial 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ClipBoard Stuff..."
   ClientHeight    =   1680
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2850
   LinkTopic       =   "Form8"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1680
   ScaleWidth      =   2850
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Clear"
      Height          =   375
      Left            =   1320
      TabIndex        =   4
      Top             =   600
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Paste"
      Height          =   375
      Left            =   0
      TabIndex        =   3
      Top             =   840
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   0
      TabIndex        =   2
      Text            =   "Text2"
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Copy"
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   360
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   0
      TabIndex        =   0
      Text            =   "Hello"
      Top             =   0
      Width           =   1215
   End
End
Attribute VB_Name = "frmSpecial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Clipboard.SetText Text1.Text 'Copy's what is currently in text1
End Sub

Private Sub Command2_Click()
    Text2.Text = "" 'Clears text2's text or sets the value to 0
    Text2.SelText = Clipboard.GetText ' Pastes what you copied
End Sub

Private Sub Command3_Click()
    Clipboard.Clear 'If you have anything copied this will clear it.
End Sub
