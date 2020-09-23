VERSION 5.00
Begin VB.Form frmCaption 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Change captions!"
   ClientHeight    =   2175
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2175
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Change labels caption!"
      Height          =   495
      Left            =   2880
      TabIndex        =   6
      Top             =   720
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   2880
      TabIndex        =   5
      Text            =   "KLoWnZ"
      Top             =   360
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Change this caption!"
      Height          =   495
      Left            =   0
      TabIndex        =   3
      Top             =   1320
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   0
      TabIndex        =   2
      Text            =   "KLoWnZ"
      Top             =   960
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Change it!"
      Height          =   375
      Left            =   720
      TabIndex        =   1
      Top             =   360
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   0
      TabIndex        =   0
      Text            =   "Change the forms crappy caption!"
      Top             =   0
      Width           =   2535
   End
   Begin VB.Line Line6 
      X1              =   4560
      X2              =   2640
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Line Line5 
      X1              =   2640
      X2              =   2640
      Y1              =   840
      Y2              =   1320
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Left            =   2760
      TabIndex        =   4
      Top             =   0
      Width           =   1215
   End
   Begin VB.Line Line4 
      X1              =   0
      X2              =   1320
      Y1              =   1920
      Y2              =   1920
   End
   Begin VB.Line Line3 
      X1              =   1320
      X2              =   1320
      Y1              =   840
      Y2              =   1920
   End
   Begin VB.Line Line2 
      X1              =   2640
      X2              =   2640
      Y1              =   0
      Y2              =   840
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   2640
      Y1              =   840
      Y2              =   840
   End
End
Attribute VB_Name = "frmCaption"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    frmCaption.Caption = Text1.Text 'The form's caption is going to be what text1's texts is.
End Sub

Private Sub Command2_Click()
    Command2.Caption = Text2.Text 'Command2's caption is goint to be text2's text.
End Sub

Private Sub Command3_Click()
    Label1.Caption = Text3.Text 'Label 1's caption is going to be text 3's text.
End Sub

Private Sub Command4_Click()
    frmCaption.Hide 'Hides frmCaption
    frmMain.Show 'Shows frmMain
End Sub

