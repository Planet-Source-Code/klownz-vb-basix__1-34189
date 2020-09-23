VERSION 5.00
Begin VB.Form frmMenu 
   Caption         =   "There are...Menus on here!"
   ClientHeight    =   1440
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   1440
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Label Label1 
      Caption         =   "Right Click form while in design mode and then click menu editor."
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   4455
   End
   Begin VB.Menu mnuKLoWnZ 
      Caption         =   "KLoWnZ"
      Begin VB.Menu mnuWebsite 
         Caption         =   "Websites"
         Begin VB.Menu mnuMysite 
            Caption         =   "My website"
         End
         Begin VB.Menu mnuGPS 
            Caption         =   "Good prog site"
         End
         Begin VB.Menu mnuFreeProgz 
            Caption         =   "You wish it was still here"
         End
      End
   End
   Begin VB.Menu mnuAOL 
      Caption         =   "AOL"
      Begin VB.Menu mnuis 
         Caption         =   "is"
      End
      Begin VB.Menu mnuVery 
         Caption         =   "very"
         Begin VB.Menu mnuBoring 
            Caption         =   "Boring"
         End
      End
      Begin VB.Menu mnuI 
         Caption         =   "I"
         Begin VB.Menu mnulike 
            Caption         =   "like"
         End
         Begin VB.Menu mnuOrangesSuck 
            Caption         =   "Cherries"
         End
      End
   End
End
Attribute VB_Name = "frmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub mnuFreeProgz_Click()
    MsgBox "www.freeprogz.com" 'A messagebox
End Sub

Private Sub mnuGPS_Click()
    MsgBox "www.lenshell.com" 'A messagebox
End Sub

Private Sub mnuMysite_Click()
    MsgBox "klownz2k2.topcities.com" 'A messagebox
End Sub
