VERSION 5.00
Begin VB.Form frmSelSudoku 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7065
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10995
   LinkTopic       =   "Form1"
   Picture         =   "frmSelSudoku.frx":0000
   ScaleHeight     =   7065
   ScaleWidth      =   10995
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Image Image1 
      Height          =   4575
      Left            =   720
      Stretch         =   -1  'True
      Top             =   1200
      Width           =   5055
   End
End
Attribute VB_Name = "frmSelSudoku"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
MakeWindow Me
End Sub

