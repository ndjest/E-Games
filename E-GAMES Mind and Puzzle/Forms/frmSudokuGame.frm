VERSION 5.00
Begin VB.Form frmSudokuGame 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2370
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7890
   LinkTopic       =   "Form1"
   ScaleHeight     =   2370
   ScaleWidth      =   7890
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Congratulation, you just solve the puzzle!"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1335
      Left            =   0
      TabIndex        =   1
      Top             =   120
      Width           =   7815
   End
   Begin VB.Label lbl1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H003576E1&
      Height          =   375
      Left            =   3240
      MouseIcon       =   "frmSudokuGame.frx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   1680
      Width           =   1575
   End
   Begin VB.Image Image3 
      Height          =   1035
      Left            =   2880
      Picture         =   "frmSudokuGame.frx":030A
      Top             =   1320
      Width           =   2250
   End
End
Attribute VB_Name = "frmSudokuGame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub lbl1_Click()
Unload Me
End Sub
