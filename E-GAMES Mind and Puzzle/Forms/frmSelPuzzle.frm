VERSION 5.00
Begin VB.Form frmSelPuzzle 
   BorderStyle     =   0  'None
   Caption         =   "&H00C00000&"
   ClientHeight    =   7080
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12360
   LinkTopic       =   "Form1"
   Picture         =   "frmSelPuzzle.frx":0000
   ScaleHeight     =   7080
   ScaleWidth      =   12360
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label lblPercent 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "SUDOKU"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H007C2E10&
      Height          =   495
      Left            =   6600
      TabIndex        =   2
      Top             =   960
      Width           =   1815
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Back"
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
      Left            =   480
      MouseIcon       =   "frmSelPuzzle.frx":15D62
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   6120
      Width           =   1575
   End
   Begin VB.Image Image5 
      Height          =   1035
      Left            =   120
      Picture         =   "frmSelPuzzle.frx":1606C
      Top             =   5760
      Width           =   2250
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Mini- Games"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   855
      Left            =   2280
      TabIndex        =   0
      Top             =   5760
      Width           =   10455
   End
   Begin VB.Image Image3 
      Height          =   1815
      Left            =   6480
      Stretch         =   -1  'True
      Top             =   240
      Width           =   1935
   End
   Begin VB.Image Image2 
      Height          =   1935
      Left            =   3120
      Stretch         =   -1  'True
      Top             =   720
      Width           =   2175
   End
   Begin VB.Image Image1 
      Height          =   1815
      Left            =   9720
      Stretch         =   -1  'True
      Top             =   720
      Width           =   1815
   End
   Begin VB.Image Image4 
      Height          =   7350
      Left            =   -120
      Picture         =   "frmSelPuzzle.frx":16302
      Top             =   -120
      Width           =   12570
   End
End
Attribute VB_Name = "frmSelPuzzle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
MakeWindow Me
End Sub

Private Sub Image1_Click()
fGame.Show
Unload Me
End Sub

Private Sub Image2_Click()
frmSelPicture.Show
Unload Me
End Sub

Private Sub Image3_Click()
FMain.Show
Unload Me
End Sub

Private Sub Label5_Click()
frmMain.Show
Unload Me
End Sub

Private Sub lblPercent_Click()
FMain.Show
Unload Me
End Sub
