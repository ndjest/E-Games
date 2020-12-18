VERSION 5.00
Begin VB.Form frmInstruction3 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7125
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12315
   LinkTopic       =   "Form1"
   ScaleHeight     =   7125
   ScaleWidth      =   12315
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmInstruction3.frx":0000
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H007C2E10&
      Height          =   3975
      Left            =   2040
      TabIndex        =   1
      Top             =   1200
      Width           =   8175
   End
   Begin VB.Label Label5 
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
      Left            =   10320
      MouseIcon       =   "frmInstruction3.frx":0110
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   6000
      Width           =   1575
   End
   Begin VB.Image Image1 
      Height          =   7350
      Left            =   -120
      Picture         =   "frmInstruction3.frx":041A
      Top             =   -120
      Width           =   12555
   End
End
Attribute VB_Name = "frmInstruction3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
MakeWindow Me
End Sub

Private Sub Label5_Click()
FMain.Show
Unload Me
End Sub
