VERSION 5.00
Begin VB.Form frmInstruction2 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7140
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12390
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   7140
   ScaleWidth      =   12390
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      MouseIcon       =   "frmInstruction2.frx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   6000
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmInstruction2.frx":030A
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
      Height          =   3735
      Left            =   1920
      TabIndex        =   0
      Top             =   960
      Width           =   8175
   End
   Begin VB.Image Image1 
      Height          =   7350
      Left            =   -120
      Picture         =   "frmInstruction2.frx":0466
      Top             =   -120
      Width           =   12555
   End
End
Attribute VB_Name = "frmInstruction2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Label5_Click()
fGame.Show
Unload Me
End Sub
