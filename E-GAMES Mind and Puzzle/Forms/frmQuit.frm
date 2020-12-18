VERSION 5.00
Begin VB.Form frmQuit 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2760
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5220
   LinkTopic       =   "Form1"
   Picture         =   "frmQuit.frx":0000
   ScaleHeight     =   2760
   ScaleWidth      =   5220
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label lbl2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "No"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H001BF4FA&
      Height          =   375
      Left            =   2880
      MouseIcon       =   "frmQuit.frx":15D62
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   2040
      Width           =   855
   End
   Begin VB.Label lbl1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Yes"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H001BF4FA&
      Height          =   375
      Left            =   1560
      MouseIcon       =   "frmQuit.frx":1606C
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   2040
      Width           =   855
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Do you want to exit?"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   735
      Left            =   480
      TabIndex        =   0
      Top             =   960
      Width           =   4440
   End
   Begin VB.Image Image1 
      Height          =   570
      Left            =   1440
      Picture         =   "frmQuit.frx":16376
      Stretch         =   -1  'True
      Top             =   1920
      Width           =   1125
   End
   Begin VB.Image Image2 
      Height          =   570
      Left            =   2760
      Picture         =   "frmQuit.frx":16CA0
      Stretch         =   -1  'True
      Top             =   1920
      Width           =   1125
   End
End
Attribute VB_Name = "frmQuit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
MakeWindow Me
End Sub


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lbl1.ForeColor = &H1BF4FA
lbl2.ForeColor = &H1BF4FA
End Sub

Private Sub lbl1_Click()
End
End Sub

Private Sub lbl1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lbl1.ForeColor = &HFF&
lbl2.ForeColor = &H1BF4FA
End Sub

Private Sub lbl2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lbl1.ForeColor = &H1BF4FA
lbl2.ForeColor = &HFF&
End Sub

Private Sub lbl2_Click()
Unload Me
End Sub
