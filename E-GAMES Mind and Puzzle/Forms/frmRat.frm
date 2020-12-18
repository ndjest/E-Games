VERSION 5.00
Begin VB.Form frmRat 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7095
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12315
   LinkTopic       =   "Form1"
   ScaleHeight     =   7095
   ScaleWidth      =   12315
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Proceed"
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
      MouseIcon       =   "frmRat.frx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   6000
      Width           =   1575
   End
   Begin VB.Image Image3 
      Height          =   1035
      Left            =   9960
      Picture         =   "frmRat.frx":030A
      Top             =   5640
      Width           =   2250
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Enemies of Plants and Trees"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   855
      Left            =   1560
      TabIndex        =   1
      Top             =   5880
      Width           =   8520
   End
   Begin VB.Image Image1 
      Height          =   7350
      Left            =   -120
      Picture         =   "frmRat.frx":05A0
      Top             =   -120
      Width           =   12555
   End
End
Attribute VB_Name = "frmRat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
MakeWindow Me
End Sub

Private Sub Label1_Click()
frmRat1.Show
Unload Me
End Sub
