VERSION 5.00
Begin VB.Form frmBee1 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7035
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12210
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial Rounded MT Bold"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   7035
   ScaleWidth      =   12210
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
      Left            =   10080
      MouseIcon       =   "frmBee1.frx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   6000
      Width           =   1575
   End
   Begin VB.Label lbl1 
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
      Left            =   600
      MouseIcon       =   "frmBee1.frx":030A
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   6000
      Width           =   1575
   End
   Begin VB.Image Image3 
      Height          =   1035
      Left            =   9720
      Picture         =   "frmBee1.frx":0614
      Top             =   5640
      Width           =   2250
   End
   Begin VB.Image Image2 
      Height          =   1035
      Left            =   240
      Picture         =   "frmBee1.frx":08AA
      Top             =   5640
      Width           =   2250
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "A Bee Family"
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
      Left            =   1800
      TabIndex        =   0
      Top             =   6000
      Width           =   8520
   End
   Begin VB.Image Image1 
      Height          =   7350
      Left            =   -240
      Picture         =   "frmBee1.frx":0B22
      Top             =   -120
      Width           =   22770
   End
End
Attribute VB_Name = "frmBee1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
MakeWindow Me
End Sub

Private Sub lbl1_Click()
Unload Me
Category.Show
End Sub
