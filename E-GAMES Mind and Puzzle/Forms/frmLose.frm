VERSION 5.00
Begin VB.Form frmLose 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2595
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8400
   LinkTopic       =   "Form1"
   ScaleHeight     =   2595
   ScaleWidth      =   8400
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      Left            =   3120
      MouseIcon       =   "frmLose.frx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   1800
      Width           =   1575
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Congrats! "" & sName & vbCr & "" You are just nerdy enough to make it onto the high score list!"" "
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
      Height          =   1575
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   7815
   End
   Begin VB.Image Image3 
      Height          =   1035
      Left            =   2760
      Picture         =   "frmLose.frx":030A
      Top             =   1440
      Width           =   2250
   End
End
Attribute VB_Name = "frmLose"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
MakeWindow Me
End Sub

Private Sub lbl1_Click()
Set rs = New ADODB.Recordset
rs.Open "Select * from mindscore", cn, 1, 2
With rs
    .AddNew
    .Fields("name").Value = sName
    .Fields("category").Value = "Drop"
    .Fields("Score").Value = fGame.lblScore.Caption
    .Update
End With
Unload Me
Unload fGame
frmSelPuzzle.Show
End Sub
