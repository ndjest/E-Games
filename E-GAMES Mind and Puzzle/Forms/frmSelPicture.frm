VERSION 5.00
Begin VB.Form frmSelPicture 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7905
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12495
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmSelPicture.frx":0000
   ScaleHeight     =   7905
   ScaleWidth      =   12495
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "How to play"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   735
      Left            =   8040
      MouseIcon       =   "frmSelPicture.frx":15D62
      MousePointer    =   99  'Custom
      TabIndex        =   8
      Top             =   6480
      Width           =   1575
   End
   Begin VB.Image Image5 
      Height          =   1035
      Left            =   7680
      Picture         =   "frmSelPicture.frx":1606C
      Top             =   6360
      Width           =   2250
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Picture Puzzle"
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
      Left            =   240
      TabIndex        =   7
      Top             =   6600
      Width           =   4920
   End
   Begin VB.Label lbl3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Level 3"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H001BF4FA&
      Height          =   615
      Left            =   4440
      MouseIcon       =   "frmSelPicture.frx":162E4
      MousePointer    =   99  'Custom
      TabIndex        =   6
      Top             =   3000
      Width           =   1695
   End
   Begin VB.Label lbl4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Level 4"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H001BF4FA&
      Height          =   615
      Left            =   6360
      MouseIcon       =   "frmSelPicture.frx":165EE
      MousePointer    =   99  'Custom
      TabIndex        =   5
      Top             =   1200
      Width           =   1695
   End
   Begin VB.Label lbl5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Level 5"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H001BF4FA&
      Height          =   615
      Left            =   8640
      MouseIcon       =   "frmSelPicture.frx":168F8
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Top             =   2880
      Width           =   1575
   End
   Begin VB.Label lbl7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Back"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H001BF4FA&
      Height          =   735
      Left            =   10440
      MouseIcon       =   "frmSelPicture.frx":16C02
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   6720
      Width           =   1575
   End
   Begin VB.Label lbl6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Level 6"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H001BF4FA&
      Height          =   615
      Left            =   10440
      MouseIcon       =   "frmSelPicture.frx":16F0C
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   1080
      Width           =   1575
   End
   Begin VB.Label lbl2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Level 2"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H001BF4FA&
      Height          =   615
      Left            =   2880
      MouseIcon       =   "frmSelPicture.frx":17216
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   1080
      Width           =   1695
   End
   Begin VB.Label lbl1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Level 1"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H001BF4FA&
      Height          =   855
      Left            =   600
      MouseIcon       =   "frmSelPicture.frx":17520
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   1920
      Width           =   1575
   End
   Begin VB.Image Image1 
      Height          =   8070
      Left            =   -120
      Picture         =   "frmSelPicture.frx":1782A
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12675
   End
   Begin VB.Image Image7 
      Height          =   570
      Left            =   6840
      MouseIcon       =   "frmSelPicture.frx":1D301
      MousePointer    =   99  'Custom
      Picture         =   "frmSelPicture.frx":1D60B
      Stretch         =   -1  'True
      Top             =   5160
      Width           =   3165
   End
   Begin VB.Image Image8 
      Height          =   570
      Left            =   6840
      MouseIcon       =   "frmSelPicture.frx":1DF35
      MousePointer    =   99  'Custom
      Picture         =   "frmSelPicture.frx":1E23F
      Stretch         =   -1  'True
      Top             =   5880
      Width           =   3165
   End
End
Attribute VB_Name = "frmSelPicture"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
Dim i_level As Integer
 Set rs = New ADODB.Recordset
 With rs
    .Open "Select * from [User] where [name ] like '" & sName & "'", cn, 1, 2
    If .EOF = False Then
        i_level = CInt(.Fields("picpuzz").Value)
    End If
End With

Select Case i_level
    Case 1:
        Me.lbl2.Caption = "Locked"
        Me.lbl3.Caption = "Locked"
        Me.lbl4.Caption = "Locked"
        Me.lbl5.Caption = "Locked"
        Me.lbl6.Caption = "Locked"
        Me.lbl2.Enabled = False
        Me.lbl3.Enabled = False
        Me.lbl4.Enabled = False
        Me.lbl5.Enabled = False
        Me.lbl6.Enabled = False
    Case 2:
        Me.lbl2.Caption = "Level 2"
        Me.lbl3.Caption = "Locked"
        Me.lbl4.Caption = "Locked"
        Me.lbl5.Caption = "Locked"
        Me.lbl6.Caption = "Locked"
        Me.lbl2.Enabled = True
        Me.lbl3.Enabled = False
        Me.lbl4.Enabled = False
        Me.lbl5.Enabled = False
        Me.lbl6.Enabled = False
    Case 3:
        Me.lbl2.Caption = "Level 2"
        Me.lbl3.Caption = "Level 3"
        Me.lbl4.Caption = "Locked"
        Me.lbl5.Caption = "Locked"
        Me.lbl6.Caption = "Locked"
        Me.lbl2.Enabled = True
        Me.lbl3.Enabled = True
        Me.lbl4.Enabled = False
        Me.lbl5.Enabled = False
        Me.lbl6.Enabled = False
    Case 4:
        Me.lbl2.Caption = "Level 2"
        Me.lbl3.Caption = "Level 3"
        Me.lbl4.Caption = "Level 4"
        Me.lbl5.Caption = "Locked"
        Me.lbl6.Caption = "Locked"
        Me.lbl2.Enabled = True
        Me.lbl3.Enabled = True
        Me.lbl4.Enabled = True
        Me.lbl5.Enabled = False
        Me.lbl6.Enabled = False
    Case 5:
        Me.lbl2.Caption = "Level 2"
        Me.lbl3.Caption = "Level 3"
        Me.lbl4.Caption = "Level 4"
        Me.lbl5.Caption = "Level 5"
        Me.lbl6.Caption = "Locked"
        Me.lbl2.Enabled = True
        Me.lbl3.Enabled = True
        Me.lbl4.Enabled = True
        Me.lbl5.Enabled = True
        Me.lbl6.Enabled = False
    Case 6:
        Me.lbl2.Caption = "Level 2"
        Me.lbl3.Caption = "Level 3"
        Me.lbl4.Caption = "Level 4"
        Me.lbl5.Caption = "Level 5"
        Me.lbl6.Caption = "Level 6"
        Me.lbl2.Enabled = True
        Me.lbl3.Enabled = True
        Me.lbl4.Enabled = True
        Me.lbl5.Enabled = True
        Me.lbl6.Enabled = True
    Case Else
        Me.lbl2.Caption = "Locked"
        Me.lbl3.Caption = "Locked"
        Me.lbl4.Caption = "Locked"
        Me.lbl5.Caption = "Locked"
        Me.lbl6.Caption = "Locked"
        Me.lbl2.Enabled = False
        Me.lbl3.Enabled = False
        Me.lbl4.Enabled = False
        Me.lbl5.Enabled = False
        Me.lbl6.Enabled = False
End Select
End Sub

Private Sub Form_Load()
 MakeWindow Me
 
End Sub

Private Sub Label6_Click()
frmInstruction4.Show
Unload Me
End Sub

Private Sub lbl1_Click()
frmPicEasy1.Show
frmSelPicture.Hide
End Sub

Private Sub lbl2_Click()
frmPicEasy2.Show
frmSelPicture.Hide
End Sub

Private Sub lbl3_Click()
frmPicNormal1.Show
frmSelPicture.Hide
End Sub

Private Sub lbl4_Click()
frmPicNormal2.Show
frmSelPicture.Hide
End Sub

Private Sub lbl5_Click()
frmPicHard1.Show
frmSelPicture.Hide
End Sub

Private Sub lbl6_Click()
frmPicHard2.Show
frmSelPicture.Hide
End Sub

Private Sub lbl7_Click()
frmSelPuzzle.Show
Unload Me
End Sub
