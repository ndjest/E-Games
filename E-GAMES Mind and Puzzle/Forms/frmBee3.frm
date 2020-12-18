VERSION 5.00
Begin VB.Form frmBee3 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7020
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12360
   LinkTopic       =   "Form1"
   ScaleHeight     =   7020
   ScaleWidth      =   12360
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label lblQuestion 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   1785
      Left            =   5160
      TabIndex        =   0
      Top             =   720
      Width           =   6525
   End
   Begin VB.Shape Shape5 
      BackColor       =   &H0080FFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00A27E4F&
      BorderWidth     =   2
      Height          =   2535
      Left            =   4920
      Shape           =   4  'Rounded Rectangle
      Top             =   360
      Width           =   7095
   End
   Begin VB.Label lbl1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Menu"
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
      Left            =   360
      MouseIcon       =   "frmBee3.frx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   8
      Top             =   6000
      Width           =   1575
   End
   Begin VB.Image Image3 
      Height          =   1035
      Left            =   120
      Picture         =   "frmBee3.frx":030A
      Top             =   5640
      Width           =   2250
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FF0000&
      BackStyle       =   0  'Transparent
      Caption         =   "Life"
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
      Height          =   555
      Left            =   4800
      TabIndex        =   7
      Top             =   6120
      Width           =   1275
   End
   Begin VB.Image imgLife 
      Height          =   480
      Index           =   3
      Left            =   7440
      Picture         =   "frmBee3.frx":0582
      Top             =   6120
      Width           =   480
   End
   Begin VB.Image imgLife 
      Height          =   480
      Index           =   2
      Left            =   6720
      Picture         =   "frmBee3.frx":0E4C
      Top             =   6120
      Width           =   480
   End
   Begin VB.Image imgLife 
      Height          =   480
      Index           =   1
      Left            =   6000
      Picture         =   "frmBee3.frx":1716
      Top             =   6120
      Width           =   480
   End
   Begin VB.Label lblScore 
      BackColor       =   &H00FF0000&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   555
      Left            =   9960
      TabIndex        =   6
      Top             =   6120
      Width           =   1875
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FF0000&
      BackStyle       =   0  'Transparent
      Caption         =   "Score"
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
      Height          =   555
      Left            =   8640
      TabIndex        =   5
      Top             =   6120
      Width           =   1275
   End
   Begin VB.Label Answer 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      BackStyle       =   0  'Transparent
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
      Height          =   435
      Index           =   2
      Left            =   6240
      TabIndex        =   3
      Top             =   4920
      Width           =   4755
   End
   Begin VB.Label Answer 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      BackStyle       =   0  'Transparent
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
      Height          =   555
      Index           =   3
      Left            =   6240
      TabIndex        =   2
      Top             =   4080
      Width           =   4695
   End
   Begin VB.Label Answer 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      BackStyle       =   0  'Transparent
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
      Height          =   555
      Index           =   1
      Left            =   6240
      TabIndex        =   1
      Top             =   3240
      Width           =   4635
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00AE813C&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H001BF4FA&
      BorderWidth     =   2
      Height          =   735
      Left            =   5280
      Shape           =   4  'Rounded Rectangle
      Top             =   4800
      Width           =   6495
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00AE813C&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H001BF4FA&
      BorderWidth     =   2
      Height          =   735
      Left            =   5280
      Shape           =   4  'Rounded Rectangle
      Top             =   3120
      Width           =   6495
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00AE813C&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H001BF4FA&
      BorderWidth     =   2
      Height          =   735
      Left            =   5280
      Shape           =   4  'Rounded Rectangle
      Top             =   3960
      Width           =   6495
   End
   Begin VB.Image Image1 
      Height          =   7350
      Left            =   -120
      Picture         =   "frmBee3.frx":1FE0
      Top             =   -120
      Width           =   12555
   End
   Begin VB.Label lblAns 
      Height          =   255
      Left            =   1680
      TabIndex        =   4
      Top             =   6360
      Width           =   375
   End
End
Attribute VB_Name = "frmBee3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim q_count As Integer

Private Sub Answer_Click(Index As Integer)
sCateg = "Category3"
sQ = "Question4"
If Index = CInt(Me.lblAns.Caption) Then
    frmMsg.lblCorect.Visible = True
    frmMsg.imgCorrect.Visible = True
    i_check = 0
    lblScore.Caption = Format(CDbl(lblScore.Caption) + 2000, "###,###,##0")
    frmMsg.Show 1
Else
'    frmMsg1.lblCorect.Visible = False
'    frmMsg1.imgCorrect.Visible = False
'    frmMsg1.lblScore.Caption = Me.lblScore.Caption
'    frmMsg1.lblRound.Caption = sRound
'    frmMsg1.Show 1
    sLife = sLife - 1
    frmMsg1.lblCorect.Visible = False
    frmMsg1.imgCorrect.Visible = False
    If sLife = 0 Then
        frmMsg1.Label1.Caption = "GAME OVER"
        frmMsg1.lblScore.Visible = True
        frmMsg1.lblScore.Caption = Me.lblScore.Caption
        frmMsg1.lblRound.Caption = sRound
        frmMsg1.Label2.Caption = "Your Score"
        frmMsg1.Kinabuhi.Visible = False
    Else
        frmMsg1.Label1.Caption = "Incorrect"
        frmMsg1.lblScore.Visible = False
        frmMsg1.lblRound.Caption = sRound
        frmMsg1.Label2.Caption = "You lose one"
        frmMsg1.Kinabuhi.Visible = True
    End If
    frmMsg1.Show 1
End If
End Sub

Private Sub Form_Activate()
    Select Case sLife
        Case 0:
            Me.imgLife(1).Visible = False
            Me.imgLife(2).Visible = False
            Me.imgLife(3).Visible = False
        Case 1:
            Me.imgLife(1).Visible = True
            Me.imgLife(2).Visible = False
            Me.imgLife(3).Visible = False
        Case 2:
            Me.imgLife(1).Visible = True
            Me.imgLife(2).Visible = True
            Me.imgLife(3).Visible = False
        Case 3:
            Me.imgLife(1).Visible = True
            Me.imgLife(2).Visible = True
            Me.imgLife(3).Visible = True
    End Select
    Me.lblScore.Caption = Format(sScore, "###,##0")
End Sub

Private Sub Form_Load()
MakeWindow Me
sRound = "Difficult Round"
q_count = 0
displayQuestions
End Sub

Sub displayQuestions()
q_count = q_count + 1
If q_count <= 5 Then
    Set rs = New ADODB.Recordset
    With rs
        .Open "Select * from hard1 where questionId = " & q_count & "", cn, 1, 2
        If .EOF = False Then
            Me.lblQuestion.Caption = .Fields("question").Value
            Answer(1).Caption = .Fields("ChoiceA").Value
            Answer(2).Caption = .Fields("ChoiceB").Value
            Answer(3).Caption = .Fields("ChoiceC").Value
            Me.lblAns.Caption = .Fields("Answer").Value
        End If
    End With
Else
    Unload Me
    frmRat.Show
'    frmMsg2.lblRond.Caption = "Difficult Round Completed."
'    frmMsg2.Show
End If
End Sub

Private Sub lbl1_Click()
Unload Me
Category.Show
End Sub

Private Sub lblScore_Change()
sScore = Me.lblScore.Caption
End Sub

