VERSION 5.00
Begin VB.Form frmMsg1 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2745
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6315
   LinkTopic       =   "Form1"
   ScaleHeight     =   2745
   ScaleWidth      =   6315
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   360
      Top             =   1800
   End
   Begin VB.Image Kinabuhi 
      Height          =   480
      Left            =   3720
      Picture         =   "frmMsg1.frx":0000
      Top             =   1920
      Width           =   480
   End
   Begin VB.Label lblRound 
      BackStyle       =   0  'Transparent
      Caption         =   "Difficult Round"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   615
      Left            =   3360
      TabIndex        =   6
      Top             =   1320
      Width           =   2895
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Category"
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
      Height          =   615
      Left            =   1080
      TabIndex        =   5
      Top             =   1320
      Width           =   2295
   End
   Begin VB.Label lblScore 
      BackStyle       =   0  'Transparent
      Caption         =   "Your Score"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   615
      Left            =   3360
      TabIndex        =   4
      Top             =   1920
      Width           =   2295
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "You lose one"
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
      Height          =   615
      Left            =   1080
      TabIndex        =   3
      Top             =   1920
      Width           =   3135
   End
   Begin VB.Image Image1 
      Height          =   825
      Left            =   240
      Picture         =   "frmMsg1.frx":08CA
      Stretch         =   -1  'True
      Top             =   240
      Width           =   810
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "GAME OVER"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   855
      Left            =   1320
      TabIndex        =   2
      Top             =   240
      Width           =   7215
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H001BF4FA&
      BackStyle       =   1  'Opaque
      Height          =   3255
      Left            =   -120
      Top             =   -120
      Width           =   6495
   End
   Begin VB.Image imgCorrect 
      Height          =   825
      Left            =   360
      Picture         =   "frmMsg1.frx":1357
      Stretch         =   -1  'True
      Top             =   600
      Width           =   810
   End
   Begin VB.Label lblCorect 
      BackStyle       =   0  'Transparent
      Caption         =   "You got it!"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF34&
      Height          =   855
      Left            =   1440
      TabIndex        =   1
      Top             =   600
      Width           =   8295
   End
   Begin VB.Image imgWrong 
      Height          =   825
      Left            =   360
      Picture         =   "frmMsg1.frx":216C
      Stretch         =   -1  'True
      Top             =   600
      Width           =   810
   End
   Begin VB.Label lblWrong 
      BackStyle       =   0  'Transparent
      Caption         =   "GAME OVER"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   855
      Left            =   1440
      TabIndex        =   0
      Top             =   600
      Width           =   7215
   End
End
Attribute VB_Name = "frmMsg1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i As Integer
Private Sub Form_Load()
MakeWindow Me
i = 0
End Sub

Private Sub Timer1_Timer()
i = i + 1
If i >= 20 Then
    If sLife = 0 Then
        Set rs = New ADODB.Recordset
        With rs
            .Open "Select * from Mindscore", cn, 1, 2
            .AddNew
            .Fields("name").Value = sName
            .Fields("category").Value = Me.lblRound.Caption
            .Fields("score").Value = Me.lblScore.Caption
            .Update
        End With
        Timer1.Enabled = False
        Unload Me
        If sQ = "Question1" Then
            Unload Question
        ElseIf sQ = "Question2" Then
            Unload Question1
        ElseIf sQ = "Question3" Then
            Unload Question2
        ElseIf sQ = "Question4" Then
            Unload frmBee3
        ElseIf sQ = "Question4.4" Then
            Unload frmRat2
        End If
        Category.Show
    ElseIf sQ = "Question1" Then
        Timer1.Enabled = False
        Unload Me
        Question.displayQuestions
    ElseIf sQ = "Question2" Then
        Timer1.Enabled = False
        Unload Me
        Question1.displayQuestions
    ElseIf sQ = "Question3" Then
        Timer1.Enabled = False
        Unload Me
        Question2.displayQuestions
    ElseIf sQ = "Question4" Then
        Timer1.Enabled = False
        Unload Me
        frmBee3.displayQuestions
    ElseIf sQ = "Question4.4" Then
        Timer1.Enabled = False
        Unload Me
        frmRat2.displayQuestions
    End If
End If
End Sub
