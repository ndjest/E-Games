VERSION 5.00
Begin VB.Form Question 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   0  'None
   ClientHeight    =   7095
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12405
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Question.frx":0000
   ScaleHeight     =   473
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   827
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
      Left            =   720
      TabIndex        =   4
      Top             =   1200
      Width           =   7245
   End
   Begin VB.Shape Shape5 
      BackColor       =   &H0080FFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00A27E4F&
      BorderWidth     =   2
      Height          =   2535
      Left            =   480
      Shape           =   4  'Rounded Rectangle
      Top             =   840
      Width           =   7815
   End
   Begin VB.Image imgLife 
      Height          =   480
      Index           =   3
      Left            =   11640
      Picture         =   "Question.frx":15D62
      Top             =   1800
      Width           =   480
   End
   Begin VB.Image imgLife 
      Height          =   480
      Index           =   2
      Left            =   10920
      Picture         =   "Question.frx":1662C
      Top             =   1800
      Width           =   480
   End
   Begin VB.Image imgLife 
      Height          =   480
      Index           =   1
      Left            =   10200
      Picture         =   "Question.frx":16EF6
      Top             =   1800
      Width           =   480
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
      Left            =   8880
      TabIndex        =   8
      Top             =   1800
      Width           =   1275
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
      Left            =   8880
      TabIndex        =   7
      Top             =   960
      Width           =   1275
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
      Left            =   10200
      TabIndex        =   6
      Top             =   960
      Width           =   1875
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
      Left            =   720
      TabIndex        =   2
      Top             =   5040
      Width           =   3435
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
      Left            =   4680
      TabIndex        =   1
      Top             =   4200
      Width           =   3375
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
      Index           =   4
      Left            =   4680
      TabIndex        =   0
      Top             =   5040
      Width           =   3405
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
      Left            =   720
      TabIndex        =   3
      Top             =   4200
      Width           =   3435
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00AE813C&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H001BF4FA&
      BorderWidth     =   2
      Height          =   735
      Left            =   600
      Shape           =   4  'Rounded Rectangle
      Top             =   4920
      Width           =   3735
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00AE813C&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H001BF4FA&
      BorderWidth     =   2
      Height          =   735
      Left            =   4560
      Shape           =   4  'Rounded Rectangle
      Top             =   4080
      Width           =   3735
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00AE813C&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H001BF4FA&
      BorderWidth     =   2
      Height          =   735
      Left            =   600
      Shape           =   4  'Rounded Rectangle
      Top             =   4080
      Width           =   3735
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00AE813C&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H001BF4FA&
      BorderWidth     =   2
      Height          =   735
      Left            =   4560
      Shape           =   4  'Rounded Rectangle
      Top             =   4920
      Width           =   3735
   End
   Begin VB.Image Image1 
      Height          =   7395
      Left            =   -120
      Picture         =   "Question.frx":177C0
      Top             =   -120
      Width           =   12555
   End
   Begin VB.Label lblAns 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Left            =   5160
      TabIndex        =   5
      Top             =   7920
      Width           =   375
   End
End
Attribute VB_Name = "Question"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim q_count As Integer
Dim idx As Integer
Dim X() As Long
Private Sub Answer_Click(Index As Integer)
sCateg = "Category1"
sQ = "Question1"
If Index = CInt(Me.lblAns.Caption) Then
    frmMsg.lblCorect.Visible = True
    frmMsg.imgCorrect.Visible = True
    i_check = 0
    lblScore.Caption = Format(CDbl(lblScore.Caption) + 1000, "###,###,##0")
    frmMsg.Show 1
Else
    sLife = sLife - 1
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
Dim max As Long
MakeWindow Me
q_count = 0

Set rs = New ADODB.Recordset
rs.Open "Select * from Easy1", cn, 1, 2
Do Until rs.EOF
    max = max + 1
    rs.MoveNext
Loop
GenerateNo (max)

displayQuestions
sRound = "Easy Round"
End Sub

Sub displayQuestions()
On Error GoTo err
Dim no, max As Long

Dim strCriteria
    
    idx = idx + 1

Set rs = New ADODB.Recordset
rs.Open "Select * from Easy1", cn, 1, 2
Do Until rs.EOF
    max = max + 1
    rs.MoveNext
Loop

q_count = q_count + 1

If q_count <= 10 Then
    strCriteria = X(idx)
    Set rs = New ADODB.Recordset
    With rs
        .Open "Select * from Easy1 where generalId = " & strCriteria & "", cn, 1, 2
        If .EOF = False Then
            Me.lblQuestion.Caption = .Fields("question").Value
            Answer(1).Caption = .Fields("ChoiceA").Value
            Answer(2).Caption = .Fields("ChoiceB").Value
            Answer(3).Caption = .Fields("ChoiceC").Value
            Answer(4).Caption = .Fields("ChoiceD").Value
            Me.lblAns.Caption = .Fields("Answer").Value
        End If
    End With
Else
    Question1.Show
    Unload Me
End If

Exit Sub
err:
    Question1.Show
    Unload Me
End Sub

Sub GenerateNo(ByVal MaxNo As Long)
On Error GoTo err_GenerateNo

    Dim no As Long
    Dim i As Long
    Dim bfound As Boolean
    
    ReDim X(1 To 1)
    
    While UBound(X) < MaxNo
        Randomize
        no = CLng((MaxNo * Rnd + 0.5))
        
        For i = LBound(X) To UBound(X)
            If X(i) = no Then
                bfound = True
                Exit For
            End If
        Next i
        
        If Not bfound Then
            If UBound(X) = LBound(X) And X(LBound(X)) = 0 Then
                X(LBound(X)) = no
            Else
                ReDim Preserve X(LBound(X) To UBound(X) + 1)
                X(UBound(X)) = no
            End If
        End If
    bfound = False
    Wend
exit_GenerateNo:
    Exit Sub
    
err_GenerateNo:
    MsgBox err.Description, vbInformation
    Resume exit_GenerateNo
End Sub

