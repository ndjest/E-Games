VERSION 5.00
Begin VB.Form frmMsg 
   BackColor       =   &H00AE813C&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2130
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5760
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   2130
   ScaleWidth      =   5760
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   4560
      Top             =   2640
   End
   Begin VB.Label lblEarned 
      BackStyle       =   0  'Transparent
      Caption         =   "1,000"
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
      Height          =   495
      Left            =   3720
      TabIndex        =   3
      Top             =   1440
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "You earned"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF34&
      Height          =   495
      Left            =   1440
      TabIndex        =   2
      Top             =   1440
      Width           =   2175
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
      Left            =   1320
      TabIndex        =   0
      Top             =   360
      Width           =   8295
   End
   Begin VB.Image imgCorrect 
      Height          =   825
      Left            =   240
      Picture         =   "frmMsg.frx":0000
      Stretch         =   -1  'True
      Top             =   360
      Width           =   810
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H001BF4FA&
      BackStyle       =   1  'Opaque
      Height          =   2655
      Left            =   -120
      Top             =   -240
      Width           =   6495
   End
   Begin VB.Label lblRem 
      Height          =   135
      Left            =   1200
      TabIndex        =   1
      Top             =   1200
      Width           =   255
   End
End
Attribute VB_Name = "frmMsg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i As Integer

Private Sub Form_Load()
MakeWindow Me
i = 0
If sCateg = "Category1" Or sCateg = "Category1.1" Then
    Me.lblEarned.Caption = "1,000"
    sScore = Format(CInt(sScore) + 1000, "###,##0")
ElseIf sCateg = "Category2" Then
    Me.lblEarned.Caption = "1,500"
    sScore = Format(CInt(sScore) + 1500, "###,##0")
ElseIf sCateg = "Category3" Then
    Me.lblEarned.Caption = "2,000"
    sScore = Format(CInt(sScore) + 2000, "###,##0")
End If
End Sub

Private Sub Timer1_Timer()
i = i + 1
If i >= 20 Then
    If sCateg = "Category1" Then
        Timer1.Enabled = False
        Unload Me
        Question.displayQuestions
        If i_check = 1 Then
            Unload Question
            frmMain.Show
        End If
    ElseIf sCateg = "Category1.1" Then
        Timer1.Enabled = False
        Unload Me
        Question1.displayQuestions
        If i_check = 1 Then
            Unload Question1
            frmMain.Show
        End If
    ElseIf sCateg = "Category2" Then
        Timer1.Enabled = False
        Unload Me
        Question2.displayQuestions
        If i_check = 1 Then
            Unload Question2
            frmMain.Show
        End If
    ElseIf sCateg = "Category3" Then
        Timer1.Enabled = False
        Unload Me
        frmBee3.displayQuestions
        If i_check = 1 Then
            Unload frmBee3
            frmMain.Show
        End If
    ElseIf sCateg = "Category3.1" Then
        Timer1.Enabled = False
        Unload Me
        frmRat2.displayQuestions
        If i_check = 1 Then
            Unload frmRat2
            frmMain.Show
        End If
    End If
End If
End Sub
