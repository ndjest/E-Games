VERSION 5.00
Begin VB.Form frmMsg2 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2340
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6705
   LinkTopic       =   "Form1"
   ScaleHeight     =   2340
   ScaleWidth      =   6705
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   600
      Top             =   1800
   End
   Begin VB.Label lblRond 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Difficult Round Completed"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   1455
      Left            =   120
      TabIndex        =   1
      Top             =   1200
      Width           =   6255
   End
   Begin VB.Label lblCorect 
      BackStyle       =   0  'Transparent
      Caption         =   "Congratulations"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   855
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   8295
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H001BF4FA&
      BackStyle       =   1  'Opaque
      Height          =   2775
      Left            =   -120
      Top             =   -120
      Width           =   6855
   End
End
Attribute VB_Name = "frmMsg2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i As Integer

Private Sub Form_Activate()
If sCateg = "Category1" Or sCateg = "Category1.1" Then
    Me.lblRond.Caption = "Easy Round Completed"
ElseIf sCateg = "Category2" Then
    Me.lblRond.Caption = "Normal Round Completed"
ElseIf sCateg = "Category3" Or sCateg = "Category3" Then
    Me.lblRond.Caption = "Difficult Round Completed"
End If
End Sub

Private Sub Form_Load()
MakeWindow Me
End Sub

Private Sub Timer1_Timer()
i = i + 1
If i >= 5 Then
    Set rs = New ADODB.Recordset
    With rs
        .Open "Select * from Mindscore", cn, 1, 2
        .AddNew
        .Fields("name").Value = sName
        If sCateg = "Category1" Or sCateg = "Category1.1" Then
            .Fields("category").Value = "Easy Round"
            Unload Question
            Unload Question1
        ElseIf sCateg = "Category2" Then
            .Fields("category").Value = "Normal Round"
            Unload Question2
        ElseIf sCateg = "Category3" Or sCateg = "Category3" Then
            .Fields("category").Value = "Difficult Round"
            Unload frmBee3
            Unload frmRat2
        End If
        .Fields("score").Value = sScore
        .Update
    End With
    
    Timer1.Enabled = False
    Unload Me
    Category.Show
End If
End Sub
