VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmScore 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7050
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12345
   LinkTopic       =   "Form1"
   ScaleHeight     =   7050
   ScaleWidth      =   12345
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   3720
      TabIndex        =   11
      Top             =   4320
      Width           =   3735
      Begin VB.OptionButton Option3 
         Caption         =   "Option1"
         Height          =   195
         Left            =   2760
         TabIndex        =   16
         Top             =   240
         Width           =   255
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Option1"
         Height          =   195
         Left            =   1200
         TabIndex        =   14
         Top             =   240
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Option1"
         Height          =   195
         Left            =   0
         TabIndex        =   12
         Top             =   240
         Width           =   255
      End
      Begin VB.Label Label4 
         Caption         =   "Hard"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E25CCB&
         Height          =   255
         Left            =   3000
         TabIndex        =   17
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "Normal"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E25CCB&
         Height          =   255
         Left            =   1440
         TabIndex        =   15
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Easy"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E25CCB&
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.OptionButton Option6 
      Caption         =   "Option1"
      Height          =   195
      Left            =   10080
      TabIndex        =   9
      Top             =   4560
      Width           =   255
   End
   Begin VB.OptionButton Option5 
      Caption         =   "Option1"
      Height          =   195
      Left            =   8880
      TabIndex        =   7
      Top             =   4560
      Width           =   255
   End
   Begin VB.OptionButton Option4 
      Caption         =   "Option1"
      Height          =   195
      Left            =   7680
      TabIndex        =   5
      Top             =   4560
      Width           =   255
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   3975
      Left            =   3720
      TabIndex        =   0
      Top             =   360
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   7011
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   0
      BackColor       =   16761087
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Rounded MT Bold"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin MSComctlLib.ListView ListView2 
      Height          =   3975
      Left            =   7680
      TabIndex        =   1
      Top             =   360
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   7011
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   15779492
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Rounded MT Bold"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin VB.Label Label7 
      Caption         =   "Drop Text"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00AE813C&
      Height          =   255
      Left            =   10320
      TabIndex        =   10
      Top             =   4560
      Width           =   1095
   End
   Begin VB.Label Label6 
      Caption         =   "Sudoku"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00AE813C&
      Height          =   255
      Left            =   9120
      TabIndex        =   8
      Top             =   4560
      Width           =   855
   End
   Begin VB.Label Label5 
      Caption         =   "Picture"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00AE813C&
      Height          =   255
      Left            =   7920
      TabIndex        =   6
      Top             =   4560
      Width           =   855
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
      Left            =   480
      MouseIcon       =   "frmScore.frx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Top             =   6000
      Width           =   1575
   End
   Begin VB.Image Image3 
      Height          =   1035
      Left            =   120
      Picture         =   "frmScore.frx":030A
      Top             =   5640
      Width           =   2250
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Puzzle Category"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H007C2E10&
      Height          =   375
      Left            =   8160
      TabIndex        =   3
      Top             =   4920
      Width           =   2775
   End
   Begin VB.Label Mind 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Mind Category"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   375
      Left            =   4080
      TabIndex        =   2
      Top             =   4920
      Width           =   2655
   End
   Begin VB.Image Image1 
      Height          =   7350
      Left            =   -120
      Picture         =   "frmScore.frx":0582
      Top             =   -120
      Width           =   12555
   End
End
Attribute VB_Name = "frmScore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
If Option1.Value = True Then
    ViewEasy
ElseIf Option3.Value = True Then
    ViewNormal
ElseIf Option2.Value = True Then
    ViewDifficult
End If
End Sub

Private Sub Form_Load()
MakeWindow Me
Me.Option1.Value = True
Me.Option4.Value = True
ViewEasy
ViewPicture
End Sub

Sub ViewEasy()
With ListView1
    .ColumnHeaders.Clear
    .ListItems.Clear
    .ColumnHeaders.Add , , "ID", 1300
    .ColumnHeaders.Add , , "Name", 1300, 1
End With
Set rs = New ADODB.Recordset
With rs
    .Open "Select * from qeasy", cn, 1, 2
    Do Until rs.EOF
        Set lst = Me.ListView1.ListItems.Add(, , rs.Fields("name").Value)
            lst.ListSubItems.Add , , Format(rs.Fields("score").Value, "###,##0")
    .MoveNext
    Loop
End With
End Sub

Sub ViewNormal()
With ListView1
    .ColumnHeaders.Clear
    .ListItems.Clear
    .ColumnHeaders.Add , , "ID", 1300
    .ColumnHeaders.Add , , "Name", 1300, 1
End With
Set rs = New ADODB.Recordset
With rs
    .Open "Select * from qnormal", cn, 1, 2
    Do Until rs.EOF
        Set lst = Me.ListView1.ListItems.Add(, , rs.Fields("name").Value)
            lst.ListSubItems.Add , , Format(rs.Fields("score").Value, "###,##0")
    .MoveNext
    Loop
End With
End Sub
Sub ViewDifficult()
With ListView1
    .ColumnHeaders.Clear
    .ListItems.Clear
    .ColumnHeaders.Add , , "ID", 1300
    .ColumnHeaders.Add , , "Name", 1300, 1
End With
Set rs = New ADODB.Recordset
With rs
    .Open "Select * from qdifficult", cn, 1, 2
    Do Until rs.EOF
        Set lst = Me.ListView1.ListItems.Add(, , rs.Fields("name").Value)
            lst.ListSubItems.Add , , Format(rs.Fields("score").Value, "###,##0")
    .MoveNext
    Loop
End With
End Sub

Sub ViewPicture()
With ListView2
    .ColumnHeaders.Clear
    .ListItems.Clear
    .ColumnHeaders.Add , , "ID", 1300
    .ColumnHeaders.Add , , "Name", 1300, 1
End With
Set rs = New ADODB.Recordset
With rs
    .Open "Select * from picture", cn, 1, 2
    Do Until rs.EOF
        Set lst = Me.ListView2.ListItems.Add(, , rs.Fields("name").Value)
            lst.ListSubItems.Add , , Format(rs.Fields("score").Value, "###,##0")
    .MoveNext
    Loop
End With
End Sub

Sub ViewDropText()
With ListView2
    .ColumnHeaders.Clear
    .ListItems.Clear
    .ColumnHeaders.Add , , "ID", 1300
    .ColumnHeaders.Add , , "Name", 1300, 1
End With
Set rs = New ADODB.Recordset
With rs
    .Open "Select * from droptext", cn, 1, 2
    Do Until rs.EOF
        Set lst = Me.ListView2.ListItems.Add(, , rs.Fields("name").Value)
            lst.ListSubItems.Add , , Format(rs.Fields("score").Value, "###,##0")
    .MoveNext
    Loop
End With
End Sub

Sub ViewSudoku()
With ListView2
    .ColumnHeaders.Clear
    .ListItems.Clear
    .ColumnHeaders.Add , , "ID", 1300
    .ColumnHeaders.Add , , "Name", 1300, 1
End With
Set rs = New ADODB.Recordset
With rs
    .Open "Select * from sudoku", cn, 1, 2
    Do Until rs.EOF
        Set lst = Me.ListView2.ListItems.Add(, , rs.Fields("name").Value)
            lst.ListSubItems.Add , , Format(rs.Fields("score").Value, "###,##0")
    .MoveNext
    Loop
End With
End Sub

Private Sub lbl1_Click()
Unload Me
frmMain.Show
End Sub

Private Sub Option1_Click()
If Option1.Value = True Then
    ViewEasy
ElseIf Option2.Value = True Then
    ViewNormal
ElseIf Option3.Value = True Then
    ViewDifficult
End If
End Sub

Private Sub Option2_Click()
If Option1.Value = True Then
    ViewEasy
ElseIf Option2.Value = True Then
    ViewNormal
ElseIf Option3.Value = True Then
    ViewDifficult
End If
End Sub

Private Sub Option3_Click()
If Option1.Value = True Then
    ViewEasy
ElseIf Option2.Value = True Then
    ViewNormal
ElseIf Option3.Value = True Then
    ViewDifficult
End If
End Sub

Private Sub Option4_Click()
If Option4.Value = True Then
    ViewPicture
ElseIf Option5.Value = True Then
    ViewSudoku
ElseIf Option6.Value = True Then
    ViewDropText
End If
End Sub

Private Sub Option5_Click()
If Option4.Value = True Then
    ViewPicture
ElseIf Option5.Value = True Then
    ViewSudoku
ElseIf Option6.Value = True Then
    ViewDropText
End If
End Sub

Private Sub Option6_Click()
If Option4.Value = True Then
    ViewPicture
ElseIf Option5.Value = True Then
    ViewSudoku
ElseIf Option6.Value = True Then
    ViewDropText
End If
End Sub
