VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7065
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12375
   LinkTopic       =   "Form1"
   Picture         =   "frmMain.frx":0000
   ScaleHeight     =   7065
   ScaleWidth      =   12375
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Image Image4 
      Height          =   855
      Left            =   9240
      ToolTipText     =   "Noahs Ark ( Christian Animated Cartoon Movie )"
      Top             =   4200
      Width           =   855
   End
   Begin VB.Image Image3 
      Height          =   855
      Left            =   7920
      ToolTipText     =   "Beginners Bible For Children Jesus Christs Life Story"
      Top             =   3000
      Width           =   855
   End
   Begin VB.Image Image2 
      Height          =   855
      Left            =   6120
      ToolTipText     =   "The Ten Commandments ( Christian Animated Cartoon Movie )"
      Top             =   3000
      Width           =   855
   End
   Begin VB.Image Image1 
      Height          =   855
      Left            =   4800
      ToolTipText     =   "The Story Of Adam And Eve ( Christian Animated Cartoon Movie)"
      Top             =   4320
      Width           =   975
   End
   Begin VB.Label lblName 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   975
      Left            =   480
      MouseIcon       =   "frmMain.frx":15D62
      MousePointer    =   99  'Custom
      TabIndex        =   6
      Top             =   1080
      Width           =   4455
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00F0C6A4&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      Height          =   975
      Left            =   360
      Shape           =   4  'Rounded Rectangle
      Top             =   1080
      Width           =   4695
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Click Here"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   3000
      MouseIcon       =   "frmMain.frx":1606C
      MousePointer    =   99  'Custom
      TabIndex        =   8
      Top             =   2160
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "If this is not you,"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   960
      MouseIcon       =   "frmMain.frx":16376
      MousePointer    =   99  'Custom
      TabIndex        =   7
      Top             =   2160
      Width           =   2175
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Welcome back"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00A27E4F&
      Height          =   735
      Left            =   720
      MouseIcon       =   "frmMain.frx":16680
      MousePointer    =   99  'Custom
      TabIndex        =   5
      Top             =   480
      Width           =   3855
   End
   Begin VB.Label lbl5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H001BF4FA&
      Height          =   375
      Left            =   10200
      MouseIcon       =   "frmMain.frx":1698A
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Top             =   6120
      Width           =   1815
   End
   Begin VB.Label lbl4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Puzzle"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H001BF4FA&
      Height          =   375
      Left            =   3360
      MouseIcon       =   "frmMain.frx":16C94
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   6120
      Width           =   1815
   End
   Begin VB.Label lbl3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Hall of Fame"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H001BF4FA&
      Height          =   375
      Left            =   7800
      MouseIcon       =   "frmMain.frx":16F9E
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   6120
      Width           =   2175
   End
   Begin VB.Label lbl2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "New User"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H001BF4FA&
      Height          =   375
      Left            =   5640
      MouseIcon       =   "frmMain.frx":172A8
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   6120
      Width           =   1815
   End
   Begin VB.Label lbl1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Mind"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H001BF4FA&
      Height          =   375
      Left            =   840
      MouseIcon       =   "frmMain.frx":175B2
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   6120
      Width           =   2175
   End
   Begin VB.Image Image6 
      Height          =   7755
      Left            =   -120
      Picture         =   "frmMain.frx":178BC
      Top             =   -600
      Width           =   27690
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
 MakeWindow Me
'     Set rs = New ADODB.Recordset
'    rs.Open "Select * from [User]", cn, 1, 2
'    If rs.EOF = False Then
'        rs.MoveLast
'        frmMain.lblName.Caption = rs.Fields("name").Value
'        frmMain.Label4.Visible = True
'        frmMain.lblName.Visible = True
'    Else
'        frmMain.Label4.Visible = False
'        frmMain.lblName.Visible = False
'    End If
lblName.Caption = sName
getLife
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lbl1.ForeColor = &H1BF4FA
lbl2.ForeColor = &H1BF4FA
lbl3.ForeColor = &H1BF4FA
lbl4.ForeColor = &H1BF4FA
lbl5.ForeColor = &H1BF4FA
End Sub

Private Sub Form_Resize()
'imgTitleMain.Width = ScaleWidth
End Sub

'Private Sub Image1_Click()
'frmStory.WindowsMediaPlayer1.URL = App.Path & "\The Story Of Adam And Eve ( Christian Animated Cartoon Movie).mp4"
'frmStory.WindowsMediaPlayer1.Controls.play
'frmStory.Show
'Unload Me
'End Sub
'
'Private Sub Image2_Click()
'frmStory.WindowsMediaPlayer1.URL = App.Path & "\The Ten Commandments ( Christian Animated Cartoon Movie ).mp4"
'frmStory.WindowsMediaPlayer1.Controls.play
'frmStory.Show
'Unload Me
'End Sub
'Private Sub Image3_Click()
'frmStory.WindowsMediaPlayer1.URL = App.Path & "\Beginners Bible For Children Jesus Christs Life Story.mp4"
'frmStory.WindowsMediaPlayer1.Controls.play
'frmStory.Show
'Unload Me
'End Sub
'Private Sub Image4_Click()
'frmStory.WindowsMediaPlayer1.URL = App.Path & "\Noahs Ark ( Christian Animated Cartoon Movie ).mp4"
'frmStory.WindowsMediaPlayer1.Controls.play
'frmStory.Show
'Unload Me
'End Sub

Private Sub Image6_Click()

End Sub

Private Sub Label2_Click()
frmUser.Show
Unload Me
End Sub

Private Sub lbl1_Click()
Category.Show
Unload Me
sLife = 3
End Sub

Private Sub lbl1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lbl1.ForeColor = &HFF&
lbl2.ForeColor = &H1BF4FA
lbl3.ForeColor = &H1BF4FA
lbl4.ForeColor = &H1BF4FA
lbl5.ForeColor = &H1BF4FA
End Sub

Private Sub lbl2_Click()
frmCreateNew.Show
Unload Me
End Sub

Private Sub lbl2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lbl1.ForeColor = &H1BF4FA
lbl2.ForeColor = &HFF&
lbl3.ForeColor = &H1BF4FA
lbl4.ForeColor = &H1BF4FA
lbl5.ForeColor = &H1BF4FA
End Sub

Private Sub lbl3_Click()
frmScore.Show
Unload Me
End Sub

Private Sub lbl3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lbl1.ForeColor = &H1BF4FA
lbl2.ForeColor = &H1BF4FA
lbl3.ForeColor = &HFF&
lbl4.ForeColor = &H1BF4FA
lbl5.ForeColor = &H1BF4FA
End Sub

Private Sub lbl4_Click()
frmSelPuzzle.Show
Unload Me
End Sub

Private Sub lbl4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lbl1.ForeColor = &H1BF4FA
lbl2.ForeColor = &H1BF4FA
lbl3.ForeColor = &H1BF4FA
lbl4.ForeColor = &HFF&
lbl5.ForeColor = &H1BF4FA
End Sub

Private Sub lbl5_Click()
frmQuit.Show
End Sub

Private Sub lbl5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lbl1.ForeColor = &H1BF4FA
lbl2.ForeColor = &H1BF4FA
lbl3.ForeColor = &H1BF4FA
lbl4.ForeColor = &H1BF4FA
lbl5.ForeColor = &HFF&
End Sub
