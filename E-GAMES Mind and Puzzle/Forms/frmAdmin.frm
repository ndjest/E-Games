VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAdmin 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7530
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   13995
   LinkTopic       =   "Form1"
   Picture         =   "frmAdmin.frx":0000
   ScaleHeight     =   7530
   ScaleWidth      =   13995
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cboCorrect 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   390
      Left            =   9840
      TabIndex        =   18
      Top             =   6120
      Width           =   1095
   End
   Begin VB.TextBox txtD 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   390
      Left            =   9840
      TabIndex        =   16
      Top             =   5640
      Width           =   3855
   End
   Begin VB.TextBox txtC 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   390
      Left            =   9840
      TabIndex        =   14
      Top             =   5160
      Width           =   3855
   End
   Begin VB.TextBox txtB 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   390
      Left            =   9840
      TabIndex        =   12
      Top             =   4680
      Width           =   3855
   End
   Begin VB.TextBox txtA 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   390
      Left            =   9840
      TabIndex        =   10
      Top             =   4200
      Width           =   3855
   End
   Begin VB.ComboBox cboCategory 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   390
      Left            =   8520
      TabIndex        =   6
      Top             =   1800
      Width           =   5175
   End
   Begin VB.TextBox txtQuestion 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   1350
      Left            =   8520
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Top             =   2760
      Width           =   5175
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   5055
      Left            =   240
      TabIndex        =   1
      Top             =   1440
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   8916
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Rounded MT Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Correct"
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
      Height          =   375
      Left            =   8520
      MouseIcon       =   "frmAdmin.frx":15D62
      MousePointer    =   99  'Custom
      TabIndex        =   17
      Top             =   6120
      Width           =   2175
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Choice D"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   8520
      MouseIcon       =   "frmAdmin.frx":1606C
      MousePointer    =   99  'Custom
      TabIndex        =   15
      Top             =   5640
      Width           =   2175
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Choice C"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   8520
      MouseIcon       =   "frmAdmin.frx":16376
      MousePointer    =   99  'Custom
      TabIndex        =   13
      Top             =   5160
      Width           =   2175
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Choice B"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   8520
      MouseIcon       =   "frmAdmin.frx":16680
      MousePointer    =   99  'Custom
      TabIndex        =   11
      Top             =   4680
      Width           =   2175
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Choice A"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   8520
      MouseIcon       =   "frmAdmin.frx":1698A
      MousePointer    =   99  'Custom
      TabIndex        =   9
      Top             =   4200
      Width           =   2175
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Question"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   8520
      MouseIcon       =   "frmAdmin.frx":16C94
      MousePointer    =   99  'Custom
      TabIndex        =   8
      Top             =   2400
      Width           =   2175
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Category"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   8520
      MouseIcon       =   "frmAdmin.frx":16F9E
      MousePointer    =   99  'Custom
      TabIndex        =   7
      Top             =   1440
      Width           =   2175
   End
   Begin VB.Label lbl3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H001BF4FA&
      Height          =   375
      Left            =   12000
      MouseIcon       =   "frmAdmin.frx":172A8
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Top             =   6795
      Width           =   1335
   End
   Begin VB.Label lbl2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Edit"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H001BF4FA&
      Height          =   375
      Left            =   10320
      MouseIcon       =   "frmAdmin.frx":175B2
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   6795
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Administrator "
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H001BF4FA&
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   3255
   End
   Begin VB.Label lbl1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "New"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H001BF4FA&
      Height          =   375
      Left            =   8640
      MouseIcon       =   "frmAdmin.frx":178BC
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   6795
      Width           =   1335
   End
   Begin VB.Image Image1 
      Height          =   450
      Left            =   8520
      MouseIcon       =   "frmAdmin.frx":17BC6
      MousePointer    =   99  'Custom
      Picture         =   "frmAdmin.frx":17ED0
      Stretch         =   -1  'True
      Top             =   6720
      Width           =   1605
   End
   Begin VB.Image Image2 
      Height          =   450
      Left            =   10200
      MouseIcon       =   "frmAdmin.frx":187FA
      MousePointer    =   99  'Custom
      Picture         =   "frmAdmin.frx":18B04
      Stretch         =   -1  'True
      Top             =   6720
      Width           =   1605
   End
   Begin VB.Image Image3 
      Height          =   450
      Left            =   11880
      MouseIcon       =   "frmAdmin.frx":1942E
      MousePointer    =   99  'Custom
      Picture         =   "frmAdmin.frx":19738
      Stretch         =   -1  'True
      Top             =   6720
      Width           =   1605
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H001BF4FA&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H001BF4FA&
      Height          =   8295
      Left            =   -240
      Top             =   435
      Width           =   14415
   End
End
Attribute VB_Name = "frmAdmin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cboCategory_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub cboCorrect_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub Form_Load()
MakeWindow Me
cboCategory.AddItem "Easy"
cboCategory.AddItem "Normal"
cboCategory.AddItem "Hard"
cboCorrect.AddItem "A"
cboCorrect.AddItem "B"
cboCorrect.AddItem "C"
cboCorrect.AddItem "D"
Call connection
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lbl1.ForeColor = &H1BF4FA
lbl2.ForeColor = &H1BF4FA
lbl3.ForeColor = &H1BF4FA
End Sub

Private Sub lbl1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lbl1.ForeColor = &HFF&
lbl2.ForeColor = &H1BF4FA
lbl3.ForeColor = &H1BF4FA
End Sub

Private Sub lbl2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lbl2.ForeColor = &HFF&
lbl1.ForeColor = &H1BF4FA
lbl3.ForeColor = &H1BF4FA
End Sub

Private Sub lbl3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lbl3.ForeColor = &HFF&
lbl2.ForeColor = &H1BF4FA
lbl1.ForeColor = &H1BF4FA
End Sub
