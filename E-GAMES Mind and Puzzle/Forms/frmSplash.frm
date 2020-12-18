VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSplash 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7095
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12435
   Icon            =   "frmSplash.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7095
   ScaleWidth      =   12435
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   6480
      Top             =   4440
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   480
      TabIndex        =   0
      Top             =   6360
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "https://www.facebook.com/sirsuspect"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   9000
      TabIndex        =   9
      Top             =   6840
      Width           =   8415
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "http://www.ipaidsoftware.blogspot.com | http://www.sirsuspect.blogspot.com"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   480
      TabIndex        =   8
      Top             =   6840
      Width           =   8415
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "KIDS"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   72
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H009F2E1A&
      Height          =   2295
      Left            =   5880
      TabIndex        =   7
      Top             =   2280
      Width           =   6255
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Games"
      BeginProperty Font 
         Name            =   "Berlin Sans FB Demi"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00EA8040&
      Height          =   975
      Left            =   1560
      TabIndex        =   6
      Top             =   1920
      Width           =   3855
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "for"
      BeginProperty Font 
         Name            =   "Berlin Sans FB"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00D37CB9&
      Height          =   1095
      Left            =   4560
      TabIndex        =   5
      Top             =   2640
      Width           =   2895
   End
   Begin VB.Label lblPercent 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Loading"
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
      Left            =   9720
      TabIndex        =   2
      Top             =   5880
      Width           =   2175
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "START"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   5400
      MouseIcon       =   "frmSplash.frx":058A
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Top             =   5280
      Width           =   1815
   End
   Begin VB.Image Image1 
      Height          =   855
      Left            =   5040
      Picture         =   "frmSplash.frx":0894
      Stretch         =   -1  'True
      Top             =   5040
      Width           =   2415
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Mind && Puzzle"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   72
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H007C2E10&
      Height          =   2295
      Left            =   120
      TabIndex        =   3
      Top             =   240
      Width           =   12255
   End
   Begin VB.Label lblLoading 
      BackStyle       =   0  'Transparent
      Caption         =   "Loading"
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
      Left            =   480
      TabIndex        =   1
      Top             =   5880
      Width           =   2175
   End
   Begin VB.Image imgTitleMain 
      Height          =   7350
      Left            =   -120
      Picture         =   "frmSplash.frx":0B0C
      Top             =   -120
      Width           =   12735
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim X As Integer
Private Sub Form_Load()
 MakeWindow Me
 Main
 Image1.Visible = False
    Label5.Visible = False
End Sub



Private Sub Label5_Click()
Unload Me
frmMain.Show
End Sub

Private Sub Timer1_Timer()
X = X + 1
If X > 20 Then
Me.ProgressBar1.Value = Me.ProgressBar1.Value + 1
Me.lblPercent.Caption = Me.ProgressBar1.Value & "%"
Select Case Me.ProgressBar1.Value
    Case 1 To 5
        Me.lblLoading.Caption = "Loading"
    Case 6 To 7
        Me.lblLoading.Caption = "Loading."
    Case 8 To 10
        Me.lblLoading.Caption = "Loading.."
    Case 11 To 13
        Me.lblLoading.Caption = "Loading..."
    Case 14 To 17
        Me.lblLoading.Caption = "Loading...."
    Case 18 To 25
        Me.lblLoading.Caption = "Loading....."
    Case 26 To 27
        Me.lblLoading.Caption = "Loading...."
    Case 28 To 30
        Me.lblLoading.Caption = "Loading..."
    Case 31 To 33
        Me.lblLoading.Caption = "Loading.."
    Case 34 To 37
        Me.lblLoading.Caption = "Loading."
    Case 38 To 45
        Me.lblLoading.Caption = "Loading"
    Case 46 To 47
        Me.lblLoading.Caption = "Loading."
    Case 48 To 50
        Me.lblLoading.Caption = "Loading.."
    Case 51 To 53
        Me.lblLoading.Caption = "Loading..."
    Case 54 To 57
        Me.lblLoading.Caption = "Loading...."
    Case 58 To 65
        Me.lblLoading.Caption = "Loading....."
    Case 66 To 67
        Me.lblLoading.Caption = "Loading...."
    Case 68 To 70
        Me.lblLoading.Caption = "Loading..."
    Case 71 To 73
        Me.lblLoading.Caption = "Loading.."
    Case 74 To 77
        Me.lblLoading.Caption = "Loading."
    Case 78 To 85
        Me.lblLoading.Caption = "Loading"
    Case 86 To 87
        Me.lblLoading.Caption = "Loading."
    Case 88 To 90
        Me.lblLoading.Caption = "Loading.."
    Case 91 To 93
        Me.lblLoading.Caption = "Loading..."
    Case 94 To 97
        Me.lblLoading.Caption = "Loading...."
    Case 98 To 100
        Me.lblLoading.Caption = "Loading....."
End Select

If Me.ProgressBar1.Value <= 10 Then
    Timer1.Interval = 10
ElseIf Me.ProgressBar1.Value <= 20 Then
    Timer1.Interval = 40
ElseIf Me.ProgressBar1.Value <= 50 Then
    Timer1.Interval = 100
ElseIf Me.ProgressBar1.Value <= 100 Then
    Timer1.Interval = 150
End If

If Me.ProgressBar1.Value >= 100 Then
    Timer1.Enabled = False
    'Unload Me
    'frmMain.Show
    Image1.Visible = True
    Label5.Visible = True
End If
End If
End Sub
