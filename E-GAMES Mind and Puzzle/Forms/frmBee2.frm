VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form frmBee2 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7095
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12360
   BeginProperty Font 
      Name            =   "Arial Rounded MT Bold"
      Size            =   14.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   7095
   ScaleWidth      =   12360
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtBee 
      BorderStyle     =   0  'None
      Height          =   5055
      Left            =   5040
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   480
      Width           =   6735
   End
   Begin WMPLibCtl.WindowsMediaPlayer WindowsMediaPlayer1 
      Height          =   135
      Left            =   2760
      TabIndex        =   4
      Top             =   9480
      Width           =   255
      URL             =   ""
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   50
      mute            =   0   'False
      uiMode          =   "full"
      stretchToFit    =   0   'False
      windowlessVideo =   0   'False
      enabled         =   -1  'True
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   450
      _cy             =   238
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Enable sound"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   855
      Left            =   5640
      TabIndex        =   3
      Top             =   6120
      Width           =   2520
   End
   Begin VB.Image Image4 
      Height          =   480
      Left            =   8280
      Picture         =   "frmBee2.frx":0000
      Stretch         =   -1  'True
      Top             =   6120
      Width           =   555
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
      MouseIcon       =   "frmBee2.frx":070A
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   6000
      Width           =   1575
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Proceed"
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
      Left            =   10200
      MouseIcon       =   "frmBee2.frx":0A14
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   6000
      Width           =   1575
   End
   Begin VB.Image Image3 
      Height          =   1035
      Left            =   120
      Picture         =   "frmBee2.frx":0D1E
      Top             =   5640
      Width           =   2250
   End
   Begin VB.Image Image2 
      Height          =   1035
      Left            =   9840
      Picture         =   "frmBee2.frx":0F96
      Top             =   5640
      Width           =   2250
   End
   Begin VB.Image Image1 
      Height          =   7350
      Left            =   -120
      Picture         =   "frmBee2.frx":122C
      Top             =   -120
      Width           =   12555
   End
End
Attribute VB_Name = "frmBee2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sirsuspect As New SpVoice
Private Sub Form_Load()
Dim content As String
MakeWindow Me
Open App.Path & "\database\hard.txt" For Input As #1
Do Until EOF(1)
    Input #1, content
    Me.txtBee.Text = Me.txtBee.Text & content & vbNewLine
Loop
Close (1)
End Sub

Private Sub Image4_Click()
On Error GoTo err
'sirsuspect.Speak Me.txtBee.Text
WindowsMediaPlayer1.URL = App.Path & "\bees.avi"
WindowsMediaPlayer1.Controls.play
Exit Sub
err:
sirsuspect.Speak "No data."
End Sub

Private Sub Label1_Click()
frmBee3.Show
Unload Me
End Sub

Private Sub lbl1_Click()
frmBee1.Show
Unload Me
End Sub
