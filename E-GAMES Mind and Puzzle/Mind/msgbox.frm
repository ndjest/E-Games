VERSION 5.00
Begin VB.Form frmMsgbox 
   BackColor       =   &H00DF8446&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   1380
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   3720
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1380
   ScaleWidth      =   3720
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label lblMsg 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   315
      Left            =   330
      TabIndex        =   0
      Top             =   210
      Width           =   3195
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   1080
      Stretch         =   -1  'True
      Top             =   720
      Width           =   1335
   End
End
Attribute VB_Name = "frmMsgbox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    lblMsg.Caption = strMsgbox

End Sub

'Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Me.Label.BackColor = &H80000008
'End Sub

'Private Sub Label_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Me.Label.BackColor = &H8000000D
'End Sub

Private Sub Image1_Click()
Unload Me
End Sub
