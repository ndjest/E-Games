VERSION 5.00
Begin VB.Form frmCreateNew 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6975
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9375
   LinkTopic       =   "Form1"
   ScaleHeight     =   6975
   ScaleWidth      =   9375
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtName 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
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
      Height          =   615
      Left            =   1920
      TabIndex        =   4
      Top             =   2640
      Width           =   5055
   End
   Begin VB.Label lblMsg 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   375
      Left            =   1920
      MouseIcon       =   "frmCreateNew.frx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   5
      Top             =   3720
      Width           =   5055
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Enter Your name"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00A27E4F&
      Height          =   375
      Left            =   1920
      MouseIcon       =   "frmCreateNew.frx":030A
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   1800
      Width           =   5055
   End
   Begin VB.Label Label2 
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
      ForeColor       =   &H007C2E10&
      Height          =   375
      Left            =   5280
      MouseIcon       =   "frmCreateNew.frx":0614
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   5640
      Width           =   1575
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Clear"
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
      Left            =   3360
      MouseIcon       =   "frmCreateNew.frx":091E
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   5640
      Width           =   1575
   End
   Begin VB.Label lbl1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Save"
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
      Left            =   1440
      MouseIcon       =   "frmCreateNew.frx":0C28
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   5640
      Width           =   1575
   End
   Begin VB.Image Image1 
      Height          =   7350
      Left            =   -120
      Picture         =   "frmCreateNew.frx":0F32
      Top             =   -120
      Width           =   12555
   End
   Begin VB.Label lblID 
      Height          =   255
      Left            =   1440
      TabIndex        =   6
      Top             =   3960
      Width           =   495
   End
End
Attribute VB_Name = "frmCreateNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
MakeWindow Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
Me.lblID.Caption = ""
End Sub

Private Sub Label1_Click()
Me.txtName.Text = ""
Me.lblMsg.Caption = ""
End Sub

Private Sub Label2_Click()
frmMain.Show
Unload Me
End Sub

Private Sub lbl1_Click()

Set rs = New ADODB.Recordset
With rs
    .Open "Select * from [User] where id like '" & Me.lblID.Caption & "'", cn, 1, 2
    If .EOF = False Then
        .Fields("Name").Value = Me.txtName.Text
        .Update
        Me.lblMsg.Caption = "New name successfully saved."
        Exit Sub
    End If
End With
Set rs = New ADODB.Recordset
With rs
    .Open "Select * from [User] where name like '" & Me.txtName.Text & "'", cn, 1, 2
    If .EOF = False Then
        Me.lblMsg.Caption = "Duplicate name."
    Else
        .AddNew
        .Fields("Name").Value = Me.txtName.Text
        .Update
        Me.lblMsg.Caption = "New name successfully saved."
        sName = Me.txtName.Text
    End If
End With
End Sub
