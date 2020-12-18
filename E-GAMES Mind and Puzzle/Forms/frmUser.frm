VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmUser 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7110
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12375
   LinkTopic       =   "Form1"
   ScaleHeight     =   7110
   ScaleWidth      =   12375
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ListView ListView1 
      Height          =   5055
      Left            =   1440
      TabIndex        =   4
      Top             =   960
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   8916
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
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
   Begin VB.Label Label3 
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
      Left            =   9840
      MouseIcon       =   "frmUser.frx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   5880
      Width           =   1575
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Delete"
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
      Left            =   9840
      MouseIcon       =   "frmUser.frx":030A
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   4200
      Width           =   1575
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Rename"
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
      MouseIcon       =   "frmUser.frx":0614
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   2400
      Width           =   1575
   End
   Begin VB.Label lbl1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Change"
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
      MouseIcon       =   "frmUser.frx":091E
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   720
      Width           =   1575
   End
   Begin VB.Image Image1 
      Height          =   7350
      Left            =   -120
      Picture         =   "frmUser.frx":0C28
      Top             =   -120
      Width           =   12555
   End
End
Attribute VB_Name = "frmUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
MakeWindow Me
ViewRec
End Sub

Private Sub Label1_Click()
frmCreateNew.lblID.Caption = Me.ListView1.SelectedItem.Text
frmCreateNew.txtName.Text = Me.ListView1.SelectedItem.ListSubItems(1).Text
frmCreateNew.Show
Unload Me
End Sub

Private Sub Label2_Click()
On Error GoTo err
Set rs = New ADODB.Recordset
With rs
    .Open "Select * from [User] where id like '" & Me.ListView1.SelectedItem.Text & "'", cn, 1, 2
    If rs.EOF = False Then
        .Delete
        MsgBox "Successfully deleted.", vbInformation
    End If
End With
ViewRec
Exit Sub
err:
    MsgBox err.Description, vbCritical
End Sub

Private Sub Label3_Click()
frmMain.Show
Unload Me
End Sub

Private Sub lbl1_Click()
sName = Me.ListView1.SelectedItem.ListSubItems(1).Text
Unload Me
frmMain.Show
End Sub

Sub ViewRec()
With ListView1
    .ColumnHeaders.Clear
    .ListItems.Clear
    .ColumnHeaders.Add , , "ID", 0
    .ColumnHeaders.Add , , "Name", 4815
End With
Set rs = New ADODB.Recordset
With rs
    .Open "Select * from [User]", cn, 1, 2
    Do Until rs.EOF
        Set lst = Me.ListView1.ListItems.Add(, , rs.Fields("id").Value)
            lst.ListSubItems.Add , , rs.Fields("name").Value
    .MoveNext
    Loop
End With
End Sub
