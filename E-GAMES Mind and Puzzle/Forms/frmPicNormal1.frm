VERSION 5.00
Object = "{27395F88-0C0C-101B-A3C9-08002B2F49FB}#1.1#0"; "PICCLP32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmPicNormal1 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7905
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11115
   DrawWidth       =   8
   LinkTopic       =   "Form1"
   ScaleHeight     =   527
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   741
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox pic 
      BackColor       =   &H80000010&
      Height          =   1200
      Index           =   0
      Left            =   360
      MouseIcon       =   "frmPicNormal1.frx":0000
      MousePointer    =   99  'Custom
      ScaleHeight     =   76
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   76
      TabIndex        =   20
      Top             =   960
      Width           =   1200
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H001BF4FA&
      Height          =   7095
      Left            =   7680
      TabIndex        =   0
      Top             =   600
      Width           =   3060
      Begin VB.OptionButton Opt 
         Appearance      =   0  'Flat
         BackColor       =   &H00008000&
         Caption         =   "10*10"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   285
         Index           =   8
         Left            =   5265
         TabIndex        =   13
         Top             =   1200
         Width           =   1000
      End
      Begin VB.Timer Timer 
         Interval        =   1000
         Left            =   1920
         Top             =   0
      End
      Begin VB.OptionButton Opt 
         Appearance      =   0  'Flat
         BackColor       =   &H00008000&
         Caption         =   "06*06"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   285
         Index           =   4
         Left            =   3900
         TabIndex        =   10
         Top             =   1200
         Width           =   1000
      End
      Begin VB.OptionButton Opt 
         Appearance      =   0  'Flat
         BackColor       =   &H00008000&
         Caption         =   "07*07"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   285
         Index           =   5
         Left            =   5265
         TabIndex        =   9
         Top             =   300
         Width           =   1000
      End
      Begin VB.OptionButton Opt 
         Appearance      =   0  'Flat
         BackColor       =   &H00008000&
         Caption         =   "08*08"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   285
         Index           =   6
         Left            =   5265
         TabIndex        =   8
         Top             =   600
         Width           =   1000
      End
      Begin VB.OptionButton Opt 
         Appearance      =   0  'Flat
         BackColor       =   &H00008000&
         Caption         =   "09*09"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   285
         Index           =   7
         Left            =   5280
         TabIndex        =   7
         Top             =   900
         Width           =   1000
      End
      Begin VB.OptionButton Opt 
         Appearance      =   0  'Flat
         BackColor       =   &H00008000&
         Caption         =   "03*03"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   285
         Index           =   1
         Left            =   3900
         TabIndex        =   6
         Top             =   300
         Value           =   -1  'True
         Width           =   1000
      End
      Begin VB.OptionButton Opt 
         Appearance      =   0  'Flat
         BackColor       =   &H00008000&
         Caption         =   "04*04"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   285
         Index           =   2
         Left            =   3900
         TabIndex        =   5
         Top             =   600
         Width           =   1000
      End
      Begin VB.OptionButton Opt 
         Appearance      =   0  'Flat
         BackColor       =   &H00008000&
         Caption         =   "05*05"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   285
         Index           =   3
         Left            =   3900
         TabIndex        =   4
         Top             =   900
         Width           =   1000
      End
      Begin VB.ComboBox Combo1 
         BackColor       =   &H00008000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000009&
         Height          =   405
         ItemData        =   "frmPicNormal1.frx":030A
         Left            =   4440
         List            =   "frmPicNormal1.frx":032C
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1800
         Width           =   975
      End
      Begin VB.PictureBox picme 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00008000&
         BorderStyle     =   0  'None
         Height          =   3135
         Left            =   75
         ScaleHeight     =   209
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   193
         TabIndex        =   1
         Top             =   2805
         Width           =   2895
         Begin VB.Label finish 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Congratulations"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   14.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   495
            Left            =   240
            TabIndex        =   2
            Top             =   1560
            Visible         =   0   'False
            Width           =   2415
         End
         Begin VB.Image Image1 
            Height          =   3075
            Left            =   15
            MouseIcon       =   "frmPicNormal1.frx":035B
            MousePointer    =   99  'Custom
            Stretch         =   -1  'True
            ToolTipText     =   "Nabeel Hosny Cairo / 2007 Click to Exit"
            Top             =   15
            Width           =   2835
         End
      End
      Begin MSComDlg.CommonDialog CmDlg 
         Left            =   1200
         Top             =   0
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.ListBox Solve 
         Height          =   2985
         Left            =   1800
         TabIndex        =   12
         Top             =   3600
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.ListBox MoveD 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   270
         Left            =   840
         TabIndex        =   11
         ToolTipText     =   "DblClick To Solve Step By Step"
         Top             =   3645
         Width           =   855
      End
      Begin VB.Label Menu 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Back"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H001BF4FA&
         Height          =   375
         Left            =   195
         MouseIcon       =   "frmPicNormal1.frx":04AD
         MousePointer    =   99  'Custom
         TabIndex        =   19
         Top             =   6240
         Width           =   2640
      End
      Begin VB.Label SolvMe 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Give up"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H001BF4FA&
         Height          =   375
         Left            =   195
         MouseIcon       =   "frmPicNormal1.frx":07B7
         MousePointer    =   99  'Custom
         TabIndex        =   18
         Top             =   1560
         Width           =   2640
      End
      Begin VB.Label NewGame 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "&Play Game"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H001BF4FA&
         Height          =   375
         Left            =   195
         MouseIcon       =   "frmPicNormal1.frx":0AC1
         MousePointer    =   99  'Custom
         TabIndex        =   17
         Top             =   480
         Width           =   2640
      End
      Begin VB.Image Image6 
         Height          =   1035
         Left            =   360
         Picture         =   "frmPicNormal1.frx":0DCB
         Top             =   120
         Width           =   2250
      End
      Begin VB.Image Image2 
         Height          =   1035
         Left            =   360
         Picture         =   "frmPicNormal1.frx":1043
         Top             =   1200
         Width           =   2250
      End
      Begin VB.Image Image7 
         Height          =   1035
         Left            =   480
         Picture         =   "frmPicNormal1.frx":12D9
         Top             =   5880
         Width           =   2250
      End
      Begin VB.Label lblElapsed 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   2400
         Width           =   2415
      End
      Begin VB.Shape Shape4 
         BorderColor     =   &H00000080&
         BorderWidth     =   4
         FillColor       =   &H00008000&
         FillStyle       =   0  'Solid
         Height          =   495
         Left            =   3240
         Shape           =   4  'Rounded Rectangle
         Top             =   1755
         Width           =   2505
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   360
         TabIndex        =   15
         Top             =   0
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Scramble"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   360
         Left            =   3075
         TabIndex        =   14
         Top             =   1800
         Width           =   1320
      End
      Begin VB.Shape Shape3 
         BorderColor     =   &H00000080&
         BorderWidth     =   4
         FillColor       =   &H00008000&
         FillStyle       =   0  'Solid
         Height          =   1335
         Left            =   3825
         Shape           =   4  'Rounded Rectangle
         Top             =   240
         Width           =   2505
      End
      Begin VB.Shape Shape2 
         FillColor       =   &H00C0C0C0&
         FillStyle       =   0  'Solid
         Height          =   375
         Left            =   3225
         Top             =   2325
         Width           =   2505
      End
      Begin VB.Image Image5 
         Height          =   7350
         Left            =   0
         Picture         =   "frmPicNormal1.frx":1551
         Top             =   -120
         Width           =   12555
      End
   End
   Begin PicClip.PictureClip Clip1 
      Left            =   0
      Top             =   4200
      _ExtentX        =   4948
      _ExtentY        =   7117
      _Version        =   393216
      Rows            =   5
      Cols            =   5
      Picture         =   "frmPicNormal1.frx":455D
   End
End
Attribute VB_Name = "frmPicNormal1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hrgn As Long, ByVal bRedraw As Boolean) As Long
Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Dim i As Integer, dummy As Integer
Dim X As Integer, Y As Integer
Dim mymy As Integer
Dim MixMoves, MaxPicControl As Integer
Dim MyTime As Long, Level As Integer, movment As Integer
Dim RealP As Integer, Mindix As Integer
Dim Grid(25, 25) As Integer
Dim A As Integer, B As Integer, C As Integer
Dim Old As Integer, po As Single
Dim DestX As Integer, DestY As Integer
Dim t(3) As String, dd As String
 Const MainPicHeight = 480
 Const MainPicWidth = 480

Private Sub Form_Load()
MakeWindow Me
 mymy = 0
End Sub

Private Sub Combo1_Click()
MixMoves = Val(Combo1.Text)
End Sub

Private Sub Menu_Click()
Unload Me
frmSelPicture.Show
End Sub

Private Sub NewGame_Click()
MoveD.Clear: Solve.Clear
Opt_Click (Val(Label1.Caption))
 ResetScreen
 Scramble
SolvMe.Enabled = True
'NewGame.Enabled = False
'Label2.Caption = MaxPicControl
End Sub

Private Sub Form_Activate()
SetWindowRgn Frame1.hWnd, CreateRoundRectRgn(0, 0, Frame1.Width, Frame1.Height, 50, 50), True
SetWindowRgn picme.hWnd, CreateRoundRectRgn(0, 0, picme.Width / Screen.TwipsPerPixelX, picme.Height / Screen.TwipsPerPixelY, 50, 50), True
SetWindowRgn Me.hWnd, CreateRoundRectRgn(0, 0, Me.Width / Screen.TwipsPerPixelX, Me.Height / Screen.TwipsPerPixelY, 75, 75), True
'SetWindowRgn MoveD.hwnd, CreateRoundRectRgn(0, 0, MoveD.Width / Screen.TwipsPerPixelX, MoveD.Height / Screen.TwipsPerPixelY, 15, 15), True
Combo1.ListIndex = 1 'pila ka moves

t(0) = "U"
t(1) = "D"
t(2) = "L"
t(3) = "R"



'Picture1.Picture = LoadPicture(App.Path & "\default.jpg")
Opt_Click (2) 'Level option ni xa
ResetScreen
End Sub

Private Sub Image1_Click()
End
End Sub

Private Sub SolvMe_Click()

Solve.Clear: dd = ""
For i = MoveD.ListCount - 1 To 0 Step -1
Select Case Right(MoveD.List(i), 1)
Case Is = "U": dd = "D"
Case Is = "D": dd = "U"
Case Is = "L": dd = "R"
Case Is = "R": dd = "L"
End Select
Solve.AddItem Left(MoveD.List(i), 2) & " " & dd
Next


MoveD.Clear
For i = 0 To Solve.ListCount - 1
MoveD.AddItem Solve.List(i)
movment = movment - 1
moveme
Next
Solve.Clear: MoveD.Clear: pic(MaxPicControl - 1).Visible = True
finish.Visible = True ': SolvMe.Enabled = False
sndPlaySound App.Path & "\pic\won.wav", 1
End Sub


Private Sub Opt_Click(Index As Integer)
Select Case Index
    Case 1
    Level = 0
    Case 2, 3, 4
    Level = 1
    Case 5, 6, 7, 8
    Level = 2
    End Select

Label1.Caption = Index
 MoveD.Clear: Solve.Clear
  movment = 0
    For i = 1 To MaxPicControl - 1
       Unload pic(i)
    Next i
    pic(0).Visible = True
    MaxPicControl = (Index + 2) ^ 2
        
    dummy = Sqr(MaxPicControl)
    
    For i = 0 To MaxPicControl - 1
        If i <> 0 Then Load pic(i)
        pic(i).Width = MainPicWidth / dummy
        pic(i).Height = MainPicHeight / dummy
        pic(i).Left = 8 + (i Mod dummy) * pic(i).Width
        pic(i).Top = (8 + (i \ dummy) * MainPicHeight / dummy) + 30
        pic(i).Visible = True
        pic(i).BorderStyle = 1
        pic(i).Tag = Format(pic(i).Left, "0000") & "  " & Format(pic(i).Top, "0000")
    Next i
   ' pic(MaxPicControl - 1).Move 1000, 1000
    Clip1.Rows = dummy
    Clip1.Cols = dummy
    Image1.Picture = Clip1.Picture


For X = 0 To MaxPicControl - 1
        Clip1.StretchX = pic(X).ScaleWidth
        Clip1.StretchY = pic(X).ScaleHeight
       pic(X) = Clip1.GraphicCell(X)
      Next X
        pic(MaxPicControl - 1).Visible = False
       PSet (408, 408), &HC991A1
       finish.Visible = False ': SolvMe.Enabled = False
       
       
 '       End If
End Sub



Private Sub pic_Click(Index As Integer)
'Label2.Caption = pic(Index).Index & "  " & pic(Index).Tag
If Point(pic(Index).Left + 1 * pic(0).Width / 2, pic(Index).Top - 1 * pic(0).Height / 2) = &HC991A1 Then dd = "U"
If Point(pic(Index).Left + 1 * pic(0).Width / 2, pic(Index).Top + 3 * pic(0).Height / 2) = &HC991A1 Then dd = "D"
If Point(pic(Index).Left - 1 * pic(0).Width / 2, pic(Index).Top + pic(0).Height / 2) = &HC991A1 Then dd = "L"
If Point(pic(Index).Left + 3 * pic(0).Width / 2, pic(Index).Top + pic(0).Height / 2) = &HC991A1 Then dd = "R"
If dd = "" Then
Exit Sub
Else
MoveD.AddItem pic(Index).Index & " " & dd
MoveD.Selected(MoveD.ListCount - 1) = True
movment = movment + 1
moveme
End If
dd = ""
checkend
End Sub
Sub checkend()
Dim mymy As Integer
For i = 0 To MaxPicControl - 1
If pic(i).Left = Left(pic(i).Tag, 4) And pic(i).Top = Right(pic(i).Tag, 4) Then mymy = mymy + 1
Next
If mymy = MaxPicControl Then
Solve.Clear: MoveD.Clear: pic(MaxPicControl - 1).Visible = True
finish.Visible = True ': SolvMe.Enabled = False
sndPlaySound App.Path & "\pic\won.wav", 1
mymy = 0
If finish.Visible = True Then
Set rs = New ADODB.Recordset
With rs
    .Open "Select * from [User] where [name] like '" & sName & "'", cn, 1, 2
    If .EOF = False Then
        .Fields("picpuzz").Value = 4
        .Update
    End If
End With
End If
puzzScore = puzzScore + 2000
       
       Set rs = New ADODB.Recordset
       rs.Open "Select * from mindscore where category like 'Picture' and [name] like '" & sName & "'", cn, 1, 2
'       If rs.EOF = False Then
'            rs.Fields("score").Value = puzzScore
'        Else
            rs.AddNew
            rs.Fields("name").Value = sName
            rs.Fields("category").Value = "Picture"
            rs.Fields("score").Value = puzzScore
            rs.Update
End If
End Sub

Private Sub Timer_Timer()
    Dim t As Date
    Dim M, s As Integer
    
    MyTime = MyTime + 1
       
    t = TimeSerial(0, 0, MyTime)
    lblElapsed.Caption = IIf(Level = 0, "Easy", IIf(Level = 1, "Normal", "Hard")) & " - " & Format(movment, "000") & " - " & Format(t, "hh:nn:ss")

End Sub
Sub Scramble()


X = dummy
Y = dummy

Old = 100

Randomize Timer
For A = 1 To MixMoves


Do

DestX = X
DestY = Y

Jer:

B = Int(Rnd * 4)
If Old + B = 1 Or Old + B = 5 Then GoTo Jer

Select Case B

Case Is = 0
    DestY = Y + 1

Case Is = 1
    DestY = Y - 1

Case Is = 2
    DestX = X + 1

Case Is = 3
    DestX = X - 1

End Select

Loop While Grid(DestX, DestY) = 0


Grid(X, Y) = Grid(DestX, DestY)
MoveD.AddItem Grid(DestX, DestY) - 1 & "  " & t(B)
MoveD.Selected(MoveD.ListCount - 1) = True

Grid(DestX, DestY) = 0

X = DestX
Y = DestY

Old = B
movment = movment + 1
moveme
Next A

End Sub
Sub ResetScreen()
Erase Grid


C = 0
For B = 1 To dummy
For A = 1 To dummy
C = C + 1
Grid(A, B) = C
Next A
Next B

X = dummy
Y = dummy
Grid(X, Y) = 0
End Sub

Sub moveme()
po = pic(0).Width / 2
If MoveD.ListCount = 0 Then Exit Sub
Me.Cls
i = MoveD.ListCount - 1
Mindix = (Mid$(MoveD.List(i), 1, 2))
Select Case Right(MoveD.List(i), 1)
Case Is = "U": RealP = pic(Mindix).Top - pic(0).Height
Case Is = "D": RealP = pic(Mindix).Top + pic(0).Height
Case Is = "L": RealP = pic(Mindix).Left - pic(0).Width
Case Is = "R": RealP = pic(Mindix).Left + pic(0).Width
End Select

Select Case Right(MoveD.List(i), 1)
Case Is = "U"
While (Int(pic(Mindix).Top) <> Int(RealP))
    pic(Mindix).Top = pic(Mindix).Top - 1: Refresh
Wend
PSet (pic(Mindix).Left + po, pic(Mindix).Top + 3 * po), &HC991A1

Case Is = "D"
While (Int(pic(Mindix).Top) <> Int(RealP))
    pic(Mindix).Top = pic(Mindix).Top + 1: Refresh
Wend
PSet (pic(Mindix).Left + po, pic(Mindix).Top - po), &HC991A1

Case Is = "L"
While (Int(pic(Mindix).Left) <> Int(RealP))
    pic(Mindix).Left = pic(Mindix).Left - 1: Refresh
Wend
PSet (pic(Mindix).Left + 3 * po, pic(Mindix).Top + po), &HC991A1

Case Is = "R"
While (Int(pic(Mindix).Left) <> Int(RealP))
    pic(Mindix).Left = pic(Mindix).Left + 1: Refresh
Wend
PSet (pic(Mindix).Left - po, pic(Mindix).Top + po), &HC991A1

End Select
sndPlaySound App.Path & "\pic\move.wav", 1

End Sub

Private Sub MoveD_DblClick()
dodo
End Sub

Sub dodo()
On Error Resume Next
po = pic(0).Width / 2
If MoveD.ListCount = 0 Then Exit Sub
Me.Cls
i = MoveD.ListCount - 1
Mindix = (Mid$(MoveD.List(i), 1, 2))
Select Case Right(MoveD.List(i), 1)
Case Is = "D": RealP = pic(Mindix).Top - pic(0).Height
Case Is = "U": RealP = pic(Mindix).Top + pic(0).Height
Case Is = "R": RealP = pic(Mindix).Left - pic(0).Width
Case Is = "L": RealP = pic(Mindix).Left + pic(0).Width
End Select

Select Case Right(MoveD.List(i), 1)
Case Is = "D"
While (Int(pic(Mindix).Top) <> Int(RealP))
    pic(Mindix).Top = pic(Mindix).Top - 1: Refresh
Wend
PSet (pic(Mindix).Left + po, pic(Mindix).Top + 3 * po), &HC991A1

Case Is = "U"
While (Int(pic(Mindix).Top) <> Int(RealP))
    pic(Mindix).Top = pic(Mindix).Top + 1: Refresh
Wend
PSet (pic(Mindix).Left + po, pic(Mindix).Top - po), &HC991A1

Case Is = "R"
While (Int(pic(Mindix).Left) <> Int(RealP))
    pic(Mindix).Left = pic(Mindix).Left - 1: Refresh
Wend
PSet (pic(Mindix).Left + 3 * po, pic(Mindix).Top + po), &HC991A1

Case Is = "L"
While (Int(pic(Mindix).Left) <> Int(RealP))
    pic(Mindix).Left = pic(Mindix).Left + 1: Refresh
Wend
PSet (pic(Mindix).Left - po, pic(Mindix).Top + po), &HC991A1

End Select
sndPlaySound App.Path & "\pic\move.wav", 1
MoveD.RemoveItem (MoveD.ListCount - 1)
If MoveD.ListCount = 0 Then
Solve.Clear: MoveD.Clear: pic(MaxPicControl - 1).Visible = True
finish.Visible = True: SolvMe.Enabled = False
sndPlaySound App.Path & "\pic\won.wav", 1
End If
End Sub


