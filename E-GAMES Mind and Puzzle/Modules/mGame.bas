Attribute VB_Name = "mGame"

Option Explicit

Public Type udtGame

    ' Canvas properties.
    Canvas          As PictureBox
    Rows            As Integer
    Cols            As Integer
    Size            As Integer
    BGColor         As Long
    
    ' Score properties.
    Level           As Integer
    Score           As Long
    Words           As Long
    
    ' Audio properties.
    Music           As Long
    MusicFile       As String
    Effects         As Long
    
    ' Game-play properties.
    ReverseSearch   As Long
    
End Type

Public g_Game   As udtGame

Public Sub Start()

    Call LoadSettings(App.Path & "\Data\Settings.ini")
    Call fGame.InitalizeGUI
    
    ' Reset game properties.
    g_Game.Score = 0
    g_Game.Words = 0
    
    ' Initialize the block array.
    Call mBlocks.Initialize
        
    ' ToDo: Nix the default arrow cursor during game play (over game field).
    g_Game.Canvas.SetFocus
    
    Call mSound.PlayMusic
    

    fGame.tmrPlay.Enabled = Not fGame.tmrPlay.Enabled
    
End Sub

Public Sub ApplySettings()

' Apply settings that refer to controls (i.e. PictureBox).

    With g_Game

        .Canvas.Cls

        ' Set the picturebox properties.
        .Canvas.BackColor = .BGColor
        .Canvas.Width = .Cols * .Size
        .Canvas.Height = .Rows * .Size

        ' Set the game level.
        pSetLevel .Level

    End With

End Sub

Public Sub LoadSettings(ByVal sFile As String)

' Load game settings from INI file.

Dim sSect   As String

    With g_Game
        
        sSect = "Game"
    
        .Rows = mINI.GetValue(sSect, "Rows", sFile, "12")
        .Cols = mINI.GetValue(sSect, "Cols", sFile, "8")
        .Size = mINI.GetValue(sSect, "Size", sFile, "50")
        .BGColor = mINI.GetValue(sSect, "BGColor", sFile, "0")
        .Level = mINI.GetValue(sSect, "Level", sFile, "5")
        
        sSect = "Audio"
        
        .Music = mINI.GetValue(sSect, "Music", sFile, "1")
        .MusicFile = mINI.GetValue(sSect, "MusicFile", sFile, "")
        .Effects = mINI.GetValue(sSect, "Effects", sFile, "1")
        
    End With
    
End Sub

Public Sub SaveSettings(ByVal sFile As String)

' Save settings to INI file.  Called when user presses "OK" on fSettings.frm.
' Perhaps create a mSettings.bas?

Dim sSect   As String

    With g_Game
    
        sSect = "Game"
    
        Call mINI.PutValue(sSect, "Rows", .Rows, sFile)
        Call mINI.PutValue(sSect, "Cols", .Cols, sFile)
        Call mINI.PutValue(sSect, "Size", .Size, sFile)
        Call mINI.PutValue(sSect, "BGColor", .BGColor, sFile)
        Call mINI.PutValue(sSect, "Level", .Level, sFile)
        
        sSect = "Audio"
        
        Call mINI.PutValue(sSect, "Music", .Music, sFile)
        Call mINI.PutValue(sSect, "MusicFile", .MusicFile, sFile)
        Call mINI.PutValue(sSect, "Effects", .Effects, sFile)
    
    End With

End Sub

Private Sub pSetLevel(ByVal iLevel As Integer)

' Level    :    1,    2,    3,   4,   5,   6,   7,  8.
' Interval : 1500, 1250, 1000, 750, 500, 250, 100, 50.
    
    Select Case iLevel
    
        Case Is < 7
            fGame.tmrPlay.Interval = 1500 - ((iLevel - 1) * 250)
            
        Case 7
            fGame.tmrPlay.Interval = 100
            
        Case 8
            fGame.tmrPlay.Interval = 50
            
    End Select
    
    g_Game.Level = iLevel
    
    fGame.lblLevel.Caption = iLevel & " "
    
End Sub

Public Sub SetScore(ByVal lAmount As Long)

    With g_Game
    
        .Score = .Score + lAmount
        
        If .Score > (.Level * 1000) Then
            
            If .Level < 8 Then Call pSetLevel(.Level + 1)
            
        End If
        
        fGame.lblScore.Caption = Format$(.Score, "#,#") & " "

    End With

End Sub

Public Sub Lose()

    Call mSound.StopMusic

    fGame.tmrPlay.Enabled = False

    If mScores.IsValid(g_Game.Score) Then
    
        'sName = InputBox("Congrats!  You are just nerdy enough to make it onto the high score list!", "We have a winner!")
        frmLose.Label1.Caption = " G A M E   O V E R " & vbCr & "Player Name: " & UCase(sName) & vbCr & "Score: " & fGame.lblScore.Caption
        frmLose.Show
        Call mScores.Add(g_Game.Score, Left$(sName, 10), g_Game.Level, g_Game.Words)
        
        'fHiScores.Show vbModal, fGame
        
    Else
        frmLose.Label1.Caption = " G A M E   O V E R " & vbCr & "Player Name: " & UCase(sName) & vbCr & "Score: " & fGame.lblScore.Caption
        
    End If
    
    fGame.lblPlay.Caption = "PLAY"
    
End Sub
