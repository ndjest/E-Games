Attribute VB_Name = "Module2"
Public cn As New ADODB.Connection
Public rs As New ADODB.Recordset
Public intCategory As Integer
Public i_check As Integer
Public lst As MSComctlLib.ListItem
Public Const HiLyt = "{HOME}+{END}"
Public sCateg As String
Public sRound As String
Public sName As String
Public sScore As Integer
Public sLife As Integer
Public sQ As String
Public puzzScore As Long

Sub Main()
Set cn = New ADODB.Connection
    With cn
        .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Database\mind.mdb;Persist Security Info=False"
        .Open
    End With
    
     Set rs = New ADODB.Recordset
    rs.Open "Select * from [User] order by id", cn, 1, 2
    If rs.EOF = False Then
        rs.MoveLast
        sName = rs.Fields("name").Value
    End If
    'frmMain.Show
    frmSplash.Show
End Sub

Sub getLife()
sLife = 3
End Sub

Public Sub Delay(PauseTime)
    Dim Start, finish, TotalTime

    Start = Timer   ' Set start time.
    Do While Timer < Start + PauseTime
        DoEvents    ' Yield to other processes.
    Loop
End Sub

