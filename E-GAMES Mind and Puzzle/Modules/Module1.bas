Attribute VB_Name = "Module1"
' Used to set the shape of the form
Public Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hrgn As Long, ByVal bRedraw As Boolean) As Long
' Used to create the rounded rectangle region
Public Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
' Used to make the form draggable
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
' Also used to make the form draggable
Public Declare Function ReleaseCapture Lib "user32" () As Long
' Used to make the window always on top
Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, Y, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
' Various constants used by the above functions
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1
Public Const SWP_NOACTIVATE = &H10
Public Const SWP_SHOWWINDOW = &H40
Public Const TOPMOST_FLAGS = SWP_NOMOVE Or SWP_NOSIZE

Public Sub DoDrag(TheForm As Form)
' TheForm:  The form you want to start dragging
    
    ReleaseCapture
    SendMessage TheForm.hWnd, &HA1, 2, 0&
End Sub

Public Sub MakeWindow(TheForm As Form)
' TheForm:  The form you want to make graphical
    
   ' TheForm.BackColor = RGB(255, 217, 0)
    ' TheForm.BackColor = RGB(205, 201, 201)
    TheForm.BackColor = RGB(250, 244, 27)
     frmQuit.BackColor = RGB(245, 217, 0)
    'TheForm.Caption = TheForm!lblTitle.Caption
    'TheForm!lblTitle.Left = 16
    'TheForm!lblTitle.Top = 7
'
'    With TheForm!imgTitleLeft
'        .Top = 0
'        .Left = 0
'    End With
    
'    With TheForm!imgTitleRight
'        .Top = 0
'        .Left = (TheForm.Width / Screen.TwipsPerPixelX) - 19
'    End With
'
'    With TheForm!imgTitleMain
'        .Top = 0
'        .Left = 19
'        .Width = (TheForm.Width / Screen.TwipsPerPixelX) - 19
'    End With
'

    
    DoTransparency TheForm
    DoTransparency frmQuit
End Sub

Public Sub DoTransparency(TheForm As Form)
' TheForm:  The form you want to be rounded rectangle shape
    
    Dim TempRegions(6) As Long
    Dim FormWidthInPixels As Long
    Dim FormHeightInPixels As Long
    Dim A
    
' Convert the form's height and width from twips to pixels
    FormWidthInPixels = TheForm.Width / Screen.TwipsPerPixelX
    FormHeightInPixels = TheForm.Height / Screen.TwipsPerPixelY
    
' Make a rounded rectangle shaped region with the dimentions of the form
    A = CreateRoundRectRgn(0, 0, FormWidthInPixels, FormHeightInPixels, 24, 24)
    
' Set this region as the shape for "TheForm"
    A = SetWindowRgn(TheForm.hWnd, A, True)
End Sub



