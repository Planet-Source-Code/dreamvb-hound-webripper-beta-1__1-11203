Attribute VB_Name = "Utils"
Declare Function DrawIconEx Lib "user32" (ByVal hDC As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyWidth As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Long
Declare Function LoadImage Lib "user32" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpsz As String, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long
Declare Function DestroyWindow Lib "user32" (ByVal hWnd As Long) As Long
Declare Function DestroyCursor Lib "user32" (ByVal hCursor As Long) As Long
Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long

Public Const LB_FINDSTRINGEXACT = &H1A2
Public Const ANI_CURSOR = 2
Public Const LOADFROMFILE = &H10
Function CenterForm(Frm As Form)
' Centers a form

With Frm
    .Top = (Screen.Height - Frm.Height) / 2
    .Left = (Screen.Width - Frm.Width) / 2
End With

End Function
Function RemoveDupeItems(LstBox As ListBox) As Integer
Dim Counter As Integer
Dim StrLine As String
Dim XPos As Integer
Dim YPos As Integer

For Counter = 0 To LstBox.ListCount - 1
Do
    DoEvents
    StrLine = LstBox.List(Counter)
    XPos = SendMessageByString(LstBox.hWnd, LB_FINDSTRINGEXACT, Counter, StrLine)
    YPos = SendMessageByString(LstBox.hWnd, LB_FINDSTRINGEXACT, XPos + 1, StrLine)
        If YPos = -1 Or YPos = XPos Then Exit Do
            LstBox.RemoveItem YPos
Loop
Next Counter

    End Function
Function GetName(Filename) As String
' used to ripout filename form url

Dim StartPos, I As Integer
Dim NewFilename As String

For I = 1 To Len(Filename)
    ch = Mid(Filename, I, 1)
        If ch = "/" Then
            StartPos = I
        End If
Next

NewFilename = Mid(Filename, StartPos + 1, Len(Filename))
GetName = NewFilename
End Function
Function CheckFolderExits(FolderPath As String) As Integer
' This just checks if the Foldername has a backslash

If Right(FolderPath, 1) = "\" Then
    FolderPath = FolderPath
 Else
    FolderPath = FolderPath + "\"
 End If
 
' This will check if folder is here or not

If Dir(FolderPath, vbDirectory) = "" Then
    CheckFolderExits = 0
    Else
    CheckFolderExits = 1
End If

End Function
