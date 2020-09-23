VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Hound Website Ripper Beta 1"
   ClientHeight    =   5190
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7920
   Icon            =   "Strip-Ulr.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5190
   ScaleWidth      =   7920
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "E&xit"
      Height          =   350
      Left            =   6915
      TabIndex        =   16
      Top             =   1950
      Width           =   945
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Hide List"
      Height          =   350
      Left            =   6900
      TabIndex        =   14
      Top             =   1515
      Width           =   945
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Start"
      Height          =   350
      Left            =   6900
      TabIndex        =   11
      Top             =   615
      Width           =   945
   End
   Begin TabDlg.SSTab Main 
      Height          =   2565
      Left            =   75
      TabIndex        =   2
      Top             =   285
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   4524
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Web Ripper"
      TabPicture(0)   =   "Strip-Ulr.frx":27A2
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Image1(7)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblInfo"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label4"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Text1"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Option1"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Option2"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Option3"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Picture1"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).ControlCount=   10
      TabCaption(1)   =   "Save Options"
      TabPicture(1)   =   "Strip-Ulr.frx":27BE
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label3"
      Tab(1).Control(1)=   "txtPath"
      Tab(1).ControlCount=   2
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   555
         Left            =   5715
         ScaleHeight     =   37
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   39
         TabIndex        =   12
         Top             =   465
         Width           =   585
      End
      Begin VB.TextBox txtPath 
         Height          =   300
         Left            =   -73785
         TabIndex        =   10
         Text            =   "C:\Pages"
         Top             =   525
         Width           =   3180
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Achives"
         Height          =   195
         Left            =   3030
         TabIndex        =   8
         Top             =   915
         Width           =   915
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Pages"
         Height          =   195
         Left            =   2100
         TabIndex        =   7
         Top             =   915
         Width           =   750
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Images"
         Height          =   195
         Left            =   1170
         TabIndex        =   6
         Top             =   915
         Width           =   810
      End
      Begin VB.TextBox Text1 
         Height          =   300
         Left            =   1170
         TabIndex        =   4
         Text            =   "http://www.clipart.com"
         Top             =   450
         Width           =   3510
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Hound Ripper Beta 1"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   1950
         TabIndex        =   15
         Top             =   1215
         Width           =   3510
      End
      Begin VB.Label lblInfo 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0 Current Downloads"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   105
         TabIndex        =   13
         Top             =   2160
         Width           =   6495
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Save Items to :"
         Height          =   195
         Left            =   -74925
         TabIndex        =   9
         Top             =   555
         Width           =   1065
      End
      Begin VB.Image Image1 
         Height          =   915
         Index           =   7
         Left            =   150
         Picture         =   "Strip-Ulr.frx":27DA
         Top             =   1170
         Width           =   6165
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Rip Options:"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   870
         Width           =   870
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Website URL:"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   510
         Width           =   1005
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   150
      Left            =   7065
      Top             =   4830
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Download"
      Height          =   350
      Left            =   6900
      TabIndex        =   1
      Top             =   1065
      Width           =   945
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   7020
      Top             =   4185
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.ListBox List1 
      Height          =   2205
      ItemData        =   "Strip-Ulr.frx":14EA0
      Left            =   75
      List            =   "Strip-Ulr.frx":14EA2
      TabIndex        =   0
      Top             =   2895
      Width           =   6765
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      Index           =   1
      X1              =   0
      X2              =   1680
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   0
      X1              =   0
      X2              =   1680
      Y1              =   15
      Y2              =   15
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   6
      Left            =   45
      Picture         =   "Strip-Ulr.frx":14EA4
      Top             =   7995
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   5
      Left            =   60
      Picture         =   "Strip-Ulr.frx":151AE
      Top             =   7995
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   4
      Left            =   60
      Picture         =   "Strip-Ulr.frx":154B8
      Top             =   7995
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   3
      Left            =   60
      Picture         =   "Strip-Ulr.frx":157C2
      Top             =   7995
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   2
      Left            =   60
      Picture         =   "Strip-Ulr.frx":15ACC
      Top             =   7995
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   1
      Left            =   60
      Picture         =   "Strip-Ulr.frx":15DD6
      Top             =   7995
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   0
      Left            =   60
      Picture         =   "Strip-Ulr.frx":160E0
      Top             =   7995
      Visible         =   0   'False
      Width           =   480
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Web Hound Ripper Beta 1

' Name Ben Jones

'Ok welcome to my little website ripper. With this program you can rip pages form website on the net.
'Ok fisrt I like to say that this program will not rip anything of a site. It will only grab the pages with in the main index page or what ever page you point it to.
'There is also a little drawback with different site that you use this program on.
'It does not like site that are hosted by other companys eg site like geosites.
'so it won't download a pages if something like http://www.freesites/YourSite/index.htm it will not work un less you change the code.
'I am still to work on this part .It will only pull down pages that are on there own servers.
'Like yahoo.com or say something like http://www.clipart.com.

'Ok anyway I hope you like it anyway please vote if you like it.

Dim FileBuffer() As Byte
Dim ResData As String
Dim TCur As Long
Dim M_Height As Boolean
Dim OLD_Height As Long
Const CURSOR_FRAMES As Integer = 11
Private Sub GetHref(html As String, ResourceItem As String)

Dim strRefs() As String
Dim l As Long

strRefs = Split(html, ResourceItem & "=""")


For l = 1 To UBound(strRefs)
On Error Resume Next
    Data = Left(strRefs(l), InStr(1, strRefs(l), """") - 1)
   
   If InStr(Data, "mail") Then          ' We don't need emails
   ElseIf InStr(Data, "java") Then      'Don't want any java scipt
   ElseIf InStr(Data, "http://") Then   'This program only serachs one page
   ElseIf InStr(Data, "HTTP") Then      'This program only serachs one page
   Else
       
        Data = Replace(Data, "..", "")
        List1.AddItem Inet1.RemoteHost & "/" & Data ' Add in slash you can remove if you like
       End If
Next

Utils.RemoveDupeItems List1 ' Remove any repeated items
 
End Sub

Private Sub Command1_Click()
Dim Data As String
Dim StrUrl As String
Dim Answer

 If Utils.CheckFolderExits(txtPath) = 0 Then ' Checks for folder returns true if found
 
    Answer = MsgBox("Can't find folder " & txtPath & " do you wish to Create the Folder now", _
    vbYesNo)
    
    
    If Answer = vbYes Then
        MkDir txtPath ' Create the download folder
    End If
        
        If Answer = vbNo Then
        Exit Sub ' Skip rest of program
        Unload Form1 ' Exit program
    End If
        
        
    Else
    If Len(ResData) = 0 Then ' Gets the length of the resData
        MsgBox "Please select your download type [Images, Pages, Achives]" ' just displays the message asking you to select a download type
    Else
    
        List1.Clear ' Clear any items in the list box
        StrUrl = Text1.Text
        Data = Inet1.OpenURL(StrUrl, 0) ' Download the main page
        
        GetHref Data, ResData  'Pass the data to be progressed
        
        StrUrl = "" ' Clear out string
     End If
    End If
   
End Sub

Private Sub Command2_Click()
On Error Resume Next
Dim I As Integer
Dim Filename As String
Dim Url As String

Timer1.Enabled = True ' Set the timer on

For I = 0 To List1.ListCount - 1 ' Gets the nummbers of items in the list box

lblInfo.Caption = "Downloading  " & I + 1 & "  of " & List1.ListCount & "  Items"
    
List1.ListIndex = I
    Url = List1.List(I)
    Filename = Utils.GetName(List1.List(I)) ' Strips the filename for the link
    FileBuffer() = Inet1.OpenURL(Url, 1) ' Used for data
     
    Open txtPath & "\" & Filename For Binary Access Write As #1 ' Save the files
    Put #1, , FileBuffer()
    Close #1
    
Next

    Timer1.Enabled = False ' Ture timer off
    Picture1.Refresh
    lblInfo.Caption = List1.ListCount & "  Items have been downloaded"
    
    
    ' Clears out any stuff we don't need any more
    I = 0
    Filename = ""
    Url = ""
    
End Sub

Private Sub Command3_Click()

Select Case M_Height
    Case True
        Command3.Caption = "&Show List"
        Me.Height = List1.Top + 500
        List1.Visible = False
        M_Height = False
    Case False
        Command3.Caption = "&Hide List"
        Me.Height = OLD_Height
        List1.Visible = True
        M_Height = True
    End Select
    

End Sub

Private Sub Command4_Click()
 MsgBox "Hound Website Ripper by Ben Jones" & vbCrLf & "Please Vote if you like it", vbInformation
 End
 
End Sub

Private Sub Form_Load()
 ' Center the form on the screen
 
 Utils.CenterForm Form1
 OLD_Height = Form1.Height ' Get the forms Current height
 M_Height = True
 Command3_Click

' Used for the loading of the download progress picture
    TCur = Utils.LoadImage(App.hInstance, App.Path & "\Bussy.ani", ANI_CURSOR, 32, 32, LOADFROMFILE)

  
End Sub

Private Sub Form_Resize()
' Just adds a 3D line under the title bar

Line1(0).X2 = Form1.Width
Line1(1).X2 = Form1.Width
 
End Sub

Private Sub Form_Unload(Cancel As Integer)
 Utils.DestroyCursor TCur ' Destroy cursor allways do this
  Unload Form1
  
End Sub

Private Sub Option1_Click()
ResData = "src"

End Sub

Private Sub Option2_Click()
ResData = "href"

End Sub

Private Sub Option3_Click()
ResData = "href"

End Sub

Private Sub Timer1_Timer()
Static Counter As Integer
' useed for displaying Animated cursors in picture box

 If Counter = CURSOR_FRAMES Then Counter = 0: Picture1.Refresh
    Utils.DrawIconEx Picture1.hDC, 0, 0, TCur, 32, 32, Counter, &H0, &H3
     Counter = Counter + 1
     
End Sub
