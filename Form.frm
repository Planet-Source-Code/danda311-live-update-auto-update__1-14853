VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Live Update"
   ClientHeight    =   3015
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6615
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   6615
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox List1 
      Height          =   2055
      IntegralHeight  =   0   'False
      ItemData        =   "Form.frx":0000
      Left            =   4680
      List            =   "Form.frx":0002
      TabIndex        =   3
      Top             =   360
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Height          =   2055
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Text            =   "Form.frx":0004
      Top             =   360
      Width           =   4455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Download Updates"
      Enabled         =   0   'False
      Height          =   345
      Left            =   4680
      TabIndex        =   1
      Top             =   2520
      Width           =   1815
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   120
      Top             =   2400
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Check for Update"
      Height          =   345
      Left            =   2880
      TabIndex        =   0
      Top             =   2520
      Width           =   1695
   End
   Begin VB.Label Label2 
      Caption         =   "Available Update:"
      Height          =   255
      Left            =   4680
      TabIndex        =   5
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Information:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public DLWhat As String

Private Function HyperJump(ByVal URL As String) As Long
    HyperJump = ShellExecute(0&, vbNullString, URL, vbNullString, vbNullString, vbNormalFocus)
End Function

Private Sub Command1_Click()
On Error GoTo Skip
'This function assume files "application.ver", "news.txt" and "app.zip"
'on server http://danda311.50megs.com/app (change "http://danda311.50megs.com/app" to your server name and path)
'Inspect contain of files "news.txt" and "application.ver" at examples
Dim Version As String, News As String
Dim Site As String

'Website and Directory.  For this we will use one that I have created!
Site = "http://danda311.50megs.com/app/"

'Clear Text Box and listbox
List1.Clear
Text1.Text = ""



    Me.MousePointer = 11
    'now assign content of file app.ver to variable Version
    Version = Inet1.OpenURL(Site & "app.txt")
    'You can try this function on Your Hard Drive, but You must change adresses:
    'for example: "file://c:\path\app.ver"
    If Trim(Version) = "" Then GoTo Skip
    
    If Trim(Version) > CDbl(App.Major & "." & App.Minor & App.Revision) Then
        'Add's Version to list to show user
        List1.AddItem "App v" & CDbl(Version)
        'assigns where to download (EX.  www.abc.com/app31.zip) It gets rid of the   .   in 1.3
        'One problem is that my server wont allow you to download it from there!
        DLWhat = Site & "app" & Replace(Version, ".", "") & ".zip"
    Else
        List1.AddItem "No New Version's!"
        DLWhat = "none"
    End If
    
    'Download news and put it into text1, Fix Enter problem
    Text1.Text = Replace(Inet1.OpenURL(Site & "news.txt"), Chr(10), vbCrLf)
    
    'Enable Command2 if information is downloaded
    Command2.Enabled = True
    Me.MousePointer = 0
Exit Sub


'If no information is sent or they are not online
Skip:
    Text1.FontSize = 12
    Text1.FontBold = True
    Text1.Text = "Error Connecting to App's Website!"
    Me.MousePointer = 0
    Exit Sub
End Sub

Private Sub Command2_Click()

'if download is available download it!
If DLWhat <> "none" Then
HyperJump DLWhat
Else
MsgBox "You Are Up To Date or Have Not Checked!"
End If
End Sub

Private Sub Form_Load()
    Me.Caption = "Live Update " & App.Major & "." & App.Minor
End Sub
