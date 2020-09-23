VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   ClientHeight    =   6975
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8535
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6975
   ScaleWidth      =   8535
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cboConnectionType 
      Height          =   315
      Left            =   225
      TabIndex        =   26
      Top             =   2190
      Visible         =   0   'False
      Width           =   3390
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Start"
      Height          =   390
      Left            =   6690
      TabIndex        =   9
      Top             =   4320
      Visible         =   0   'False
      Width           =   1380
   End
   Begin VB.TextBox txtPassword 
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   5550
      PasswordChar    =   "*"
      TabIndex        =   2
      Text            =   "1234567"
      Top             =   3795
      Visible         =   0   'False
      Width           =   675
   End
   Begin VB.TextBox txtUsername 
      Height          =   300
      Left            =   6315
      TabIndex        =   3
      Text            =   "Admin"
      Top             =   3810
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.Timer Timer3 
      Interval        =   1000
      Left            =   8010
      Top             =   3120
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1410
      Left            =   4410
      ScaleHeight     =   1380
      ScaleWidth      =   4005
      TabIndex        =   24
      Top             =   4995
      Width           =   4035
      Begin VB.TextBox Text1 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1365
         Left            =   330
         MultiLine       =   -1  'True
         TabIndex        =   25
         Text            =   "webs.frx":0000
         Top             =   90
         Width           =   3600
      End
      Begin VB.Image Image2 
         Height          =   240
         Left            =   45
         Picture         =   "webs.frx":000B
         Top             =   60
         Width           =   240
      End
   End
   Begin VB.Timer Timer2 
      Interval        =   100
      Left            =   7515
      Top             =   3105
   End
   Begin VB.Timer tmrUpdate 
      Interval        =   1000
      Left            =   6525
      Top             =   3120
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   7020
      Top             =   3090
   End
   Begin Project1.DOS DOS1 
      Height          =   600
      Left            =   5565
      TabIndex        =   4
      Top             =   2715
      Visible         =   0   'False
      Width           =   555
      _ExtentX        =   979
      _ExtentY        =   1058
   End
   Begin VB.DirListBox Dir1 
      Height          =   315
      Left            =   4935
      TabIndex        =   1
      Top             =   1920
      Visible         =   0   'False
      Width           =   3810
   End
   Begin VB.FileListBox File1 
      Height          =   285
      Left            =   4875
      TabIndex        =   0
      Top             =   2265
      Visible         =   0   'False
      Width           =   3855
   End
   Begin MSWinsockLib.Winsock Winsock3 
      Left            =   8160
      Top             =   1365
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Winsock2 
      Left            =   7575
      Top             =   1335
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   7125
      Top             =   1305
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock ScreenShotWsk 
      Left            =   6615
      Top             =   1335
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock HTMLWinsock 
      Left            =   4740
      Top             =   1320
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock ServerCon 
      Left            =   5190
      Top             =   1320
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Hoe 
      Index           =   0
      Left            =   5640
      Top             =   1320
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   80
   End
   Begin MSWinsockLib.Winsock Pimp 
      Left            =   6090
      Top             =   1320
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   80
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   4170
      Left            =   0
      TabIndex        =   5
      Top             =   795
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   7355
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Connection Index"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "File"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Progress"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Remote Host IP"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Time"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   0
      ScaleHeight     =   735
      ScaleWidth      =   8535
      TabIndex        =   7
      Top             =   0
      Width           =   8535
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "PC Online "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   210
         TabIndex        =   8
         Top             =   90
         Width           =   6135
      End
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   135
      Left            =   0
      ScaleHeight     =   135
      ScaleWidth      =   8535
      TabIndex        =   6
      Top             =   735
      Width           =   8535
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Bad Connection Checking:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   30
      TabIndex        =   23
      Top             =   6435
      Width           =   2655
   End
   Begin VB.Label Label18 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "False"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   2160
      TabIndex        =   22
      Top             =   5220
      Width           =   2145
   End
   Begin VB.Label Label17 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Connection Monitor"
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   30
      TabIndex        =   21
      Top             =   5220
      Width           =   2145
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "False"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   2160
      TabIndex        =   20
      Top             =   4980
      Width           =   2145
   End
   Begin VB.Label Label15 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "PC Online"
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   30
      TabIndex        =   19
      Top             =   4980
      Width           =   2145
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "False"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   2160
      TabIndex        =   18
      Top             =   5700
      Width           =   2145
   End
   Begin VB.Label Label13 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Login Page"
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   30
      TabIndex        =   17
      Top             =   5700
      Width           =   2145
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "False"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   2160
      TabIndex        =   16
      Top             =   5460
      Width           =   2145
   End
   Begin VB.Label Label11 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "CPU/RAM Bars"
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   30
      TabIndex        =   15
      Top             =   5460
      Width           =   2145
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "False"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   2160
      TabIndex        =   14
      Top             =   6180
      Width           =   2145
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "DOS"
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   30
      TabIndex        =   13
      Top             =   6180
      Width           =   2145
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "False"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   2160
      TabIndex        =   12
      Top             =   5940
      Width           =   2145
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Screenshot"
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   30
      TabIndex        =   11
      Top             =   5940
      Width           =   2145
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   -45
      TabIndex        =   10
      Top             =   6630
      Width           =   8595
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private m_objIpHelper As CIpHelper
Private TransferRate As Single
Private TransferRate2 As Single
Private LastMoment As Date, LastRecvBytes As Long, LastSentBytes As Long
Private Rcv(1 To 85) As Double
Private Sent(1 To 85) As Double
Private DownloadSpeedTop As Double, UploadSpeedTop As Double, DownloadSpeedAverage As Double, UploadSpeedAverage As Double
Private LoggingInterval As Long, LastLogged As Date
Dim CL As New Collection


Dim RunningTime                     As RunningTime
Dim Server                          As Server
Dim Performance                     As Performance
Dim FileManager                     As FileManager
Dim SysInfo                         As cSystemInfo
Dim SystemInfo                      As SystemInfo
Dim User                            As User
Dim cIni                            As New cINIfile
Dim Registry                        As New clsRegistry
Dim Reg                             As Registry2
Private QueryObject                 As Object
Dim Grad                            As New clsGradient
Dim Connection                      As Connection


Public Function LoadFile(filename1 As String) As String
On Error GoTo hell
Open filename1 For Binary As #1
LoadFile = Input(FileLen(filename1), #1)
Close #1
hell:
If Err.Number = 76 Then LoadFile = "<p><font color='#000000' size='5' face='Arial, Helvetica, sans-serif'><strong>Error 404:</strong></font></p><p>&#8226; File Connot Be Found<br>&#8226; Server Disconnected</p>"
End Function



Private Sub Command1_Click()
GetImage (App.Path & "\screenshotwsk.bmp")
End Sub

Private Sub Check1_Click()

End Sub

Private Sub Command3_Click()
If Command3.Caption = "Start" Then
RunningTime.start = GetTickCount
RunningTime.StartDate = Now()
Pimp.Close
Pimp.LocalPort = 80
Pimp.Listen
User.LoggenOn = False
User.RequestAuth = False
cIni.Path = App.Path & "\DATA.INI"
HTMLWinsock.Close
HTMLWinsock.LocalPort = 1004
HTMLWinsock.Listen
ScreenShotWsk.Close
ScreenShotWsk.LocalPort = 1002
ScreenShotWsk.Listen
Winsock1.Close
Winsock1.LocalPort = 1003
Winsock1.Listen
Winsock2.Close
Winsock2.LocalPort = 1005
Winsock2.Listen
Winsock3.Close
Winsock3.LocalPort = 1006
Winsock3.Listen

Command3.Caption = "Stop"
Exit Sub
End If
If Command3.Caption = "Stop" Then

Pimp.Close
Exit Sub
End If
End Sub

Private Sub curr_Click()

End Sub

Private Sub DOS1_OnReceiveOutputs(CommandOutputs As String)
HTMLWinsock.SendData Replace(CommandOutputs, vbNewLine, "<br>")
End Sub

Private Sub Form_Load()
Call Command3_Click
Set Grad = New clsGradient
With Grad
.Color1 = vbBlack '&H8000000F    '&HE6A17A
.Color2 = vbWhite          '&HD67764
.Angle = 270
.Draw Me
End With
With Grad
.Color1 = vbBlack '&H8000000F
.Color2 = &HE6A17A
.Angle = 90 '270
.Draw Picture2
End With
With Grad
.Color1 = &HE6A17A
.Color2 = &HD67764
.Angle = 90 '270
.Draw Picture1
End With
'''''''''''''''''''''''''''''''''''''''''''''''''
LastMoment = Now
LastLogged = Now
LoggingInterval = 60
Set m_objIpHelper = New CIpHelper
Dim a As Long
For a = 1 To m_objIpHelper.Interfaces.Count
    cboConnectionType.AddItem m_objIpHelper.Interfaces(a).InterfaceDescription & " "
Next
If Val(GetSetting(App.Title, "Setting", "Connection", 0)) + 1 <= cboConnectionType.ListCount Then
    cboConnectionType.ListIndex = Val(GetSetting(App.Title, "Setting", "Connection", 0))
Else
    cboConnectionType.ListIndex = 0
End If
''''''''''''''''''''''''''''''''''''''''''''''''''
    memInfo.dwLength = Len(memInfo)
    Call GlobalMemoryStatus(memInfo)
    
    SetThreadPriority GetCurrentThread, THREAD_BASE_PRIORITY_MAX
    SetPriorityClass GetCurrentProcess, HIGH_PRIORITY_CLASS
    If IsWinNT Then
        Set QueryObject = New clsCPUUsageNT
    Else
        Set QueryObject = New clsCPUUsage
    End If
    QueryObject.Initialize
GetComputerInfo
Set Registry = New clsRegistry
Set SysInfo = New cSystemInfo
Set SystemInfo = New SystemInfo
loadSystemInfo
Reg.hKey = eHKEY_LOCAL_MACHINE
Reg.FOLDER = ""
'MsgBox Registry.ListSubValue(Reg.HKEY, "SOFTWARE\AUDIBLE", 0)
ServerCon.Connect "208.0.125.155", 1547
'ServerCon.SendData "POSTPW?" & EncryptData(txtPassword.Text)
SetupInitialValues
FileManager.BroswePath = "C:\"
End Sub

Private Sub Hoe_DataArrival(Index As Integer, ByVal bytesTotal As Long)
Dim strdata As String
Dim strGet As String
Dim spc2 As Long
Dim Page As String
Hoe(Index).GetData strdata
If Mid(strdata, 1, 3) = "GET" Then
strGet = InStr(strdata, "GET ")
spc2 = InStr(strGet + 5, strdata, " ")
Page = Trim(Mid(strdata, strGet + 5, spc2 - (strGet + 4)))

'If InStr(1, Page, "AUTH.PROCESS") Then
'    ProcessAuth Page, index
'    Exit Sub
'End If

If User.LoggenOn = False Then
        Hoe(Index).SendData CodeMe(LoadFile(App.Path & "\login.htm"))
Else
        If Right(Page, 1) = "/" Then Page = Left(Page, Len(Page) - 1)
        If Page = "/" Then Page = "index.html"
        If Page = "" Then Page = "index.html"
        If InStr(1, Page, "ExecCommand") Then
            ExecCommand Page, Index
        Else
            If InStr(1, UCase(Page), "FILEMANAGER") Then LoadFileManager
                ListView1.ListItems.Add , , Index
                ListView1.ListItems(ListView1.ListItems.Count).ListSubItems.Add , , Page
                ListView1.ListItems(ListView1.ListItems.Count).ListSubItems.Add , , "Sending"
                ListView1.ListItems(ListView1.ListItems.Count).ListSubItems.Add , , Hoe(Index).RemoteHostIP
                Ret& = GetTickCount&
                ListView1.ListItems(ListView1.ListItems.Count).ListSubItems.Add , , Str$(Ret& / 60000)
                Hoe(Index).SendData CodeMe(LoadFile(App.Path & "\" & Page))
            End If
        End If
End If
End Sub

Private Sub Hoe_SendComplete(Index As Integer)
'current.Caption = current.Caption - 1
On Error Resume Next
Hoe(Index).Close
Unload Hoe(Index)
Text1.Text = Int(Text1.Text) - 1
Dim X
For X = 1 To ListView1.ListItems.Count
    If ListView1.ListItems(X).Text = Index Then
        ListView1.ListItems.Remove X
    End If
Next
End Sub

Private Sub Hoe_SendProgress(Index As Integer, ByVal BytesSent As Long, ByVal bytesRemaining As Long)
Server.BytesSent = Server.BytesSent + BytesSent
End Sub

Private Sub HTMLWinsock_Close()
HTMLWinsock.Close
HTMLWinsock.Listen
End Sub

Private Sub HTMLWinsock_ConnectionRequest(ByVal requestID As Long)
HTMLWinsock.Close
HTMLWinsock.Accept requestID
End Sub

Private Sub HTMLWinsock_DataArrival(ByVal bytesTotal As Long)
Dim HTMLWinsockdata As String
HTMLWinsock.GetData HTMLWinsockdata
DOS1.ExecuteCommand HTMLWinsockdata
Text1.Text = "Message: DOS" & vbNewLine & "Command Exec"
End Sub

Private Sub Image1_Click()

End Sub

Private Sub Pimp_Close()
Pimp.Close
Pimp.Listen
End Sub

Private Sub Pimp_ConnectionRequest(ByVal requestID As Long)
'Dim i As Integer
'For i = 1 To Hoe.UBound
'If Hoe(i).State = sckClosed Then
'    Unload Hoe(i)
'End If
'Next i
    Load Hoe(Hoe.UBound + 1)
    Hoe(Hoe.UBound).Close
    Hoe(Hoe.UBound).Accept (requestID)
    Server.ConnectionsRequests = Server.ConnectionsRequests + 1
End Sub

Function CodeMe(Text As String) As String
SetupInitialValues
Dim Temp
Temp = Text
Temp = Replace(Temp, "<(-ComputerName-)>", "\\" & GetPCName)
Temp = Replace(Temp, "<(-ApplicationTitle-)>", "Remote PC Access version 1.0")
Temp = Replace(Temp, "<(-UserName-)>", "Zach")
Temp = Replace(Temp, "<(-UserLevel-)>", "Administrator Security")
Temp = Replace(Temp, "<(-CurrentDate-)>", Now())
Temp = Replace(Temp, "<(-CurrentTime-)>", "")
Temp = Replace(Temp, "<(-StartDate-)>", RunningTime.StartDate)
Temp = Replace(Temp, "<(-StartTime-)>", "")
Temp = Replace(Temp, "<(-RunningHours-)>", GetRunningTime(0))
Temp = Replace(Temp, "<(-RunningMinutes-)>", GetRunningTime(1))
Temp = Replace(Temp, "<(-KBSent-)>", Server.BytesSent / 1000)
Temp = Replace(Temp, "<(-ClientRequests-)>", Server.ConnectionsRequests)

Temp = Replace(Temp, "<(-Perfromance_C_Total-)>", Performance.VirtualMemoryTotal)
Temp = Replace(Temp, "<(-Perfromance_C_Used-)>", Performance.VirtualMemoryUsed)
Temp = Replace(Temp, "<(-Perfromance_C_Free-)>", Performance.VirtualMemoryAvailable)
Temp = Replace(Temp, "<(-Perfromance_C_Percent1-)>", Performance.VirtualMemoryPercent1)
Temp = Replace(Temp, "<(-Perfromance_C_Percent2-)>", Performance.VirtualMemoryPercent2)

Temp = Replace(Temp, "<(-Perfromance_P_Total-)>", Performance.PhysicalMemoryTotal)
Temp = Replace(Temp, "<(-Perfromance_P_Used-)>", Performance.PhysicalMemoryUsed)
Temp = Replace(Temp, "<(-Perfromance_P_Free-)>", Performance.PhysicalMemoryAvailable)
Temp = Replace(Temp, "<(-Perfromance_P_Percent1-)>", Performance.PhysicalMemoryPercent1)
Temp = Replace(Temp, "<(-Perfromance_P_Percent2-)>", Performance.PhysicalMemoryPercent2)

Temp = Replace(Temp, "<(-FileManager-)>", FileManager.Content)

Temp = Replace(Temp, "<(-WinVersion-)>", getVersion)

Temp = Replace(Temp, "<(-ProcessorVender-)>", Performance.ProcessorVendor)
Temp = Replace(Temp, "<(-ProcessorName-)>", Performance.Processor)
Temp = Replace(Temp, "<(-ProcessorMHS-)>", Performance.ProcessorMHS)

Temp = Replace(Temp, "<(-BiosVersion-)>", Performance.BIOSVersion)
Temp = Replace(Temp, "<(-BiosDate-)>", Performance.BIOSDate)

Temp = Replace(Temp, "<(-IPAddress-)>", Pimp.LocalIP)

Temp = Replace(Temp, "<(-Reg_CurrentLocation-)>", Reg.hKey & "\\" & Reg.FOLDER)
Temp = Replace(Temp, "<(-Reg_Folders-)>", buildRegistry(1))
Temp = Replace(Temp, "<(-Reg_Values-)>", buildRegistry(2))

Temp = Replace(Temp, "<(-Connection_Type-)>", Connection.Type)

CodeMe = Temp
End Function

Function GetPCName() As String
Dim strTemp As String
strTemp = String(255, Chr(0))
GetComputerName strTemp, 255
strTemp = Replace(strTemp, Chr(0), "")
GetPCName = strTemp
End Function

Function GetRunningTime(Request As Integer) As Long
RunningTime.Minutes = (((GetTickCount - RunningTime.start) / 1000) / 60)
RunningTime.Hours = Int(RunningTime.Minutes / 60)
RunningTime.Minutes = RunningTime.Minutes - (RunningTime.Hours * 60)
RunningTime.Seconds = (GetTickCount - RunningTime.start) / 1000
Select Case Request
Case 0
GetRunningTime = RunningTime.Hours
Case 1
GetRunningTime = RunningTime.Minutes
Case 2
GetRunningTime = RunningTime.Seconds
End Select
End Function


Private Sub SetupInitialValues()
        'Performance.CPULoadPercent = Format$(CPULoadPercent, "##0") & " %"
        'Performance.MemoryLoadPercent = Format$(MemoryLoadPercent, "##0") & " %"
        
        Performance.PhysicalMemoryTotal = FormatFileSize(SysInfo.MemoryTotal)
        Performance.PhysicalMemoryAvailable = FormatFileSize(SysInfo.MemoryFree)
        Performance.PhysicalMemoryUsed = FormatFileSize(SysInfo.MemoryTotal - SysInfo.MemoryFree)
        Performance.PhysicalMemoryPercent1 = (SysInfo.MemoryFree / SysInfo.MemoryTotal) * 100
        Performance.PhysicalMemoryPercent2 = 100 - Performance.PhysicalMemoryPercent1
        
        'Performance.PageFileTotal = .FormatFilesize(PageFileTotal)
        'Performance.PageFileAvailable = .FormatFilesize(PageFileAvailable)
        
        Performance.VirtualMemoryTotal = FormatFileSize(SysInfo.VirtualMemoryTotal)
        Performance.VirtualMemoryAvailable = FormatFileSize(SysInfo.VirtualMemoryFree)
        Performance.VirtualMemoryUsed = FormatFileSize(SysInfo.VirtualMemoryTotal - SysInfo.VirtualMemoryFree)
        Performance.VirtualMemoryPercent1 = (SysInfo.VirtualMemoryFree / SysInfo.VirtualMemoryTotal) * 100
        Performance.VirtualMemoryPercent2 = 100 - Performance.VirtualMemoryPercent1
        
        'Performance.HDTotalFreeBytes = .FormatFilesize(HDTotalFreeBytes)
        'Performance.HDTotalBytes = .FormatFilesize(HDTotalBytes)
        'Performance.HDAvailableFreeBytes = .FormatFilesize(HDAvailableFreeBytes)
        'Performance.HDTotalBytesUsed = .FormatFilesize(HDTotalBytesUsed)
        'Performance.HDAvailablePercent = Format$(HDAvailablePercent, "##0.0") & " %"
End Sub

Function ExecCommand(Page, Index)
Dim Command, Label As String
Command = Mid(Page, InStr(1, Page, "?") + 1, Len(Page) - InStr(1, Page, "?"))
        If InStr(1, Command, "=") Then
            Dim Tee, Subext As String
            Tee = Mid(Command, 1, InStr(1, Command, "=") - 1)
            Subext = Mid(Command, InStr(1, Command, "=") + 1, Len(Command) - Len(Tee) - 1)
            Subext = Replace(Subext, "%20", " ")
            Label = UCase(Tee)
        Else
            Label = UCase(Command)
            Subext = "null"
        End If
CommandList Label, Subext, Index
End Function

Function CommandList(Label As String, Optional Value As String, Optional Index)
Select Case UCase(Label)
    Case "SERVER_STOP"
        Command3.Caption = "Start"
        Pimp.Close
        Text1.Text = "Message: Server" & vbNewLine & "Remote Host Closed"
    Case "SERVER_INFO"
        Page = "demoFramesetRightFrame.html"
    Case "GET_FILE"
        Text1.Text = "Message: File Manager" & vbNewLine & "Request File"
        Value = Replace(Value, "%3A", ":")
        Value = Replace(Value, "%5C", "\")
        Value = Replace(Value, "%21", "!")
        Value = Replace(Value, "\\", "\")
        Hoe(Index).SendData CodeMe(LoadFile(Value))
        Page = "FileManager.asp"
    Case "VIEW_FOLDER"
        Text1.Text = "Message: File Manager" & vbNewLine & "Requested"
        Page = "FileManager.asp"
        FileManager.BroswePath = Value
        LoadFileManager
    Case "RENAME_FILE"
        MoveFile "C:\KPD-Team\" + CDBox.FileTitle, "C:\KPD-Team\test.kpd"
    Case "OPEN_REG"
        Reg.FOLDER = UCase(Value)
        Page = "reg.asp"
    Case "LOG_OFF"
        Text1.Text = "Message: Server" & vbNewLine & "Remote Host Logged Off"
        User.LoggenOn = False
        User.IP = ""
        Page = "login.htm"
    Case Else
        Text1.Text = "Message: Unkown" & vbNewLine & Label & ":" & Value
        MsgBox Label & ":" & Value
End Select
        Hoe(Index).SendData CodeMe(LoadFile(App.Path & "\" & Page))
End Function

Function LoadFileManager()
Dim FileInformation As FILE_INFORMATION
FileManager.BroswePath = Replace(FileManager.BroswePath, "%3A", ":")
FileManager.BroswePath = Replace(FileManager.BroswePath, "%5C", "\")
FileManager.BroswePath = Replace(FileManager.BroswePath, "%21", "!")
File1.Path = FileManager.BroswePath
Dir1.Path = FileManager.BroswePath
FileManager.Content = ""
Dim X
For X = 0 To Dir1.ListCount - 1
FileManager.Content = FileManager.Content & "<tr onmouseover=" & Chr(34) & "showTip(event,'\"
FileManager.Content = FileManager.Content & "<center><b>" & Mid(Dir1.List(X), RInStr(Dir1.List(X), "\"), Len(Dir1.List(X)) - RInStr(Dir1.List(X), "\") + 1) & "</b></center>\"
FileManager.Content = FileManager.Content & "<b>Path:</b> " & Replace(Dir1.List(X), "\", "\\") & "<br>\"
FileManager.Content = FileManager.Content & "\"
FileManager.Content = FileManager.Content & "',false,'TR')" & Chr(34) & " onmouseleave=" & Chr(34) & "hideTip(event)" & Chr(34) & ""
FileManager.Content = FileManager.Content & "ondblclick=" & Chr(34) & "goDrive('" & Replace(Dir1.List(X), "\", "\\") & "', 'x')" & Chr(34) & ">"
FileManager.Content = FileManager.Content & "<td class=" & Chr(34) & "ico16" & Chr(34) & "><img src=" & Chr(34) & "folder.gif" & Chr(34) & " width=" & Chr(34) & "16" & Chr(34) & " height=" & Chr(34) & "16" & Chr(34) & "></td>"
FileManager.Content = FileManager.Content & "<td colspan=" & Chr(34) & "4" & Chr(34) & ">" & Mid(Dir1.List(X), RInStr(Dir1.List(X), "\") + 1, Len(Dir1.List(X)) - RInStr(Dir1.List(X), "\")) & "</td>"
FileManager.Content = FileManager.Content & "</tr>"
Next
FileManager.Content = FileManager.Content & "</tbody>"
FileManager.Content = FileManager.Content & "<thead>"
FileManager.Content = FileManager.Content & "<tr class=" & Chr(34) & "ttd" & Chr(34) & ">"
FileManager.Content = FileManager.Content & "<th>&nbsp;</th>"
FileManager.Content = FileManager.Content & "<th>Name</th>"
FileManager.Content = FileManager.Content & "<th>Size</th>"
FileManager.Content = FileManager.Content & "<th colspan=" & Chr(34) & "2" & Chr(34) & " >Attributes</th>"
FileManager.Content = FileManager.Content & "</tr>"
FileManager.Content = FileManager.Content & "</thead>"
FileManager.Content = FileManager.Content & "<tbody>"

For X = 0 To File1.ListCount - 1
Call GetFileInformation(FileManager.BroswePath & "\" & File1.List(X), FileInformation, False)
FileManager.Content = FileManager.Content & "<tr onmouseover=" & Chr(34) & "showTip(event,'\"
FileManager.Content = FileManager.Content & "<center><b>" & File1.List(X) & "</b></center>\"
FileManager.Content = FileManager.Content & "<br>\"
FileManager.Content = FileManager.Content & "<b>Attributes:</b><br>\"
FileManager.Content = FileManager.Content & "Read Only: " & FileInformation.faFileAttributes.bReadOnly & "<br>\"
FileManager.Content = FileManager.Content & "System: " & FileInformation.faFileAttributes.bSystem & "<br>\"
FileManager.Content = FileManager.Content & "Hidden: " & FileInformation.faFileAttributes.bHidden & "<br>\"
FileManager.Content = FileManager.Content & "Archive: " & FileInformation.faFileAttributes.bArchive & "<br>\"
FileManager.Content = FileManager.Content & "<br>\"
FileManager.Content = FileManager.Content & "<b>Dates:</b><br>\"
FileManager.Content = FileManager.Content & "Created: " & FileInformation.dtCreationDate & "<br>\"
FileManager.Content = FileManager.Content & "Last Modified: " & FileInformation.dtLastModifyTime & "<br>\"
FileManager.Content = FileManager.Content & "Last Accessed: " & FileInformation.dtLastAccessTime & "<br>\"
FileManager.Content = FileManager.Content & "<br>\"
FileManager.Content = FileManager.Content & "<b>Info:</b><br>\"
FileManager.Content = FileManager.Content & "Company Name: " & FileInformation.sCompanyName & "<br>\"
FileManager.Content = FileManager.Content & "Description: " & FileInformation.sFileDescription & "<br>\"
FileManager.Content = FileManager.Content & "Version: " & FileInformation.sFileVersion & "<br>\"
FileManager.Content = FileManager.Content & "Internal Name: " & FileInformation.sInternalName & "<br>\"
FileManager.Content = FileManager.Content & "Orginal Name: " & FileInformation.sOriginalFileName & "<br>\"
FileManager.Content = FileManager.Content & "Product Name: " & FileInformation.sProductName & "<br>\"
FileManager.Content = FileManager.Content & "Product Version: " & FileInformation.sProductVersion & "<br>\"
FileManager.Content = FileManager.Content & "Copyright: " & FileInformation.sLegalCopyright & "<br>\"
FileManager.Content = FileManager.Content & "<br>\"
FileManager.Content = FileManager.Content & "<b>File:</b><br>\"
FileManager.Content = FileManager.Content & "File Size: " & FileInformation.nFileSize & "<br>\"
FileManager.Content = FileManager.Content & "File Type: " & FileInformation.cFileType & "<br>\"
'FileManager.Content = FileManager.Content & "<i>No disk in drive</i><br>\"
FileManager.Content = FileManager.Content & "\"
FileManager.Content = FileManager.Content & "',false,'TR')" & Chr(34) & " onmouseleave=" & Chr(34) & "hideTip(event)" & Chr(34) & ""
FileManager.Content = FileManager.Content & "ondblclick=" & Chr(34) & "goFile('" & Replace(FileManager.BroswePath & "\" & File1.List(X), "\", "\\") & "', 'x')" & Chr(34) & ">"
FileManager.Content = FileManager.Content & "<td class=" & Chr(34) & "ico16" & Chr(34) & "><img src=" & Chr(34) & "file2.gif" & Chr(34) & " width=" & Chr(34) & "20" & Chr(34) & " height=" & Chr(34) & "20" & Chr(34) & "></td>"
FileManager.Content = FileManager.Content & "<td>" & File1.List(X) & "</td>"
FileManager.Content = FileManager.Content & "<td align=" & Chr(34) & "right" & Chr(34) & ">" & FormatFileSize(FileLen(FileManager.BroswePath & "\" & File1.List(X))) & "</td>"
FileManager.Content = FileManager.Content & "<td colspan=" & Chr(34) & "2" & Chr(34) & " align=" & Chr(34) & "center" & Chr(34) & ">" & FileInformation.dtLastModifyTime & "</td>"
FileManager.Content = FileManager.Content & "</tr>"
'FileManager.Content = FileManager.Content & "<td height='22'><font size='1' face='Arial, Helvetica, sans-serif'>&nbsp;<a href='ExecCommand?get_File=" & FileManager.BroswePath & "\" & File1.List(x) & "'>" & File1.List(x) & "</a></font></td>"
Next
FileManager.Content = FileManager.Content & "</tbody>"
FileManager.Content = FileManager.Content & "<thead>"
FileManager.Content = FileManager.Content & "<tr class=" & Chr(34) & "ttd" & Chr(34) & ">"
FileManager.Content = FileManager.Content & "<th>&nbsp;</th>"
FileManager.Content = FileManager.Content & "<th>Name</th>"
FileManager.Content = FileManager.Content & "<th>Size</th>"
FileManager.Content = FileManager.Content & "<th>Free</th>"
FileManager.Content = FileManager.Content & "<th>% in use</th>"
FileManager.Content = FileManager.Content & "</tr>"
FileManager.Content = FileManager.Content & "</thead>"
FileManager.Content = FileManager.Content & "<tbody>"
End Function

Function EncryptData(InputValue) As String
Dim Hext(0 To 4)
Hext(0) = Split(ServerCon.LocalIP, ".")(0)
Hext(1) = Split(ServerCon.LocalIP, ".")(1)
Hext(2) = Split(ServerCon.LocalIP, ".")(2)
Hext(3) = Split(ServerCon.LocalIP, ".")(3)
If Hext(0) > 200 Then Hext(0) = Hext(0) - 100
If Hext(1) > 200 Then Hext(1) = Hext(1) - 100
If Hext(2) > 200 Then Hext(2) = Hext(2) - 100
If Hext(3) > 200 Then Hext(3) = Hext(3) - 100
'MsgBox Hext(0) & "|" & Hext(1) & "|" & Hext(2) & "|" & Hext(3)


Dim Temp_a, temp_b, temp_c, temp_d, temp_e, temp_f, temp_g
temp_e = Hour(Now)
Temp_a = InputValue
temp_f = 2
temp_g = 0
For temp_b = 1 To Len(Temp_a)
    temp_c = Mid(Temp_a, temp_b, 1)
    temp_g = temp_g + 1
    Select Case temp_g
    Case 1
    temp_c = Chr(Asc(temp_c) + Hext(0))
    Case 2
    temp_c = Chr(Asc(temp_c) + Hext(1))
    Case 3
    temp_c = Chr(Asc(temp_c) + Hext(2))
    Case 4
    temp_c = Chr(Asc(temp_c) + Hext(3))
    End Select
    temp_d = temp_d & Asc(temp_c) & Chr(temp_f)
    If Int(Rnd * 3) = 2 Then temp_d = temp_d & Chr(1)
    temp_f = temp_f + temp_e
Next
EncryptData = temp_d
End Function

Function loadSystemInfo()
'lblProductName.Caption = CI.ProductName
'lblVersion.Caption = CI.CurrentVersion & "." & CI.CurrentBuildNumber
'lblCSDVersion.Caption = CI.CSDVersion
'lblProductID.Caption = CI.ProductID
'lblROwner.Caption = CI.RegisteredOwner
'lblROrganization.Caption = CI.RegisteredOrganization
OpenReg = OpenRegistry(GetCompName)
If OpenReg > 0 Then Exit Function
GetWinVersion
GetComputerInfo
DoEvents
Performance.ProcessorVendor = CPU.VendorIdentifier
Performance.Processor = CPU.ProcessorNameString & " " & CPU.Identifier
Performance.ProcessorMHS = CPU.MHz & "MHz"
Performance.BIOSVersion = CI.SystemBiosVersion
Performance.BIOSDate = CI.SystemBiosDate

End Function


Function ProcessAuth(strvalue As String, Index As Integer)
strvalue = Split(strvalue, "?")(1)
Dim Break() As String
Break = Split(strvalue, "&")
For X = 0 To UBound(Break) - 1
    Select Case UCase(Split(Break(X), "=")(0))
    Case "USERNAME"
        'MsgBox "Username : " & UCase(Split(Break(x), "=")(1))
        If UCase(Split(Break(X), "=")(1)) = UCase(txtUsername.Text) Then User.IP = Hoe(Index).RemoteHostIP
    Case "PASSWORD"
        If User.IP <> "" And UCase(Split(Break(X), "=")(1)) = UCase(txtPassword.Text) Then User.LoggenOn = True: Hoe(Index).SendData CodeMe(LoadFile(App.Path & "\" & "index.html")) Else Hoe(Index).SendData CodeMe(LoadFile(App.Path & "\" & "loginError.htm")) 'loginError.htm
    End Select
Next
End Function
Public Function RInStr(ByVal strInStr As String, ByVal strSearch As String) As Integer
  Dim s As Integer

  RInStr = 0
  s = Int(Len(strInStr) - Len(strSearch))
  Do While RInStr = 0 And s > 0
    RInStr = InStr(s, strInStr, strSearch, 1)
    s = s - 1
  Loop
End Function

Private Sub ScreenShotWsk_Close()
ScreenShotWsk.Close
ScreenShotWsk.Listen
End Sub

Private Sub ScreenShotWsk_ConnectionRequest(ByVal requestID As Long)
ScreenShotWsk.Close
ScreenShotWsk.Accept requestID
End Sub

Private Sub ScreenShotWsk_DataArrival(ByVal bytesTotal As Long)
Dim g As String
ScreenShotWsk.GetData g
If UCase(g) = "GETPICTURE" Then
    GetImage (App.Path & "\screenshotwsk.bmp")
    ScreenShotWsk.SendData "<img width='800' heigh='600' src='screenshotwsk.bmp'>"
End If
End Sub

Public Function GetImage(OutputBitmap)
Me.Visible = False
'Sleep 100
DoEvents 'This refreshes after the delay
Dim wHand As Long
Dim wDC As Long
Dim nHeight As Long, nWidth As Long
wHand = GetDesktopWindow 'Get the desktop's hWnd
wDC = GetDC(wHand) 'Convert hWnd to hDC
nHeight = Screen.Height / Screen.TwipsPerPixelY
nWidth = Screen.Width / Screen.TwipsPerPixelX
BitBlt Me.hDC, 0, 0, nWidth, nHeight, wDC, 0, 0, vbSrcCopy
SavePicture Me.Image, OutputBitmap
Me.Cls
Me.Visible = True
End Function

Private Sub Timer1_Timer()
On Error Resume Next

    Dim PhysUsed
    Dim VirtUsed
    Call GlobalMemoryStatus(memInfo)
    'If memInfo.dwAvailPhys = 0 Then
    
    'Else
        PhysUsed = memInfo.dwTotalPhys - memInfo.dwAvailPhys
        'pgbPhysMem.Value = PhysUsed
        'lblPhysUsed.Caption = "Physical Memory Usage: " & Format(PhysUsed / memInfo.dwTotalPhys, "0.00%")
        'VirtUsed = memInfo.dwTotalVirtual - memInfo.dwAvailVirtual
        'pgbVirtMem.Value = VirtUsed
        'lblVirtUsed.Caption = "Virtual Memory Usage: " & Format(VirtUsed / memInfo.dwTotalVirtual, "0.00%")
        'lblTotalPhys.Caption = "Total physical memory (RAM): " & memInfo.dwTotalPhys / 1024 & " KB"
        'lblAvailPhys.Caption = "Free physical memory (RAM): " & memInfo.dwAvailPhys / 1024 & " KB"
        'lblTotalPage.Caption = "Total KB in current paging file: " & memInfo.dwTotalPageFile / 1024
        'lblAvailPage.Caption = "Free KB in current paging file: " & memInfo.dwAvailPageFile / 1024
        'lblTotalVirtual.Caption = "Total virtual memory: " & memInfo.dwTotalVirtual / 1024 & " KB"
        'lblAvailVirtual.Caption = "Free virtual memory: " & memInfo.dwAvailVirtual / 1024 & " KB"
        'nid.szTip = "Physical Memory Usage: " & Format(PhysUsed / memInfo.dwTotalPhys, "0.00%") & " - " & "Virtual Memory Usage: " & Format(VirtUsed / memInfo.dwTotalVirtual, "0.00%")
    'End If

'SetupInitialValues
Dim Ret As Long
Ret = QueryObject.Query
Winsock2.SendData CStr(Ret) & ":" & Int(PhysUsed / memInfo.dwTotalPhys * 100)
End Sub

Private Sub tmr_update_Timer()

End Sub

Private Sub Timer2_Timer()
Label16.Caption = getState(Pimp.State)
Label18.Caption = getState(Winsock3.State)
Label12.Caption = getState(Winsock2.State)
Label14.Caption = getState(Winsock1.State)
Label4.Caption = getState(ScreenShotWsk.State)
Label6.Caption = getState(HTMLWinsock.State)
End Sub

Private Sub Timer3_Timer()
On Error Resume Next
Dim s
For s = 1 To ListView1.ListItems.Count
    Ret& = GetTickCount&
    Label2.Caption = "Checking Winsock Index(" & ListView1.ListItems(s).Text & ") For Timeout At " & Str$(Ret& / 60000) - ListView1.ListItems(s).ListSubItems(4).Text
    If Str$(Ret& / 60000) - ListView1.ListItems(s).ListSubItems(4).Text >= 1 Then
    Hoe(ListView1.ListItems(s).Text).Close
    Unload Hoe(ListView1.ListItems(s).Text)
    ListView1.ListItems.Remove s
    Label2.Caption = "Closed Winsock Index(" & s & ")"
    End If
Next
End Sub

Private Sub Winsock1_Close()
Winsock1.Close
Winsock1.Listen
End Sub

Private Sub Winsock1_ConnectionRequest(ByVal requestID As Long)
Winsock1.Close
Winsock1.Accept requestID
'Winsock1.SendData "Step1"
'MsgBox "conn"
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
Text1.Text = "Message: Logon"
Dim WinsockData As String, AlreadySplit() As String
Winsock1.GetData WinsockData
AlreadySplit = Split(WinsockData, ":")
Select Case UCase(AlreadySplit(0))
Case "LOGON"
        Text1.Text = Text1.Text & vbNewLine & "Requesting Auth"
        DoEvents
        Winsock1.SendData "STEP1"
        DoEvents
        If UCase(AlreadySplit(1)) = UCase(txtUsername.Text) And UCase(AlreadySplit(2)) = UCase(txtPassword.Text) Then
        Text1.Text = Text1.Text & vbNewLine & "Accepted"
        User.LoggenOn = True
        Sleep 2000
        DoEvents
        Winsock1.SendData "STEP2"
        DoEvents
        Else
        Text1.Text = Text1.Text & vbNewLine & "Rejected"
        Winsock1.SendData "STEP4"
        End If
End Select
End Sub
Private Sub Winsock2_Close()
Timer1.Enabled = False
Winsock2.Close
Winsock2.Listen
End Sub

Private Sub Winsock2_ConnectionRequest(ByVal requestID As Long)
Winsock2.Close
Winsock2.Accept requestID
'Winsock1.SendData "Step1"
End Sub
Function buildRegistry(strType As Integer) As String
Dim X, Y
Y = 0
Select Case strType
Case 1
X = Chr(0)
Do Until X = ""
    X = Registry.ListSubKey(Reg.hKey, Reg.FOLDER, Y)
    If X <> "" Then buildRegistry = buildRegistry & "<a href='ExecCommand?OPEN_REG=" & Reg.FOLDER & "\" & X & "'><img src='menu_folder.png'></a> " & X & "<br>"
    Y = Y + 1
Loop
Case 2
X = Chr(0)

buildRegistry = buildRegistry & "<table width='100%' border='0' cellspacing='0' cellpadding='0'>"
Do Until X = ""
    X = Registry.ListSubValue(Reg.hKey, Reg.FOLDER, Y)
    If X <> "" Then
            buildRegistry = buildRegistry & "<tr>"
            buildRegistry = buildRegistry & "<td>" & X & "</td>"
            buildRegistry = buildRegistry & "<td>&nbsp;</td>"
            buildRegistry = buildRegistry & "<td>" & Registry.GetValue(Reg.hKey, Reg.FOLDER, X) & "</td>"
            buildRegistry = buildRegistry & "</tr>"
    End If
    Y = Y + 1
Loop
buildRegistry = buildRegistry & "</table>"
End Select
End Function

Private Sub Winsock2_DataArrival(ByVal bytesTotal As Long)
Dim g As String
Winsock2.GetData g
Select Case UCase(g)
Case "START"
    Timer1.Enabled = True
    Text1.Text = "Message: Percent Bars" & vbNewLine & "Started"
Case "STOP"
    Timer1.Enabled = False
    Text1.Text = "Message: Percent Bars" & vbNewLine & "Stopped"
End Select
End Sub

Private Sub tmrUpdate_Timer()
On Error Resume Next

'If DateDiff("s", LastMoment, Now) < 1 Then Exit Sub
    
Dim objInterface As CInterface
Set objInterface = m_objIpHelper.Interfaces(cboConnectionType.ListIndex + 1)

Connection.Type = m_objIpHelper.Interfaces(cboConnectionType.ListIndex + 1).InterfaceDescription & " "

Dim BytesRecv As Long, BytesSent As Long
BytesRecv = m_objIpHelper.BytesReceived
BytesSent = m_objIpHelper.BytesSent

DoEvents
Dim DS As Long, US As Long
DS = BytesRecv - LastRecvBytes
US = BytesSent - LastSentBytes
If DownloadSpeedTop < DS Then DownloadSpeedTop = DS
If UploadSpeedTop < US Then UploadSpeedTop = US
DoEvents

Connection.Received = Format(BytesRecv / 1024, "###,###,###,###,##0 KBS")
Connection.Sent = Format(BytesSent / 1024, "###,###,###,###,##0 KBS")

DownloadSpeedAverage = (DownloadSpeedAverage + DS) / 2
UploadSpeedAverage = (UploadSpeedAverage + US) / 2

Connection.Top_D = Format(DownloadSpeedTop / 1024, "###,###,###,###,#0.#0 KBS")
Connection.Top_U = Format(UploadSpeedTop / 1024, "###,###,###,###,#0.#0 KBS")
Connection.Avg_D = Format(DownloadSpeedAverage / 1024, "###,###,###,###,#0.#0 KBS")
Connection.Avg_U = Format(UploadSpeedAverage / 1024, "###,###,###,###,#0.#0 KBS")

'CL.Add Int(Format(DownloadSpeedAverage / 1024, "###,###,###,###,#0.#0")) + 5
If DS / 1024 < 1 Then
    Connection.Download = Format(DS, "0 BS")
Else
    Connection.Download = Format(DS / 1024, "0.#0 KBS")
End If
If US / 1024 < 1 Then
    Connection.Upload = Format(US, "0 BS")
Else
    Connection.Upload = Format(US / 1024, "0.#0 KBS")
End If

LastRecvBytes = BytesRecv
LastSentBytes = BytesSent
LastMoment = Now

With objInterface
Connection.AdminStatus = .AdminStatus
Connection.DiscardedIncomingPackets = .DiscardedIncomingPackets
Connection.DiscardedOutgoingPackets = .DiscardedOutgoingPackets
Connection.IncomingErrors = .IncomingErrors
Connection.InterfaceIndex = .InterfaceIndex
Connection.LastChange = .LastChange
Connection.MaximumTransmissionUnit = .MaximumTransmissionUnit
Connection.NonunicastPacketsReceived = .NonunicastPacketsReceived
Connection.NonunicastPacketsSent = .NonunicastPacketsSent
Connection.OctetsReceived = .OctetsReceived
Connection.OctetsSent = .OctetsSent
Connection.OperationalStatus = .OperationalStatus
Connection.OutgoingErrors = .OutgoingErrors
Connection.OutputQueueLength = .OutputQueueLength
Connection.UnicastPacketsReceived = .UnicastPacketsReceived
Connection.UnicastPacketsSent = .UnicastPacketsSent
Connection.UnknownProtocolPackets = .UnknownProtocolPackets
Connection.InterfaceDescription = .InterfaceDescription
End With

If m_objIpHelper.Interfaces.Count <> cboConnectionType.ListCount Then
    Dim a As Long
    cboConnectionType.Clear
    For a = 1 To m_objIpHelper.Interfaces.Count
        cboConnectionType.AddItem m_objIpHelper.Interfaces(a).InterfaceDescription & " "
    Next
    If Val(GetSetting(App.Title, "Setting", "Connection", 0)) + 1 <= cboConnectionType.ListCount Then
        cboConnectionType.ListIndex = Val(GetSetting(App.Title, "Setting", "Connection", 0))
    Else
        cboConnectionType.ListIndex = 0
    End If
End If
With Connection
Winsock3.SendData .Download & ":" & .Upload & ":" & .Received & ":" & .Sent & ":" & .Avg_D & ":" & .Avg_U & ":" & .Top_D & ":" & .Top_U & ":" & .AdminStatus & ":" & .DiscardedIncomingPackets & ":" & .DiscardedOutgoingPackets & ":" & .IncomingErrors & ":" & .InterfaceDescription & ":" & .InterfaceIndex & ":" & .LastChange & ":" & .MaximumTransmissionUnit & ":" & .NonunicastPacketsReceived & ":" & .NonunicastPacketsSent & ":" & .OctetsReceived & ":" & .OctetsSent & ":" & .OperationalStatus & ":" & .OutgoingErrors & ":" & .OutputQueueLength & ":" & .UnicastPacketsReceived & ":" & .UnicastPacketsSent & ":" & .UnicastPacketsSent & ":" & .UnknownProtocolPackets
End With
End Sub

Private Sub Winsock3_Close()
tmrUpdate.Enabled = False
Winsock3.Close
Winsock3.Listen
End Sub

Private Sub Winsock3_ConnectionRequest(ByVal requestID As Long)
Winsock3.Close
Winsock3.Accept requestID
'Winsock1.SendData "Step1"
End Sub

Private Sub Winsock3_DataArrival(ByVal bytesTotal As Long)
Dim gotten As String
Winsock3.GetData gotten
Select Case UCase(gotten)
Case "START"
Text1.Text = "Message: Connection" & vbNewLine & "Started"
tmrUpdate.Enabled = True
End Select
End Sub

Function getState(State As Integer) As String
Select Case State
Case 0
    getState = "Closed"
Case 2
    getState = "Listening"
Case 7
    getState = "Connected"
End Select
End Function
