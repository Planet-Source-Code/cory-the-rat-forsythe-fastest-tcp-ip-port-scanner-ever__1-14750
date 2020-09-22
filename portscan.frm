VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form portscan 
   BorderStyle     =   0  'None
   Caption         =   "PortScanner"
   ClientHeight    =   8370
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11745
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8370
   ScaleWidth      =   11745
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.Frame Frame1 
      BackColor       =   &H00BFBFBF&
      Caption         =   "Port Scanner"
      Height          =   7695
      Left            =   4800
      TabIndex        =   0
      Top             =   600
      Width           =   9975
      Begin VB.TextBox Text6a 
         Height          =   285
         Left            =   1560
         TabIndex        =   13
         Top             =   1080
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox Text6 
         Height          =   285
         Left            =   1560
         TabIndex        =   10
         Top             =   480
         Width           =   1095
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00BFBFBF&
         Caption         =   "IP Range"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   1080
         Width           =   1215
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00BFBFBF&
         Caption         =   "Single IP"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   480
         Value           =   1  'Checked
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         Appearance      =   0  'Flat
         Caption         =   "Start Scan"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   600
         MaskColor       =   &H000000FF&
         TabIndex        =   4
         Top             =   3840
         Width           =   1935
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   2640
         MaxLength       =   4
         TabIndex        =   3
         Text            =   "85"
         Top             =   2040
         Width           =   855
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   3720
         MaxLength       =   4
         TabIndex        =   2
         Text            =   "20"
         Top             =   1800
         Width           =   855
      End
      Begin VB.ListBox List1 
         Appearance      =   0  'Flat
         BackColor       =   &H00BFBFBF&
         Height          =   1200
         ItemData        =   "portscan.frx":0000
         Left            =   2760
         List            =   "portscan.frx":0002
         TabIndex        =   1
         Top             =   3240
         Width           =   3015
      End
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   255
         Index           =   0
         Left            =   2280
         TabIndex        =   5
         Top             =   1560
         Visible         =   0   'False
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Stop Ip:"
         Height          =   195
         Left            =   1800
         TabIndex        =   12
         Top             =   840
         Visible         =   0   'False
         Width           =   555
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Start Ip:"
         Height          =   255
         Left            =   1800
         TabIndex        =   11
         Top             =   240
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "From Port:"
         Height          =   195
         Left            =   1200
         TabIndex        =   7
         Top             =   2280
         Width           =   720
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "To Port:"
         Height          =   195
         Left            =   1560
         TabIndex        =   6
         Top             =   3120
         Width           =   570
      End
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Index           =   0
      Left            =   -480
      Top             =   -2520
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Winsock2 
      Left            =   -120
      Top             =   -480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
End
Attribute VB_Name = "portscan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' dim the variables so they may be used later
Public iport As Integer
Public upip
Public start_ipf
Public stop_ipf
Public start_port As Integer
Public stop_port As Integer
Dim start_ip As String
Dim stop_ip As String
Public go2 As Boolean
Public Sub TimeOut(Duration As Double)
    ' standard timeout sub, causes a short pause in the code
    Dim StartTime As Double, X As Integer
    StartTime = Timer
    Do While Timer - StartTime < Duration
        X = DoEvents()
    Loop
End Sub
Public Sub GetURLip(address As String)

End Sub
Public Sub verify(target As String)
End Sub

Private Sub Command1_Click()
GetURLip (Text6.Text)
If ProgressBar1.Count = 2 Then Unload ProgressBar1(1)
List1.clear
'Text6.Locked = True
portscan.MousePointer = 11
Command1.Enabled = False
start_port = Text1.Text
stop_port = Text2.Text
Call porit("127.0.0.1")
Command1.Enabled = True
portscan.MousePointer = 0
'Text6.Locked = False
status "Ready "
End Sub

Public Sub porit(ipip)
iport = start_port
Load ProgressBar1(1)
ProgressBar1(1).Visible = True
ProgressBar1(1).Left = ProgressBar1(0).Left
ProgressBar1(1).Top = ProgressBar1(0).Top
ProgressBar1(1).Max = portscan.stop_port
ProgressBar1(1).Min = portscan.start_port
If start_port = stop_port Then start_port = start_port - 1
For iport = start_port To stop_port
Load Winsock1(iport + 1)
ProgressBar1(1).Value = iport
If Winsock1(iport + 1).State <> sckClosed Then Winsock1(iport + 1).Close
Winsock1(iport + 1).Connect ipip, iport
status "Scan: " & iport
Next iport
End Sub

Private Sub Command2_Click()
Command2.Enabled = False
Winsock1(0).SendData Text7.Text
End Sub

Private Sub Command3_Click()
Command3.Enabled = False
Winsock1(0).Close
End Sub

Private Sub endbutt_Click()
End
End Sub

Private Sub Form_Load()
Open App.Path & "\scanlog.txt" For Append As #1
Write #1, Int(0)
Write #1, Int(0)
Write #1, Int(0)
Write #1, Int(0)
Close #1
End Sub

Private Sub Frame4_DragDrop(Source As Control, X As Single, Y As Single)

End Sub

Private Sub Frame2_DragDrop(Source As Control, X As Single, Y As Single)

End Sub

Private Sub List1_Click()
go2 = True
If Winsock2.State <> sckClosed Then Winsock2.Close
Winsock2.Connect Text6.Text, Left(List1.List(List1.ListIndex), 4)
List1.clear
End Sub

Private Sub List2_Click()

End Sub

Private Sub Winsock1_Close(Index As Integer)
List1.AddItem Winsock1(Index).RemotePort & "     Closed"
Close #1
End Sub

Private Sub Winsock1_DataArrival(Index As Integer, ByVal bytesTotal As Long)
Winsock1(Index).GetData Data, vbString

List1.AddItem Winsock1(Index).RemotePort & "    RETURNS:: " & Data
List1.AddItem ""
End Sub

Private Sub Winsock1_Connect(Index As Integer)

List1.AddItem Winsock1(Index).RemotePort & "     "
TimeOut 1.2
If Winsock1(Index).State <> sckClosed Then Winsock1(Index).Close
destroy Winsock1(Index)
End Sub

Private Sub Winsock1_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
portscan.MousePointer = 0
Command1.Enabled = True
'Text6.Locked = False
destroy Winsock1(Index)
End Sub
Public Sub destroy(thing As Object)
Unload thing
End Sub

Private Sub Winsock1_SendComplete(Index As Integer)
If Index = 0 Then Command2.Enabled = True
End Sub

Private Sub Winsock2_Connect()
If go2 = True Then

List1.AddItem ""
List1.AddItem "DATA OF: " & Winsock2.RemotePort
List1.AddItem "-------------"
List1.AddItem "HostName: " & Winsock2.RemoteHost
List1.AddItem "HostIP: " & Winsock2.RemoteHostIP
List1.AddItem "SckHandle: " & Winsock2.SocketHandle


TimeOut 0.5
Winsock2.Close
End If
End Sub

Private Sub Winsock2_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
List1.AddItem Description
End Sub
