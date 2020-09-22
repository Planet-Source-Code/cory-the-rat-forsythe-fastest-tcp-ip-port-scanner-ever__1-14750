VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form mainform 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00BFBFBF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "TCP/IP STATION"
   ClientHeight    =   6495
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8760
   ControlBox      =   0   'False
   Icon            =   "Form3.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "Form3.frx":0442
   Picture         =   "Form3.frx":074C
   ScaleHeight     =   6495
   ScaleWidth      =   8760
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      BackColor       =   &H00BFBFBF&
      Caption         =   "Quik Scan"
      Default         =   -1  'True
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   3960
      Width           =   975
   End
   Begin VB.ListBox List2 
      BackColor       =   &H00BFBFBF&
      Height          =   255
      ItemData        =   "Form3.frx":79EB
      Left            =   1800
      List            =   "Form3.frx":79ED
      TabIndex        =   10
      Top             =   1080
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BackColor       =   &H00BFBFBF&
      CausesValidation=   0   'False
      ForeColor       =   &H00000000&
      Height          =   1200
      ItemData        =   "Form3.frx":79EF
      Left            =   120
      List            =   "Form3.frx":79F1
      MouseIcon       =   "Form3.frx":79F3
      TabIndex        =   9
      Top             =   5160
      Width           =   5295
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Index           =   0
      Left            =   1680
      Top             =   360
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   6240
      Top             =   360
   End
   Begin VB.ListBox ops 
      Appearance      =   0  'Flat
      BackColor       =   &H00BFBFBF&
      CausesValidation=   0   'False
      Height          =   1785
      ItemData        =   "Form3.frx":7CFD
      Left            =   2640
      List            =   "Form3.frx":7CFF
      TabIndex        =   6
      Top             =   1920
      Width           =   3255
   End
   Begin VB.ListBox tasks 
      Appearance      =   0  'Flat
      BackColor       =   &H00BFBFBF&
      CausesValidation=   0   'False
      Height          =   1785
      ItemData        =   "Form3.frx":7D01
      Left            =   240
      List            =   "Form3.frx":7D03
      TabIndex        =   5
      Top             =   1920
      Width           =   2175
   End
   Begin VB.ListBox tasktimes 
      Appearance      =   0  'Flat
      BackColor       =   &H00BFBFBF&
      CausesValidation=   0   'False
      Height          =   1785
      ItemData        =   "Form3.frx":7D05
      Left            =   6120
      List            =   "Form3.frx":7D07
      TabIndex        =   4
      Top             =   1920
      Width           =   2295
   End
   Begin VB.CommandButton runbutt 
      BackColor       =   &H00BFBFBF&
      Caption         =   "Run task"
      Height          =   375
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3960
      Width           =   975
   End
   Begin VB.CommandButton addbutt 
      BackColor       =   &H00BFBFBF&
      Caption         =   "Add Scheduled Task"
      Height          =   375
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3960
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00BFBFBF&
      Caption         =   "Remove Task"
      Height          =   375
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3960
      Width           =   1215
   End
   Begin VB.CommandButton endbutt 
      Appearance      =   0  'Flat
      BackColor       =   &H00BFBFBF&
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8280
      MaskColor       =   &H00BFBFBF&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   240
      Width           =   255
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ".... More Stuff to go in here laters"
      Height          =   195
      Left            =   6120
      TabIndex        =   13
      Top             =   5280
      Width           =   2310
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Open Ports / Data Returned"
      Height          =   195
      Left            =   600
      TabIndex        =   12
      Top             =   4800
      Width           =   2010
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00E0E0E0&
      X1              =   120
      X2              =   8520
      Y1              =   4690
      Y2              =   4690
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   8520
      Y1              =   4680
      Y2              =   4680
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00:00:00 MM"
      Height          =   195
      Left            =   7200
      TabIndex        =   8
      Top             =   4080
      Width           =   945
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Current Time:"
      Height          =   195
      Left            =   6000
      TabIndex        =   7
      Top             =   4080
      Width           =   945
   End
End
Attribute VB_Name = "mainform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public itemno As Integer
Public Go As Boolean
Public GoS As Boolean
Public oldx
Public oldy
Public iporT



Public Sub TimeOut(Duration As Double)
    ' standard timeout sub, causes a short pause in the code
    Dim StartTime As Double, x As Integer
    StartTime = Timer
    Do While Timer - StartTime < Duration
        x = DoEvents()
    Loop
End Sub
Private Sub addbutt_Click()
addbutt.Enabled = False
List1.clear
List1.AddItem "", 0
List2.clear
Load addtask
addtask.Visible = True
End Sub

Private Sub Command1_Click()
List1.clear
qs.Text1.Text = ""
qs.Text2.Text = ""
qs.Text3.Text = ""
addtask.startport = Null
addtask.stopport = Null
Load qs
qs.Visible = True
End Sub

Private Sub Command2_Click()
changerr = tasks.ListIndex
If changerr = -1 Then
MsgBox "Select a valid task before trying to remove it!", vbOKOnly, "IDIOT!"
ElseIf changerr = 0 Then
MsgBox "Select a valid task before trying to remove it!", vbOKOnly, "IDIOT!"
Else
tasks.RemoveItem changerr
ops.RemoveItem changerr
tasktimes.RemoveItem changerr
End If
End Sub

Private Sub endbutt_Click()
If tasks.ListCount > 1 Then
decide = MsgBox("Exiting will delete your scheduled tasks.  Do you want to continue?", vbYesNo, "Exit?")
If decide = 6 Then End
Else
End
End If
End Sub

Private Sub Form_Load()
ops.AddItem "Code Master: Cory Forsythe", 0
tasks.AddItem "TCP/IP King V:1.0", 0
tasktimes.AddItem "Written@December,2000", 0
itemno = 0
GoS = False
End Sub

Private Sub List1_Click()
MsgBox List1.List(List1.ListIndex), vbOKOnly, " "
End Sub
Private Sub Winsock1_DataArrival(Index As Integer, ByVal bytesTotal As Long)
If login_data = True Then
Winsock1(Index).GetData Data, vbString
List1.AddItem Winsock1(Index).RemotePort & "    RETURNS:: " & Data
List2.AddItem ""
List1.AddItem ""
List2.AddItem ""
Open App.Path & "\Scanner log.txt" For Append As #2
Write #2, Winsock1(Index).RemotePort & " RETURNS::   " & Data
Close #2
End If
End Sub

Private Sub Winsock1_Connect(Index As Integer)
If Winsock1(Index).RemotePort = 8080 Then
List1.AddItem "PROXY SERVER FOUND"
Else
List1.AddItem "Port " & Winsock1(Index).RemotePort & " is open"
Open App.Path & "\Scanner log.txt" For Append As #4
Write #4, "Port " & Winsock1(Index).RemotePort & " is open"
Close #4
End If
End Sub

Private Sub Winsock1_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
destroy Winsock1(Index)
End Sub
Public Sub destroy(obj As Object)
TimeOut 1
Unload obj
TimeOut 1
End Sub
Private Sub ops_Click()
tasks.Selected(ops.ListIndex) = True
tasktimes.Selected(ops.ListIndex) = True
End Sub

Private Sub runbutt_Click()
changer = tasks.ListIndex
If changer = -1 Then
MsgBox "Select a valid task before trying to run it!", vbOKOnly, "IDIOT!"
ElseIf changer = 0 Then
MsgBox "Select a valid task before trying to run it!", vbOKOnly, "IDIOT!"
Else
tasktimes.RemoveItem changer
tasktimes.AddItem "Egaged", changer
RUNTASK changer, tasks.List(changer), ops.List(chagner)
tasks.RemoveItem changer
ops.RemoveItem changer
tasktimes.RemoveItem changer
End If
End Sub

Private Sub tasks_Click()
tasktimes.Selected(tasks.ListIndex) = True
ops.Selected(tasks.ListIndex) = True
End Sub

Private Sub tasktimes_Click()
tasks.Selected(tasktimes.ListIndex) = True
ops.Selected(tasktimes.ListIndex) = True
End Sub

Private Sub Timer1_Timer()
If IsNetConnectOnline() = True Then

If Not List1.List(0) = "External Scan Ready..." Then
'List1.RemoveItem 0
List1.AddItem "External Scan Ready...", 0
End If

ElseIf IsNetConnectOnline() = False Then

If Not List1.List(0) = "External Scan NOT Ready..." Then
List1.AddItem "External Scan NOT Ready...", 0
End If

End If
Label2.Caption = Time
i = 1
For i = 1 To tasktimes.ListCount

'' this will run the current procedure
If tasktimes.List(i) = Time Then
RUNTASK i, tasks.List(i), ops.List(i)
tasks.RemoveItem i
ops.RemoveItem i
tasktimes.RemoveItem i
ElseIf tasktimes.List(i) = "N/A" Then
tasktimes.RemoveItem i
tasktimes.AddItem "Engaged", i
RUNTASK i, tasks.List(i), ops.List(i)
tasks.RemoveItem i
ops.RemoveItem i
tasktimes.RemoveItem i
End If
Next i
End Sub

Private Sub Winsock2_Connect()
Go = False
List1.AddItem "Connected to:" & Winsock2.RemoteHost & "(" & Winsock2.RemoteHostIP & ")" & ":" & Winsock2.RemotePort
End Sub

Private Sub Winsock2_DataArrival(ByVal bytesTotal As Long)
Winsock2.GetData Data, vbString
List1.AddItem Data
End Sub

Private Sub Winsock2_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
If Number = Number Then
End If
End Sub
