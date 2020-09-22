VERSION 5.00
Begin VB.Form addtask 
   BackColor       =   &H00BFBFBF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add Task Wizard"
   ClientHeight    =   4035
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6120
   Icon            =   "addtask.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4035
   ScaleWidth      =   6120
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00BFBFBF&
      Height          =   3975
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6135
      Begin VB.Frame Frame7 
         BackColor       =   &H00BFBFBF&
         Caption         =   "Step 4"
         Height          =   3615
         Left            =   1560
         TabIndex        =   41
         Top             =   600
         Visible         =   0   'False
         Width           =   3855
         Begin VB.TextBox Text5 
            Height          =   285
            Left            =   960
            TabIndex        =   49
            Top             =   1440
            Width           =   1215
         End
         Begin VB.TextBox Text4 
            Height          =   285
            Left            =   960
            TabIndex        =   47
            Top             =   840
            Width           =   1215
         End
         Begin VB.CommandButton Command19 
            BackColor       =   &H00BFBFBF&
            Caption         =   "Next"
            Height          =   375
            Left            =   2880
            Style           =   1  'Graphical
            TabIndex        =   44
            Top             =   3120
            Width           =   855
         End
         Begin VB.CommandButton Command18 
            BackColor       =   &H00BFBFBF&
            Caption         =   "Back"
            Height          =   375
            Left            =   1920
            Style           =   1  'Graphical
            TabIndex        =   43
            Top             =   3120
            Width           =   855
         End
         Begin VB.CommandButton Command17 
            BackColor       =   &H00BFBFBF&
            Caption         =   "Cancel"
            Height          =   375
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   42
            Top             =   3120
            Width           =   855
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "To:"
            Height          =   195
            Left            =   240
            TabIndex        =   48
            Top             =   1440
            Width           =   240
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "From:"
            Height          =   195
            Left            =   240
            TabIndex        =   46
            Top             =   840
            Width           =   390
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "What Ports would you like to scan?"
            Height          =   195
            Left            =   240
            TabIndex        =   45
            Top             =   240
            Width           =   2505
         End
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00BFBFBF&
         Caption         =   "Step 3"
         Height          =   3615
         Left            =   4440
         TabIndex        =   31
         Top             =   360
         Visible         =   0   'False
         Width           =   3855
         Begin VB.CommandButton Command16 
            BackColor       =   &H00BFBFBF&
            Caption         =   "RUN NOW"
            Height          =   375
            Left            =   1560
            Style           =   1  'Graphical
            TabIndex        =   40
            Top             =   1200
            Width           =   1215
         End
         Begin VB.TextBox Text3 
            Height          =   285
            Left            =   1560
            MaxLength       =   11
            TabIndex        =   36
            Text            =   "10:34:00 PM"
            Top             =   2280
            Width           =   1095
         End
         Begin VB.CommandButton Command15 
            BackColor       =   &H00BFBFBF&
            Caption         =   "DONE"
            Height          =   375
            Left            =   2880
            Style           =   1  'Graphical
            TabIndex        =   34
            Top             =   3120
            Width           =   855
         End
         Begin VB.CommandButton Command14 
            BackColor       =   &H00BFBFBF&
            Caption         =   "Back"
            Height          =   375
            Left            =   1920
            Style           =   1  'Graphical
            TabIndex        =   33
            Top             =   3120
            Width           =   855
         End
         Begin VB.CommandButton Command13 
            BackColor       =   &H00BFBFBF&
            Caption         =   "Cancel"
            Height          =   375
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   32
            Top             =   3120
            Width           =   855
         End
         Begin VB.Line Line1 
            X1              =   120
            X2              =   3720
            Y1              =   720
            Y2              =   720
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Run This Task Immediately"
            Height          =   195
            Left            =   1200
            TabIndex        =   39
            Top             =   960
            Width           =   1920
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "EG:  10:34:00 PM"
            Height          =   195
            Left            =   1440
            TabIndex        =   38
            Top             =   2640
            Width           =   1275
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Schedule this task for:"
            Height          =   195
            Left            =   1320
            TabIndex        =   37
            Top             =   2040
            Width           =   1575
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "When would you like to run this task?"
            Height          =   195
            Left            =   240
            TabIndex        =   35
            Top             =   240
            Width           =   2655
         End
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00BFBFBF&
         Caption         =   "Step 3"
         Height          =   3615
         Left            =   480
         TabIndex        =   23
         Top             =   1560
         Visible         =   0   'False
         Width           =   3855
         Begin VB.CheckBox Check3 
            BackColor       =   &H00BFBFBF&
            Caption         =   "View Incomming Server Messages"
            Height          =   255
            Left            =   600
            TabIndex        =   30
            Top             =   1920
            Width           =   3135
         End
         Begin VB.CheckBox Check2 
            BackColor       =   &H00BFBFBF&
            Caption         =   "Scan for Proxy servers"
            Height          =   255
            Left            =   600
            TabIndex        =   29
            Top             =   1440
            Width           =   2295
         End
         Begin VB.CheckBox Check1 
            BackColor       =   &H00BFBFBF&
            Caption         =   "Scan Ports for connectivety"
            Height          =   375
            Left            =   600
            TabIndex        =   28
            Top             =   840
            Width           =   3015
         End
         Begin VB.CommandButton Command12 
            BackColor       =   &H00BFBFBF&
            Caption         =   "Cancel"
            Height          =   375
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   26
            Top             =   3120
            Width           =   855
         End
         Begin VB.CommandButton Command11 
            BackColor       =   &H00BFBFBF&
            Caption         =   "Back"
            Height          =   375
            Left            =   1920
            Style           =   1  'Graphical
            TabIndex        =   25
            Top             =   3120
            Width           =   855
         End
         Begin VB.CommandButton Command10 
            BackColor       =   &H00BFBFBF&
            Caption         =   "Next"
            Height          =   375
            Left            =   2880
            Style           =   1  'Graphical
            TabIndex        =   24
            Top             =   3120
            Width           =   855
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "What would you like to do?"
            Height          =   195
            Left            =   240
            TabIndex        =   27
            Top             =   240
            Width           =   1935
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00BFBFBF&
         Caption         =   "Step 2"
         Height          =   3615
         Left            =   6120
         TabIndex        =   16
         Top             =   -3480
         Visible         =   0   'False
         Width           =   3855
         Begin VB.CommandButton Command9 
            BackColor       =   &H00BFBFBF&
            Caption         =   "Next"
            Height          =   375
            Left            =   2880
            Style           =   1  'Graphical
            TabIndex        =   20
            Top             =   3120
            Width           =   855
         End
         Begin VB.CommandButton Command8 
            BackColor       =   &H00BFBFBF&
            Caption         =   "Back"
            Height          =   375
            Left            =   1920
            Style           =   1  'Graphical
            TabIndex        =   19
            Top             =   3120
            Width           =   855
         End
         Begin VB.CommandButton Command7 
            BackColor       =   &H00BFBFBF&
            Caption         =   "Cancel"
            Height          =   375
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   18
            Top             =   3120
            Width           =   855
         End
         Begin VB.TextBox Text2 
            Height          =   375
            Left            =   1080
            MaxLength       =   15
            ScrollBars      =   1  'Horizontal
            TabIndex        =   17
            Top             =   1320
            Width           =   1455
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Enter The IP you wish to taget"
            Height          =   195
            Left            =   240
            TabIndex        =   22
            Top             =   240
            Width           =   2145
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "IP ADDRESS"
            Height          =   195
            Left            =   1320
            TabIndex        =   21
            Top             =   840
            Width           =   975
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00BFBFBF&
         Caption         =   "Step 1"
         Height          =   3615
         Left            =   -3600
         TabIndex        =   5
         Top             =   -360
         Visible         =   0   'False
         Width           =   3855
         Begin VB.ComboBox target_type 
            Height          =   315
            ItemData        =   "addtask.frx":0442
            Left            =   240
            List            =   "addtask.frx":044C
            Style           =   2  'Dropdown List
            TabIndex        =   9
            Top             =   1560
            Width           =   2895
         End
         Begin VB.CommandButton Command1 
            BackColor       =   &H00BFBFBF&
            Caption         =   "Next"
            Height          =   375
            Left            =   2880
            Style           =   1  'Graphical
            TabIndex        =   8
            Top             =   3120
            Width           =   855
         End
         Begin VB.CommandButton Command2 
            BackColor       =   &H00BFBFBF&
            Caption         =   "Back"
            Height          =   375
            Left            =   1920
            Style           =   1  'Graphical
            TabIndex        =   7
            Top             =   3120
            Width           =   855
         End
         Begin VB.CommandButton Command3 
            BackColor       =   &H00BFBFBF&
            Caption         =   "Cancel"
            Height          =   375
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   6
            Top             =   3120
            Width           =   855
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Please Select the type of target....."
            Height          =   195
            Left            =   120
            TabIndex        =   12
            Top             =   240
            Width           =   2445
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "yahoo.com  -->  Domain"
            Height          =   195
            Left            =   120
            TabIndex        =   11
            Top             =   720
            Width           =   1680
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "127.0.0.1    -->  IP ADDRESS"
            Height          =   195
            Left            =   120
            TabIndex        =   10
            Top             =   1080
            Width           =   2100
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00BFBFBF&
         Caption         =   "Step 2"
         Height          =   3615
         Left            =   -3360
         TabIndex        =   1
         Top             =   0
         Visible         =   0   'False
         Width           =   3855
         Begin VB.TextBox Text1 
            Height          =   375
            Left            =   240
            ScrollBars      =   1  'Horizontal
            TabIndex        =   15
            Top             =   1080
            Width           =   3375
         End
         Begin VB.CommandButton Command4 
            BackColor       =   &H00BFBFBF&
            Caption         =   "Cancel"
            Height          =   375
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   4
            Top             =   3120
            Width           =   855
         End
         Begin VB.CommandButton Command5 
            BackColor       =   &H00BFBFBF&
            Caption         =   "Back"
            Height          =   375
            Left            =   1920
            Style           =   1  'Graphical
            TabIndex        =   3
            Top             =   3120
            Width           =   855
         End
         Begin VB.CommandButton Command6 
            BackColor       =   &H00BFBFBF&
            Caption         =   "Next"
            Height          =   375
            Left            =   2880
            Style           =   1  'Graphical
            TabIndex        =   2
            Top             =   3120
            Width           =   855
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Domain Name"
            Height          =   195
            Left            =   1320
            TabIndex        =   14
            Top             =   840
            Width           =   1005
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Enter The Domain Name you wish to taget"
            Height          =   195
            Left            =   240
            TabIndex        =   13
            Top             =   240
            Width           =   3000
         End
      End
      Begin VB.Image Image1 
         Height          =   3750
         Left            =   120
         Picture         =   "addtask.frx":0464
         Top             =   120
         Width           =   1950
      End
   End
End
Attribute VB_Name = "addtask"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public runtime As String
Public targettype As String
Public target As String
Public proxyscan As Boolean
Public portscan As Boolean
Public logindata As Boolean
Public clear As Boolean
Public portsc2 As Integer
Public proxysc2 As Integer
Public loginda2 As Integer
Public startport
Public stopport


'framlocs are
'Top 240
'Left 2160
'
'
'
' user passes first form   ->  target type
Private Sub Command1_Click()
If target_type.ListIndex = -1 Then
MsgBox "You must select a target type to proceed", vbOKOnly, "Reminder"
ElseIf target_type.ListIndex = 0 Then
'goto domain selector
targettype = "domain"
Frame2.Visible = False
Frame3.Left = 2160
Frame3.Top = 240
Frame3.Visible = True
ElseIf target_type.ListIndex = 1 Then
''goto ip selector
targettype = "ip"
Frame2.Visible = False
Frame4.Left = 2160
Frame4.Top = 240
Frame4.Visible = True
End If
End Sub

Private Sub Command10_Click()
'get options
If Check1.Value = 1 Then portscan = True
If Check2.Value = 1 Then proxyscan = True
If Check3.Value = 1 Then logindata = True
If portscan = False And proxyscan = False Then MsgBox "You must scan for something, you cant just log data", vbOKOnly, "Uh...huh..."
Frame5.Visible = False
If portscan = True Then
Frame7.Left = 2160
Frame7.Top = 240
Frame7.Visible = True
ElseIf portscan = False Then
Frame6.Left = 2160
Frame6.Top = 240
Frame6.Visible = True
End If
End Sub

Private Sub Command11_Click()
If targettype = "domain" Then
Frame5.Visible = False
Frame3.Visible = True
Else
Frame5.Visible = False
Frame4.Visible = True
End If
End Sub

Private Sub Command12_Click()
Unload addtask
End Sub

Private Sub Command13_Click()
Unload addtask
End Sub

Private Sub Command14_Click()
Frame5.Top = 240
Frame5.Left = 2160
Frame5.Visible = True
Frame6.Visible = False
End Sub

Private Sub Command15_Click()
runtime = Text3.Text
finsihed target, mainform.itemno, targettype, portscan, proxyscan, logindata, runtime
addtask.Visible = False
End Sub

Private Sub Command16_Click()
runtime = "N/A"
finsihed target, mainform.itemno, targettype, portscan, proxyscan, logindata, runtime
addtask.Visible = False
End Sub

Private Sub Command17_Click()
Unload addtask
End Sub

Private Sub Command18_Click()
Frame7.Visible = False
Frame5.Left = 2160
Frame5.Top = 240
Frame5.Visible = True
End Sub

Private Sub Command19_Click()
If Text5.Text < Text4.Text Then MsgBox "Stop port must be larger than Start Port!", vbOKOnly, "ASS ALERT!"
If Text5.Text = Text4.Text Then MsgBox "Start & Stop ports cannot be the same!", vbOKOnly, "Reminder"
startport = Text4.Text
stopport = Text5.Text
Frame7.Visible = False
Frame6.Left = 2160
Frame6.Top = 240
Frame6.Visible = True
End Sub

Private Sub Command3_Click()
Unload addtask
End Sub

Private Sub Command4_Click()
Unload addtask
End Sub

Private Sub Command5_Click()
Frame3.Visible = False
Frame2.Visible = True
End Sub

Private Sub Command6_Click()
Dim comparison2
comparison2 = Text1 Like "*.*"
If Text1.Text = "" Then
MsgBox "Enter Something!", vbOKOnly, "IDIOT!"
ElseIf comparison2 = True Then
target = Text1.Text
Call ops
Else
MsgBox "Not a domain name!", vbOKOnly, "IDIOT!:"
End If

End Sub

Private Sub Command7_Click()
Unload addtask
End Sub

Private Sub Command8_Click()
Frame4.Visible = False
Frame2.Visible = True
End Sub

Private Sub Command9_Click()
Dim comparison1 As Boolean
comparison1 = Text2.Text Like "*.*.*"
If comparison1 = False Then
MsgBox "thats not an ip address", vbOKOnly, "Reminder"
Else:
target = Text2.Text
Call ops
End If
End Sub

Private Sub Form_Load()
clear = False
target = vbNullString
targettype = vbNullString
runtime = vbNullString
portscan = False
proxyscan = False
logindata = False
Text1.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text2.Text = ""
Check1.Value = 0
Check2.Value = 0
Check3.Value = 0
Frame3.Visible = False
Frame4.Visible = False
Frame5.Visible = False
Frame6.Visible = False
mainform.itemno = mainform.itemno + 1
Command2.Visible = False
Frame2.Top = 240
Frame2.Left = 2160
Frame2.Visible = True
End Sub

Sub ops()
If targettype = "domain" Then
Frame3.Visible = False
Frame5.Top = 240
Frame5.Left = 2160
Frame5.Visible = True
Text3.Text = Time
ElseIf targettype = "ip" Then
Frame4.Visible = False
Frame5.Top = 240
Frame5.Left = 2160
Frame5.Visible = True
Text3.Text = Time
End If
End Sub
Sub finsihed(targetaddy As String, itemnos As Integer, targetty As String, portsc, proxysc, loginda, times As String)
If portsc = True Then portsc2 = 1
If proxysc = True Then proxysc2 = 1
If loginda = True Then loginda2 = 1
itemnos2 = mainform.tasktimes.ListCount
mainform.tasks.AddItem targetaddy, itemnos2
mainform.ops.AddItem portsc2 & proxysc2 & loginda2, itemnos2
mainform.tasktimes.AddItem times, itemnos2
Unload addtask
End Sub

