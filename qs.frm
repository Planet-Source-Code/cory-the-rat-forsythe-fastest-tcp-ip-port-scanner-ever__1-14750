VERSION 5.00
Begin VB.Form qs 
   BackColor       =   &H00BFBFBF&
   Caption         =   "Quick Scan Feature..."
   ClientHeight    =   2355
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3360
   Icon            =   "qs.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2355
   ScaleWidth      =   3360
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      BackColor       =   &H00BFBFBF&
      Caption         =   "ENGAGE QUICK SCAN"
      Height          =   495
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1800
      Width           =   3135
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   2160
      TabIndex        =   6
      Top             =   1200
      Width           =   735
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   360
      MaxLength       =   6
      TabIndex        =   4
      Top             =   1200
      Width           =   735
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   840
      TabIndex        =   1
      Top             =   210
      Width           =   2415
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Stop Port"
      Height          =   195
      Left            =   2160
      TabIndex        =   5
      Top             =   960
      Width           =   660
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Start Port"
      Height          =   195
      Left            =   360
      TabIndex        =   3
      Top             =   960
      Width           =   660
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "*Domain or Ip"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   1440
      TabIndex        =   2
      Top             =   600
      Width           =   945
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Address"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   570
   End
End
Attribute VB_Name = "qs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
addtask.startport = Text2.Text
addtask.stopport = Text3.Text
mainform.tasks.AddItem Text1.Text
mainform.ops.AddItem 111
mainform.tasktimes.AddItem "N/A"
qs.Visible = False
End Sub
