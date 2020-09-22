VERSION 5.00
Begin VB.Form running 
   BackColor       =   &H00BFBFBF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   " "
   ClientHeight    =   645
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4545
   ControlBox      =   0   'False
   Icon            =   "running.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   645
   ScaleWidth      =   4545
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Running Operations:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1080
      TabIndex        =   0
      Top             =   120
      Width           =   2490
   End
End
Attribute VB_Name = "running"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
mainform.Visible = False
End Sub

