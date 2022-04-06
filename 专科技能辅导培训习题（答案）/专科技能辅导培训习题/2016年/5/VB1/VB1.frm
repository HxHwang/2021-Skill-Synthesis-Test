VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command2 
      Caption         =   "重置"
      Height          =   615
      Left            =   2760
      TabIndex        =   3
      Top             =   1680
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "求解"
      Height          =   615
      Left            =   1080
      TabIndex        =   2
      Top             =   1680
      Width           =   1095
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2040
      TabIndex        =   1
      Top             =   240
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "平方根"
      Height          =   495
      Left            =   720
      TabIndex        =   0
      Top             =   360
      Width           =   975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Label2.Caption = Sqr(10)
End Sub

Private Sub Command2_Click()
    Label2.Caption = "平方根"
End Sub
