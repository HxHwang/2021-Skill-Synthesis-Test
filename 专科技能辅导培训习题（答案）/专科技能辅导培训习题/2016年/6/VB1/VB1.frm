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
      Caption         =   "退出"
      Height          =   735
      Left            =   3000
      TabIndex        =   3
      Top             =   1560
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "显示时间"
      Height          =   735
      Left            =   840
      TabIndex        =   2
      Top             =   1560
      Width           =   1335
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label2"
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   1920
      TabIndex        =   1
      Top             =   240
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "当前时间"
      Height          =   495
      Left            =   600
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
    Label2.Caption = Time
End Sub

Private Sub Command2_Click()
    End
End Sub
