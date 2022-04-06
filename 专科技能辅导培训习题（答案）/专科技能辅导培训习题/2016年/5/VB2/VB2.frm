VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "成绩判断"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command2 
      Caption         =   "判断"
      Height          =   615
      Left            =   2760
      TabIndex        =   3
      Top             =   1800
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "输入"
      Height          =   615
      Left            =   1080
      TabIndex        =   2
      Top             =   1800
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   735
      Left            =   2040
      TabIndex        =   1
      Top             =   360
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "输入成绩"
      Height          =   495
      Left            =   840
      TabIndex        =   0
      Top             =   480
      Width           =   975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Label2.Caption = InputBox("输入一个成绩")
End Sub

Private Sub Command2_Click()
    Dim x As Integer
    x = Label2.Caption
    If x >= 90 Then
        MsgBox "优秀"
    Else
        MsgBox "一般"
    End If
End Sub
