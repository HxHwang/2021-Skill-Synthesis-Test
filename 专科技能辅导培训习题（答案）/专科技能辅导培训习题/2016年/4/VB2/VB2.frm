VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "分段函数计算"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command2 
      Caption         =   "计算"
      Height          =   735
      Left            =   2760
      TabIndex        =   3
      Top             =   1680
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "输入"
      Height          =   735
      Left            =   960
      TabIndex        =   2
      Top             =   1680
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   735
      Left            =   1800
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   240
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "输入变量x"
      Height          =   495
      Left            =   600
      TabIndex        =   0
      Top             =   360
      Width           =   855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Text1.Text = InputBox("输入一个数")
End Sub

Private Sub Command2_Click()
    Dim x As Integer
    Dim y As Integer
    x = Text1.Text
    If x >= 5 Then
        y = x + 3
    Else
        y = x - 3
    End If
    MsgBox y
End Sub
