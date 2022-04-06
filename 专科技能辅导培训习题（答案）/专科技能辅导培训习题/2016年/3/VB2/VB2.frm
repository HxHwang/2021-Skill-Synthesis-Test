VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "判断是否是两位数"
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
      Left            =   2640
      TabIndex        =   3
      Top             =   1560
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "输入"
      Height          =   615
      Left            =   960
      TabIndex        =   2
      Top             =   1560
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   735
      Left            =   1920
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   240
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "输入整数"
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
    Text1.Text = InputBox("输入整数")
End Sub

Private Sub Command2_Click()
    Dim x As Integer
    x = Text1.Text
    If x >= 10 And x <= 99 Then
        MsgBox "是"
    Else
        MsgBox "不是"
    End If
End Sub
