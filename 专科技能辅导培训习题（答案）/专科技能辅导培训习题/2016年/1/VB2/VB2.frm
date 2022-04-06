VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "判断能否被3整除"
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
      Left            =   2520
      TabIndex        =   3
      Top             =   1440
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "输入"
      Height          =   615
      Left            =   960
      TabIndex        =   2
      Top             =   1440
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   615
      Left            =   1920
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   240
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "输入整数"
      Height          =   495
      Left            =   600
      TabIndex        =   0
      Top             =   360
      Width           =   1095
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
    If x Mod 3 = 0 Then
        MsgBox "可以"
    Else
        MsgBox "不可以"
    End If
End Sub
