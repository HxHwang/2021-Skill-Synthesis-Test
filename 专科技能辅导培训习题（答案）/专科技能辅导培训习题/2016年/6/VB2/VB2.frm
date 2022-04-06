VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "判断是否是负数"
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
      Height          =   495
      Left            =   2520
      TabIndex        =   3
      Top             =   1800
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "输入"
      Height          =   495
      Left            =   960
      TabIndex        =   2
      Top             =   1800
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   735
      Left            =   1920
      TabIndex        =   1
      Top             =   360
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "输入整数"
      Height          =   615
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
    Label2.Caption = InputBox("输入整数")
End Sub

Private Sub Command2_Click()
    Dim x As Integer
    x = Label2.Caption
    If x >= 0 Then
        MsgBox "非负数"
    Else
        MsgBox "负数"
    End If
End Sub
