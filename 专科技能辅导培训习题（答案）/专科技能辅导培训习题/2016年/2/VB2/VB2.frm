VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "�жϸ�λ���Ƿ���7"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton Command2 
      Caption         =   "�ж�"
      Height          =   615
      Left            =   2520
      TabIndex        =   3
      Top             =   1560
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "����"
      Height          =   615
      Left            =   960
      TabIndex        =   2
      Top             =   1560
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   735
      Left            =   1800
      TabIndex        =   1
      Top             =   240
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   "��������"
      Height          =   495
      Left            =   480
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
    Label2.Caption = InputBox("��������")
End Sub

Private Sub Command2_Click()
    Dim x As Integer
    Dim a As Integer
    x = Label2.Caption
    a = x Mod 10
    If a = 7 Then
        MsgBox "��"
    Else
        MsgBox "����"
    End If
End Sub
