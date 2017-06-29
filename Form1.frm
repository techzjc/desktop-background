VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "密码校验"
   ClientHeight    =   3105
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7320
   LinkTopic       =   "Form1"
   ScaleHeight     =   3105
   ScaleWidth      =   7320
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command1 
      Caption         =   "确定(&Y)"
      Height          =   975
      Left            =   2520
      TabIndex        =   2
      Top             =   1920
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   42
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   975
      Left            =   1680
      TabIndex        =   1
      Top             =   840
      Width           =   5055
   End
   Begin VB.Label password 
      Caption         =   "密码："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   24
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   0
      TabIndex        =   0
      Top             =   960
      Width           =   1575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Label1_Click()

End Sub

Private Sub Command1_Click()
Dim p As Integer
If Text1.Text = "nUmlock123987" Then
    Form2.Show
    Unload Me
Else
    p = MsgBox("密码错误！", 5 + 48, "输入密码")
    If p = 4 Then
        Text1.SetFocus
        Text1.Text = ""
    Else
        End
    End If
End If
End Sub

Private Sub Form_Load()
Text1.PasswordChar = "6"
Text1.Text = ""
End Sub

