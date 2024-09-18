VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5160
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6855
   LinkTopic       =   "Form1"
   ScaleHeight     =   5160
   ScaleWidth      =   6855
   StartUpPosition =   3  '窗口缺省
   Begin VB.CheckBox Check3 
      Caption         =   "Check3"
      Height          =   210
      Left            =   3120
      TabIndex        =   14
      Top             =   3705
      Width           =   210
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Check2"
      Height          =   210
      Left            =   4680
      TabIndex        =   10
      Top             =   4095
      Width           =   210
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   210
      Left            =   4680
      TabIndex        =   9
      Top             =   3705
      Value           =   1  'Checked
      Width           =   210
   End
   Begin VB.Timer Timer6 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   3300
      Top             =   4800
   End
   Begin VB.ListBox List2 
      Height          =   420
      Left            =   2700
      TabIndex        =   7
      Top             =   600
      Width           =   800
   End
   Begin VB.Timer Timer5 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   3900
      Top             =   4800
   End
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   4500
      Top             =   4800
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   5100
      Top             =   4800
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   5700
      Top             =   4800
   End
   Begin VB.CommandButton Command1 
      Caption         =   "开始"
      Height          =   615
      Left            =   4200
      MaskColor       =   &H8000000D&
      TabIndex        =   5
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      Caption         =   "变形"
      Enabled         =   0   'False
      Height          =   615
      Left            =   4500
      TabIndex        =   4
      Top             =   2100
      Width           =   615
   End
   Begin VB.CommandButton Command4 
      Caption         =   "下降"
      Enabled         =   0   'False
      Height          =   615
      Left            =   4500
      TabIndex        =   3
      Top             =   2700
      Width           =   615
   End
   Begin VB.CommandButton Command3 
      Caption         =   "右移"
      Enabled         =   0   'False
      Height          =   615
      Left            =   5100
      TabIndex        =   2
      Top             =   2100
      Width           =   615
   End
   Begin VB.CommandButton Command2 
      Caption         =   "左移"
      Enabled         =   0   'False
      Height          =   615
      Left            =   3900
      TabIndex        =   1
      Top             =   2100
      Width           =   615
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   6360
      Top             =   4800
   End
   Begin VB.ListBox List1 
      BackColor       =   &H00C0FFFF&
      ForeColor       =   &H00C0C000&
      Height          =   3660
      Left            =   600
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   360
      Width           =   1935
   End
   Begin VB.Label Label5 
      Caption         =   "是否启用格子标记（默认不启用）"
      Height          =   600
      Left            =   3510
      TabIndex        =   13
      Top             =   3705
      Width           =   990
   End
   Begin VB.Label Label4 
      Caption         =   "大方块▇"
      Height          =   405
      Left            =   5070
      TabIndex        =   12
      Top             =   4095
      Width           =   600
   End
   Begin VB.Label Label3 
      Caption         =   "小方块■"
      Height          =   405
      Left            =   5070
      TabIndex        =   11
      Top             =   3510
      Width           =   600
   End
   Begin VB.Label Label2 
      Caption         =   "下一个："
      Height          =   315
      Left            =   2700
      TabIndex        =   8
      Top             =   300
      Width           =   915
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000D&
      BeginProperty Font 
         Name            =   "楷体"
         Size            =   26.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   4200
      TabIndex        =   6
      Top             =   300
      Width           =   1515
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'bug日志
'L型变形的时候会出错'原因不明，改不来'好像也改好了？希望
'结束判定有问题’再想想’大概改好了?'好吧还是有问题，I型在刚出现就变形会直接导致结束
'在落到下一个方块前，极限左右移动会出现浮空bug’改不来，懒得改了’这真的是极端情况，改不来改不来
'长宽会变的形状，变形的时候在最左（右）边，会卡到另一边’大概改了，也许没改'这倒是改了
Dim f As Integer, n As Integer, k As Integer
Dim a1 As Integer, a2 As Integer, a3 As Integer, a4 As Integer '从上到下，从左到右，每个方块依次定义为a1,a2,a3,a4
Dim a(-20 To 230) As String, s As String, bj As String, m As String

Private Sub Check1_Click()
If Check1.Value = 1 Then m = "■": Check2.Value = 0
If Check1.Value = 0 Then m = "▇": Check2.Value = 1
End Sub

Private Sub Check2_Click()
If Check2.Value = 1 Then m = "▇": Check1.Value = 0
If Check2.Value = 0 Then m = "■": Check1.Value = 1
End Sub

Private Sub Check3_Click()
If Check3.Value = 0 Then
    bj = "  ": Label5.Caption = "已停用格子标记"
    ElseIf Check3.Value = 1 Then
        bj = "。": Label5.Caption = "已启用格子标记"
End If
End Sub

Private Sub Command1_Click() '开始
n = 0: k = 0
Label1.Caption = CStr(n)
Command1.Enabled = False
Command2.Enabled = True
Command3.Enabled = True
Command4.Enabled = True
Command5.Enabled = True
Check1.Enabled = False
Check2.Enabled = False
Check3.Enabled = False
If Check1.Value = 1 Then m = "■"
If Check2.Value = 1 Then m = "▇"
If Check3.Value = 0 Then bj = "  "
If Check3.Value = 1 Then bj = "。"
Call Form_Load
Timer1.Enabled = True
Timer2.Enabled = True
Timer3.Interval = 500
Timer4.Enabled = True
End Sub

Private Sub Command2_Click() '左移': a1 = a1 - 1: a2 = a2 - 1: a3 = a3 - 1: a4 = a4 - 1
    If a1 Mod 10 <> 1 And a2 Mod 10 <> 1 And a3 Mod 10 <> 1 And a4 Mod 10 <> 1 And (a1 > 20 Or a2 > 20 Or a3 > 20 Or a4 > 20) Then
        If f = 1 Then
            If a(a1 - 1) = bj Then a(a1 - 1) = m: a(a4) = bj: a1 = a1 - 1: a2 = a2 - 1: a3 = a3 - 1: a4 = a4 - 1
            ElseIf f = 11 Then
                If a(a1 - 1) = bj And a(a2 - 1) = bj And a(a3 - 1) = bj And a(a4 - 1) = bj Then a(a1 - 1) = m: a(a2 - 1) = m: a(a3 - 1) = m: a(a4 - 1) = m: a(a1) = bj: a(a2) = bj: a(a3) = bj: a(a4) = bj: a1 = a1 - 1: a2 = a2 - 1: a3 = a3 - 1: a4 = a4 - 1
        ElseIf f = 2 Then
            If a(a1 - 1) = bj And a(a3 - 1) = bj Then a(a1 - 1) = m: a(a3 - 1) = m: a(a2) = bj: a(a4) = bj: a1 = a1 - 1: a2 = a2 - 1: a3 = a3 - 1: a4 = a4 - 1
            ElseIf f = 21 Then
                If a(a1 - 1) = bj And a(a2 - 1) = bj And a(a4 - 1) = bj Then a(a1 - 1) = m: a(a2 - 1) = m: a(a4 - 1) = m: a(a1) = bj: a(a3) = bj: a(a4) = bj: a1 = a1 - 1: a2 = a2 - 1: a3 = a3 - 1: a4 = a4 - 1
        ElseIf f = 3 Then
            If a(a1 - 1) = bj And a(a3 - 1) = bj Then a(a1 - 1) = m: a(a3 - 1) = m: a(a2) = bj: a(a4) = bj: a1 = a1 - 1: a2 = a2 - 1: a3 = a3 - 1: a4 = a4 - 1
            ElseIf f = 31 Then
                If a(a1 - 1) = bj And a(a2 - 1) = bj And a(a4 - 1) = bj Then a(a1 - 1) = m: a(a2 - 1) = m: a(a4 - 1) = m: a(a1) = bj: a(a3) = bj: a(a4) = bj: a1 = a1 - 1: a2 = a2 - 1: a3 = a3 - 1: a4 = a4 - 1
        ElseIf f = 4 Then
            If a(a1 - 1) = bj And a(a3 - 1) = bj Then a(a1 - 1) = m: a(a3 - 1) = m: a(a2) = bj: a(a4) = bj: a1 = a1 - 1: a2 = a2 - 1: a3 = a3 - 1: a4 = a4 - 1
        ElseIf f = 5 Then
            If a(a1 - 1) = bj And a(a2 - 1) = bj Then a(a1 - 1) = m: a(a2 - 1) = m: a(a1) = bj: a(a4) = bj: a1 = a1 - 1: a2 = a2 - 1: a3 = a3 - 1: a4 = a4 - 1
            ElseIf f = 51 Then
                If a(a1 - 1) = bj And a(a2 - 1) = bj And a(a4 - 1) = bj Then a(a1 - 1) = m: a(a2 - 1) = m: a(a4 - 1) = m: a(a1) = bj: a(a3) = bj: a(a4) = bj: a1 = a1 - 1: a2 = a2 - 1: a3 = a3 - 1: a4 = a4 - 1
            ElseIf f = 52 Then
                If a(a1 - 1) = bj And a(a4 - 1) = bj Then a(a1 - 1) = m: a(a4 - 1) = m: a(a3) = bj: a(a4) = bj: a1 = a1 - 1: a2 = a2 - 1: a3 = a3 - 1: a4 = a4 - 1
            ElseIf f = 53 Then
                If a(a1 - 1) = bj And a(a2 - 1) = bj And a(a4 - 1) = bj Then a(a1 - 1) = m: a(a2 - 1) = m: a(a4 - 1) = m: a(a1) = bj: a(a3) = bj: a(a4) = bj: a1 = a1 - 1: a2 = a2 - 1: a3 = a3 - 1: a4 = a4 - 1
        ElseIf f = 6 Then
            If a(a1 - 1) = bj And a(a2 - 1) = bj Then a(a1 - 1) = m: a(a2 - 1) = m: a(a1) = bj: a(a4) = bj: a1 = a1 - 1: a2 = a2 - 1: a3 = a3 - 1: a4 = a4 - 1
            ElseIf f = 61 Then
                If a(a1 - 1) = bj And a(a3 - 1) = bj And a(a4 - 1) = bj Then a(a1 - 1) = m: a(a3 - 1) = m: a(a4 - 1) = m: a(a2) = bj: a(a3) = bj: a(a4) = bj: a1 = a1 - 1: a2 = a2 - 1: a3 = a3 - 1: a4 = a4 - 1
            ElseIf f = 62 Then
                If a(a1 - 1) = bj And a(a4 - 1) = bj Then a(a1 - 1) = m: a(a4 - 1) = m: a(a3) = bj: a(a4) = bj: a1 = a1 - 1: a2 = a2 - 1: a3 = a3 - 1: a4 = a4 - 1
            ElseIf f = 63 Then
                If a(a1 - 1) = bj And a(a2 - 1) = bj And a(a3 - 1) = bj Then a(a1 - 1) = m: a(a2 - 1) = m: a(a3 - 1) = m: a(a1) = bj: a(a2) = bj: a(a4) = bj: a1 = a1 - 1: a2 = a2 - 1: a3 = a3 - 1: a4 = a4 - 1
        ElseIf f = 7 Then
            If a(a1 - 1) = bj And a(a2 - 1) = bj Then a(a1 - 1) = m: a(a2 - 1) = m: a(a1) = bj: a(a4) = bj: a1 = a1 - 1: a2 = a2 - 1: a3 = a3 - 1: a4 = a4 - 1
            ElseIf f = 71 Then
                If a(a1 - 1) = bj And a(a2 - 1) = bj And a(a3 - 1) = bj Then a(a1 - 1) = m: a(a2 - 1) = m: a(a3 - 1) = m: a(a1) = bj: a(a2) = bj: a(a4) = bj: a1 = a1 - 1: a2 = a2 - 1: a3 = a3 - 1: a4 = a4 - 1
            ElseIf f = 72 Then
                If a(a1 - 1) = bj And a(a4 - 1) = bj Then a(a1 - 1) = m: a(a4 - 1) = m: a(a3) = bj: a(a4) = bj: a1 = a1 - 1: a2 = a2 - 1: a3 = a3 - 1: a4 = a4 - 1
            ElseIf f = 73 Then
                If a(a1 - 1) = bj And a(a3 - 1) = bj And a(a4 - 1) = bj Then a(a1 - 1) = m: a(a3 - 1) = m: a(a4 - 1) = m: a(a2) = bj: a(a3) = bj: a(a4) = bj: a1 = a1 - 1: a2 = a2 - 1: a3 = a3 - 1: a4 = a4 - 1
        End If
    End If
End Sub

Private Sub Command3_Click() '右移': a1 = a1 + 1: a2 = a2 + 1: a3 = a3 + 1: a4 = a4 + 1
    If a1 Mod 10 <> 0 And a2 Mod 10 <> 0 And a3 Mod 10 <> 0 And a4 Mod 10 <> 0 And (a1 > 20 Or a2 > 20 Or a3 > 20 Or a4 > 20) Then
        If f = 1 Then
            If a(a4 + 1) = bj Then a(a4 + 1) = m: a(a1) = bj: a1 = a1 + 1: a2 = a2 + 1: a3 = a3 + 1: a4 = a4 + 1
            ElseIf f = 11 Then
                If a(a1 + 1) = bj And a(a2 + 1) = bj And a(a3 + 1) = bj And a(a4 + 1) = bj Then a(a1 + 1) = m: a(a2 + 1) = m: a(a3 + 1) = m: a(a4 + 1) = m: a(a1) = bj: a(a2) = bj: a(a3) = bj: a(a4) = bj: a1 = a1 + 1: a2 = a2 + 1: a3 = a3 + 1: a4 = a4 + 1
        ElseIf f = 2 Then
            If a(a2 + 1) = bj And a(a4 + 1) = bj Then a(a2 + 1) = m: a(a4 + 1) = m: a(a1) = bj: a(a3) = bj: a1 = a1 + 1: a2 = a2 + 1: a3 = a3 + 1: a4 = a4 + 1
            ElseIf f = 21 Then
                If a(a1 + 1) = bj And a(a3 + 1) = bj And a(a4 + 1) = bj Then a(a1 + 1) = m: a(a3 + 1) = m: a(a4 + 1) = m: a(a1) = bj: a(a2) = bj: a(a4) = bj: a1 = a1 + 1: a2 = a2 + 1: a3 = a3 + 1: a4 = a4 + 1
        ElseIf f = 3 Then
            If a(a2 + 1) = bj And a(a4 + 1) = bj Then a(a2 + 1) = m: a(a4 + 1) = m: a(a1) = bj: a(a3) = bj: a1 = a1 + 1: a2 = a2 + 1: a3 = a3 + 1: a4 = a4 + 1
            ElseIf f = 31 Then
                If a(a1 + 1) = bj And a(a3 + 1) = bj And a(a4 + 1) = bj Then a(a1 + 1) = m: a(a3 + 1) = m: a(a4 + 1) = m: a(a1) = bj: a(a2) = bj: a(a4) = bj: a1 = a1 + 1: a2 = a2 + 1: a3 = a3 + 1: a4 = a4 + 1
        ElseIf f = 4 Then
            If a(a2 + 1) = bj And a(a4 + 1) = bj Then a(a2 + 1) = m: a(a4 + 1) = m: a(a1) = bj: a(a3) = bj: a1 = a1 + 1: a2 = a2 + 1: a3 = a3 + 1: a4 = a4 + 1
        ElseIf f = 5 Then
            If a(a1 + 1) = bj And a(a4 + 1) = bj Then a(a1 + 1) = m: a(a4 + 1) = m: a(a1) = bj: a(a2) = bj: a1 = a1 + 1: a2 = a2 + 1: a3 = a3 + 1: a4 = a4 + 1
            ElseIf f = 51 Then
                If a(a1 + 1) = bj And a(a3 + 1) = bj And a(a4 + 1) = bj Then a(a1 + 1) = m: a(a3 + 1) = m: a(a4 + 1) = m: a(a1) = bj: a(a2) = bj: a(a4) = bj: a1 = a1 + 1: a2 = a2 + 1: a3 = a3 + 1: a4 = a4 + 1
            ElseIf f = 52 Then
                If a(a3 + 1) = bj And a(a4 + 1) = bj Then a(a3 + 1) = m: a(a4 + 1) = m: a(a1) = bj: a(a4) = bj: a1 = a1 + 1: a2 = a2 + 1: a3 = a3 + 1: a4 = a4 + 1
            ElseIf f = 53 Then
                If a(a1 + 1) = bj And a(a3 + 1) = bj And a(a4 + 1) = bj Then a(a1 + 1) = m: a(a3 + 1) = m: a(a4 + 1) = m: a(a1) = bj: a(a2) = bj: a(a4) = bj: a1 = a1 + 1: a2 = a2 + 1: a3 = a3 + 1: a4 = a4 + 1
        ElseIf f = 6 Then
            If a(a1 + 1) = bj And a(a4 + 1) = bj Then a(a1 + 1) = m: a(a4 + 1) = m: a(a1) = bj: a(a2) = bj: a1 = a1 + 1: a2 = a2 + 1: a3 = a3 + 1: a4 = a4 + 1
            ElseIf f = 61 Then
                If a(a2 + 1) = bj And a(a3 + 1) = bj And a(a4 + 1) = bj Then a(a2 + 1) = m: a(a3 + 1) = m: a(a4 + 1) = m: a(a1) = bj: a(a3) = bj: a(a4) = bj: a1 = a1 + 1: a2 = a2 + 1: a3 = a3 + 1: a4 = a4 + 1
            ElseIf f = 62 Then
                If a(a3 + 1) = bj And a(a4 + 1) = bj Then a(a3 + 1) = m: a(a4 + 1) = m: a(a1) = bj: a(a4) = bj: a1 = a1 + 1: a2 = a2 + 1: a3 = a3 + 1: a4 = a4 + 1
            ElseIf f = 63 Then
                If a(a1 + 1) = bj And a(a2 + 1) = bj And a(a4 + 1) = bj Then a(a1 + 1) = m: a(a2 + 1) = m: a(a4 + 1) = m: a(a1) = bj: a(a2) = bj: a(a3) = bj: a1 = a1 + 1: a2 = a2 + 1: a3 = a3 + 1: a4 = a4 + 1
        ElseIf f = 7 Then
            If a(a1 + 1) = bj And a(a4 + 1) = bj Then a(a1 + 1) = m: a(a4 + 1) = m: a(a1) = bj: a(a2) = bj: a1 = a1 + 1: a2 = a2 + 1: a3 = a3 + 1: a4 = a4 + 1
            ElseIf f = 71 Then
                If a(a1 + 1) = bj And a(a2 + 1) = bj And a(a4 + 1) = bj Then a(a1 + 1) = m: a(a2 + 1) = m: a(a4 + 1) = m: a(a1) = bj: a(a2) = bj: a(a3) = bj: a1 = a1 + 1: a2 = a2 + 1: a3 = a3 + 1: a4 = a4 + 1
            ElseIf f = 72 Then
                If a(a3 + 1) = bj And a(a4 + 1) = bj Then a(a3 + 1) = m: a(a4 + 1) = m: a(a1) = bj: a(a4) = bj: a1 = a1 + 1: a2 = a2 + 1: a3 = a3 + 1: a4 = a4 + 1
            ElseIf f = 73 Then
                If a(a2 + 1) = bj And a(a3 + 1) = bj And a(a4 + 1) = bj Then a(a2 + 1) = m: a(a3 + 1) = m: a(a4 + 1) = m: a(a1) = bj: a(a3) = bj: a(a4) = bj: a1 = a1 + 1: a2 = a2 + 1: a3 = a3 + 1: a4 = a4 + 1
        End If
    End If
End Sub

Private Sub Command4_Click() '下降'只能说这方法非常的逃课，还顺便解决了直接到底的方法
Dim i As Integer
For i = 1 To 5
    Call Timer3_Timer
Next i
End Sub

Private Sub Command5_Click() '变形’盲猜弄不出来'枚举所有可能？？'枚举！'应该是可以了的’不过L型的有点问题，要加点限制条件，感觉都要加边界限制，不然能从左边变形到右边'长的都要限制
    If f = 1 Then Call bx11: Exit Sub
        If f = 11 Then Call bx1: Exit Sub
    If f = 2 Then Call bx21: Exit Sub
        If f = 21 Then Call bx2: Exit Sub
    If f = 3 Then Call bx31: Exit Sub
        If f = 31 Then Call bx3: Exit Sub
    If f = 5 Then Call bx51: Exit Sub
        If f = 51 Then Call bx52: Exit Sub
        If f = 52 Then Call bx53: Exit Sub
        If f = 53 Then Call bx5: Exit Sub
    If f = 6 Then Call bx61: Exit Sub
        If f = 61 Then Call bx62: Exit Sub
        If f = 62 Then Call bx63: Exit Sub
        If f = 63 Then Call bx6: Exit Sub
    If f = 7 Then Call bx71: Exit Sub
        If f = 71 Then Call bx72: Exit Sub
        If f = 72 Then Call bx73: Exit Sub
        If f = 73 Then Call bx7: Exit Sub
End Sub
Private Sub bx11() '以第三块为中心
'■
'■
'■
'■
If a(a3 - 20) = bj And a(a3 - 10) = bj And a(a3 + 10) = bj Then f = 11: a(a3 - 20) = m: a(a3 - 10) = m: a(a3 + 10) = m: a(a1) = bj: a(a2) = bj: a(a4) = bj: a1 = a1 + 2 - 20: a2 = a2 + 1 - 10: a4 = a4 - 1 + 10
End Sub
Private Sub bx1() '以第三块为中心
'■■■■
If a3 Mod 10 >= 3 And a3 Mod 10 <> 0 And a(a3 - 2) = bj And a(a3 - 1) = bj And a(a3 + 1) = bj Then f = 1: a(a3 - 2) = m: a(a3 - 1) = m: a(a3 + 1) = m: a(a1) = bj: a(a2) = bj: a(a4) = bj: a1 = a1 - 2 + 20: a2 = a2 - 1 + 10: a4 = a4 + 1 - 10
End Sub
Private Sub bx21() '以最上方两块为基准
'■
'■■
'  ■
If a(a1 - 10) = bj And a(a2 + 10) = bj Then f = 21: a(a1 - 10) = m: a(a2 + 10) = m: a(a3) = bj: a(a4) = bj: a1 = a1 - 10: a2 = a2 - 1: a3 = a3 + 2 - 10: a4 = a4 + 1
End Sub
Private Sub bx2()  '以中心横着的两块为基准
'  ■■
'■■
If a2 Mod 10 >= 2 And a(a2 - 1 + 10) = bj And a(a2 + 10) = bj Then f = 2: a(a2 - 1 + 10) = m: a(a2 + 10) = m: a(a1) = bj: a(a4) = bj: a1 = a1 + 10: a2 = a2 + 1: a3 = a3 - 2 + 10: a4 = a4 - 1
End Sub
Private Sub bx31() '以最上方两块为基准
'  ■
'■■
'■
If a(a1 + 10) = bj And a(a2 - 10) = bj Then f = 31: a(a1 + 10) = m: a(a2 - 10) = m: a(a3) = bj: a(a4) = bj: a1 = a1 + 1 - 10: a2 = a2 - 1: a3 = a3 - 10: a4 = a4 - 2
End Sub
Private Sub bx3() '以中心横着的两块为基准
'■■
'  ■■
If a3 Mod 10 <> 0 And a(a3 + 10) = bj And a(a3 + 1 + 10) = bj Then f = 3: a(a3 + 10) = m: a(a3 + 1 + 10) = m: a(a1) = bj: a(a4) = bj: a1 = a1 - 1 + 10: a2 = a2 + 1: a3 = a3 + 10: a4 = a4 + 2
End Sub
Private Sub bx51() '5均以最中心点为基准
'■
'■■
'■
If a(a3 + 10) = bj Then f = 51: a(a3 + 10) = m: a(a2) = bj: a2 = a2 + 1: a3 = a3 + 1: a4 = a4 - 1 + 10
End Sub
Private Sub bx52() '5均以最中心点为基准
'■■■
'  ■
If a2 Mod 10 >= 2 And a(a2 - 1) = bj Then f = 52: a(a2 - 1) = m: a(a1) = bj: a1 = a1 - 1 + 10
End Sub
Private Sub bx53() '5均以最中心点为基准
'  ■
'■■
'  ■
If a(a2 - 10) = bj Then f = 53: a(a2 - 10) = m: a(a3) = bj: a1 = a1 + 1 - 10: a2 = a2 - 1: a3 = a3 - 1
End Sub
Private Sub bx5() '5均以最中心点为基准
'  ■
'■■■
If a3 Mod 10 <> 0 And a(a3 + 1) = bj Then f = 5: a(a3 + 1) = m: a(a4) = bj: a4 = a4 + 1 - 10
End Sub
Private Sub bx61() '6,7均以折点与长边中间点为基准
'■■
'■
'■
If a(a2 + 10) = bj And a(a2 + 20) = bj Then f = 61: a(a2 + 10) = m: a(a2 + 20) = m: a(a1) = bj: a(a4) = bj: a1 = a1 + 10: a2 = a2 + 1: a3 = a3 - 1 + 10: a4 = a4 - 2 + 20
End Sub
Private Sub bx62() '6,7均以折点与长边中间点为基准
'■■■
'    ■
If a1 Mod 10 >= 3 And a(a1 - 1) = bj And a(a1 - 2) = bj Then f = 62: a(a1 - 1) = m: a(a1 - 2) = m: a(a2) = bj: a(a4) = bj: a1 = a1 - 2: a2 = a2 - 2: a3 = a3 - 10: a4 = a4 - 10
End Sub
Private Sub bx63() '6,7均以折点与长边中间点为基准
'  ■
'  ■
'■■
If a(a3 - 10) = bj And a(a3 - 20) = bj Then f = 63: a(a3 - 10) = m: a(a3 - 20) = m: a(a1) = bj: a(a4) = bj: a1 = a1 + 2 - 20: a2 = a2 + 1 - 10: a3 = a3 - 1: a4 = a4 - 10
End Sub
Private Sub bx6() '6,7均以折点与长边中间点为基准'????'6,7似乎都有问题？？？'6,7似乎都没问题
'■
'■■■
If a3 Mod 10 <= 7 And a(a4 + 1) = bj And a(a4 + 2) = bj Then f = 6: a(a4 + 1) = m: a(a4 + 2) = m: a(a1) = bj: a(a3) = bj: a1 = a1 + 10: a2 = a2 + 10: a3 = a3 + 2: a4 = a4 + 2
End Sub
Private Sub bx71() '6,7均以折点与长边中间点为基准'卡右边时会出错
'■
'■
'■■
If a(a3 - 10) = bj And a(a3 - 20) = bj Then f = 71: a(a3 - 10) = m: a(a3 - 20) = m: a(a1) = bj: a(a2) = bj: a1 = a1 - 1 - 10: a2 = a2 + 1 - 10
End Sub
Private Sub bx72() '6,7均以折点与长边中间点为基准
'■■■
'■
If a2 Mod 10 <= 8 And a(a2 + 1) = bj And a(a2 + 2) = bj Then f = 72: a(a2 + 1) = m: a(a2 + 2) = m: a(a1) = bj: a(a4) = bj: a1 = a1 + 10: a2 = a2 + 1: a3 = a3 + 2 - 10: a4 = a4 - 1
End Sub
Private Sub bx73() '6,7均以折点与长边中间点为基准
'■■
'  ■
'  ■
If a(a2 + 10) = bj And a(a2 + 20) = bj Then f = 73: a(a2 + 10) = m: a(a2 + 20) = m: a(a3) = bj: a(a4) = bj: a3 = a3 - 1 + 10: a4 = a4 + 1 + 10
End Sub
Private Sub bx7() '6,7均以折点与长边中间点为基准
'    ■
'■■■
If a3 Mod 10 >= 3 And a(a3 - 1) = bj And a(a3 - 2) = bj Then f = 7: a(a3 - 1) = m: a(a3 - 2) = m: a(a1) = bj: a(a4) = bj: a1 = a1 + 1: a2 = a2 - 2 + 10: a3 = a3 - 1: a4 = a4 - 10
End Sub
Private Sub Form_Load()
Dim i As Integer
    For i = -20 To 230
        a(i) = bj
    Next i
End Sub



Private Sub Timer1_Timer() '刷新
Dim i As Integer
    List1.Clear
    For i = 21 To 220
        s = s + a(i)
        If i Mod 10 = 0 Then List1.AddItem s: s = ""
    Next i
End Sub

Private Sub Timer2_Timer() '生成'只能说伪随机数真的伪随机，老是出现很多相同的'好家伙隔了一天我就不记得xz是什么意思了，虽然不影响
Randomize
If k = 0 Then k = Int(Rnd * 7 + 1)
If k = 1 Then Call xz1
If k = 2 Then Call xz2
If k = 3 Then Call xz3
If k = 4 Then Call xz4
If k = 5 Then Call xz5
If k = 6 Then Call xz6
If k = 7 Then Call xz7
k = Int(Rnd * 7 + 1)
Call xyg
Timer3.Enabled = True
Timer2.Enabled = False
End Sub
Private Sub xyg() '下一个
Dim c As String
List2.Clear
If k = 1 Then c = "■■■■": List2.AddItem c
If k = 2 Then c = "  ■■": List2.AddItem c: c = "■■  ": List2.AddItem c
If k = 3 Then c = "■■  ": List2.AddItem c: c = "  ■■": List2.AddItem c
If k = 4 Then c = "■■  ": List2.AddItem c: c = "■■  ": List2.AddItem c
If k = 5 Then c = "  ■  ": List2.AddItem c: c = "■■■": List2.AddItem c
If k = 6 Then c = "■    ": List2.AddItem c: c = "■■■": List2.AddItem c
If k = 7 Then c = "    ■": List2.AddItem c: c = "■■■": List2.AddItem c
End Sub
Private Sub xz1()
'■■■■
a1 = 4: a2 = 5: a3 = 6: a4 = 7: f = 1
a(a1) = m: a(a2) = m: a(a3) = m: a(a4) = m
End Sub
Private Sub xz2()
'  ■■
'■■
a1 = 6: a2 = 7: a3 = 15: a4 = 16: f = 2
a(a1) = m: a(a2) = m: a(a3) = m: a(a4) = m
End Sub
Private Sub xz3()
'■■
'  ■■
a1 = 5: a2 = 6: a3 = 16: a4 = 17: f = 3
a(a1) = m: a(a2) = m: a(a3) = m: a(a4) = m
End Sub
Private Sub xz4()
'■■
'■■
a1 = 6: a2 = 7: a3 = 16: a4 = 17: f = 4
a(a1) = m: a(a2) = m: a(a3) = m: a(a4) = m
End Sub
Private Sub xz5()
'  ■
'■■■
a1 = 6: a2 = 15: a3 = 16: a4 = 17: f = 5
a(a1) = m: a(a2) = m: a(a3) = m: a(a4) = m
End Sub
Private Sub xz6()
'■
'■■■
a1 = 5: a2 = 15: a3 = 16: a4 = 17: f = 6
a(a1) = m: a(a2) = m: a(a3) = m: a(a4) = m
End Sub
Private Sub xz7()
'    ■
'■■■
a1 = 7: a2 = 15: a3 = 16: a4 = 17: f = 7
a(a1) = m: a(a2) = m: a(a3) = m: a(a4) = m
End Sub

Private Sub Timer3_Timer() '下落':a1 = a1 + 10: a2 = a2 + 10: a3 = a3 + 10: a4 = a4 + 10: flag = True Else flag = False
Dim flag As Boolean
 flag = True
    If a1 > 210 Or a2 > 210 Or a3 > 210 Or a4 > 210 Then
        flag = False
    ElseIf f = 1 Then
        If a(a1 + 10) = bj And a(a2 + 10) = bj And a(a3 + 10) = bj And a(a4 + 10) = bj Then a(a1 + 10) = m: a(a2 + 10) = m: a(a3 + 10) = m: a(a4 + 10) = m: a(a1) = bj: a(a2) = bj: a(a3) = bj: a(a4) = bj: a1 = a1 + 10: a2 = a2 + 10: a3 = a3 + 10: a4 = a4 + 10: flag = True Else flag = False
        ElseIf f = 11 Then
            If a(a4 + 10) = bj Then a(a4 + 10) = m: a(a1) = bj: a1 = a1 + 10: a2 = a2 + 10: a3 = a3 + 10: a4 = a4 + 10: flag = True Else flag = False
    ElseIf f = 2 Then
        If a(a2 + 10) = bj And a(a3 + 10) = bj And a(a4 + 10) = bj Then a(a2 + 10) = m: a(a3 + 10) = m: a(a4 + 10) = m: a(a1) = bj: a(a2) = bj: a(a3) = bj: a1 = a1 + 10: a2 = a2 + 10: a3 = a3 + 10: a4 = a4 + 10: flag = True Else flag = False
        ElseIf f = 21 Then
            If a(a2 + 10) = bj And a(a4 + 10) = bj Then a(a2 + 10) = m: a(a4 + 10) = m: a(a1) = bj: a(a3) = bj: a1 = a1 + 10: a2 = a2 + 10: a3 = a3 + 10: a4 = a4 + 10: flag = True Else flag = False
    ElseIf f = 3 Then
        If a(a1 + 10) = bj And a(a3 + 10) = bj And a(a4 + 10) = bj Then a(a1 + 10) = m: a(a3 + 10) = m: a(a4 + 10) = m: a(a1) = bj: a(a2) = bj: a(a4) = bj: a1 = a1 + 10: a2 = a2 + 10: a3 = a3 + 10: a4 = a4 + 10: flag = True Else flag = False
        ElseIf f = 31 Then
            If a(a3 + 10) = bj And a(a4 + 10) = bj Then a(a3 + 10) = m: a(a4 + 10) = m: a(a1) = bj: a(a2) = bj: a1 = a1 + 10: a2 = a2 + 10: a3 = a3 + 10: a4 = a4 + 10: flag = True Else flag = False
    ElseIf f = 4 Then
        If a(a3 + 10) = bj And a(a4 + 10) = bj Then a(a3 + 10) = m: a(a4 + 10) = m: a(a1) = bj: a(a2) = bj: a1 = a1 + 10: a2 = a2 + 10: a3 = a3 + 10: a4 = a4 + 10: flag = True Else flag = False
    ElseIf f = 5 Then
        If a(a2 + 10) = bj And a(a3 + 10) = bj And a(a4 + 10) = bj Then a(a2 + 10) = m: a(a3 + 10) = m: a(a4 + 10) = m: a(a1) = bj: a(a2) = bj: a(a4) = bj: a1 = a1 + 10: a2 = a2 + 10: a3 = a3 + 10: a4 = a4 + 10: flag = True Else flag = False
        ElseIf f = 51 Then
            If a(a3 + 10) = bj And a(a4 + 10) = bj Then a(a3 + 10) = m: a(a4 + 10) = m: a(a1) = bj: a(a3) = bj: a1 = a1 + 10: a2 = a2 + 10: a3 = a3 + 10: a4 = a4 + 10: flag = True Else flag = False
        ElseIf f = 52 Then
            If a(a1 + 10) = bj And a(a3 + 10) = bj And a(a4 + 10) = bj Then a(a1 + 10) = m: a(a3 + 10) = m: a(a4 + 10) = m: a(a1) = bj: a(a2) = bj: a(a3) = bj: a1 = a1 + 10: a2 = a2 + 10: a3 = a3 + 10: a4 = a4 + 10: flag = True Else flag = False
        ElseIf f = 53 Then
            If a(a2 + 10) = bj And a(a4 + 10) = bj Then a(a2 + 10) = m: a(a4 + 10) = m: a(a1) = bj: a(a2) = bj: a1 = a1 + 10: a2 = a2 + 10: a3 = a3 + 10: a4 = a4 + 10: flag = True Else flag = False
    ElseIf f = 6 Then
        If a(a2 + 10) = bj And a(a3 + 10) = bj And a(a4 + 10) = bj Then a(a2 + 10) = m: a(a3 + 10) = m: a(a4 + 10) = m: a(a1) = bj: a(a3) = bj: a(a4) = bj: a1 = a1 + 10: a2 = a2 + 10: a3 = a3 + 10: a4 = a4 + 10: flag = True Else flag = False
        ElseIf f = 61 Then
            If a(a2 + 10) = bj And a(a4 + 10) = bj Then a(a2 + 10) = m: a(a4 + 10) = m: a(a1) = bj: a(a2) = bj: a1 = a1 + 10: a2 = a2 + 10: a3 = a3 + 10: a4 = a4 + 10: flag = True Else flag = False
        ElseIf f = 62 Then
            If a(a1 + 10) = bj And a(a2 + 10) = bj And a(a4 + 10) = bj Then a(a1 + 10) = m: a(a2 + 10) = m: a(a4 + 10) = m: a(a1) = bj: a(a2) = bj: a(a3) = bj: a1 = a1 + 10: a2 = a2 + 10: a3 = a3 + 10: a4 = a4 + 10: flag = True Else flag = False
        ElseIf f = 63 Then
            If a(a3 + 10) = bj And a(a4 + 10) = bj Then a(a3 + 10) = m: a(a4 + 10) = m: a(a1) = bj: a(a3) = bj: a1 = a1 + 10: a2 = a2 + 10: a3 = a3 + 10: a4 = a4 + 10: flag = True Else flag = False
    ElseIf f = 7 Then
        If a(a2 + 10) = bj And a(a3 + 10) = bj And a(a4 + 10) = bj Then a(a2 + 10) = m: a(a3 + 10) = m: a(a4 + 10) = m: a(a1) = bj: a(a2) = bj: a(a3) = bj: a1 = a1 + 10: a2 = a2 + 10: a3 = a3 + 10: a4 = a4 + 10: flag = True Else flag = False
        ElseIf f = 71 Then
            If a(a3 + 10) = bj And a(a4 + 10) = bj Then a(a3 + 10) = m: a(a4 + 10) = m: a(a1) = bj: a(a4) = bj: a1 = a1 + 10: a2 = a2 + 10: a3 = a3 + 10: a4 = a4 + 10: flag = True Else flag = False
        ElseIf f = 72 Then
            If a(a2 + 10) = bj And a(a3 + 10) = bj And a(a4 + 10) = bj Then a(a2 + 10) = m: a(a3 + 10) = m: a(a4 + 10) = m: a(a1) = bj: a(a2) = bj: a(a3) = bj: a1 = a1 + 10: a2 = a2 + 10: a3 = a3 + 10: a4 = a4 + 10: flag = True Else flag = False
        ElseIf f = 73 Then
            If a(a1 + 10) = bj And a(a4 + 10) = bj Then a(a1 + 10) = m: a(a4 + 10) = m: a(a1) = bj: a(a2) = bj: a1 = a1 + 10: a2 = a2 + 10: a3 = a3 + 10: a4 = a4 + 10: flag = True Else flag = False
    End If
    If flag = False Then Timer2.Enabled = True: Timer4.Enabled = True: Timer5.Enabled = True
'a(a1 + 10) = a(a1): a(a2 + 10) = a(a2): a(a3 + 10) = a(a3): a(a4 + 10) = a(a4): a(a1) = bj: a(a2) = bj: a(a3) = bj: a(a4) = bj: a1 = a1 + 10: a2 = a2 + 10: a3 = a3 + 10: a4 = a4 + 10
End Sub

Private Sub Timer4_Timer() '结束
If (a(5) = m Or a(6) = m Or a(7) = m) And (a(15) = m Or a(16) = m Or a(17) = m) And (a(21) = m Or a(22) = m Or a(23) = m Or a(24) = m Or a(25) = m Or a(26) = m Or a(27) = m Or a(28) = m Or a(29) = m Or a(30) = m) And (a(31) = m Or a(32) = m Or a(33) = m Or a(34) = m Or a(35) = m Or a(36) = m Or a(37) = m Or a(38) = m Or a(39) = m Or a(40) = m) Then
    Timer1.Enabled = False
    Timer2.Enabled = False
    Timer3.Enabled = False
    Timer4.Enabled = False
    Timer6.Enabled = False
    Command1.Enabled = True
    Command1.Caption = "重新开始"
    Command2.Enabled = False
    Command3.Enabled = False
    Command4.Enabled = False
    Command5.Enabled = False
    Check1.Enabled = True
    Check2.Enabled = True
    Check3.Enabled = True
Else
    Timer4.Enabled = False
    Timer3.Enabled = True
End If
End Sub

Private Sub Timer5_Timer() '消除+计数
Dim i As Integer, j As Integer, num As Integer
For i = 21 To 220
    If a(i) = m Then num = num + 1
    If i Mod 10 = 0 Then
        If num = 10 Then
            For j = i - 9 To i
                a(j) = bj
            Next j
            For j = i To 31 Step -1
                a(j) = a(j - 10)
            Next j
            For j = 21 To 30
                a(j) = bj
            Next j
            n = n + 1
        End If
        num = 0
    End If
Next i
Label1.Caption = CStr(n)
Timer6.Enabled = True
Timer5.Enabled = False
End Sub

Private Sub Timer6_Timer() '调速
If Timer3.Interval >= 250 Then Timer3.Interval = 500 - 50 * (n \ 20)
Timer6.Enabled = False
End Sub
