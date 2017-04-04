VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "羽～☆英文時態測驗★"
   ClientHeight    =   5610
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7830
   LinkTopic       =   "Form1"
   ScaleHeight     =   5610
   ScaleWidth      =   7830
   StartUpPosition =   3  '系統預設值
   Begin VB.CommandButton Command4 
      Caption         =   "打分數"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   14.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   6435
      TabIndex        =   12
      Top             =   195
      Width           =   1200
   End
   Begin VB.OptionButton Option5 
      Caption         =   "現在完成式／現在完成進行式"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   2925
      TabIndex        =   11
      Top             =   3510
      Width           =   3720
   End
   Begin VB.OptionButton Option4 
      Caption         =   "現在進行式"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   390
      TabIndex        =   10
      Top             =   3510
      Width           =   1770
   End
   Begin VB.OptionButton Option3 
      Caption         =   "未來簡單式"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5655
      TabIndex        =   8
      Top             =   2730
      Width           =   1770
   End
   Begin VB.OptionButton Option2 
      Caption         =   "現在簡單式"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2925
      TabIndex        =   7
      Top             =   2730
      Width           =   1770
   End
   Begin VB.OptionButton Option1 
      Caption         =   "過去簡單式"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   390
      TabIndex        =   6
      Top             =   2730
      Width           =   1770
   End
   Begin VB.Frame Frame1 
      Caption         =   "選擇一個吧～！"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   195
      TabIndex        =   5
      Top             =   2535
      Width           =   7425
   End
   Begin VB.CommandButton Command3 
      Caption         =   "離　開"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   14.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   6435
      TabIndex        =   3
      Top             =   1365
      Width           =   1200
   End
   Begin VB.CommandButton Command2 
      Caption         =   "重　來"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   14.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   195
      TabIndex        =   2
      Top             =   1365
      Width           =   1200
   End
   Begin VB.CommandButton Command1 
      Caption         =   "確　定"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   14.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   195
      TabIndex        =   1
      Top             =   195
      Width           =   1200
   End
   Begin VB.Label Label4 
      Caption         =   "此程式由 陳敬翔 製作　e-mail：oneleo760823@yahoo.com.tw　指導老師：補教界的巨人─齊涵老師　歡迎來信指教，讀書加油啦～^^~"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   8.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF80FF&
      Height          =   405
      Left            =   195
      TabIndex        =   13
      Top             =   5265
      Width           =   7230
   End
   Begin VB.Label Label3 
      Alignment       =   2  '置中對齊
      Caption         =   "§ 請點選此英文的時態用法 §"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   15.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   1560
      TabIndex        =   9
      Top             =   195
      Width           =   4695
   End
   Begin VB.Label Label2 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   960
      Left            =   195
      TabIndex        =   4
      Top             =   4290
      Width           =   7425
   End
   Begin VB.Label Label1 
      Alignment       =   2  '置中對齊
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   24
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1185
      Left            =   1560
      TabIndex        =   0
      Top             =   975
      Width           =   4695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim E(45), C(45), Y As String
Dim X, R, W, P, Pout, Delay As Integer

Private Sub Form_Load()

E(0) = "yesterday":                     C(0) = "昨天"
E(1) = "the day before yesterday":      C(1) = "前天"
E(2) = "the other day":                 C(2) = "前幾天"
E(3) = "ago":                           C(3) = "以前"
E(4) = "in＋過去年代":                  C(4) = "從前"
E(5) = "once upon a time":              C(5) = "從前"
E(6) = "once":                          C(6) = "曾經"
E(7) = "in the past":                   C(7) = "過去"
E(8) = "just now":                      C(8) = "剛才"
E(9) = "last～":                        C(9) = "上一..."
E(10) = "then":                         C(10) = "當時"
E(11) = "always":                       C(11) = "總是"
E(12) = "usually":                      C(12) = "通常"
E(13) = "sometimes":                    C(13) = "有時"
E(14) = "occasionally":                 C(14) = "有時"
E(15) = "at times":                     C(15) = "有時"
E(16) = "every now and then":           C(16) = "有時"
E(17) = "from time to time":            C(17) = "有時"
E(18) = "once in a while":              C(18) = "有時"
E(19) = "seldom":                       C(19) = "不常"
E(20) = "every～":                      C(20) = "每..."
E(21) = "tomorrow":                     C(21) = "明天"
E(22) = "the day after tomorrow":       C(22) = "後天"
E(23) = "in＋一段時間":                 C(23) = "在過多久"
E(24) = "in＋未來年代":                 C(24) = "下一..."
E(25) = "next～":                       C(25) = "下一..."
E(26) = "soon":                         C(26) = "不久"
E(27) = "before long":                  C(27) = "不久"
E(28) = "someday":                      C(28) = "有一天"
E(29) = "in the future":                C(29) = "將來"
E(30) = "now":                          C(30) = "現在"
E(31) = "at present":                   C(31) = "現在"
E(32) = "for the time being":           C(32) = "現在"
E(33) = "Look！":                       C(33) = "快看！"
E(34) = "Listen！":                     C(34) = "快聽！"
E(35) = "recently":                     C(35) = "最近"
E(36) = "lately":                       C(36) = "最近"
E(37) = "of late":                      C(37) = "最近"
E(38) = "so far":                       C(38) = "迄今"
E(39) = "until now":                    C(39) = "迄今"
E(40) = "up to now":                    C(40) = "迄今"
E(41) = "as yet":                       C(41) = "迄今"
E(42) = "by now":                       C(42) = "迄今"
E(43) = "already":                      C(43) = "已經"
E(44) = "yet":                          C(44) = "尚未"
E(45) = "just":                         C(45) = "剛剛"
Randomize
X = Int(46 * Rnd)
Label1.Caption = E(X)
R = 0: W = 0: P = 0
End Sub

Private Sub Command2_Click()

If (Option1.Value = False) And (Option2.Value = False) And (Option3.Value = False) And (Option4.Value = False) And (Option5.Value = False) = True Then
    Label2.Caption = "等等～！你這題還沒答完！"
ElseIf (Label2.Caption = "等等～！你這題還沒答完！") Or (Label2.Caption = "你還沒答完任何一題！") Or (Label2.Caption = "") = True Then
    Label2.Caption = "等等～！你這題還沒答完！"
Else
    Label2.Caption = ""
    Randomize
    X = Int(46 * Rnd)
    Label1.Caption = E(X)
End If

End Sub

Private Sub Command1_Click()

If (Option1.Value = False) And (Option2.Value = False) And (Option3.Value = False) And (Option4.Value = False) And (Option5.Value = False) = True Then
    Label2.Caption = "請點選一個最適當的時態～！"
ElseIf (Label2.Caption = "") Or (Label2.Caption = "請點選一個最適當的時態～！") Or (Label2.Caption = "等等～！你這題還沒答完！") Or (Label2.Caption = "你還沒答完任何一題！") = True Then
    Select Case X
        Case 0 To 10
            If Option1.Value = True Then
                Label2.Caption = "答對了！" + Chr(10) + "此英文解釋為－＞" + C(X) + Chr(10) + "恭喜妳！"
                R = R + 1
            Else
                Label2.Caption = "妳答錯了喔！" + Chr(10) + "此英文正確解釋為－＞" + C(X) + Chr(10) + "所以為－＞" + Option1.Caption
                W = W + 1
            End If
        Case 11 To 20
            If Option2.Value = True Then
                Label2.Caption = "答對了！" + Chr(10) + "此英文解釋為－＞" + C(X) + Chr(10) + "恭喜妳！"
                R = R + 1
            Else
                Label2.Caption = "妳答錯了喔！" + Chr(10) + "此英文正確解釋為－＞" + C(X) + Chr(10) + "所以為－＞" + Option2.Caption
                W = W + 1
            End If
        Case 21 To 29
            If Option3.Value = True Then
                Label2.Caption = "答對了！" + Chr(10) + "此英文解釋為－＞" + C(X) + Chr(10) + "恭喜妳！"
                R = R + 1
            Else
                Label2.Caption = "妳答錯了喔！" + Chr(10) + "此英文正確解釋為－＞" + C(X) + Chr(10) + "所以為－＞" + Option3.Caption
                W = W + 1
            End If
        Case 30 To 34
            If Option4.Value = True Then
                Label2.Caption = "答對了！" + Chr(10) + "此英文解釋為－＞" + C(X) + Chr(10) + "恭喜妳！"
                R = R + 1
            Else
                Label2.Caption = "妳答錯了喔！" + Chr(10) + "此英文正確解釋為－＞" + C(X) + Chr(10) + "所以為－＞" + Option4.Caption
                W = W + 1
            End If
        Case 35 To 45
            If Option5.Value = True Then
                Label2.Caption = "答對了！" + Chr(10) + "此英文解釋為－＞" + C(X) + Chr(10) + "恭喜妳！"
                R = R + 1
            Else
                Label2.Caption = "妳答錯了喔！" + Chr(10) + "此英文正確解釋為－＞" + C(X) + Chr(10) + "所以為－＞" + Option5.Caption
                W = W + 1
            End If
    End Select
    P = P + 1
End If

End Sub

Private Sub Command3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label2.Caption = "拜拜囉～請慢走～^^~"

End Sub

Private Sub Command3_Click()

End

End Sub

Private Sub Command4_Click()
If P = 0 Then
    Label2.Caption = "你還沒答完任何一題！"
Else
    Select Case Int(R / P * 100)
           Case 0 To 59
               Y = "娃～有待加強捏～在英文上頭要加把勁喔～加油加油～╰（＞△＜╰)"
           Case 60 To 79
               Y = "呼嚕嚕～成績不錯～努力往九十分以上邁進～≧△≦"
           Case 80 To 100
               Y = "哈哈～好厲害好厲害～要繼續保持喔～=≧﹏≦="
    End Select
    Label2.Caption = "分數揭曉囉～妳總共答了" + Str(P) + "題！～" + "對了" + Str(R) + "題，以及錯了" + Str(W) + "題！" + Chr(10) + "答對率是" + Str(Int(R / P * 100)) + "℅" + Chr(10) + Y
End If
R = 0: W = 0: P = 0

End Sub

Private Sub Form_Resize()
If Form1.WindowState = 0 Then
    If (Form1.Height <> 6120) Or (Form1.Width <> 7950) = True Then
        Form1.Height = 6120
        Form1.Width = 7950
    End If
ElseIf Form1.WindowState = 2 Then
    Form1.WindowState = 0
End If

End Sub
