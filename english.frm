VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "�С㡸�^��ɺA���硹"
   ClientHeight    =   5610
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7830
   LinkTopic       =   "Form1"
   ScaleHeight     =   5610
   ScaleWidth      =   7830
   StartUpPosition =   3  '�t�ιw�]��
   Begin VB.CommandButton Command4 
      Caption         =   "������"
      BeginProperty Font 
         Name            =   "�s�ө���"
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
      Caption         =   "�{�b���������{�b�����i�榡"
      BeginProperty Font 
         Name            =   "�s�ө���"
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
      Caption         =   "�{�b�i�榡"
      BeginProperty Font 
         Name            =   "�s�ө���"
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
      Caption         =   "����²�榡"
      BeginProperty Font 
         Name            =   "�s�ө���"
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
      Caption         =   "�{�b²�榡"
      BeginProperty Font 
         Name            =   "�s�ө���"
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
      Caption         =   "�L�h²�榡"
      BeginProperty Font 
         Name            =   "�s�ө���"
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
      Caption         =   "��ܤ@�ӧa��I"
      BeginProperty Font 
         Name            =   "�s�ө���"
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
      Caption         =   "���@�}"
      BeginProperty Font 
         Name            =   "�s�ө���"
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
      Caption         =   "���@��"
      BeginProperty Font 
         Name            =   "�s�ө���"
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
      Caption         =   "�T�@�w"
      BeginProperty Font 
         Name            =   "�s�ө���"
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
      Caption         =   "���{���� ���q�� �s�@�@e-mail�Goneleo760823@yahoo.com.tw�@���ɦѮv�G�ɱЬɪ����H�w���[�Ѯv�@�w��ӫH���СAŪ�ѥ[�o�ա�^^~"
      BeginProperty Font 
         Name            =   "�s�ө���"
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
      Alignment       =   2  '�m�����
      Caption         =   "�� ���I�惡�^�媺�ɺA�Ϊk ��"
      BeginProperty Font 
         Name            =   "�s�ө���"
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
         Name            =   "�s�ө���"
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
      Alignment       =   2  '�m�����
      BeginProperty Font 
         Name            =   "�s�ө���"
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

E(0) = "yesterday":                     C(0) = "�Q��"
E(1) = "the day before yesterday":      C(1) = "�e��"
E(2) = "the other day":                 C(2) = "�e�X��"
E(3) = "ago":                           C(3) = "�H�e"
E(4) = "in�ϹL�h�~�N":                  C(4) = "�q�e"
E(5) = "once upon a time":              C(5) = "�q�e"
E(6) = "once":                          C(6) = "���g"
E(7) = "in the past":                   C(7) = "�L�h"
E(8) = "just now":                      C(8) = "��~"
E(9) = "last��":                        C(9) = "�W�@..."
E(10) = "then":                         C(10) = "���"
E(11) = "always":                       C(11) = "�`�O"
E(12) = "usually":                      C(12) = "�q�`"
E(13) = "sometimes":                    C(13) = "����"
E(14) = "occasionally":                 C(14) = "����"
E(15) = "at times":                     C(15) = "����"
E(16) = "every now and then":           C(16) = "����"
E(17) = "from time to time":            C(17) = "����"
E(18) = "once in a while":              C(18) = "����"
E(19) = "seldom":                       C(19) = "���`"
E(20) = "every��":                      C(20) = "�C..."
E(21) = "tomorrow":                     C(21) = "����"
E(22) = "the day after tomorrow":       C(22) = "���"
E(23) = "in�Ϥ@�q�ɶ�":                 C(23) = "�b�L�h�["
E(24) = "in�ϥ��Ӧ~�N":                 C(24) = "�U�@..."
E(25) = "next��":                       C(25) = "�U�@..."
E(26) = "soon":                         C(26) = "���["
E(27) = "before long":                  C(27) = "���["
E(28) = "someday":                      C(28) = "���@��"
E(29) = "in the future":                C(29) = "�N��"
E(30) = "now":                          C(30) = "�{�b"
E(31) = "at present":                   C(31) = "�{�b"
E(32) = "for the time being":           C(32) = "�{�b"
E(33) = "Look�I":                       C(33) = "�֬ݡI"
E(34) = "Listen�I":                     C(34) = "��ť�I"
E(35) = "recently":                     C(35) = "�̪�"
E(36) = "lately":                       C(36) = "�̪�"
E(37) = "of late":                      C(37) = "�̪�"
E(38) = "so far":                       C(38) = "����"
E(39) = "until now":                    C(39) = "����"
E(40) = "up to now":                    C(40) = "����"
E(41) = "as yet":                       C(41) = "����"
E(42) = "by now":                       C(42) = "����"
E(43) = "already":                      C(43) = "�w�g"
E(44) = "yet":                          C(44) = "�|��"
E(45) = "just":                         C(45) = "���"
Randomize
X = Int(46 * Rnd)
Label1.Caption = E(X)
R = 0: W = 0: P = 0
End Sub

Private Sub Command2_Click()

If (Option1.Value = False) And (Option2.Value = False) And (Option3.Value = False) And (Option4.Value = False) And (Option5.Value = False) = True Then
    Label2.Caption = "������I�A�o�D�٨S�����I"
ElseIf (Label2.Caption = "������I�A�o�D�٨S�����I") Or (Label2.Caption = "�A�٨S��������@�D�I") Or (Label2.Caption = "") = True Then
    Label2.Caption = "������I�A�o�D�٨S�����I"
Else
    Label2.Caption = ""
    Randomize
    X = Int(46 * Rnd)
    Label1.Caption = E(X)
End If

End Sub

Private Sub Command1_Click()

If (Option1.Value = False) And (Option2.Value = False) And (Option3.Value = False) And (Option4.Value = False) And (Option5.Value = False) = True Then
    Label2.Caption = "���I��@�ӳ̾A���ɺA��I"
ElseIf (Label2.Caption = "") Or (Label2.Caption = "���I��@�ӳ̾A���ɺA��I") Or (Label2.Caption = "������I�A�o�D�٨S�����I") Or (Label2.Caption = "�A�٨S��������@�D�I") = True Then
    Select Case X
        Case 0 To 10
            If Option1.Value = True Then
                Label2.Caption = "����F�I" + Chr(10) + "���^��������С�" + C(X) + Chr(10) + "���ߩp�I"
                R = R + 1
            Else
                Label2.Caption = "�p�����F��I" + Chr(10) + "���^�奿�T�������С�" + C(X) + Chr(10) + "�ҥH���С�" + Option1.Caption
                W = W + 1
            End If
        Case 11 To 20
            If Option2.Value = True Then
                Label2.Caption = "����F�I" + Chr(10) + "���^��������С�" + C(X) + Chr(10) + "���ߩp�I"
                R = R + 1
            Else
                Label2.Caption = "�p�����F��I" + Chr(10) + "���^�奿�T�������С�" + C(X) + Chr(10) + "�ҥH���С�" + Option2.Caption
                W = W + 1
            End If
        Case 21 To 29
            If Option3.Value = True Then
                Label2.Caption = "����F�I" + Chr(10) + "���^��������С�" + C(X) + Chr(10) + "���ߩp�I"
                R = R + 1
            Else
                Label2.Caption = "�p�����F��I" + Chr(10) + "���^�奿�T�������С�" + C(X) + Chr(10) + "�ҥH���С�" + Option3.Caption
                W = W + 1
            End If
        Case 30 To 34
            If Option4.Value = True Then
                Label2.Caption = "����F�I" + Chr(10) + "���^��������С�" + C(X) + Chr(10) + "���ߩp�I"
                R = R + 1
            Else
                Label2.Caption = "�p�����F��I" + Chr(10) + "���^�奿�T�������С�" + C(X) + Chr(10) + "�ҥH���С�" + Option4.Caption
                W = W + 1
            End If
        Case 35 To 45
            If Option5.Value = True Then
                Label2.Caption = "����F�I" + Chr(10) + "���^��������С�" + C(X) + Chr(10) + "���ߩp�I"
                R = R + 1
            Else
                Label2.Caption = "�p�����F��I" + Chr(10) + "���^�奿�T�������С�" + C(X) + Chr(10) + "�ҥH���С�" + Option5.Caption
                W = W + 1
            End If
    End Select
    P = P + 1
End If

End Sub

Private Sub Command3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label2.Caption = "�����o��кC����^^~"

End Sub

Private Sub Command3_Click()

End

End Sub

Private Sub Command4_Click()
If P = 0 Then
    Label2.Caption = "�A�٨S��������@�D�I"
Else
    Select Case Int(R / P * 100)
           Case 0 To 59
               Y = "���㦳�ݥ[�j����b�^��W�Y�n�[��l���[�o�[�o�㢢�]�֡��բ�)"
           Case 60 To 79
               Y = "�I�P�P�㦨�Z������V�O���E�Q���H�W�ڶi��١���"
           Case 80 To 100
               Y = "������n�F�`�n�F�`��n�~��O�����=�١\��="
    End Select
    Label2.Caption = "���ƴ����o��p�`�@���F" + Str(P) + "�D�I��" + "��F" + Str(R) + "�D�A�H�ο��F" + Str(W) + "�D�I" + Chr(10) + "����v�O" + Str(Int(R / P * 100)) + "��" + Chr(10) + Y
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
