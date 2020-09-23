VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5985
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6225
   LinkTopic       =   "Form1"
   ScaleHeight     =   5985
   ScaleWidth      =   6225
   StartUpPosition =   3  'Windows Default
   Begin Project1.Progbar Progbar1 
      Height          =   1215
      Left            =   1320
      TabIndex        =   4
      Top             =   1680
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   2143
      ForeColor       =   255
      BackColor       =   0
      Max             =   100
      Mode            =   0
      Border          =   1
      Mark            =   -1  'True
      MarkThicness    =   3
      MarkColor       =   65535
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   735
      Left            =   2280
      TabIndex        =   3
      Top             =   4920
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   2280
      TabIndex        =   2
      Top             =   4200
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1200
      TabIndex        =   1
      Text            =   "0"
      Top             =   4200
      Width           =   975
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label1"
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   4200
      Width           =   975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Dim X As Long
Dim per As Double
Dim cx As Single

per = Val(Text1.Text)

X = Me.ScaleWidth


cx = (per * X) / 100

Label1.Caption = cx
Form_MouseMove 1, 0, cx, 0

End Sub
Private Sub Form_DblClick()
Form1.Cls
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim x1 As Long
Dim per As Double
Form1.Caption = "X :" + Str(X) + " Y :" + Str(Y) + " Button " + Str(Button)
If Button = 1 Then
    Line (0, 0)-(X, 4200), vbBlue, BF
End If
x1 = X
per = (x1 / Me.ScaleWidth) * 100
Label1.Caption = per
Form1.Caption = Form1.Caption + Str(per)
End Sub


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Form1.Caption = "X :" + Str(X) + " Y :" + Str(Y) + " Button " + Str(Button)
Dim x1 As Long
Dim per As Double
Form1.Caption = "X :" + Str(X) + " Y :" + Str(Y) + " Button " + Str(Button)
If Button = 1 Then
    Line (0, 0)-(X, 4200), vbBlue, BF
End If
x1 = X
per = (x1 / Me.ScaleWidth) * 100
Label1.Caption = per
Form1.Caption = Form1.Caption + Str(per)

End Sub

