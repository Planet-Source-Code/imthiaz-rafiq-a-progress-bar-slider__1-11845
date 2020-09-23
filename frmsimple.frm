VERSION 5.00
Object = "{1D67E38D-FDED-11D3-AF3A-F7C129AE2B4E}#2.0#0"; "Grapher.ocx"
Object = "{EE75CAB5-FF11-11D3-AF3A-D60535FEA24D}#6.0#0"; "Progressor.ocx"
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3660
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   4635
   ControlBox      =   0   'False
   Icon            =   "frmsimple.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3660
   ScaleWidth      =   4635
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin GrapherControl.Grapher g2 
      Height          =   615
      Left            =   3150
      Top             =   4530
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   1085
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Max             =   100
      BarWidth        =   2
      Flat            =   0   'False
      Inverted        =   0   'False
      Bstyle          =   2
   End
   Begin ProgyBar.Progbar p4 
      Height          =   915
      Left            =   240
      TabIndex        =   7
      Top             =   4140
      Width           =   2355
      _ExtentX        =   4154
      _ExtentY        =   1614
      ForeColor       =   255
      BackColor       =   0
      Max             =   100
      Mode            =   0
      Border          =   1
      Mark            =   -1  'True
      MarkThicness    =   3
      MarkColor       =   65535
   End
   Begin Project1.Progbar Progbar2 
      Height          =   285
      Left            =   30
      TabIndex        =   5
      Top             =   2010
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   503
      ForeColor       =   255
      BackColor       =   0
      Max             =   1000
      Mode            =   1
      Border          =   1
      Mark            =   -1  'True
      MarkThicness    =   3
      MarkColor       =   65535
   End
   Begin Project1.Progbar Progbar1 
      Height          =   285
      Left            =   30
      TabIndex        =   4
      Top             =   1680
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   503
      ForeColor       =   255
      BackColor       =   0
      Max             =   1000
      Mode            =   0
      Border          =   1
      Mark            =   -1  'True
      MarkThicness    =   3
      MarkColor       =   65535
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Stop"
      Height          =   435
      Left            =   3150
      TabIndex        =   3
      Top             =   2700
      Width           =   1305
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   1590
      Top             =   930
   End
   Begin VB.CommandButton Command3 
      Caption         =   "End"
      Height          =   435
      Left            =   3150
      TabIndex        =   2
      Top             =   3150
      Width           =   1305
   End
   Begin VB.CommandButton Command2 
      Caption         =   "About Grapher"
      Height          =   435
      Left            =   120
      TabIndex        =   1
      Top             =   3150
      Width           =   1305
   End
   Begin VB.CommandButton Command1 
      Caption         =   "About ProgYbar"
      Height          =   435
      Left            =   120
      TabIndex        =   0
      Top             =   2700
      Width           =   1305
   End
   Begin Project1.Progbar Progbar3 
      Height          =   1605
      Left            =   4440
      TabIndex        =   6
      Top             =   30
      Width           =   195
      _ExtentX        =   344
      _ExtentY        =   2831
      ForeColor       =   255
      BackColor       =   0
      Max             =   1000
      Mode            =   2
      Border          =   0
      Mark            =   -1  'True
      MarkThicness    =   3
      MarkColor       =   65535
   End
   Begin Project1.Grapher Grapher1 
      Height          =   1605
      Left            =   30
      Top             =   30
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   2831
      BackColor       =   0
      ForeColor       =   255
      Max             =   1000
      BarWidth        =   2
      Flat            =   -1  'True
      Inverted        =   0   'False
      Bstyle          =   2
   End
   Begin VB.Image Image1 
      Height          =   1110
      Left            =   1680
      Picture         =   "frmsimple.frx":08CA
      Stretch         =   -1  'True
      Top             =   2460
      Width           =   1320
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Mode As Boolean

Private Sub Command1_Click()
p4.AboutBox
End Sub

Private Sub Command2_Click()
g2.DisplayAboutBox
End Sub

Private Sub Command3_Click()
End
End Sub

Private Sub Command4_Click()
If Timer1.Enabled = True Then
    Timer1.Enabled = False
    Mode = False
    Command4.Caption = "Start"
Else
    Timer1.Enabled = True
    Mode = True
    Command4.Caption = "Stop"
End If
End Sub

Private Sub Form_Load()
    Mode = True
End Sub

Private Sub Image1_Click()
If Grapher1.Drawstyle = bar Then
    Grapher1.Drawstyle = Dots
ElseIf Grapher1.Drawstyle = Dots Then
    Grapher1.Drawstyle = Lines
ElseIf Grapher1.Drawstyle = Lines Then
    Grapher1.Drawstyle = bar
End If
End Sub

Private Sub Progbar1_click(Value As Double)
Dim Y As Double
Dim Z As Long
Y = Value
Z = Y
Grapher1.Update Z
Progbar2.DrawBar Y
Progbar3.DrawBar Y
End Sub

Private Sub Progbar2_click(Value As Double)
Dim Y As Double
Dim Z As Long
Y = Value
Z = Y
Grapher1.Update Z
Progbar1.DrawBar Y
Progbar3.DrawBar Y
End Sub

Private Sub Progbar3_click(Value As Double)
Dim Y As Double
Dim Z As Long
Y = Value
Z = Y
Grapher1.Update Z
Progbar1.DrawBar Y
Progbar2.DrawBar Y
End Sub

Private Sub Timer1_Timer()
Dim Y As Double
Dim Z As Long
Y = Int(Rnd() * 1000)
Z = Y
Grapher1.Update Z
Progbar1.DrawBar Y
If Mode = True Then
    Progbar2.DrawBar Y
    Progbar3.DrawBar Y
End If
    
End Sub
