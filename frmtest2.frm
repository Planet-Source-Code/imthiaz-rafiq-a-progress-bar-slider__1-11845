VERSION 5.00
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.1#0"; "COMCT232.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmtest2 
   BackColor       =   &H8000000A&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6780
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   6780
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmtest2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6780
   ScaleWidth      =   6780
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtjump 
      Height          =   315
      Left            =   5280
      TabIndex        =   19
      Text            =   "1"
      Top             =   5160
      Width           =   375
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   480
      Top             =   1800
   End
   Begin VB.TextBox txttime 
      Height          =   315
      Left            =   3240
      TabIndex        =   15
      Text            =   "100"
      Top             =   5160
      Width           =   615
   End
   Begin VB.TextBox txtmt 
      Height          =   315
      Left            =   4320
      TabIndex        =   13
      Text            =   "3"
      Top             =   5160
      Width           =   375
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Mark"
      Height          =   255
      Left            =   6000
      TabIndex        =   12
      Top             =   5160
      Value           =   1  'Checked
      Width           =   975
   End
   Begin MSComDlg.CommonDialog cmd1 
      Left            =   5880
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.ComboBox cmbmode 
      Height          =   315
      Left            =   1440
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   5160
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Exit"
      Height          =   315
      Left            =   5640
      TabIndex        =   4
      Top             =   6240
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Start"
      Height          =   315
      Left            =   4800
      TabIndex        =   3
      Top             =   6240
      Width           =   855
   End
   Begin ProgYYbar.ProgYbar Progbar1 
      Height          =   4095
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   7223
      ForeColor       =   255
      BackColor       =   0
      Max             =   100
      Mode            =   1
      Border          =   1
      Mark            =   -1  'True
      MarkThicness    =   3
      MarkColor       =   65535
   End
   Begin ComCtl2.UpDown UpDown1 
      Height          =   315
      Left            =   840
      TabIndex        =   5
      Top             =   5160
      Width           =   240
      _ExtentX        =   344
      _ExtentY        =   556
      _Version        =   327681
      BuddyControl    =   "txtmax"
      BuddyDispid     =   196619
      OrigLeft        =   6360
      OrigTop         =   2040
      OrigRight       =   6600
      OrigBottom      =   2325
      Increment       =   10
      Max             =   5000000
      SyncBuddy       =   -1  'True
      BuddyProperty   =   65547
      Enabled         =   -1  'True
   End
   Begin ComCtl2.UpDown UpDown2 
      Height          =   315
      Left            =   4680
      TabIndex        =   14
      Top             =   5160
      Width           =   240
      _ExtentX        =   344
      _ExtentY        =   556
      _Version        =   327681
      Value           =   1
      BuddyControl    =   "txtmt"
      BuddyDispid     =   196613
      OrigLeft        =   6360
      OrigTop         =   2040
      OrigRight       =   6600
      OrigBottom      =   2325
      Max             =   7
      Min             =   1
      SyncBuddy       =   -1  'True
      BuddyProperty   =   65547
      Enabled         =   -1  'True
   End
   Begin ComCtl2.UpDown UpDown3 
      Height          =   315
      Left            =   3840
      TabIndex        =   16
      Top             =   5160
      Width           =   240
      _ExtentX        =   344
      _ExtentY        =   556
      _Version        =   327681
      Value           =   1
      BuddyControl    =   "txttime"
      BuddyDispid     =   196612
      OrigLeft        =   6360
      OrigTop         =   2040
      OrigRight       =   6600
      OrigBottom      =   2325
      Increment       =   10
      Max             =   10000
      SyncBuddy       =   -1  'True
      BuddyProperty   =   65547
      Enabled         =   -1  'True
   End
   Begin VB.TextBox txtmax 
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Text            =   "100"
      Top             =   5160
      Width           =   735
   End
   Begin ComCtl2.UpDown UpDown4 
      Height          =   315
      Left            =   5640
      TabIndex        =   20
      Top             =   5160
      Width           =   240
      _ExtentX        =   344
      _ExtentY        =   556
      _Version        =   327681
      Value           =   1
      BuddyControl    =   "txtjump"
      BuddyDispid     =   196610
      OrigLeft        =   6360
      OrigTop         =   2040
      OrigRight       =   6600
      OrigBottom      =   2325
      Increment       =   10
      Max             =   999
      Min             =   1
      SyncBuddy       =   -1  'True
      BuddyProperty   =   65547
      Enabled         =   -1  'True
   End
   Begin VB.Label Label12 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   $"frmtest2.frx":164A
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   120
      TabIndex        =   24
      Top             =   6120
      Width           =   4575
   End
   Begin VB.Label Label11 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3120
      TabIndex        =   23
      Top             =   480
      Width           =   3615
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   22
      Top             =   480
      Width           =   2895
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "Jump"
      Height          =   195
      Left            =   5280
      TabIndex        =   21
      Top             =   4920
      Width           =   375
   End
   Begin VB.Label Label9 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   120
      Width           =   6615
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Time (ms):"
      Height          =   195
      Left            =   3240
      TabIndex        =   17
      Top             =   4920
      Width           =   720
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Mark Thick "
      Height          =   195
      Left            =   4320
      TabIndex        =   11
      Top             =   4920
      Width           =   855
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "MarkColor"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   4920
      TabIndex        =   10
      Top             =   5640
      Width           =   1575
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "BackColor"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   2520
      TabIndex        =   9
      Top             =   5640
      Width           =   1575
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "ForeColor"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   5640
      Width           =   1575
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Mode"
      Height          =   195
      Left            =   1440
      TabIndex        =   7
      Top             =   4920
      Width           =   405
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Progbar max "
      Height          =   315
      Left            =   120
      TabIndex        =   2
      Top             =   4920
      Width           =   930
   End
End
Attribute VB_Name = "frmtest2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Valw As Double
Private Sub Check1_Click()
Progbar1.Mark = Check1.Value
End Sub

Private Sub cmbmode_Change()
    Progbar1.Mode = cmbmode.ListIndex
End Sub

Private Sub cmbmode_Click()
    cmbmode_Change
End Sub

Private Sub Command1_Click()
    If Timer1.Enabled = True Then
        Command1.Caption = "Start"
        Timer1.Enabled = False
    Else
        Command1.Caption = "Stop"
        Timer1.Enabled = True
    End If
End Sub

Private Sub Command2_Click()
Unload Me
'frmnote.Show vbModal
End Sub

Private Sub Command3_Click()
MsgBox "jj"
End Sub

Private Sub Form_Load()
    Progbar1.Max = Val(txtmax.Text)
    cmbmode.AddItem "Horizontal Forward"
    cmbmode.AddItem "Horizontal Backward"
    cmbmode.AddItem "Vertical Upward"
    cmbmode.AddItem "Vertical Downward"
    cmbmode.ListIndex = 0
End Sub

Private Sub Label4_Click()
    cmd1.ShowColor
    Label4.BackColor = cmd1.Color
    Progbar1.ForeColor = cmd1.Color
End Sub

Private Sub Label5_Click()
    cmd1.ShowColor
    Label5.BackColor = cmd1.Color
    Progbar1.BackColor = cmd1.Color
End Sub

Private Sub Label6_Click()
    cmd1.ShowColor
    Label6.BackColor = cmd1.Color
    Progbar1.MarkColor = cmd1.Color
End Sub

Private Sub Progbar1_click(Value As Double)
If Value <> -1 Then
    Label2.Caption = "The Clicked Value is " + Str(Int(Value))
    Valw = Value
    Progbar1.ToolTipText = " Value :" + Str(Int(Value))
Else
    Label2.Caption = "The Clicked Value is out of limits"
End If
End Sub


Private Sub Progbar1_MouseHover(Button As PgButton, X As Single, Y As Single, Value As Double)
If Value <> -1 Then
    Label11.Caption = " X :" + Str(X) + " Y :" + Str(Y) + " Value :" + Str(Int(Value))
Else
    Label11.Caption = " X :" + Str(X) + " Y :" + Str(Y) + " Value : (Out of limits)"
End If
End Sub

Private Sub Progbar1_ValueChange(Newval As Double, Oldval As Double)
    Label9.Caption = "Last Updated Value is :" + Str(Int(Oldval)) + " Current Value is :" + Str(Int(Newval))
End Sub

Private Sub Timer1_Timer()
Valw = Valw + Val(txtjump.Text)
If Valw >= Progbar1.Max Then Valw = 0
Progbar1.DrawBar Valw
DoEvents
End Sub

Private Sub UpDown1_Change()
    Progbar1.Max = Val(txtmax.Text)
End Sub

Private Sub UpDown2_Change()
    Progbar1.MarkThickness = Val(txtmt.Text)
End Sub

Private Sub UpDown3_Change()
    Timer1.Interval = Val(txttime.Text)
    If Val(txttime.Text) < 25 Then
        UpDown3.Increment = 1
    Else
        UpDown3.Increment = 10
    End If
        
End Sub

Private Sub UpDown4_Change()
    If Val(txtjump.Text) < 25 Then
        UpDown4.Increment = 1
    Else
        UpDown4.Increment = 10
    End If
End Sub
