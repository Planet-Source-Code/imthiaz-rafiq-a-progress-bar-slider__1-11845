VERSION 5.00
Begin VB.Form frmnote 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3750
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   3990
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3750
   ScaleWidth      =   3990
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   -120
      Top             =   -210
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   -150
      Top             =   -150
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   -120
      Top             =   -210
   End
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   -60
      Top             =   -150
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H000040C0&
      Height          =   3105
      Left            =   360
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "frmnote.frx":0000
      Top             =   330
      Width           =   3225
   End
   Begin ProgYYbar.ProgYbar p3 
      Height          =   345
      Left            =   0
      TabIndex        =   1
      Top             =   3390
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   609
      ForeColor       =   255
      BackColor       =   0
      Max             =   100
      Mode            =   1
      Border          =   1
      Mark            =   -1  'True
      MarkThicness    =   3
      MarkColor       =   65535
   End
End
Attribute VB_Name = "frmnote"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim I1 As Double
Dim I2 As Double
Dim I3 As Double
Dim I4 As Double

Dim M1 As Boolean
Dim M2 As Boolean
Dim M3 As Boolean
Dim M4 As Boolean


Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Then
        Unload Me
    Else
        MsgBox "Please read till the fag end and find the key to close this stuff", vbInformation + vbOKOnly, "Sorry for inconvineance"
    End If
End Sub

Private Sub Form_Load()
M1 = True
M2 = True
M3 = True
M4 = True
End Sub

Private Sub Timer1_Timer()
If M1 = True Then
    I1 = I1 + 100
    If I1 > 1000 Then
        Timer1.Enabled = False
        Timer2.Enabled = True
        M1 = False
    End If
End If
If M1 = False Then
    I1 = I1 - 100
    If I1 < 0 Then
        Timer1.Enabled = False
        Timer4.Enabled = True
        M1 = True
    End If
End If
'p1.DrawBar I1
End Sub

Private Sub Timer2_Timer()
If M2 = True Then
    I2 = I2 + 100
    If I2 > 1000 Then
        Timer2.Enabled = False
        Timer3.Enabled = True
        M2 = False
    End If
End If
If M2 = False Then
    I2 = I2 - 100
    If I2 < 0 Then
        Timer2.Enabled = False
        Timer1.Enabled = True
        M2 = True
    End If
End If
'p2.DrawBar I2
End Sub

Private Sub Timer3_Timer()
If M3 = True Then
    I3 = I3 + 100
    If I3 > 1000 Then
        Timer3.Enabled = False
        Timer4.Enabled = True
        M3 = False
    End If
End If
If M3 = False Then
    I3 = I3 - 100
    If I3 < 0 Then
        Timer3.Enabled = False
        Timer2.Enabled = True
        M3 = True
    End If
End If
p3.DrawBar I3

End Sub

Private Sub Timer4_Timer()
If M4 = True Then
    I4 = I4 + 100
    If I4 > 1000 Then
        M4 = False
    End If
End If
If M4 = False Then
    I4 = I4 - 100
    If I4 < 0 Then
        Timer4.Enabled = False
        Timer3.Enabled = True
        M4 = True
    End If
End If
'p4.DrawBar I4
End Sub

