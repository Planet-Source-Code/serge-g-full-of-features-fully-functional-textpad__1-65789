VERSION 5.00
Begin VB.Form frmSetTimer 
   ClientHeight    =   2175
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3495
   Icon            =   "frmSetTimer.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2175
   ScaleWidth      =   3495
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   0
      TabIndex        =   9
      Top             =   1560
      Width           =   1455
      Begin VB.OptionButton optAM 
         Caption         =   "AM"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   615
      End
      Begin VB.OptionButton optPM 
         Caption         =   "PM"
         Height          =   255
         Left            =   720
         TabIndex        =   11
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   2760
      Top             =   1920
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   2400
      Top             =   1920
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   2040
      Top             =   1920
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Reset"
      Height          =   255
      Left            =   1560
      TabIndex        =   8
      Top             =   1680
      Width           =   1695
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Cancel"
      Height          =   300
      Left            =   1920
      TabIndex        =   6
      Top             =   555
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Alarm"
      Height          =   375
      Left            =   2640
      TabIndex        =   5
      Top             =   1200
      Width           =   615
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Countdown"
      Height          =   375
      Left            =   1560
      TabIndex        =   4
      Top             =   1200
      Width           =   975
   End
   Begin VB.TextBox txtMinute 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   960
      MaxLength       =   2
      TabIndex        =   2
      Top             =   1200
      Width           =   495
   End
   Begin VB.TextBox txtHour 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   120
      MaxLength       =   2
      TabIndex        =   1
      Top             =   1200
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Set Timer"
      Height          =   300
      Left            =   240
      TabIndex        =   0
      Top             =   555
      Width           =   1335
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   1680
      Top             =   1920
   End
   Begin VB.Label lblTimer 
      Alignment       =   2  'Center
      Caption         =   "Timer"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   375
      Left            =   360
      TabIndex        =   7
      Top             =   120
      Width           =   2415
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      TabIndex        =   3
      Top             =   1200
      Width           =   375
   End
   Begin VB.Line Line1 
      X1              =   15
      X2              =   3495
      Y1              =   900
      Y2              =   900
   End
End
Attribute VB_Name = "frmSetTimer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim h As Integer
Dim m As Long
Dim s As Long
Dim ttl As Long
Dim theT As String
Dim cnt As Integer
Dim tempStat As Boolean

Private Sub Command1_Click()

    If Command1.Caption = "Set Timer" Then
        Command1.Caption = "Close"
        Me.Height = 2580
    Else
        Command1.Caption = "Set Timer"
        Me.Height = 1350
    End If

End Sub

Private Sub Command2_Click()

    lblTimer.ForeColor = &H8000000F

    h = 0
    m = 0
    s = 0

    If txtHour.Text <> "" Then
        h = CInt(txtHour.Text)
        m = h * 60
    End If
    
    If txtMinute.Text <> "" Then
        m = m + CLng(txtMinute.Text)
    End If
    
    If m <= 0 Then
        Command1_Click
        lblTimer.Caption = "Not Set"
        Exit Sub
    Else
        Dim start, theTime
        Dim X As Integer
        s = m * 60
        ttl = s
        Timer2.Enabled = True
'        start = Timer
'        pausetime = s
'
'            Do While Timer < start + pausetime
'                DoEvents
'                x = pausetime + start - Timer
'                theTime = Int(x / 3600) & ":" & Int(x / 60) & ":" & Int(x Mod 60)
'                lblTimer.Caption = Format(theTime, "hh:mm:ss")
'            Loop
    End If

End Sub

Private Sub Command3_Click()

    h = 0
    m = 0
    s = 0
    theT = ""

    If txtHour.Text = "" And txtMinute.Text = "" Then
        Command1_Click
        Exit Sub
    End If

    If txtHour.Text <> "" Then
        h = CInt(txtHour.Text)
    Else
        h = 0
    End If
    
    If txtMinute.Text <> "" Then
        m = CInt(txtMinute.Text)
    Else
        m = 0
    End If
    
    If h < 13 And optPM = True Then
        h = h + 12
    End If
    
    theT = h & ":" & m & ":" & s
    theT = Format(theT, "hh:mm:ss")
    
    cnt = 3
    
    Timer3.Enabled = True
    

End Sub

Private Sub Command4_Click()

    Unload Me

End Sub

Private Sub Command5_Click()

    lblTimer.ForeColor = &H8000000F
    lblTimer.Caption = "Timer"

    txtMinute.Text = ""
    txtHour.Text = ""

End Sub

Private Sub Form_Load()

    tempStat = TopMost
    TopMost = False
    SetTopMost
    
    Me.Caption = "Time : " & Time
    lblTimer.ForeColor = &H8000000F
    
    Me.Height = 1350
    
    If InStr(1, Time, "am", vbTextCompare) <> 0 Then
        optAM.Value = True
    Else
        optPM.Value = True
    End If

End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)

    If KeyAscii < 32 Then
       Exit Sub
    End If
    
    If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
       KeyAscii = 0
    End If

End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)

    If KeyAscii < 32 Then
       Exit Sub
    End If
    
    If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
       KeyAscii = 0
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)

    TopMost = tempStat
    SetTopMost
    
    With frmTextPad.StatusBar1.Panels(2)
        .Style = sbrTime
        .Bevel = sbrInset
    End With

End Sub

Private Sub Timer1_Timer()

    Me.Caption = "Time : " & Time

End Sub

Private Sub Timer2_Timer()

    ttl = ttl - 1
    theT = Int(ttl / 3600) & ":" & Int((ttl Mod 3600) / 60) & ":" & Int(ttl Mod 60)
    lblTimer.Caption = Format(theT, "hh:mm:ss")
    
    If ttl <= 10 Then
        
        With frmTextPad.StatusBar1.Panels(2)
            .Style = sbrText
            .Text = "Timer : " & ttl
        End With
        
        If frmTextPad.StatusBar1.Panels(2).Bevel = sbrInset Then
            frmTextPad.StatusBar1.Panels(2).Bevel = sbrRaised
        ElseIf frmTextPad.StatusBar1.Panels(2).Bevel = sbrRaised Then
            frmTextPad.StatusBar1.Panels(2).Bevel = sbrNoBevel
        ElseIf frmTextPad.StatusBar1.Panels(2).Bevel = sbrNoBevel Then
            frmTextPad.StatusBar1.Panels(2).Bevel = sbrInset
        End If
        
    End If
    
    If ttl <= 3 Then
        Me.WindowState = vbNormal
        Me.Show
    End If
    
    If ttl <= 0 Then
        
        Beep
        
        With frmTextPad.StatusBar1.Panels(2)
            .Style = sbrTime
            .Bevel = sbrInset
        End With
        
        Timer2.Enabled = False
        Exit Sub
    End If

End Sub

Private Sub Timer3_Timer()

    lblTimer.ForeColor = QBColor(12)
    lblTimer.Caption = theT

    If Time >= CDate(theT) Then
        Me.WindowState = vbNormal
        Me.Show
        Beep
        Timer3.Enabled = False
        Timer4.Enabled = True
    End If

End Sub

Private Sub Timer4_Timer()

    If cnt <= 0 Then
        Timer4.Enabled = False
        
        With frmTextPad.StatusBar1.Panels(2)
            .Style = sbrTime
            .Bevel = sbrInset
        End With
        
        Exit Sub
    End If
    
    With frmTextPad.StatusBar1.Panels(2)
    .Style = sbrText
    .Text = theT
    End With
    
    If frmTextPad.StatusBar1.Panels(2).Bevel = sbrInset Then
        frmTextPad.StatusBar1.Panels(2).Bevel = sbrRaised
    ElseIf frmTextPad.StatusBar1.Panels(2).Bevel = sbrRaised Then
        frmTextPad.StatusBar1.Panels(2).Bevel = sbrNoBevel
    ElseIf frmTextPad.StatusBar1.Panels(2).Bevel = sbrNoBevel Then
        frmTextPad.StatusBar1.Panels(2).Bevel = sbrInset
    End If
    
    cnt = cnt - 1

End Sub
