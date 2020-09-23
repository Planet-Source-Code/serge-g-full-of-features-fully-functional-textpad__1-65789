VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmSearchTxt 
   BackColor       =   &H80000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Find"
   ClientHeight    =   3300
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5625
   Icon            =   "frmSearchTxt.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3300
   ScaleWidth      =   5625
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ProgressBar pBar1 
      Height          =   105
      Left            =   75
      TabIndex        =   12
      Top             =   3165
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   185
      _Version        =   393216
      Appearance      =   0
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   195
      Left            =   0
      TabIndex        =   11
      Top             =   3105
      Width           =   5625
      _ExtentX        =   9922
      _ExtentY        =   344
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Done"
      Height          =   375
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1680
      Width           =   1215
   End
   Begin VB.CommandButton cmdReplaceAll 
      Caption         =   "Replace All"
      Height          =   375
      Left            =   4200
      TabIndex        =   6
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton cmdReplace 
      Caption         =   "Replace"
      Height          =   375
      Left            =   4200
      TabIndex        =   5
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "Find &Next"
      Default         =   -1  'True
      Height          =   375
      Left            =   4200
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000000&
      Caption         =   "Search Options"
      Height          =   1695
      Left            =   120
      TabIndex        =   3
      ToolTipText     =   "Click toggle betwwen search and replace options"
      Top             =   1320
      Width           =   3855
      Begin VB.CheckBox chkMCase 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Match Case"
         Height          =   375
         Left            =   2040
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   960
         Width           =   1575
      End
      Begin VB.CheckBox chkWWord 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Whole Word"
         Height          =   375
         Left            =   2040
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   360
         Width           =   1575
      End
      Begin VB.CheckBox chkStartFromTop 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Start from top"
         Height          =   375
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   960
         Width           =   1575
      End
      Begin VB.CheckBox chkReplace 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Replace"
         Height          =   375
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.TextBox txtReplace 
      Height          =   375
      Left            =   120
      TabIndex        =   2
      ToolTipText     =   "Enter the replace string"
      Top             =   840
      Width           =   3855
   End
   Begin VB.TextBox txtFind 
      Height          =   375
      Left            =   120
      TabIndex        =   1
      ToolTipText     =   "Enter a search string"
      Top             =   240
      Width           =   3855
   End
End
Attribute VB_Name = "frmSearchTxt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim firstFind As Integer
Dim onceAround As Boolean
Dim ttlOccured As Long
Dim searchOpts As FindConstants
Dim tempStat As Boolean


Private Sub Check2_Click()

End Sub

Private Sub chkReplace_Click()

    checkStatus

End Sub

Private Sub cmdCancel_Click()

    iAmActive = 0
    Unload Me

End Sub

Private Sub cmdFind_Click()

    searchOption

    Dim searchingFor, tempStr
    
    If chkStartFromTop.Value = 1 And firstFind = 0 Then
10        searchingFor = frmTextPad.rtbox1.Find(txtFind.Text, frmTextPad.rtbox1.SelStart = 1 + firstFind, , searchOpts)
    Else
        searchingFor = frmTextPad.rtbox1.Find(txtFind.Text, frmTextPad.rtbox1.SelStart + firstFind, , searchOpts)
    End If
    
    frmTextPad.rtbox1.SetFocus
     
        If searchingFor = -1 Then
      
            Me.Hide
                If onceAround = False Then
                    sft = MsgBox("Start from the top?", vbYesNo, "Not Found")
                Else
                    MsgBox ("Finished searching")
                    ttlOccured = 0
                    onceAround = False
                    Unload Me
                    Exit Sub
                End If
        Else
        End If
        
        firstFind = firstFind + 1
    If onceAround = False Then
        If sft = vbYes Then
            frmTextPad.rtbox1.SelStart = 0
            firstFind = 0
            Me.Show
            frmTextPad.rtbox1.SetFocus
            onceAround = True
            GoSub 10
        Else
            'Form_Unload (1)
        End If

    End If

End Sub

Private Sub cmdReplace_Click()

    searchOption

    If Len(frmTextPad.rtbox1.SelText) > 0 Then
        frmTextPad.rtbox1.SelText = txtReplace.Text
    End If

End Sub

Private Sub cmdReplaceAll_Click()
    
    s = 0
    ttlOccured = 0
    
    searchOption
    
    On Error Resume Next
    
    For i = 1 To Len(frmTextPad.rtbox1.Text)
     
        X = frmTextPad.rtbox1.Find(txtFind.Text, s, , searchOpts)
        If X <> -1 Then
            frmTextPad.rtbox1.SelText = txtReplace.Text
            s = s + 1
            ttlOccured = ttlOccured + 1
            pBar1.Value = pBar1.Value + 1
        Else
            Exit For
        End If
        
    Next i
        
    pBar1.Value = 100
    
 
'    Dim searchingFor
'
'    'searchingFor = frmTextPad.RTBox1.Find(frmSearchTxt.txtFind.Text, ,, rtfWholeWord)
'    frmTextPad.RTBox1.SetFocus
'    'If searchingFor <> -1 Then
'    '    frmTextPad.RTBox1.SelText = txtReplace.Text
'    'End If
'
'    Dim s
'
'    ttlOccured = 0
'
'    For i = 1 To Len(frmTextPad.RTBox1.Text)
'        searchingFor = frmTextPad.RTBox1.Find(frmSearchTxt.txtFind.Text, searchingFor + Len(txtFind.Text), , rtfWholeWord)
'        If searchingFor <> -1 Then
'            frmTextPad.RTBox1.SelText = txtReplace.Text
'            ttlOccured = ttlOccured + 1
'        End If
'    Next i
'
''    Do While Not searchingFor = -1
''        searchingFor = frmTextPad.RTBox1.Find(frmSearchTxt.txtFind.Text, searchingFor + Len(txtFind.Text), , rtfWholeWord)
''        frmTextPad.RTBox1.SelText = txtReplace.Text
''        ttlOccured = ttlOccured + 1
''    Loop
'
    Me.Hide

    If ttlOccured > 0 Then
        MsgBox ("Total of " & ttlOccured & " strings replaced"), vbInformation, "Complete"
        ttlOccured = 0
    Else
        MsgBox ("No strings replaced"), vbInformation, "Complete"
    End If

    Me.Show
    
    pBar1.Value = 0

End Sub


Private Sub Form_Load()

    tempStat = TopMost
    TopMost = False
    SetTopMost

    'iAmActive = 10
    firstFind = 0
    ttlOccured = 0
    
    Me.Top = Screen.Height / 2 - (Me.Height / 2)
    Me.Left = Screen.Width / 2 - (Me.Width / 2)
    Me.Icon = LoadPicture("")
    
    'txtFind.Text = frmTextPad.RTBox1.SelText
    
    checkStatus
    txtFind_Change
    onceAround = False
    
    Me.Show
    
    txtFind.SetFocus
    
End Sub

Sub checkStatus()

    If chkReplace.Value = 1 Then
        chkReplace.Caption = "Replace"
        Me.Caption = "Search and Replace"
        
        If txtFind.Text <> "" Then
            cmdReplace.Enabled = True
            cmdReplaceAll.Enabled = True
        End If
        txtReplace.Enabled = True
        txtReplace.BackColor = &HFFFFFF
    Else
        chkReplace.Caption = "Replace"
        Me.Caption = "Search"
        cmdReplace.Enabled = False
        cmdReplaceAll.Enabled = False
        txtReplace.Enabled = False
        txtReplace.BackColor = &HC0C0C0    '&HE0E0E0
    End If


End Sub

Private Sub Form_Resize()

    txtFind.SetFocus
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    iAmActive = 0
    'resetFrmTextPad
    TopMost = tempStat
    SetTopMost
    Unload Me

End Sub
Private Sub txtFind_Change()

    If txtFind.Text = "" Then
        cmdFind.Enabled = False
        cmdReplace.Enabled = False
        cmdReplaceAll.Enabled = False
    Else
        cmdFind.Enabled = True
        If chkReplace.Value = 1 Then
            cmdReplace.Enabled = True
            cmdReplaceAll.Enabled = True
        End If
    End If

End Sub


Sub searchOption()

    If chkWWord.Value = 1 And chkMCase.Value = 1 Then
        searchOpts = rtfMatchCase Or rtfWholeWord
    ElseIf chkWWord.Value = 1 And chkMCase.Value = 0 Then
        searchOpts = rtfWholeWord
    ElseIf chkWWord.Value = 0 And chkMCase.Value = 1 Then
        searchOpts = rtfMatchCase
    Else
        searchOpts = 0
    End If

End Sub

Private Sub txtFind_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        cmdFind_Click
    ElseIf KeyAscii = 27 Then
        cmdCancel_Click
    End If

End Sub
