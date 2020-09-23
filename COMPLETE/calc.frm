VERSION 5.00
Begin VB.Form frmCalc 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Calculator"
   ClientHeight    =   3150
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   2070
   Icon            =   "calc.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3150
   ScaleWidth      =   2070
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdEqual 
      BackColor       =   &H00E0E0E0&
      Caption         =   "="
      Height          =   375
      Left            =   645
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   2385
      Width           =   1125
   End
   Begin VB.CommandButton cmdDiv 
      BackColor       =   &H00E0E0E0&
      Caption         =   "/"
      Height          =   375
      Left            =   1395
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   1965
      Width           =   375
   End
   Begin VB.CommandButton cmdMult 
      BackColor       =   &H00E0E0E0&
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1395
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   1605
      Width           =   375
   End
   Begin VB.CommandButton cmdDecimal 
      Caption         =   "."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   975
      TabIndex        =   17
      Top             =   1980
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "9"
      Height          =   375
      Index           =   9
      Left            =   975
      TabIndex        =   9
      Top             =   1620
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "6"
      Height          =   375
      Index           =   6
      Left            =   975
      TabIndex        =   6
      Top             =   1260
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "3"
      Height          =   375
      Index           =   3
      Left            =   975
      TabIndex        =   3
      Top             =   900
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "0"
      Height          =   375
      Index           =   0
      Left            =   615
      TabIndex        =   10
      Top             =   1980
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "8"
      Height          =   375
      Index           =   8
      Left            =   615
      TabIndex        =   8
      Top             =   1620
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "5"
      Height          =   375
      Index           =   5
      Left            =   615
      TabIndex        =   5
      Top             =   1260
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "2"
      Height          =   375
      Index           =   2
      Left            =   615
      TabIndex        =   2
      Top             =   900
      Width           =   375
   End
   Begin VB.CommandButton cmdNegative 
      BackColor       =   &H00E0E0E0&
      Caption         =   "+/-"
      Height          =   375
      Left            =   255
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   1980
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "7"
      Height          =   375
      Index           =   7
      Left            =   255
      TabIndex        =   7
      Top             =   1620
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "4"
      Height          =   375
      Index           =   4
      Left            =   255
      TabIndex        =   4
      Top             =   1260
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "1"
      Height          =   375
      Index           =   1
      Left            =   255
      TabIndex        =   1
      Top             =   900
      Width           =   375
   End
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H80000018&
      Caption         =   "C"
      Height          =   375
      Left            =   255
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2385
      Width           =   375
   End
   Begin VB.CommandButton cmdMinu 
      BackColor       =   &H00E0E0E0&
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1395
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   1245
      Width           =   375
   End
   Begin VB.CommandButton cmdPlus 
      BackColor       =   &H00E0E0E0&
      Caption         =   "+"
      Height          =   375
      Left            =   1395
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   885
      Width           =   375
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00C0C0C0&
      Height          =   495
      Left            =   0
      ScaleHeight     =   435
      ScaleWidth      =   1965
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   15
      Width           =   2025
      Begin VB.TextBox txtSign 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   180
         Left            =   120
         Locked          =   -1  'True
         MaxLength       =   1
         TabIndex        =   20
         TabStop         =   0   'False
         Text            =   "?"
         Top             =   120
         Width           =   180
      End
      Begin VB.TextBox txtCalc 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   75
         MaxLength       =   15
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   75
         Width           =   1785
      End
   End
   Begin VB.PictureBox pic1 
      BackColor       =   &H00C0C0C0&
      Height          =   2520
      Left            =   15
      ScaleHeight     =   2460
      ScaleWidth      =   1965
      TabIndex        =   18
      Top             =   585
      Width           =   2025
      Begin VB.Frame Frame1 
         BackColor       =   &H00C0C0C0&
         Height          =   2295
         Left            =   75
         TabIndex        =   19
         Top             =   15
         Width           =   1800
      End
   End
   Begin VB.Menu hidMen 
      Caption         =   "Menu"
      Begin VB.Menu mnuCopyItem 
         Caption         =   "Copy"
      End
      Begin VB.Menu mnuInItem 
         Caption         =   "Insert Result"
      End
      Begin VB.Menu mnuScrapItem 
         Caption         =   "Scrap Page"
      End
      Begin VB.Menu mnuSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMinimizeItem 
         Caption         =   "Minimize"
      End
      Begin VB.Menu sep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEItem 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "frmCalc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim currTxt As Double
Dim currTtl As Double
Dim decIn As Boolean
Dim nxt As Boolean
Dim fstTime As Boolean
Dim negPos As Boolean
Dim scrapPage As Variant
Dim retChr As String

Private Sub cmdClear_Click()

    txtSign.Text = ""
    currTxt = 0
    txtCalc.Text = "0"
    currTtl = 0
    
    Form_Load
    
    addScrap (retChr)
    
    pic1.SetFocus

End Sub

Private Sub cmdClear_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = 2 Then
         PopupMenu hidMen
    End If

End Sub

Private Sub cmdDecimal_Click()

    If decIn = False Then

        If nxt = True Then
            txtCalc.Text = "0."
            addScrap ("0.")
        Else
            txtCalc.Text = txtCalc.Text & "."
            addScrap (".")
        End If
            
        currTxt = CDbl(txtCalc.Text)
        decIn = True
        nxt = False
        
    End If

    pic1.SetFocus

End Sub

Private Sub cmdDecimal_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = 2 Then
         PopupMenu hidMen
    End If


End Sub

Private Sub cmdDiv_Click()

    On Error GoTo erHand

    If fstTime = False Then
    
        Select Case txtSign.Text
        Case "-"
            currTtl = currTtl - currTxt
            txtCalc.Text = currTtl
        Case "+"
            currTtl = currTtl + currTxt
            txtCalc.Text = currTtl
        Case "/"
            currTtl = currTtl / currTxt
            txtCalc.Text = currTtl
        Case "*"
            currTtl = currTtl * currTxt
            txtCalc.Text = currTtl
        End Select
    
    Else
        currTtl = currTxt
    End If
    
    txtSign.Text = "/"
    decIn = False
    nxt = True
    fstTime = False
    
    pic1.SetFocus
    
    addScrap (retChr & txtSign.Text & " ")
    
erHand:
    If Err.Number = 0 Or Err.Number = 20 Then
        Resume Next
    Else

        MsgBox (Err.Description), , "Calculator"
        txtCalc.Text = "Error"
        txtSign.Text = "?"
        currTtl = 0
        currTxt = 0
        
        pic1.SetFocus
        Exit Sub
        
    End If

End Sub

Private Sub cmdDiv_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = 2 Then
         PopupMenu hidMen
    End If


End Sub

Private Sub cmdEqual_Click()

    On Error GoTo erHand

    If fstTime = False Then
    
        Select Case txtSign.Text
        Case "-"
            currTtl = currTtl - currTxt
            txtCalc.Text = currTtl
        Case "+"
            currTtl = currTtl + currTxt
            txtCalc.Text = currTtl
        Case "/"
            currTtl = currTtl / currTxt
            txtCalc.Text = currTtl
        Case "*"
            currTtl = currTtl * currTxt
            txtCalc.Text = currTtl
        End Select
    
    Else
        currTtl = currTxt
        txtCalc.Text = currTtl
    End If
    
    txtSign.Text = "="
    decIn = False
    nxt = True
    fstTime = False
    
    pic1.SetFocus
    
    addScrap (retChr & txtSign.Text & txtCalc.Text)

erHand:
    If Err.Number = 0 Or Err.Number = 20 Then
        Resume Next
    Else

        MsgBox (Err.Description), , "Calculator"
        txtCalc.Text = "Error"
        txtSign.Text = "?"
        currTtl = 0
        currTxt = 0
        
        pic1.SetFocus
        
        Exit Sub
        
    End If


End Sub

Private Sub cmdEqual_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = 2 Then
         PopupMenu hidMen
    End If


End Sub

Private Sub cmdMinu_Click()
    
    On Error GoTo erHand
    
    If fstTime = False Then
    
        Select Case txtSign.Text
        Case "-"
            currTtl = currTtl - currTxt
            txtCalc.Text = currTtl
        Case "+"
            currTtl = currTtl + currTxt
            txtCalc.Text = currTtl
        Case "/"
            currTtl = currTtl / currTxt
            txtCalc.Text = currTtl
        Case "*"
            currTtl = currTtl * currTxt
            txtCalc.Text = currTtl
        End Select
    
    Else
        currTtl = currTxt
    End If
    
    fstTime = False
    
    txtSign.Text = "-"
    decIn = False
    nxt = True

    pic1.SetFocus
    
    addScrap (retChr & txtSign.Text & " ")

erHand:
    If Err.Number = 0 Or Err.Number = 20 Then
        Resume Next
    Else

        MsgBox (Err.Description), , "Calculator"
        txtCalc.Text = "Error"
        txtSign.Text = "?"
        currTtl = 0
        currTxt = 0
        
        pic1.SetFocus
        
        Exit Sub
        
    End If



End Sub

Private Sub cmdMinu_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = 2 Then
         PopupMenu hidMen
    End If


End Sub

Private Sub cmdMult_Click()

    On Error GoTo erHand
    
    If fstTime = False Then
    
        Select Case txtSign.Text
        Case "-"
            currTtl = currTtl - currTxt
            txtCalc.Text = currTtl
        Case "+"
            currTtl = currTtl + currTxt
            txtCalc.Text = currTtl
        Case "/"
            currTtl = currTtl / currTxt
            txtCalc.Text = currTtl
        Case "*"
            currTtl = currTtl * currTxt
            txtCalc.Text = currTtl
        End Select
    
    Else
        currTtl = currTxt
    End If
    
    fstTime = False
    
    txtSign.Text = "*"
    decIn = False
    nxt = True
    
    pic1.SetFocus
    
    addScrap (retChr & txtSign.Text & " ")
    
erHand:
    If Err.Number = 0 Or Err.Number = 20 Then
        Resume Next
    Else

        MsgBox (Err.Description), , "Calculator"
        txtCalc.Text = "Error"
        txtSign.Text = "?"
        currTtl = 0
        currTxt = 0
        
        pic1.SetFocus
        
        Exit Sub
        
    End If


End Sub

Private Sub cmdMult_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = 2 Then
         PopupMenu hidMen
    End If


End Sub

Private Sub cmdNegative_Click()

    If nxt = True Then
        txtCalc.Text = "-"
    End If
            
    nxt = False

    If negPos = False Then
        negPos = True
        currTxt = currTxt * -1
        txtCalc.Text = currTxt
    Else
        negPos = False
        currTxt = currTxt * -1
        txtCalc.Text = currTxt
    End If

    pic1.SetFocus

End Sub

Private Sub cmdNegative_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = 2 Then
         PopupMenu hidMen
    End If


End Sub

Private Sub cmdPlus_Click()
    
    On Error GoTo erHand
     
    If fstTime = False Then
    
        Select Case txtSign.Text
        Case "-"
            currTtl = currTtl - currTxt
            txtCalc.Text = currTtl
        Case "+"
            currTtl = currTtl + currTxt
            txtCalc.Text = currTtl
        Case "/"
            currTtl = currTtl / currTxt
            txtCalc.Text = currTtl
        Case "*"
            currTtl = currTtl * currTxt
            txtCalc.Text = currTtl
        End Select
    
    Else
        currTtl = currTxt
    End If
    
    fstTime = False
    
    txtSign.Text = "+"
    decIn = False
    nxt = True
    
    pic1.SetFocus
    
    addScrap (retChr & txtSign.Text & " ")
    
erHand:
    If Err.Number = 0 Or Err.Number = 20 Then
        Resume Next
    Else

        MsgBox (Err.Description), , "Calculator"
        txtCalc.Text = "Error"
        txtSign.Text = "?"
        currTtl = 0
        currTxt = 0
        
        pic1.SetFocus
        
        Exit Sub
        
    End If

End Sub

Private Sub cmdPlus_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = 2 Then
         PopupMenu hidMen
    End If


End Sub

Private Sub Command1_Click(Index As Integer)

    On Error Resume Next

    If txtCalc.Text <> "0" And nxt = False Then
        txtCalc.Text = txtCalc.Text & Index
        currTxt = CDbl(txtCalc.Text)
        nxt = False
    ElseIf nxt = True Then
        txtCalc.Text = Index
        currTxt = CDbl(txtCalc.Text)
        nxt = False
    Else
        txtCalc.Text = Index
        currTxt = CDbl(txtCalc.Text)
    End If

    pic1.SetFocus
    
    addScrap (Index)

End Sub


Private Sub Command1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)

    subKeyDown (KeyCode)
    pic1.SetFocus

End Sub

Private Sub Command1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = 2 Then
         PopupMenu hidMen
    End If

End Sub

Private Sub Form_Load()

    txtCalc.Text = 0
    txtSign.Text = ""
    decIn = False
    nxt = False
    currTtl = 0
    fstTime = True
    negPos = False
    
    Me.Left = Screen.Width / 2 - (Me.Width / 2)
    Me.Top = Screen.Height / 2 - (Me.Height / 2)
    
    Me.Show
    
    retChr = Chr(10) & Chr(13)
    
    pic1.SetFocus

End Sub

Sub subKeyDown(cK)

    If Chr(cK) = 1 Or cK = 97 Then
        Command1_Click (1)
    ElseIf Chr(cK) = 2 Or cK = 98 Then
        Command1_Click (2)
    ElseIf Chr(cK) = 3 Or cK = 99 Then
        Command1_Click (3)
    ElseIf Chr(cK) = 4 Or cK = 100 Then
        Command1_Click (4)
    ElseIf Chr(cK) = 5 Or cK = 101 Then
        Command1_Click (5)
    ElseIf Chr(cK) = 6 Or cK = 102 Then
        Command1_Click (6)
    ElseIf Chr(cK) = 7 Or cK = 103 Then
        Command1_Click (7)
    ElseIf Chr(cK) = 8 Or cK = 104 Then
        Command1_Click (8)
    ElseIf Chr(cK) = 9 Or cK = 105 Then
        Command1_Click (9)
    ElseIf Chr(cK) = 0 Or cK = 96 Then
        Command1_Click (0)
    ElseIf cK = 110 Or cK = 190 Then
        cmdDecimal_Click
    ElseIf cK = 187 Then
        cmdEqual_Click
    ElseIf cK = 107 Then
        cmdPlus_Click
    ElseIf cK = 109 Then
        cmdMinu_Click
    ElseIf cK = 106 Then
        cmdMult_Click
    ElseIf cK = 111 Or cK = 191 Then
        cmdDiv_Click
    ElseIf cK = 13 Then
        cmdEqual_Click
    ElseIf cK = 8 Then
        
        If Len(txtCalc.Text) <> 0 Then
            ct = txtCalc.Text
            xx = Mid(ct, 1, (Len(txtCalc.Text) - 1))
            txtCalc.Text = xx
            If txtCalc.Text = "" Or txtCalc.Text = "." Then
                cmdClear_Click
            End If
        End If
        
    End If

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = 2 Then
         PopupMenu hidMen
    End If

End Sub

Private Sub Frame1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = 2 Then
         PopupMenu hidMen
    End If


End Sub

Private Sub mnuCopyItem_Click()

    Clipboard.Clear
    Clipboard.SetText txtCalc.Text

End Sub

Private Sub mnuEItem_Click()

    Unload Me

End Sub

Private Sub mnuInItem_Click()

    frmTextPad.rtbox1.SelText = txtCalc.Text

End Sub

Private Sub mnuMinimizeItem_Click()

    Me.WindowState = vbMinimized

End Sub

Private Sub mnuScrapItem_Click()

    MsgBox (scrapPage)

End Sub

Private Sub pic1_KeyDown(KeyCode As Integer, Shift As Integer)

    subKeyDown (KeyCode)

End Sub

Private Sub pic1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = 2 Then
         PopupMenu hidMen
    End If


End Sub

Private Sub Picture2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = 2 Then
         PopupMenu hidMen
    End If


End Sub

Private Sub txtCalc_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = 2 Then
         PopupMenu hidMen
    End If

End Sub

Sub addScrap(D)

    scrapPage = scrapPage & D

End Sub
