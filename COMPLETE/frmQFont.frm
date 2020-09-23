VERSION 5.00
Begin VB.Form frmQFont 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2445
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4155
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2445
   ScaleWidth      =   4155
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Height          =   2430
      Left            =   0
      ScaleHeight     =   2370
      ScaleWidth      =   4080
      TabIndex        =   1
      Top             =   0
      Width           =   4140
      Begin VB.ListBox List1 
         Height          =   2400
         Left            =   -15
         TabIndex        =   0
         Top             =   -15
         Width           =   4110
      End
   End
End
Attribute VB_Name = "frmQFont"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim tempFont As String
Dim t As Boolean

Private Sub Form_Load()

    t = TopMost
    TopMost = False
    SetTopMost

    isQFActive = True

    Me.Top = frmTextPad.Top + (frmTextPad.Height / 2) - (Me.Height / 2)
    Me.Left = frmTextPad.Left + (frmTextPad.Width / 2) - (Me.Width / 2)

    If frmTextPad.rtbox1.SelLength <> 0 Then
        tempFont = frmTextPad.rtbox1.SelFontName
    Else
        tempFont = frmTextPad.rtbox1.Font.Name
    End If
    
    For i = 0 To Screen.FontCount - 1
        List1.AddItem Screen.Fonts(i)
    Next i

End Sub

Private Sub Form_Unload(Cancel As Integer)

    isQFActive = False
    
    TopMost = t
    SetTopMost

End Sub

Private Sub List1_Click()

    frmTextPad.rtbox1.SelFontName = List1.Text

End Sub

Private Sub List1_DblClick()

    Unload Me

End Sub

Private Sub List1_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 27 Then     'Hit Escape
        
        If frmTextPad.rtbox1.SelLength <> 0 Then
            frmTextPad.rtbox1.SelFontName = tempFont
        Else
            frmTextPad.rtbox1.Font.Name = tempFont
        End If
        
        Unload Me
    End If
    
    If KeyCode = 13 Then     'Hit Enter
        Unload Me
    End If

End Sub

