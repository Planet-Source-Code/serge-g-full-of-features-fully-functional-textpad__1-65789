VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form frmTextPad_FontNames 
   BackColor       =   &H8000000A&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Select Font"
   ClientHeight    =   5520
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5310
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5520
   ScaleWidth      =   5310
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      Height          =   3375
      Left            =   360
      TabIndex        =   3
      Top             =   360
      Width           =   4575
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Select and Exit"
      Height          =   375
      Left            =   480
      TabIndex        =   2
      Top             =   4920
      Width           =   2055
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel without change"
      Height          =   375
      Left            =   2760
      TabIndex        =   0
      Top             =   4920
      Width           =   2055
   End
   Begin RichTextLib.RichTextBox rtbPreview 
      Height          =   495
      Left            =   360
      TabIndex        =   1
      Top             =   4200
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   873
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"frmTextPad_FontNames.frx":0000
   End
   Begin VB.Frame Frame1 
      Caption         =   "Font"
      Height          =   3735
      Left            =   240
      TabIndex        =   4
      Top             =   120
      Width           =   4815
   End
   Begin VB.Frame Frame2 
      Caption         =   "Preview"
      Height          =   855
      Left            =   240
      TabIndex        =   5
      Top             =   3960
      Width           =   4815
   End
End
Attribute VB_Name = "frmTextPad_FontNames"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim theX As Long
Dim theY As Long

Private Sub cmdCancel_Click()

    Unload Me

End Sub

Private Sub cmdOK_Click()

    defaultFont = List1.Text
    frmOptions.cmdApply.Enabled = True
    Unload Me

End Sub

Private Sub Form_Load()

''    Me.BackColor = &HFF0000
''
''    setTrans

    For i = 0 To Screen.FontCount - 1
        List1.AddItem Screen.Fonts(i)
    Next i

End Sub

Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    theX = X
    theY = Y

End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

      If Button = 1 Then
        Me.Left = Me.Left + (X - theX)
        Me.Top = Me.Top + (Y - theY)
      End If


End Sub

Private Sub List1_Click()

    rtbPreview.Text = List1.Text
    rtbPreview.Font.Name = List1.Text

End Sub
