VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmFontSize 
   BorderStyle     =   0  'None
   Caption         =   "Change Font Size"
   ClientHeight    =   4230
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5310
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4230
   ScaleWidth      =   5310
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3000
      TabIndex        =   4
      Top             =   3600
      Width           =   975
   End
   Begin VB.TextBox txtFS 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   2280
      TabIndex        =   3
      Top             =   840
      Width           =   735
   End
   Begin MSComctlLib.Slider sld1 
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   240
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   661
      _Version        =   393216
      BorderStyle     =   1
      Min             =   7
      Max             =   50
      SelStart        =   7
      Value           =   7
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   390
      Left            =   1320
      TabIndex        =   0
      Top             =   3600
      Width           =   960
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   5160
      Y1              =   3480
      Y2              =   3480
   End
   Begin VB.Label lblFontSize 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Sample Text"
      Height          =   1695
      Left            =   240
      TabIndex        =   1
      Top             =   1560
      Width           =   4815
   End
   Begin VB.Shape Shape1 
      Height          =   4215
      Left            =   0
      Top             =   0
      Width           =   5295
   End
End
Attribute VB_Name = "frmFontSize"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim tempF

Private Sub cmdCancel_Click()

    currFontSize = tempF
    Unload Me

End Sub

Private Sub cmdCancel_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 27 Then
        cmdCancel_Click
    ElseIf KeyCode = 13 Then
        cmdOK_Click
    End If

End Sub

Private Sub cmdOK_Click()

    currFontSize = sld1.Value
    Unload Me

End Sub

Private Sub cmdOK_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 27 Then
        cmdCancel_Click
    ElseIf KeyCode = 13 Then
        cmdOK_Click
    End If

End Sub

Private Sub Form_Load()

    Me.Top = Screen.Height / 2 - (Me.Height / 2)
    Me.Left = Screen.Width / 2 - (Me.Width / 2)
    
    tempF = currFontSize
    
    sld1.Value = CInt(currFontSize)
    lblFontSize.FontName = currFontName
    lblFontSize.Caption = currFontName
    txtFS.Text = Int(currFontSize)

End Sub

Private Sub sld1_Change()

    txtFS.Text = sld1.Value
    'If sld1.Value < 35 Then
        lblFontSize.FontSize = sld1.Value
    'End If

End Sub

Private Sub sld1_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 27 Then
        cmdCancel_Click
    ElseIf KeyCode = 13 Then
        cmdOK_Click
    End If

End Sub

Private Sub txtFS_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 27 Then
        cmdCancel_Click
    ElseIf KeyCode = 13 Then
        cmdOK_Click
    End If

End Sub
