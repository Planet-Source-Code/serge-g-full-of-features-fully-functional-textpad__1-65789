VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About"
   ClientHeight    =   3180
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4605
   Icon            =   "frmTxtPadAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3180
   ScaleWidth      =   4605
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   1785
      TabIndex        =   0
      Top             =   2670
      Width           =   1095
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      X1              =   120
      X2              =   4440
      Y1              =   2505
      Y2              =   2505
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000009&
      BorderWidth     =   2
      X1              =   120
      X2              =   4440
      Y1              =   2520
      Y2              =   2520
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmTxtPadAbout.frx":0442
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   240
      TabIndex        =   1
      Top             =   360
      Width           =   4215
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00C0C0C0&
      FillStyle       =   6  'Cross
      Height          =   2175
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   120
      Width           =   4335
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdOK_Click()

    Unload Me

End Sub

Private Sub Form_Load()

    Me.Left = Screen.Width / 2 - (Me.Width / 2)
    Me.Top = Screen.Height / 2 - (Me.Height / 2)

End Sub

