VERSION 5.00
Begin VB.Form frmMsgBoxYNC 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   1620
   ClientLeft      =   4185
   ClientTop       =   2955
   ClientWidth     =   4005
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1620
   ScaleWidth      =   4005
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2625
      TabIndex        =   0
      Top             =   1065
      Width           =   1095
   End
   Begin VB.CommandButton cmdNo 
      Caption         =   "Don't Save"
      Height          =   375
      Left            =   1425
      TabIndex        =   2
      Top             =   1065
      Width           =   1095
   End
   Begin VB.CommandButton cmdYes 
      Caption         =   "Save"
      Height          =   375
      Left            =   225
      TabIndex        =   1
      Top             =   1065
      Width           =   1095
   End
   Begin VB.Label lblBottom 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   705
      TabIndex        =   4
      Top             =   585
      Width           =   2775
   End
   Begin VB.Label lblTop 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   705
      TabIndex        =   3
      Top             =   105
      Width           =   2775
      WordWrap        =   -1  'True
   End
   Begin VB.Image Image1 
      Height          =   465
      Left            =   105
      Picture         =   "msgBox_Y_N_Cnl_TextPad.frx":0000
      Stretch         =   -1  'True
      Top             =   225
      Width           =   480
   End
End
Attribute VB_Name = "frmMsgBoxYNC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()

    quitConfirm = vbCancel
    Unload Me

End Sub

Private Sub cmdNo_Click()

    quitConfirm = vbNo
    Unload Me

End Sub

Private Sub cmdYes_Click()

    quitConfirm = vbYes
    Unload Me

End Sub

Private Sub Form_Load()

    Beep
    Me.Icon = LoadPicture("")
    Me.Top = Screen.Height / 2 - (Me.Height / 2)
    Me.Left = Screen.Width / 2 - (Me.Width / 2)
        
End Sub

