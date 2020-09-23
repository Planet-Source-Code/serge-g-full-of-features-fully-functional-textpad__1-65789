VERSION 5.00
Begin VB.Form frmHelp 
   Caption         =   "General Help"
   ClientHeight    =   6255
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   7800
   Icon            =   "frmTxtPadHelp.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6255
   ScaleWidth      =   7800
   StartUpPosition =   3  'Windows Default
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   4560
      TabIndex        =   11
      ToolTipText     =   "Line number of cursor's position"
      Top             =   4680
      Width           =   2535
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   3600
      TabIndex        =   10
      ToolTipText     =   "Save status"
      Top             =   4680
      Width           =   975
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   2760
      TabIndex        =   9
      ToolTipText     =   "See the status of NUMS and CAPS"
      Top             =   4680
      Width           =   855
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   1440
      TabIndex        =   8
      ToolTipText     =   "Time and Date. Click to hide"
      Top             =   4680
      Width           =   1335
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   480
      TabIndex        =   7
      ToolTipText     =   "Change between being on top or being dockable"
      Top             =   4680
      Width           =   975
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Always on top/Dockable"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      TabIndex        =   6
      Top             =   3480
      Width           =   2295
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Cursor Position"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5760
      TabIndex        =   5
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "File Status"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2640
      TabIndex        =   4
      Top             =   5520
      Width           =   1215
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "Time and Date"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   960
      TabIndex        =   3
      Top             =   5640
      Width           =   1455
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "CAPS and NUMS Lock"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4800
      TabIndex        =   2
      Top             =   1560
      Width           =   1815
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Status Bar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2040
      TabIndex        =   1
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Menu"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   615
      TabIndex        =   0
      Top             =   1470
      Width           =   720
   End
   Begin VB.Line Line27 
      X1              =   2130
      X2              =   2280
      Y1              =   4980
      Y2              =   4950
   End
   Begin VB.Line Line26 
      X1              =   2280
      X2              =   2310
      Y1              =   5115
      Y2              =   4950
   End
   Begin VB.Line Line25 
      X1              =   2310
      X2              =   1845
      Y1              =   4935
      Y2              =   5520
   End
   Begin VB.Line Line24 
      X1              =   1350
      X2              =   1155
      Y1              =   3105
      Y2              =   3120
   End
   Begin VB.Line Line23 
      X1              =   1215
      X2              =   1170
      Y1              =   2925
      Y2              =   3120
   End
   Begin VB.Line Line22 
      X1              =   1185
      X2              =   2415
      Y1              =   3120
      Y2              =   2085
   End
   Begin VB.Line Line21 
      X1              =   1125
      X2              =   975
      Y1              =   765
      Y2              =   1395
   End
   Begin VB.Line Line20 
      X1              =   795
      X2              =   975
      Y1              =   750
      Y2              =   1425
   End
   Begin VB.Line Line19 
      X1              =   1695
      X2              =   1830
      Y1              =   4980
      Y2              =   5160
   End
   Begin VB.Line Line18 
      X1              =   1590
      X2              =   1695
      Y1              =   5160
      Y2              =   4980
   End
   Begin VB.Line Line17 
      X1              =   3720
      X2              =   3960
      Y1              =   5025
      Y2              =   4995
   End
   Begin VB.Line Line16 
      X1              =   3870
      X2              =   3915
      Y1              =   5205
      Y2              =   4995
   End
   Begin VB.Line Line15 
      X1              =   3945
      X2              =   3270
      Y1              =   4980
      Y2              =   5430
   End
   Begin VB.Line Line14 
      X1              =   1695
      X2              =   1695
      Y1              =   4980
      Y2              =   5580
   End
   Begin VB.Line Line13 
      X1              =   3450
      X2              =   3435
      Y1              =   4365
      Y2              =   4530
   End
   Begin VB.Line Line12 
      X1              =   3615
      X2              =   3435
      Y1              =   4500
      Y2              =   4545
   End
   Begin VB.Line Line11 
      X1              =   3195
      X2              =   2985
      Y1              =   4515
      Y2              =   4560
   End
   Begin VB.Line Line10 
      X1              =   3015
      X2              =   3000
      Y1              =   4395
      Y2              =   4560
   End
   Begin VB.Line Line9 
      X1              =   3420
      X2              =   5535
      Y1              =   4560
      Y2              =   2220
   End
   Begin VB.Line Line8 
      X1              =   5550
      X2              =   3015
      Y1              =   2205
      Y2              =   4560
   End
   Begin VB.Line Line7 
      X1              =   5175
      X2              =   4920
      Y1              =   4500
      Y2              =   4560
   End
   Begin VB.Line Line6 
      X1              =   4920
      X2              =   4920
      Y1              =   4320
      Y2              =   4560
   End
   Begin VB.Line Line5 
      X1              =   4920
      X2              =   6240
      Y1              =   4560
      Y2              =   3120
   End
   Begin VB.Line Line4 
      X1              =   1320
      X2              =   1080
      Y1              =   4440
      Y2              =   4560
   End
   Begin VB.Line Line3 
      X1              =   1080
      X2              =   1080
      Y1              =   4320
      Y2              =   4560
   End
   Begin VB.Line Line2 
      X1              =   1080
      X2              =   1560
      Y1              =   4560
      Y2              =   3840
   End
   Begin VB.Line Line1 
      X1              =   600
      X2              =   960
      Y1              =   720
      Y2              =   1440
   End
   Begin VB.Image Image2 
      Height          =   225
      Left            =   480
      Picture         =   "frmTxtPadHelp.frx":08CA
      Top             =   4680
      Width           =   6750
   End
   Begin VB.Image Image1 
      Height          =   2895
      Left            =   480
      Picture         =   "frmTxtPadHelp.frx":30F9
      Stretch         =   -1  'True
      ToolTipText     =   "TextPad Window"
      Top             =   480
      Width           =   4455
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuExitItem 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuHlpTops 
      Caption         =   "&Help Topics"
      Visible         =   0   'False
      Begin VB.Menu mnuT1Item 
         Caption         =   "Topic 1"
      End
      Begin VB.Menu mnuT2Item 
         Caption         =   "Topic 2"
      End
      Begin VB.Menu mnuT3Item 
         Caption         =   "Topic 3"
      End
      Begin VB.Menu mnuT4Item 
         Caption         =   "Topic 4"
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuExitItem2 
         Caption         =   "   Exit"
         Visible         =   0   'False
      End
   End
End
Attribute VB_Name = "frmHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim tempStat As Boolean

Private Sub Form_Load()

    tempStat = TopMost
    TopMost = False
    SetTopMost

    Me.Left = Screen.Width / 2 - (Me.Width / 2)
    Me.Top = Screen.Height / 2 - (Me.Height / 2)
    
    mnuExitItem2.Visible = False
    mnuSep1.Visible = False

End Sub


Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = 2 Then
        mnuExitItem2.Visible = True
        mnuSep1.Visible = True
        PopupMenu mnuHlpTops
    Else
        mnuExitItem2.Visible = False
        mnuSep1.Visible = False
    End If

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    mnuExitItem2.Visible = False
    mnuSep1.Visible = False

End Sub

Private Sub Form_Unload(Cancel As Integer)

    TopMost = tempStat
    SetTopMost

End Sub

Private Sub Label10_Click()

    MsgBox ("Shows the status of Number Lock and Caps Lock"), , "Text Pad"

End Sub

Private Sub Label11_Click()

    MsgBox ("Shows the status of the current file Saved / Unsaved"), , "Text Pad"

End Sub

Private Sub Label12_Click()

    MsgBox ("Shows the cursor's line position"), , "Text Pad"

End Sub

Private Sub Label8_Click()

    MsgBox ("Click to always be on top"), , "Text Pad"

End Sub

Private Sub Label9_Click()

    MsgBox ("You can set an alarm by clicking time, and see the calendar by clicking the date"), , "Text Pad"

End Sub

Private Sub mnuExitItem_Click()

    Unload Me

End Sub

Private Sub mnuExitItem2_Click()

    Unload Me

End Sub
