VERSION 5.00
Begin VB.Form frmQHelp 
   BackColor       =   &H00808080&
   Caption         =   "Quick Help"
   ClientHeight    =   4935
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8055
   Icon            =   "frmTxtPadQHelp.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4935
   ScaleWidth      =   8055
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   360
      Picture         =   "frmTxtPadQHelp.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "View help in big font"
      Top             =   4020
      Width           =   465
   End
   Begin VB.TextBox txtHelp 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      Height          =   3015
      Left            =   4680
      MultiLine       =   -1  'True
      TabIndex        =   5
      Top             =   825
      Width           =   2055
   End
   Begin VB.CommandButton Command4 
      Appearance      =   0  'Flat
      Height          =   495
      Left            =   6960
      Picture         =   "frmTxtPadQHelp.frx":0BD4
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Scroll up through help"
      Top             =   1560
      Width           =   495
   End
   Begin VB.CommandButton Command3 
      Appearance      =   0  'Flat
      Height          =   495
      Left            =   6960
      Picture         =   "frmTxtPadQHelp.frx":1016
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Scroll down through help"
      Top             =   2520
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      Height          =   495
      Left            =   6960
      Picture         =   "frmTxtPadQHelp.frx":1458
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Back to the TextPad"
      Top             =   3480
      Width           =   495
   End
   Begin VB.ListBox lstSC 
      Appearance      =   0  'Flat
      Height          =   3150
      Left            =   240
      TabIndex        =   0
      Top             =   720
      Width           =   3975
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Scroll"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   6840
      TabIndex        =   9
      Top             =   1320
      Width           =   735
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Caption         =   "Bold"
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
      Left            =   330
      TabIndex        =   11
      Top             =   4515
      Width           =   585
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Caption         =   "Exit"
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
      Left            =   6960
      TabIndex        =   10
      Top             =   4080
      Width           =   495
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Quick Help"
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
      Left            =   4680
      TabIndex        =   7
      Top             =   600
      Width           =   2055
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Keys to press"
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
      TabIndex        =   6
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Shortcut"
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
      Left            =   240
      TabIndex        =   1
      Top             =   480
      Width           =   735
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Height          =   4695
      Left            =   120
      Top             =   120
      Width           =   7815
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0C0C0&
      BorderColor     =   &H80000005&
      BorderWidth     =   3
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   4695
      Left            =   120
      Top             =   120
      Width           =   7815
   End
End
Attribute VB_Name = "frmQHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
        (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam _
        As Long, lParam As Any) As Long

Const LB_SETTABSTOPS = &H192

Dim myInt As Integer
Dim tempStat As Boolean

Private Sub Check1_Click()

    lstSC.FontBold = Not lstSC.FontBold
    txtHelp.FontBold = Not txtHelp.FontBold

End Sub

Private Sub Command1_Click()

    Unload Me

End Sub

Private Sub Command2_Click()

    ShowHelp (lstSC.ListIndex)

End Sub

Private Sub Command3_Click()

    If myInt > lstSC.ListCount - 3 Then
        myInt = myInt - 1
    End If
    myInt = myInt + 1
    ShowHelp (myInt)

End Sub

Private Sub Command4_Click()

    If myInt < 1 Then
        myInt = 1
    End If
    myInt = myInt - 1
    ShowHelp (myInt)


End Sub

Private Sub Form_Load()

    tempStat = TopMost
    TopMost = False
    SetTopMost
    
    Dim lngRet As Long
    Dim aTabs As Long
   
    aTabs = 80

    lngRet = SendMessage(lstSC.hwnd, _
                        LB_SETTABSTOPS, _
                        2, _
                        aTabs)

    lstSC.AddItem "Special Copy" & vbTab & "Alt + Ctrl + C"
    lstSC.AddItem "Special Cut" & vbTab & "Alt + Ctrl + X"
    lstSC.AddItem "Private Copy" & vbTab & "Shft + Ctrl + C"
    lstSC.AddItem "Private Cut" & vbTab & "Shft + Ctrl + X"
    lstSC.AddItem "Private Paste" & vbTab & "Shft + Ctrl + V"
    lstSC.AddItem "Search / Replace" & vbTab & "Ctrl + F | F2"
    lstSC.AddItem "Undo" & vbTab & "Ctrl + Z"
    lstSC.AddItem "Redo" & vbTab & "Ctrl + Y"
    lstSC.AddItem "Shortcut Help" & vbTab & "Ctrl + Alt + H or F1"
    lstSC.AddItem "Save" & vbTab & "Ctrl + S"
    lstSC.AddItem "Insert Time" & vbTab & "Alt + Ctrl + T"
    lstSC.AddItem "Insert Date" & vbTab & "Alt + Ctrl + D"
    lstSC.AddItem "Bold" & vbTab & "Alt + Shft +B"
    lstSC.AddItem "Italic" & vbTab & "Alt + Shft +I"
    lstSC.AddItem "Underline" & vbTab & "Alt + Shft +U"
    lstSC.AddItem "Strikethrough" & vbTab & "Alt + Shft + S"
    
    Me.Left = Screen.Width / 2 - (Me.Width / 2)
    Me.Top = Screen.Height / 2 - (Me.Height / 2)
        
End Sub


Private Sub lstKey_Click()

    ShowHelp (lstKey.ListIndex)
    
End Sub

Private Sub lstKey_GotFocus()

    txtHelp.SetFocus

End Sub

Private Sub Form_Unload(Cancel As Integer)

    TopMost = tempStat
    SetTopMost

End Sub

Private Sub lstSC_Click()

    ShowHelp (lstSC.ListIndex)

End Sub

Sub ShowHelp(indx)

    Dim txt
    
    myInt = indx
    
    Select Case indx
    Case 0:
        txt = "Special Copy will append the current text to previously copied text using Special Copy or Special Cut methods"
    Case 1:
        txt = "Special Cut will append the current text to previously cut text using Special Copy or Special Cut methods"
    Case 2:
        txt = "Private Copy is used to copy and paste text without using the system's clipboard. Useful if you are holding information on the clipboard that you might want to use later"
    Case 3:
        txt = "Private Cut is  used to cut and paste text without using the system's clipboard. Useful if you are holding information on the clipboard that you might want to use later"
    Case 4:
        txt = "Private Paste is used to paste private copied / cut text"
    Case 4:
        txt = "Search or replace words in the document"
    Case 5:
        txt = "Undo, un-does your last operation, such as deletion by mistake"
    Case 6:
        txt = "Redo un-does what undo did. If undo needs to be 'un-used'"""
    Case 7:
        txt = "Help screen for the TextPad"
    Case 8:
        txt = "Saves your current file"
    Case 9:
        txt = "Inserts a time stamp into your text"
    Case 10:
        txt = "Inserts a date stamp into your text"
    Case 11:
        txt = "If any text is selected, it will be in bold font"
    Case 12:
        txt = "If any text is selected, it will be in italic font"
    Case 13:
        txt = "If any text is selected, it will be in underlined"
    Case 14:
        txt = "If any text is selected, it will have a line striking through"
    End Select
    
    txtHelp.Text = txt

End Sub

