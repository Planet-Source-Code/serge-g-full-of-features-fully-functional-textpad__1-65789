VERSION 5.00
Begin VB.Form frmCWordTxt 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4605
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4605
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2760
      TabIndex        =   2
      Top             =   2640
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Ok"
      Height          =   375
      Left            =   720
      TabIndex        =   1
      Top             =   2640
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   2175
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   240
      Width           =   4335
   End
End
Attribute VB_Name = "frmCWordTxt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()

    Unload Me

End Sub

Private Sub cmdOK_Click()

    Select Case Me.Caption
    Case "Word 1"
        frmOptions.Text1.Text = Text1.Text
    Case "Word 2"
        frmOptions.Text2.Text = Text1.Text
    Case "Word 3"
        frmOptions.Text3.Text = Text1.Text
    Case "Word 4"
        frmOptions.Text4.Text = Text1.Text
    Case "Word 5"
        frmOptions.Text5.Text = Text1.Text
    End Select
    
    Unload Me

End Sub
