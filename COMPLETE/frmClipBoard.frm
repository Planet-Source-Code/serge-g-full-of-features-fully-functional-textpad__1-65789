VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmClipBoard 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Clipboard View"
   ClientHeight    =   3135
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4710
   Icon            =   "frmClipBoard.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3135
   ScaleWidth      =   4710
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab SSTab1 
      Height          =   3150
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4725
      _ExtentX        =   8334
      _ExtentY        =   5556
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "System Board"
      TabPicture(0)   =   "frmClipBoard.frx":014A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "rtbSysBrd"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cmdEnd"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmdClear"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cmdSave"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cmdPasteIt"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      TabCaption(1)   =   "Private Board"
      TabPicture(1)   =   "frmClipBoard.frx":0166
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmdPastePrvt"
      Tab(1).Control(1)=   "cmdSave2"
      Tab(1).Control(2)=   "cmdClear2"
      Tab(1).Control(3)=   "cmdOK2"
      Tab(1).Control(4)=   "rtbPrvt"
      Tab(1).ControlCount=   5
      Begin VB.CommandButton cmdPastePrvt 
         Caption         =   "Paste"
         Height          =   375
         Left            =   -73920
         TabIndex        =   10
         Top             =   2400
         Width           =   735
      End
      Begin VB.CommandButton cmdPasteIt 
         Caption         =   "Paste"
         Height          =   375
         Left            =   1080
         TabIndex        =   9
         Top             =   2400
         Width           =   735
      End
      Begin VB.CommandButton cmdSave2 
         Caption         =   "Save Changes"
         Height          =   375
         Left            =   -71760
         TabIndex        =   8
         Top             =   2400
         Width           =   1215
      End
      Begin VB.CommandButton cmdClear2 
         Caption         =   "Clear Private"
         Height          =   375
         Left            =   -73080
         TabIndex        =   7
         Top             =   2400
         Width           =   1215
      End
      Begin VB.CommandButton cmdOK2 
         Caption         =   "Exit"
         Height          =   375
         Left            =   -74760
         TabIndex        =   6
         Top             =   2400
         Width           =   735
      End
      Begin RichTextLib.RichTextBox rtbPrvt 
         Height          =   1455
         Left            =   -74760
         TabIndex        =   5
         Top             =   720
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   2566
         _Version        =   393217
         TextRTF         =   $"frmClipBoard.frx":0182
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "Save Changes"
         Height          =   375
         Left            =   3240
         TabIndex        =   4
         Top             =   2400
         Width           =   1230
      End
      Begin VB.CommandButton cmdClear 
         Caption         =   "Clear Clipboard"
         Height          =   375
         Left            =   1920
         TabIndex        =   3
         Top             =   2400
         Width           =   1215
      End
      Begin VB.CommandButton cmdEnd 
         Caption         =   "Exit"
         Height          =   375
         Left            =   240
         TabIndex        =   2
         Top             =   2400
         Width           =   735
      End
      Begin RichTextLib.RichTextBox rtbSysBrd 
         Height          =   1455
         Left            =   240
         TabIndex        =   1
         Top             =   720
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   2566
         _Version        =   393217
         TextRTF         =   $"frmClipBoard.frx":0204
      End
   End
End
Attribute VB_Name = "frmClipBoard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim prvtChars, sysChars

Private Sub cmdClear_Click()

    Clipboard.Clear
    rtbSysBrd.Text = ""
    cmdClear.Enabled = False
    cmdSave.Enabled = False

End Sub

Private Sub cmdClear2_Click()

    publicCopy = ""
    rtbPrvt.Text = ""
    cmdClear2.Enabled = False
    cmdSave2.Enabled = False

End Sub

Private Sub cmdEnd_Click()

    Unload Me

End Sub

Private Sub cmdOK2_Click()

    Unload Me

End Sub

Private Sub cmdPasteIt_Click()

    frmTextPad.rtbox1.SelText = rtbSysBrd.Text

End Sub

Private Sub cmdPastePrvt_Click()

    frmTextPad.rtbox1.SelText = rtbPrvt.Text

End Sub

Private Sub cmdSave_Click()

    Clipboard.Clear
    Clipboard.SetText rtbSysBrd.Text
    cmdSave.Enabled = False
    sysChars = Len(rtbSysBrd.Text)

End Sub

Private Sub cmdSave2_Click()

    publicCopy = rtbPrvt.Text
    cmdSave2.Enabled = False
    prvtChars = Len(rtbPrvt.Text)

End Sub

Private Sub Form_Load()

    cmdSave.Enabled = False
    rtbPrvt.Text = publicCopy
    rtbSysBrd.Text = Clipboard.GetText
    cmdSave2.Enabled = False
    prvtChars = Len(rtbPrvt.Text)
    sysChars = Len(rtbSysBrd.Text)
    
    If Clipboard.GetText = "" Then
        cmdClear.Enabled = False
        cmdPasteIt.Enabled = False
    End If
    
    If publicCopy = "" Then
        cmdClear2.Enabled = False
        cmdPastePrvt.Enabled = False
    End If

End Sub

Private Sub rtbPrvt_Change()

    If Len(rtbPrvt.Text) > 0 Then
        cmdPastePrvt.Enabled = True
    Else
        cmdPastePrvt.Enabled = False
    End If

End Sub

Private Sub rtbPrvt_KeyDown(KeyCode As Integer, Shift As Integer)

    cmdSave2.Enabled = True
    cmdClear2.Enabled = True

End Sub

Private Sub rtbPrvt_KeyUp(KeyCode As Integer, Shift As Integer)

    If Len(rtbPrvt.Text) <> prvtChars Then
        cmdSave2.Enabled = True
        cmdClear2.Enabled = True
    End If

End Sub

Private Sub rtbSysBrd_Change()

    If Len(rtbSysBrd.Text) > 0 Then
        cmdPasteIt.Enabled = True
    Else
        cmdPasteIt.Enabled = False
    End If

End Sub

Private Sub rtbSysBrd_KeyDown(KeyCode As Integer, Shift As Integer)

    cmdSave.Enabled = True
    cmdClear.Enabled = True

End Sub

Private Sub rtbSysBrd_KeyUp(KeyCode As Integer, Shift As Integer)

    If Len(rtbSysBrd.Text) <> sysChars Then
        cmdSave.Enabled = True
        cmdClear.Enabled = True
    End If

End Sub

