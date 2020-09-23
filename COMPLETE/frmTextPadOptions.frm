VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmOptions 
   BorderStyle     =   0  'None
   Caption         =   "Settings"
   ClientHeight    =   6120
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   6150
   Icon            =   "frmTextPadOptions.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6120
   ScaleWidth      =   6150
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog cmndlg2 
      Left            =   120
      Top             =   5520
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Height          =   375
      Left            =   2280
      TabIndex        =   3
      Top             =   5520
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   3480
      TabIndex        =   2
      Top             =   5520
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4680
      TabIndex        =   0
      Top             =   5520
      Width           =   1095
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6135
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   6165
      _ExtentX        =   10874
      _ExtentY        =   10821
      _Version        =   393216
      Style           =   1
      Tab             =   2
      TabHeight       =   520
      TabCaption(0)   =   "General"
      TabPicture(0)   =   "frmTextPadOptions.frx":014A
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame6"
      Tab(0).Control(1)=   "Frame5"
      Tab(0).Control(2)=   "Frame1"
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Editor"
      TabPicture(1)   =   "frmTextPadOptions.frx":0166
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame4"
      Tab(1).Control(1)=   "Frame2"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Colors and Style"
      TabPicture(2)   =   "frmTextPadOptions.frx":0182
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "Frame3"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      Begin VB.Frame Frame6 
         Caption         =   "Auto Uppercase"
         Height          =   1335
         Left            =   -72720
         TabIndex        =   40
         Top             =   600
         Width           =   1455
         Begin VB.CheckBox chkAutoCase 
            Caption         =   "Enable"
            Height          =   255
            Left            =   120
            TabIndex        =   41
            Top             =   600
            Width           =   1095
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Custom Insert Text (will appear in Edit -> Insert ->)"
         Height          =   3135
         Left            =   -74880
         TabIndex        =   34
         Top             =   2160
         Width           =   5895
         Begin VB.CommandButton cmdW5 
            Caption         =   "..."
            Enabled         =   0   'False
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
            Left            =   5400
            TabIndex        =   46
            Top             =   2520
            Width           =   375
         End
         Begin VB.CommandButton cmdW4 
            Caption         =   "..."
            Enabled         =   0   'False
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
            Left            =   5400
            TabIndex        =   45
            Top             =   2040
            Width           =   375
         End
         Begin VB.CommandButton cmdW3 
            Caption         =   "..."
            Enabled         =   0   'False
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
            Left            =   5400
            TabIndex        =   44
            Top             =   1560
            Width           =   375
         End
         Begin VB.CommandButton cmdW2 
            Caption         =   "..."
            Enabled         =   0   'False
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
            Left            =   5400
            TabIndex        =   43
            Top             =   1080
            Width           =   375
         End
         Begin VB.CommandButton cmdW1 
            Caption         =   "..."
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
            Left            =   5400
            TabIndex        =   42
            Top             =   600
            Width           =   375
         End
         Begin VB.TextBox Text5 
            Enabled         =   0   'False
            Height          =   375
            Left            =   960
            TabIndex        =   39
            Top             =   2400
            Width           =   4335
         End
         Begin VB.TextBox Text4 
            Enabled         =   0   'False
            Height          =   375
            Left            =   960
            TabIndex        =   38
            Top             =   1920
            Width           =   4335
         End
         Begin VB.TextBox Text3 
            Enabled         =   0   'False
            Height          =   375
            Left            =   960
            TabIndex        =   37
            Top             =   1440
            Width           =   4335
         End
         Begin VB.TextBox Text2 
            Enabled         =   0   'False
            Height          =   375
            Left            =   960
            TabIndex        =   36
            Top             =   960
            Width           =   4335
         End
         Begin VB.TextBox Text1 
            Height          =   375
            Left            =   960
            TabIndex        =   35
            Top             =   480
            Width           =   4335
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Text 5"
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
            Left            =   120
            TabIndex        =   51
            Top             =   2400
            Width           =   735
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Text 4"
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
            Left            =   120
            TabIndex        =   50
            Top             =   1920
            Width           =   735
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Text 3"
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
            Left            =   120
            TabIndex        =   49
            Top             =   1440
            Width           =   735
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Text 2"
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
            Left            =   120
            TabIndex        =   48
            Top             =   960
            Width           =   735
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Text 1"
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
            Left            =   120
            TabIndex        =   47
            Top             =   480
            Width           =   735
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Indent by default"
         Height          =   1095
         Left            =   -74760
         TabIndex        =   26
         Top             =   4080
         Width           =   5415
         Begin VB.OptionButton optInBull 
            Caption         =   "Bullet Indent"
            Height          =   255
            Left            =   3480
            TabIndex        =   31
            Top             =   360
            Width           =   1695
         End
         Begin VB.OptionButton optInBoth 
            Caption         =   "Indent Right and Left"
            Height          =   375
            Left            =   1320
            TabIndex        =   30
            Top             =   600
            Width           =   1935
         End
         Begin VB.OptionButton optInRight 
            Caption         =   "Indent Right"
            Height          =   375
            Left            =   1320
            TabIndex        =   29
            Top             =   240
            Width           =   1335
         End
         Begin VB.OptionButton optInLeft 
            Caption         =   "Indent Left"
            Height          =   255
            Left            =   120
            TabIndex        =   28
            Top             =   720
            Width           =   1215
         End
         Begin VB.OptionButton optNoIn 
            Caption         =   "None"
            Height          =   255
            Left            =   120
            TabIndex        =   27
            Top             =   360
            Width           =   975
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Default Style"
         Height          =   4215
         Left            =   120
         TabIndex        =   17
         Top             =   600
         Width           =   5775
         Begin VB.CommandButton cmdReset 
            Caption         =   "Reset All Colors and Styles"
            Height          =   375
            Left            =   120
            TabIndex        =   33
            Top             =   3450
            Width           =   2295
         End
         Begin RichTextLib.RichTextBox rtbAll 
            Height          =   615
            Left            =   2640
            TabIndex        =   32
            Top             =   3240
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   1085
            _Version        =   393217
            TextRTF         =   $"frmTextPadOptions.frx":019E
         End
         Begin RichTextLib.RichTextBox rtbDefBGColor 
            Height          =   375
            Left            =   2640
            TabIndex        =   25
            Top             =   2640
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   661
            _Version        =   393217
            TextRTF         =   $"frmTextPadOptions.frx":0220
         End
         Begin RichTextLib.RichTextBox rtbDefFontColor 
            Height          =   375
            Left            =   2640
            TabIndex        =   24
            Top             =   1920
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   661
            _Version        =   393217
            TextRTF         =   $"frmTextPadOptions.frx":02AA
         End
         Begin VB.CommandButton cmdDefBGColor 
            Caption         =   "Default Background Color..."
            Height          =   375
            Left            =   120
            TabIndex        =   23
            Top             =   2640
            Width           =   2295
         End
         Begin VB.CommandButton cmdDefFontColor 
            Caption         =   "Default Font Color..."
            Height          =   375
            Left            =   120
            TabIndex        =   22
            Top             =   1920
            Width           =   2295
         End
         Begin VB.TextBox txtDefSize 
            Height          =   375
            Left            =   2640
            TabIndex        =   21
            Top             =   1200
            Width           =   2175
         End
         Begin VB.CommandButton cmdDefaultSize 
            Caption         =   "Default Font Size..."
            Height          =   375
            Left            =   120
            TabIndex        =   20
            Top             =   1200
            Width           =   2295
         End
         Begin RichTextLib.RichTextBox rtbDefFont 
            Height          =   375
            Left            =   2640
            TabIndex        =   19
            Top             =   480
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   661
            _Version        =   393217
            TextRTF         =   $"frmTextPadOptions.frx":0334
         End
         Begin VB.CommandButton cmdDefaultFont 
            Caption         =   "Default Font Name..."
            Height          =   375
            Left            =   120
            TabIndex        =   18
            Top             =   480
            Width           =   2295
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Set Indentation Size"
         Height          =   3375
         Left            =   -74760
         TabIndex        =   7
         Top             =   600
         Width           =   5415
         Begin VB.TextBox txtInBull 
            Alignment       =   2  'Center
            BackColor       =   &H00E0E0E0&
            Height          =   285
            Left            =   3960
            Locked          =   -1  'True
            TabIndex        =   16
            Top             =   2880
            Width           =   1215
         End
         Begin VB.HScrollBar hScrInBull 
            Height          =   255
            LargeChange     =   150
            Left            =   120
            Max             =   4000
            SmallChange     =   50
            TabIndex        =   15
            Top             =   2880
            Width           =   3735
         End
         Begin RichTextLib.RichTextBox rtbinBull 
            Height          =   375
            Left            =   120
            TabIndex        =   14
            Top             =   2400
            Width           =   5055
            _ExtentX        =   8916
            _ExtentY        =   661
            _Version        =   393217
            BackColor       =   14737632
            Enabled         =   -1  'True
            ReadOnly        =   -1  'True
            TextRTF         =   $"frmTextPadOptions.frx":03B6
         End
         Begin VB.TextBox txtTEMP 
            BackColor       =   &H00E0E0E0&
            Height          =   375
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   13
            Text            =   "Right Indent"
            Top             =   1320
            Width           =   5055
         End
         Begin VB.TextBox txtInRight 
            Alignment       =   2  'Center
            BackColor       =   &H00E0E0E0&
            Height          =   285
            Left            =   3960
            Locked          =   -1  'True
            TabIndex        =   12
            Top             =   1800
            Width           =   1215
         End
         Begin VB.HScrollBar hScrInRight 
            Height          =   255
            LargeChange     =   150
            Left            =   120
            Max             =   4000
            SmallChange     =   50
            TabIndex        =   11
            Top             =   1800
            Width           =   3735
         End
         Begin VB.TextBox txtInLeft 
            Alignment       =   2  'Center
            BackColor       =   &H00E0E0E0&
            Height          =   285
            Left            =   4080
            Locked          =   -1  'True
            TabIndex        =   10
            Top             =   720
            Width           =   1095
         End
         Begin VB.HScrollBar hScrInLeft 
            Height          =   255
            LargeChange     =   150
            Left            =   120
            Max             =   4000
            SmallChange     =   50
            TabIndex        =   9
            Top             =   720
            Width           =   3735
         End
         Begin RichTextLib.RichTextBox rtbIndent 
            Height          =   375
            Left            =   120
            TabIndex        =   8
            Top             =   240
            Width           =   5055
            _ExtentX        =   8916
            _ExtentY        =   661
            _Version        =   393217
            BackColor       =   14737632
            Enabled         =   -1  'True
            ReadOnly        =   -1  'True
            TextRTF         =   $"frmTextPadOptions.frx":0442
         End
         Begin VB.Line Line4 
            BorderColor     =   &H00808080&
            BorderStyle     =   6  'Inside Solid
            X1              =   0
            X2              =   5400
            Y1              =   2205
            Y2              =   2205
         End
         Begin VB.Line Line3 
            BorderColor     =   &H80000005&
            BorderWidth     =   2
            X1              =   0
            X2              =   5400
            Y1              =   2220
            Y2              =   2220
         End
         Begin VB.Line Line2 
            BorderColor     =   &H00808080&
            BorderStyle     =   6  'Inside Solid
            X1              =   0
            X2              =   5380
            Y1              =   1132
            Y2              =   1132
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00FFFFFF&
            BorderWidth     =   2
            Index           =   0
            X1              =   0
            X2              =   5380
            Y1              =   1140
            Y2              =   1140
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Status Bar Options"
         Height          =   1335
         Left            =   -74880
         TabIndex        =   4
         Top             =   600
         Width           =   1935
         Begin VB.CheckBox chkDate 
            Caption         =   "Show Date"
            Height          =   495
            Left            =   120
            TabIndex        =   6
            Top             =   720
            Width           =   1575
         End
         Begin VB.CheckBox chkTime 
            Caption         =   "Show Time"
            Height          =   495
            Left            =   120
            TabIndex        =   5
            Top             =   240
            Width           =   1695
         End
      End
   End
   Begin VB.Menu mnuHidden 
      Caption         =   "Hidden Menu"
      Visible         =   0   'False
      Begin VB.Menu mnuSize 
         Caption         =   "Select Size"
         Begin VB.Menu mnuXSItem 
            Caption         =   "Extra Small"
         End
         Begin VB.Menu mnuSItem 
            Caption         =   "Small"
         End
         Begin VB.Menu mnuMItem 
            Caption         =   "Medium"
         End
         Begin VB.Menu mnuLItem 
            Caption         =   "Large"
         End
         Begin VB.Menu mnuXLItem 
            Caption         =   "Extra Large"
         End
      End
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim doApply As Boolean
Dim tempRight As Integer
Dim tempLeft As Integer
Dim tempBull
Dim tempDefFont As String
Dim tempViewSize As Integer
Dim tempBGColor As Long
Dim tempColor As Long
Dim tempInAll As String
Dim tempStat As Boolean



Private Sub chkAutoCase_Click()

    If doApply = True Then
        cmdApply.Enabled = True
    End If

End Sub

Private Sub chkDate_Click()

    If doApply = True Then
        cmdApply.Enabled = True
    End If

End Sub

Private Sub chkTime_Click()

    If doApply = True Then
        cmdApply.Enabled = True
    End If

End Sub

Private Sub cmdApply_Click()

    applySettings
    cmdApply.Enabled = False

End Sub

Private Sub cmdCancel_Click()

    indentRight = tempRight
    indentLeft = tempLeft
    bullIndentBy = tempBull
    defaultFont = tempDefFont
    frmTextPad.rtbox1.Font.Name = defaultFont
    frmTextPad.rtbox1.SelStart = frmTextPad.rtbox1.SelStart + Len(frmTextPad.rtbox1.SelText)
    frmTextPad.rtbox1.SelLength = 0
    currViewSize = tempViewSize
    defaultBGColor = tempBGColor
    defaultFontColor = tempColor
    defaultIndentAll = tempInAll
    
    customWordCount = 0
    
    Unload Me

End Sub

Private Sub cmdDefaultFont_Click()

'    cmndlg2.Flags = cdlCFBoth Or cdlCFApply
'    cmndlg2.ShowFont
'
'    If cmndlg2.FontName <> "" Then
'        cmdApply.Enabled = True
'        defaultFont = cmndlg2.FontName
'        rtbDefFont.Font.Name = defaultFont
'        rtbDefFont.Text = defaultFont
'    End If

    Load frmTextPad_FontNames
    frmTextPad_FontNames.Show vbModal, Me
    rtbDefFont.Font.Name = defaultFont
    rtbDefFont.Font.Size = currViewSize
    rtbDefFont.Text = defaultFont
    
    rtbAll.Font.Name = defaultFont
    rtbAll.Text = defaultFont
    
    frmTextPad.rtbox1.SelStart = 0
    frmTextPad.rtbox1.SelLength = Len(frmTextPad.rtbox1.Text)
    frmTextPad.rtbox1.Font.Name = defaultFont
    
    frmTextPad.rtbox1.SelStart = frmTextPad.rtbox1.SelStart + Len(frmTextPad.rtbox1.SelText)
    frmTextPad.rtbox1.SelLength = 0

End Sub

Private Sub cmdDefaultSize_Click()
 
    PopupMenu mnuSize

End Sub

Private Sub cmdDefBGColor_Click()

    cmnDlg2.ShowColor
    
    If cmnDlg2.Color <> 0 Then
        defaultBGColor = cmnDlg2.Color
        cmdApply.Enabled = True
        rtbDefBGColor.BackColor = defaultBGColor
        
        rtbAll.BackColor = defaultBGColor
    End If

End Sub

Private Sub cmdDefFontColor_Click()

    cmnDlg2.ShowColor
    
    If cmnDlg2.Color <> 0 Then
        defaultFontColor = cmnDlg2.Color
        cmdApply.Enabled = True
          rtbDefFontColor.SelStart = 0
          rtbDefFontColor.SelLength = Len(rtbDefFontColor.Text)
          rtbDefFontColor.SelColor = defaultFontColor
          
          rtbAll.SelStart = 0
          rtbAll.SelLength = Len(rtbAll.Text)
          rtbAll.SelColor = defaultFontColor
    End If

End Sub

Private Sub cmdOK_Click()

    applySettings
    Unload Me
    
End Sub

Private Sub cmdReset_Click()

    defaultFont = "MS Sans Serif"
    currViewSize = 9
    defaultBGColor = 16777215
    defaultFontColor = 986895
    
    rtbAll.Font.Name = defaultFont
    rtbAll.Text = defaultFont
    
    frmTextPad.rtbox1.SelStart = 0
    frmTextPad.rtbox1.SelLength = Len(frmTextPad.rtbox1.Text)
    frmTextPad.rtbox1.Font.Name = defaultFont
    
    clearChecks
    mnuSItem.Checked = True
    currViewSize = 9
    rtbAll.Font.Size = currViewSize
    
    rtbDefFontColor.SelStart = 0
    rtbDefFontColor.SelLength = Len(rtbDefFontColor.Text)
    rtbDefFontColor.SelColor = defaultFontColor
    
    rtbAll.SelStart = 0
    rtbAll.SelLength = Len(rtbAll.Text)
    rtbAll.SelColor = defaultFontColor
    
    rtbDefBGColor.BackColor = defaultBGColor
    rtbAll.BackColor = defaultBGColor
    
    cmdApply.Enabled = True

End Sub



Private Sub cmdW1_Click()

    Load frmCWordTxt
    frmCWordTxt.Caption = "Word 1"
    frmCWordTxt.Text1.Text = Text1.Text
    frmCWordTxt.Text1.SelStart = frmCWordTxt.Text1.SelStart + Len(frmCWordTxt.Text1.Text)
    frmCWordTxt.Text1.SelLength = 0
    frmCWordTxt.Show vbModeless, Me

End Sub

Private Sub cmdW2_Click()

    Load frmCWordTxt
    frmCWordTxt.Caption = "Word 2"
    frmCWordTxt.Text1.Text = Text2.Text
    frmCWordTxt.Text1.SelStart = frmCWordTxt.Text1.SelStart + Len(frmCWordTxt.Text1.Text)
    frmCWordTxt.Text1.SelLength = 0
    frmCWordTxt.Show vbModeless, Me

End Sub

Private Sub cmdW3_Click()

    Load frmCWordTxt
    frmCWordTxt.Caption = "Word 3"
    frmCWordTxt.Text1.Text = Text3.Text
    frmCWordTxt.Text1.SelStart = frmCWordTxt.Text1.SelStart + Len(frmCWordTxt.Text1.Text)
    frmCWordTxt.Text1.SelLength = 0
    frmCWordTxt.Show vbModeless, Me

End Sub

Private Sub cmdW4_Click()

    Load frmCWordTxt
    frmCWordTxt.Caption = "Word 4"
    frmCWordTxt.Text1.Text = Text4.Text
    frmCWordTxt.Text1.SelStart = frmCWordTxt.Text1.SelStart + Len(frmCWordTxt.Text1.Text)
    frmCWordTxt.Text1.SelLength = 0
    frmCWordTxt.Show vbModeless, Me

End Sub

Private Sub cmdW5_Click()

    Load frmCWordTxt
    frmCWordTxt.Caption = "Word 5"
    frmCWordTxt.Text1.Text = Text5.Text
    frmCWordTxt.Text1.SelStart = frmCWordTxt.Text1.SelStart + Len(frmCWordTxt.Text1.Text)
    frmCWordTxt.Text1.SelLength = 0
    frmCWordTxt.Show vbModeless, Me

End Sub

Private Sub Form_Load()

    tempStat = TopMost
    TopMost = False
    SetTopMost

    Me.Height = 6495
    Me.Width = 6240
    Me.Top = Screen.Height / 2 - (Me.Height / 2)
    Me.Left = Screen.Width / 2 - (Me.Width / 2)
    cmdApply.Enabled = False
    doApply = False
    
    Text1.BackColor = QBColor(7)
    Text2.BackColor = QBColor(7)
    Text3.BackColor = QBColor(7)
    Text4.BackColor = QBColor(7)
    Text5.BackColor = QBColor(7)
    
    If showTime = True Then
        chkTime.Value = 1
    Else
        chkTime.Value = 0
    End If
    
    If showDate = True Then
        chkDate.Value = 1
    Else
        chkDate.Value = 0
    End If
    
    ssTab1.Tab = 0
    
    tempRight = indentRight
    tempLeft = indentLeft
    tempBull = bullIndentBy
    tempDefFont = defaultFont
    tempViewSize = currViewSize
    tempBGColor = defaultBGColor
    tempColor = defaultFontColor
    tempInAll = defaultIndentAll
    
    hScrInLeft.Value = indentLeft
    txtInLeft.Text = Str(indentLeft)
    rtbIndent.SelStart = 0
    rtbIndent.SelLength = Len(rtbIndent.Text)
    rtbIndent.SelIndent = indentLeft
    
    hScrInRight.Value = indentRight
    txtInRight.Text = Str(indentRight)
    txtTEMP.Text = Space(CInt(91 - (indentRight / 44))) & "Right indent"
    
    rtbinBull.BulletIndent = bullIndentBy
    hScrInBull.Value = bullIndentBy
    txtInBull.Text = Str(bullIndentBy)
    rtbinBull.SelStart = 0
    rtbinBull.SelLength = Len(rtbinBull.Text)
    rtbinBull.SelBullet = True
    
    rtbDefFont.Font.Name = defaultFont
    rtbDefFont.Text = defaultFont
    rtbDefFont.Font.Size = currViewSize
    
    txtDefSize.Text = currViewName
    
      rtbDefFontColor.SelStart = 0
      rtbDefFontColor.SelLength = Len(rtbDefFontColor.Text)
      rtbDefFontColor.SelColor = defaultFontColor
    
    rtbDefBGColor.BackColor = defaultBGColor
    
      rtbDefFontColor.SelStart = 0
      rtbDefFontColor.SelLength = Len(rtbDefFontColor.Text)
      rtbDefFontColor.SelColor = defaultFontColor
      
    rtbAll.Text = defaultFont
    rtbAll.Font.Name = defaultFont
    rtbAll.BackColor = defaultBGColor
    rtbAll.Font.Size = currViewSize
      rtbAll.SelStart = 0
      rtbAll.SelLength = Len(rtbAll.Text)
      rtbAll.SelColor = defaultFontColor
    
    selectSize
    
    setIndentOpts
    
    If useLastChar = True Then
        chkAutoCase.Value = 1
    Else
        chkAutoCase.Value = 0
    End If
    
    populateCustomWord
    
    doApply = True

End Sub

Sub applySettings()

    If chkTime.Value = 1 Then
        showTime = True
        frmTextPad.StatusBar1.Panels(2).Bevel = sbrInset
        frmTextPad.StatusBar1.Panels(2).Style = sbrTime
        frmTextPad.StatusBar1.Panels(2).ToolTipText = "Click to hide Time"
    ElseIf chkTime.Value = 0 Then
        showTime = False
        frmTextPad.StatusBar1.Panels(2).Bevel = sbrNoBevel
        frmTextPad.StatusBar1.Panels(2).Style = sbrText
        frmTextPad.StatusBar1.Panels(2).Text = ""
        frmTextPad.StatusBar1.Panels(2).ToolTipText = "Click to show Time"
    End If
    
    If chkDate.Value = 1 Then
        showDate = True
        frmTextPad.StatusBar1.Panels(3).Bevel = sbrInset
        frmTextPad.StatusBar1.Panels(3).Style = sbrDate
        frmTextPad.StatusBar1.Panels(3).ToolTipText = "Click to hide Date"
    ElseIf chkDate.Value = 0 Then
        showDate = False
        frmTextPad.StatusBar1.Panels(3).Bevel = sbrNoBevel
        frmTextPad.StatusBar1.Panels(3).Style = sbrText
        frmTextPad.StatusBar1.Panels(3).Text = ""
        frmTextPad.StatusBar1.Panels(3).ToolTipText = "Click to show Date"
    End If
    
    tempRight = indentRight
    tempLeft = indentLeft
    tempBull = bullIndentBy
    
    frmTextPad.rtbox1.BulletIndent = bullIndentBy
    tempDefFont = defaultFont
    
    tempViewSize = currViewSize
    frmTextPad.rtbox1.Font.Size = currViewSize
    clearChecksTextPad
    setCheckTextPad
    
    frmTextPad.rtbox1.BackColor = defaultBGColor
    tempBGColor = defaultBGColor
    
      frmTextPad.rtbox1.SelStart = 0
      frmTextPad.rtbox1.SelLength = Len(frmTextPad.rtbox1.Text)
      frmTextPad.rtbox1.SelColor = defaultFontColor
      tempColor = defaultFontColor
      
    tempInAll = defaultIndentAll
          
    Select Case defaultIndentAll
    Case "Indent Left"
        With frmTextPad.rtbox1
            .SelStart = 1
            .SelLength = Len(frmTextPad.rtbox1.Text)
            .SelIndent = indentLeft
        End With
    Case "Indent Right"
         With frmTextPad.rtbox1
            .SelStart = 1
            .SelLength = Len(frmTextPad.rtbox1.Text)
            .SelRightIndent = indentRight
         End With
    Case "Indent Left and Right"
        With frmTextPad.rtbox1
            .SelStart = 1
            .SelLength = Len(frmTextPad.rtbox1.Text)
            .SelIndent = indentLeft
            .SelRightIndent = indentRight
        End With
    Case "Bullet Indent"
        frmTextPad.rtbox1.SelStart = 0
        frmTextPad.rtbox1.SelLength = Len(frmTextPad.rtbox1.Text)
        frmTextPad.rtbox1.SelBullet = True
    Case "No Indent"
         With frmTextPad.rtbox1
             .SelStart = 1
             .SelLength = Len(frmTextPad.rtbox1.Text)
             .SelIndent = 0
             .SelRightIndent = 0
          End With
          
          frmTextPad.rtbox1.SelStart = 0
          frmTextPad.rtbox1.SelLength = Len(frmTextPad.rtbox1.Text)
          frmTextPad.rtbox1.SelBullet = False
    End Select
    
    frmTextPad.rtbox1.SelStart = frmTextPad.rtbox1.SelStart + Len(frmTextPad.rtbox1.SelText)
    frmTextPad.rtbox1.SelLength = 0
    
    customWordMenu
    
    If chkAutoCase.Value = 1 Then
        useLastChar = True
    Else
        useLastChar = False
    End If
    
    saveOptions
    
    cmdApply.Enabled = False
    
    frmTextPad.cmb1.Text = defaultFont

End Sub


Private Sub Form_Unload(Cancel As Integer)

    TopMost = tempStat
    SetTopMost
    frmTextPad.SetFocus

End Sub



Private Sub hScrInBull_Change()

    If doApply = True Then
        cmdApply.Enabled = True
    End If
    
    bullIndentBy = hScrInBull.Value
    rtbinBull.BulletIndent = bullIndentBy
    txtInBull.Text = Str(hScrInBull.Value)
    rtbinBull.SelStart = 0
    rtbinBull.SelLength = Len(rtbinBull.Text)
    rtbinBull.SelBullet = True


End Sub

Private Sub hScrInLeft_Change()

    If doApply = True Then
        cmdApply.Enabled = True
    End If
    
    indentLeft = hScrInLeft.Value
    txtInLeft.Text = Str(hScrInLeft.Value)
    rtbIndent.SelStart = 0
    rtbIndent.SelLength = Len(rtbIndent.Text)
    rtbIndent.SelIndent = indentLeft

End Sub

Private Sub hScrInRight_Change()

    If doApply = True Then
        cmdApply.Enabled = True
    End If
    
    indentRight = hScrInRight.Value
    txtInRight.Text = Str(hScrInRight.Value)
    
    txtTEMP.Text = Space(CInt(91 - (indentRight / 44))) & "Right Indent"

End Sub

Sub selectSize()

    Select Case currViewSize
    Case 8:
        mnuXSItem.Checked = True
        currViewName = "Extra Small"
    Case 9:
        mnuSItem.Checked = True
        currViewName = "Small"
    Case 11:
        mnuMItem.Checked = True
        currViewName = "Medium"
    Case 13:
        mnuLItem.Checked = True
        currViewName = "Large"
    Case 16:
        mnuXLItem.Checked = True
        currViewName = "Extra Large"
    Case Else:
        currViewSize = 10
    End Select

End Sub

Private Sub mnuLItem_Click()

    clearChecks
    mnuLItem.Checked = True
    currViewSize = 13
    currViewName = "Large"
    txtDefSize.Text = currViewName
    rtbAll.Font.Size = currViewSize
    

End Sub

Sub clearChecks()

    mnuXSItem.Checked = False
    mnuSItem.Checked = False
    mnuMItem.Checked = False
    mnuLItem.Checked = False
    mnuXLItem.Checked = False
    
    cmdApply.Enabled = True

End Sub

Private Sub mnuMItem_Click()

    clearChecks
    mnuMItem.Checked = True
    currViewSize = 11
    currViewName = "Medium"
    txtDefSize.Text = currViewName
    rtbAll.Font.Size = currViewSize

End Sub

Private Sub mnuSItem_Click()

    clearChecks
    mnuSItem.Checked = True
    currViewSize = 9
    currViewName = "Small"
    txtDefSize.Text = currViewName
    rtbAll.Font.Size = currViewSize

End Sub

Private Sub mnuXLItem_Click()

    clearChecks
    mnuXLItem.Checked = True
    currViewSize = 16
    currViewName = "Extra Large"
    txtDefSize.Text = currViewName
    rtbAll.Font.Size = currViewSize

End Sub

Private Sub mnuXSItem_Click()

    clearChecks
    mnuXSItem.Checked = True
    currViewSize = 8
    currViewName = "Extra Small"
    txtDefSize.Text = currViewName
    rtbAll.Font.Size = currViewSize

End Sub

Sub clearChecksTextPad()
    
    frmTextPad.mnuXSmallItem.Checked = False
    frmTextPad.mnuSmallItem.Checked = False
    frmTextPad.mnuMediumItem.Checked = False
    frmTextPad.mnuLargeItem.Checked = False
    frmTextPad.mnuXLargeItem.Checked = False

End Sub

Sub setCheckTextPad()

    Select Case currViewSize
    Case 8:
        frmTextPad.mnuXSmallItem.Checked = True
    Case 9:
        frmTextPad.mnuSmallItem.Checked = True
    Case 11:
        frmTextPad.mnuMediumItem.Checked = True
    Case 13:
        frmTextPad.mnuLargeItem.Checked = True
    Case 16:
        frmTextPad.mnuXLargeItem.Checked = True
    Case Else:
        currViewSize = 10
    End Select

End Sub

Sub setIndentOpts()

    Select Case defaultIndentAll
    Case "Indent Left"
        optInLeft.Value = True
    Case "Indent Right"
        optInRight.Value = True
    Case "Indent Left and Right"
        optInBoth.Value = True
    Case "Bullet Indent"
        optInBull.Value = True
    Case "No Indent"
        optNoIn.Value = True
    End Select

End Sub

Private Sub optInBoth_Click()

    If doApply = True Then
        cmdApply.Enabled = True
        defaultIndentAll = "Indent Left and Right"
    End If

End Sub

Private Sub optInBull_Click()

    If doApply = True Then
        cmdApply.Enabled = True
        defaultIndentAll = "Bullet Indent"
    End If

End Sub

Private Sub optInLeft_Click()

    If doApply = True Then
        cmdApply.Enabled = True
        defaultIndentAll = "Indent Left"
    End If
    
End Sub

Private Sub optInRight_Click()

    If doApply = True Then
        cmdApply.Enabled = True
        defaultIndentAll = "Indent Right"
    End If

End Sub

Private Sub optNoIn_Click()

    If doApply = True Then
        cmdApply.Enabled = True
        defaultIndentAll = "No Indent"
    End If

End Sub

Private Sub Text1_Change()

    If Len(Text1.Text) > 0 Then
        If doApply = True Then
            cmdApply.Enabled = True
        End If
        Text1.BackColor = QBColor(15)
        Text2.Enabled = True
        cmdW2.Enabled = True
        Text2.BackColor = QBColor(15)
        customWordCount = 1
        customWord(1) = Text1.Text
    ElseIf Len(Text1.Text) = 0 Then
        If doApply = True Then
            cmdApply.Enabled = True
        End If
        Text1.BackColor = QBColor(7)
        Text2.Enabled = False
        cmdW2.Enabled = False
        Text2.BackColor = QBColor(7)
        customWordCount = 0
        customWord(1) = ""
    End If

End Sub

Private Sub Text2_Change()

    If Len(Text2.Text) > 0 Then
        If doApply = True Then
            cmdApply.Enabled = True
        End If
        Text2.BackColor = QBColor(15)
        Text3.Enabled = True
        cmdW3.Enabled = True
        Text3.BackColor = QBColor(15)
        customWordCount = 2
        customWord(2) = Text2.Text
    ElseIf Len(Text2.Text) = 0 Then
        If doApply = True Then
            cmdApply.Enabled = True
        End If
        Text2.BackColor = QBColor(7)
        Text3.Enabled = False
        cmdW3.Enabled = False
        Text3.BackColor = QBColor(7)
        customWordCount = 1
        customWord(2) = ""
    End If

End Sub

Private Sub Text3_Change()

    If Len(Text3.Text) > 0 Then
        If doApply = True Then
            cmdApply.Enabled = True
        End If
        Text3.BackColor = QBColor(15)
        Text4.Enabled = True
        cmdW4.Enabled = True
        Text4.BackColor = QBColor(15)
        customWordCount = 3
        customWord(3) = Text3.Text
    ElseIf Len(Text3.Text) = 0 Then
        If doApply = True Then
            cmdApply.Enabled = True
        End If
        Text3.BackColor = QBColor(7)
        Text4.Enabled = False
        cmdW4.Enabled = False
        Text4.BackColor = QBColor(7)
        customWordCount = 2
        customWord(3) = ""
    End If

End Sub

Private Sub Text4_Change()

    If Len(Text4.Text) > 0 Then
        If doApply = True Then
            cmdApply.Enabled = True
        End If
        Text4.BackColor = QBColor(15)
        Text5.Enabled = True
        cmdW5.Enabled = True
        Text5.BackColor = QBColor(15)
        customWordCount = 4
        customWord(4) = Text4.Text
    ElseIf Len(Text4.Text) = 0 Then
        If doApply = True Then
            cmdApply.Enabled = True
        End If
        Text4.BackColor = QBColor(7)
        Text5.Enabled = False
        cmdW5.Enabled = False
        Text5.BackColor = QBColor(7)
        customWordCount = 3
        customWord(4) = ""
    End If

End Sub

Private Sub Text5_Change()

    If Len(Text5.Text) > 0 Then
        If doApply = True Then
            cmdApply.Enabled = True
        End If
        Text5.BackColor = QBColor(15)
        customWordCount = 5
        customWord(5) = Text5.Text
    ElseIf Len(Text5.Text) = 0 Then
        If doApply = True Then
            cmdApply.Enabled = True
        End If
        Text5.BackColor = QBColor(7)
        customWordCount = 4
        customWord(5) = ""
    End If

End Sub

Sub populateCustomWord()

    Select Case customWordCount:
    Case 1:
        Text1.Text = customWord(1)
    Case 2:
        Text1.Text = customWord(1)
        Text2.Text = customWord(2)
    Case 3:
        Text1.Text = customWord(1)
        Text2.Text = customWord(2)
        Text3.Text = customWord(3)
    Case 4:
        Text1.Text = customWord(1)
        Text2.Text = customWord(2)
        Text3.Text = customWord(3)
        Text4.Text = customWord(4)
    Case 5:
        Text1.Text = customWord(1)
        Text2.Text = customWord(2)
        Text3.Text = customWord(3)
        Text4.Text = customWord(4)
        Text5.Text = customWord(5)
    End Select

End Sub
